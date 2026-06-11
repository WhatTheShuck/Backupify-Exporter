# Throwaway — stubbed end-to-end pipeline test for the multi-service TUI.
# Monkeypatches every network-touching function; nothing real is contacted.
# Run: PYTHONPATH=. uv run --with "textual>=0.52.0" --with "playwright>=1.44.0" \
#        --with "httpx>=0.27.0" python test_stubbed_pipeline.py
import asyncio
import os
import sys
import tempfile
from pathlib import Path

import backupify_export as bx
from textual.widgets import Button, DataTable, SelectionList, TabbedContent

TMP = Path(tempfile.mkdtemp(prefix="bfy_test_"))
os.chdir(TMP)  # playwright_debug.log lands here, not in the repo

CFG = {
    "account":  {"email": "test@example.com", "totp_enabled": False},
    "server":   {"base_url": "https://example.backupify.test",
                 "customer_id": "999", "ext_customer_id": "999"},
    "defaults": {"concurrency": 4, "output_dir": str(TMP / "out")},
}

FAKE_USERS = {
    "exchange": [
        {"service_id": "101", "name": "Alice Adams", "email": "alice@x.com",
         "snapshot_id": "9001", "size_bytes": 5000, "size_label": "5 KB"},
        {"service_id": "102", "name": "Bob Brown [b]", "email": "bob@x.com",
         "snapshot_id": "9002", "size_bytes": 100, "size_label": "100 B"},
    ],
    "onedrive": [
        {"service_id": "201", "name": "Carol Chen", "email": "carol@x.com",
         "snapshot_id": "9101", "size_bytes": 777, "size_label": "777 B"},
        {"service_id": "202", "name": "Dan Diaz", "email": "dan@x.com",
         "snapshot_id": "9102", "size_bytes": None, "size_label": ""},
    ],
    "sharepoint": [
        {"service_id": "301", "name": "Intranet", "email": "https://x.sharepoint.com/",
         "snapshot_id": "9201", "size_bytes": 9999, "size_label": "9.9 KB"},
        {"service_id": "302", "name": "EH&S [archived]", "email": "https://x.sharepoint.com/ehs",
         "snapshot_id": "9202", "size_bytes": 42, "size_label": "42 B"},
    ],
}

triggered: dict[str, list[str]] = {k: [] for k in FAKE_USERS}
downloaded: dict[str, list[str]] = {k: [] for k in FAKE_USERS}
config_saves = []


class FakeContext:
    browser = None


async def fake_login(logger, debug_log, on_status=None):
    if on_status:
        on_status("stub login")
    return FakeContext()


async def fake_extract_cookies(ctx):
    return {"__backupify_session": "stub"}


async def fake_get_entities(ctx, cookies, svc, logger):
    await asyncio.sleep(0.05)
    return list(FAKE_USERS[svc.key])


async def fake_get_latest_snapshot_id(ctx, svc, sid, logger):
    raise AssertionError("should not be called — users carry snapshot_id")


# sids whose triggered job never shows up on the export page (gets an id the
# fake scan never lists) — drives the missing→prompt path
MISSING_SIDS: set[str] = set()
# sids temporarily invisible to scans (simulates "server has no export yet"
# without forgetting they were triggered) — lets phases force fresh triggers
HIDDEN_SIDS: set[str] = set()


async def fake_trigger_export(ctx, cookies, svc, sid, snapshot_id, logger):
    triggered[svc.key].append(sid)
    if sid in MISSING_SIDS:
        return f"88{sid}"
    return f"77{sid}"


NAME_BY_SID = {u["service_id"]: u["name"]
               for users in FAKE_USERS.values() for u in users}
scan_counts: dict[str, int] = {k: 0 for k in FAKE_USERS}

# Adoption scenario (onedrive): the server already holds a completed export
# for Carol and an id-less "In Progress" row for Dan that completes under a
# brand-new id a few scans later — nothing should be triggered.
ADOPT_TEST = {"active": False, "base": 0}


async def fake_scan_export_page(ctx, svc, logger):
    scan_counts[svc.key] += 1
    if ADOPT_TEST["active"] and svc.key == "onedrive":
        rows = [{"job_id": "9901", "source_name": "Carol Chen", "status": "completed",
                 "download_url": "https://example.backupify.test/dl?id=9901"}]
        if scan_counts["onedrive"] - ADOPT_TEST["base"] <= 4:
            rows.append({"job_id": None, "source_name": "Dan Diaz",
                         "status": "queued", "download_url": ""})
        else:
            rows.append({"job_id": "9902", "source_name": "Dan Diaz",
                         "status": "completed",
                         "download_url": "https://example.backupify.test/dl?id=9902"})
        return rows
    # SharePoint queued/running rows carry no ?id= action links (verified live)
    # — simulate a queue longer than max_unseen (10) scans to prove the poll's
    # source-name fallback keeps the job alive instead of declaring it missing.
    if svc.key == "sharepoint" and scan_counts[svc.key] <= 12:
        return [
            {"job_id": None, "source_name": NAME_BY_SID[sid], "status": "queued",
             "download_url": ""}
            for sid in triggered[svc.key]
        ]
    # afterwards every triggered job is "completed" (jobs for MISSING_SIDS
    # never get listed at all)
    return [
        {"job_id": f"77{sid}", "source_name": NAME_BY_SID[sid], "status": "completed",
         "download_url": f"https://example.backupify.test/dl?id=77{sid}"}
        for sid in triggered[svc.key] if sid not in MISSING_SIDS | HIDDEN_SIDS
    ]


async def fake_download_file(url, cookies, output_dir, name, sid, svc, logger,
                             on_progress=None):
    if on_progress:
        on_progress(512, 1024)
        on_progress(1024, 1024)
    fname = bx.export_file_path(output_dir, name, sid, svc.file_ext).name
    (output_dir / fname).write_bytes(b"stub")
    downloaded[svc.key].append(sid)
    return fname


def fake_save_config(cfg):
    config_saves.append(True)  # must never hit the real config file


bx.login = fake_login
bx.extract_cookies = fake_extract_cookies
bx.get_entities = fake_get_entities
bx.get_latest_snapshot_id = fake_get_latest_snapshot_id
bx.trigger_export = fake_trigger_export
bx.scan_export_page = fake_scan_export_page
bx.download_file = fake_download_file
bx.save_config = fake_save_config
bx.POLL_INTERVAL = 0.3
# keep the missing-threshold below the simulated 12-scan queued phase so the
# test still proves the source-name fallback (not just a long leash)
bx.MISSING_FRESH_SCANS = 10

failures: list[str] = []


def check(cond, msg):
    print(("  ok   " if cond else "  FAIL ") + msg)
    if not cond:
        failures.append(msg)


def reset_tab_state(app, key):
    """Forget completions + delete files so the tab can run again from scratch."""
    tab = app._tabs[key]
    tab.state.state["completed"].clear()
    tab.state.save()
    out = Path(CFG["defaults"]["output_dir"]) / key
    for p in out.iterdir():
        if "__" in p.name:
            p.unlink()


async def wait_for(pilot, pred, timeout=15.0, what="condition"):
    deadline = asyncio.get_event_loop().time() + timeout
    while asyncio.get_event_loop().time() < deadline:
        if pred():
            return True
        await pilot.pause(0.1)
    print(f"  TIMEOUT waiting for {what}")
    return False


async def run_tab(app, pilot, key):
    print(f"--- {key} tab ---")
    tab = app._tabs[key]

    ok = await wait_for(pilot, lambda: tab.users_loaded, what=f"{key} users")
    check(ok, f"{key}: users loaded")
    sl = app.query_one(f"#{key}-user-list", SelectionList)
    check(sl.option_count == len(FAKE_USERS[key]),
          f"{key}: SelectionList shows {sl.option_count} rows")
    check(tab.selected == {u["service_id"] for u in FAKE_USERS[key]},
          f"{key}: all entities selected by default")

    btn = app.query_one(f"#{key}-btn-start", Button)
    check(not btn.disabled, f"{key}: Start enabled after load")

    app._start_export_for(key)
    await pilot.pause()
    check(tab.export_running, f"{key}: export_running set")
    check(btn.disabled, f"{key}: Start disabled during run")

    ok = await wait_for(pilot, lambda: not tab.export_running, timeout=30.0,
                        what=f"{key} run finish")
    check(ok, f"{key}: run finished (export_running reset)")
    check(not btn.disabled, f"{key}: Start re-enabled")

    table = app.query_one(f"#{key}-progress-table", DataTable)
    statuses = []
    for u in FAKE_USERS[key]:
        cell = str(table.get_cell(u["service_id"], "status"))
        statuses.append(cell)
        check("Complete" in cell, f"{key}: {u['name']} row = {cell!r}")

    sids = {u["service_id"] for u in FAKE_USERS[key]}
    check(set(triggered[key]) == sids, f"{key}: all {len(sids)} exports triggered")
    check(set(downloaded[key]) == sids, f"{key}: all {len(sids)} downloads ran")
    check(set(tab.state.state["completed"]) == sids,
          f"{key}: progress.json marks all complete")
    out = Path(CFG["defaults"]["output_dir"]) / key
    files = sorted(p.name for p in out.iterdir()
                   if "__" in p.name and not p.name.endswith(".part"))
    svc_ext = bx.SERVICES[key].file_ext
    check(len(files) == len(sids) and all(f.endswith(svc_ext) for f in files),
          f"{key}: files on disk {files}")


async def main():
    app = bx.BackupifyApp(CFG)
    async with app.run_test(size=(140, 44)) as pilot:
        await pilot.pause()

        # paint sanity: non-zero regions on the active (exchange) tab
        for sel in (".user-list", ".progress-table", ".log-pane", ".settings-bar"):
            w = app.query_one(f"#tab-exchange {sel}")
            check(w.region.height > 0 and w.region.width > 0,
                  f"paint: exchange {sel} region {w.region}")

        await run_tab(app, pilot, "exchange")

        # switch tabs — entities must lazy-load on activation; the layout can
        # take a few frames to settle after the switch, so wait, don't sample.
        for key in ("onedrive", "sharepoint"):
            app.query_one(TabbedContent).active = f"tab-{key}"
            await pilot.pause()
            for sel in (".user-list", ".progress-table"):
                w = app.query_one(f"#tab-{key} {sel}")
                ok = await wait_for(pilot,
                                    lambda w=w: w.region.height > 0 and w.region.width > 0,
                                    timeout=10.0, what=f"{key} {sel} paint")
                check(ok, f"paint: {key} {sel} region {w.region}")

            await run_tab(app, pilot, key)

        # ── wind-down: no-new-jobs — nothing starts, nothing triggers ──────
        print("--- wind-down: no-new-jobs (onedrive) ---")
        reset_tab_state(app, "onedrive")
        app.query_one(TabbedContent).active = "tab-onedrive"
        await pilot.pause()
        app.winddown = "no-new-jobs"
        before_tr = len(triggered["onedrive"])
        app._start_export_for("onedrive")
        await pilot.pause()
        ok = await wait_for(pilot, lambda: not app._tabs["onedrive"].export_running,
                            timeout=30.0, what="no-new-jobs run finish")
        check(ok, "no-new-jobs: run finished")
        table = app.query_one("#onedrive-progress-table", DataTable)
        for u in FAKE_USERS["onedrive"]:
            cell = str(table.get_cell(u["service_id"], "status"))
            check("Deferred" in cell, f"no-new-jobs: {u['name']} row = {cell!r}")
        check(len(triggered["onedrive"]) == before_tr, "no-new-jobs: no new exports triggered")

        app.winddown = "off"
        app._start_export_for("onedrive")
        await pilot.pause()
        ok = await wait_for(pilot, lambda: not app._tabs["onedrive"].export_running,
                            timeout=30.0, what="onedrive re-run finish")
        check(ok, "no-new-jobs: re-run after wind-down off finished")
        for u in FAKE_USERS["onedrive"]:
            cell = str(table.get_cell(u["service_id"], "status"))
            check("Complete" in cell, f"no-new-jobs: {u['name']} re-run row = {cell!r}")

        # ── wind-down: no-new-downloads — exports queue, downloads defer ───
        print("--- wind-down: no-new-downloads (sharepoint) ---")
        reset_tab_state(app, "sharepoint")
        # forget old scans AND hide round-1's rows so workers trigger fresh
        # exports instead of adopting the old completed jobs; un-hide once
        # the triggers have fired so the polls can see them
        bx._export_caches["sharepoint"].clear()
        bx._export_pending["sharepoint"].clear()
        HIDDEN_SIDS.update({"301", "302"})
        app.query_one(TabbedContent).active = "tab-sharepoint"
        await pilot.pause()
        app.winddown = "no-new-downloads"
        sp_tab    = app._tabs["sharepoint"]
        before_tr = len(triggered["sharepoint"])
        before_dl = len(downloaded["sharepoint"])
        app._start_export_for("sharepoint")
        await pilot.pause()
        ok = await wait_for(pilot,
                            lambda: len(triggered["sharepoint"]) == before_tr + 2,
                            timeout=30.0, what="fresh triggers fired")
        check(ok, "no-new-downloads: fresh exports triggered")
        HIDDEN_SIDS.clear()  # rows become visible — polls can now find them
        ok = await wait_for(pilot, lambda: not sp_tab.export_running,
                            timeout=30.0, what="no-new-downloads run finish")
        check(ok, "no-new-downloads: run finished")
        table = app.query_one("#sharepoint-progress-table", DataTable)
        sp_sids = {u["service_id"] for u in FAKE_USERS["sharepoint"]}
        for u in FAKE_USERS["sharepoint"]:
            cell = str(table.get_cell(u["service_id"], "status"))
            check("Deferred" in cell, f"no-new-downloads: {u['name']} row = {cell!r}")
        check(len(triggered["sharepoint"]) == before_tr + 2,
              "no-new-downloads: exports WERE triggered (queued for later)")
        check(len(downloaded["sharepoint"]) == before_dl,
              "no-new-downloads: no downloads ran")
        check(set(sp_tab.state.state["in_progress"]) == sp_sids,
              "no-new-downloads: in_progress preserved for resume")
        out = Path(CFG["defaults"]["output_dir"]) / "sharepoint"
        check(not [p for p in out.iterdir() if "__" in p.name],
              "no-new-downloads: no files on disk")

        app.winddown = "off"
        app._start_export_for("sharepoint")
        await pilot.pause()
        ok = await wait_for(pilot, lambda: not sp_tab.export_running,
                            timeout=30.0, what="sharepoint resume finish")
        check(ok, "no-new-downloads: resume run finished")
        for u in FAKE_USERS["sharepoint"]:
            cell = str(table.get_cell(u["service_id"], "status"))
            check("Complete" in cell, f"resume: {u['name']} row = {cell!r}")
        check(not sp_tab.state.state["in_progress"], "resume: in_progress drained")
        check(len([p for p in out.iterdir() if "__" in p.name]) == 2,
              "resume: files downloaded")

        # ── missing-job extend dialog ───────────────────────────────────────
        print("--- missing-job extend dialog ---")
        dialog_results: list[bool] = []

        async def _ask() -> None:
            dialog_results.append(await app._ask_extend_poll("Test User", "999"))

        for btn, expected in (("extend-wait", True), ("extend-fail", False)):
            app.run_worker(_ask())
            ok = await wait_for(pilot,
                                lambda: isinstance(app.screen, bx.ExtendPollScreen),
                                what=f"dialog visible ({btn})")
            check(ok, f"dialog appeared ({btn})")
            await pilot.click(f"#{btn}")
            await pilot.pause()
        check(dialog_results == [True, False], f"dialog results {dialog_results}")

        # ── missing job → prompt → give up → failed ─────────────────────────
        print("--- missing job → prompt → fail (onedrive) ---")
        reset_tab_state(app, "onedrive")
        # forget old scans, or Dan's worker adopts his completed job from an
        # earlier phase instead of triggering the never-listed one
        bx._export_caches["onedrive"].clear()
        bx._export_pending["onedrive"].clear()
        MISSING_SIDS.add("202")  # Dan Diaz's job will never be listed
        app.query_one(TabbedContent).active = "tab-onedrive"
        await pilot.pause()
        app._start_export_for("onedrive")
        ok = await wait_for(pilot,
                            lambda: isinstance(app.screen, bx.ExtendPollScreen),
                            timeout=30.0, what="missing-job dialog")
        check(ok, "missing: dialog appeared after MISSING_FRESH_SCANS")
        await pilot.click("#extend-fail")
        ok = await wait_for(pilot, lambda: not app._tabs["onedrive"].export_running,
                            timeout=30.0, what="missing-job run finish")
        check(ok, "missing: run finished")
        table = app.query_one("#onedrive-progress-table", DataTable)
        cell = str(table.get_cell("202", "status"))
        check("Failed" in cell, f"missing: Dan Diaz row = {cell!r}")
        cell = str(table.get_cell("201", "status"))
        check("Complete" in cell, f"missing: Carol Chen row = {cell!r}")
        MISSING_SIDS.clear()

        # ── adopt exports already on the server — no duplicate triggers ─────
        print("--- adopt existing server-side exports (onedrive) ---")
        reset_tab_state(app, "onedrive")
        bx._export_caches["onedrive"].clear()
        bx._export_pending["onedrive"].clear()
        ADOPT_TEST["active"] = True
        ADOPT_TEST["base"]   = scan_counts["onedrive"]
        before_tr = len(triggered["onedrive"])
        app._start_export_for("onedrive")
        await pilot.pause()
        ok = await wait_for(pilot, lambda: not app._tabs["onedrive"].export_running,
                            timeout=30.0, what="adoption run finish")
        check(ok, "adopt: run finished")
        ADOPT_TEST["active"] = False
        table = app.query_one("#onedrive-progress-table", DataTable)
        for u in FAKE_USERS["onedrive"]:
            cell = str(table.get_cell(u["service_id"], "status"))
            check("Complete" in cell, f"adopt: {u['name']} row = {cell!r}")
        check(len(triggered["onedrive"]) == before_tr,
              "adopt: no duplicate exports triggered")

        check(len(config_saves) == 9, f"save_config stubbed & called per run ({len(config_saves)}x)")

        app.save_screenshot(str(TMP / "tui.svg"))
        print(f"  screenshot: {TMP}/tui.svg")

    print()
    if failures:
        print(f"=== {len(failures)} FAILURE(S) ===")
        for f in failures:
            print("  -", f)
        sys.exit(1)
    print("=== ALL CHECKS PASSED ===")


asyncio.run(main())
