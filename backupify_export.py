# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "playwright>=1.44.0",
#   "httpx>=0.27.0",
#   "textual>=0.52.0",
# ]
# ///
"""
Backupify M365 Bulk Exporter
============================
TUI-driven tool for bulk-exporting Backupify/Datto M365 backups to local files.
Supports Exchange PST; OneDrive, SharePoint, and Teams planned (see ROADMAP.md).

Configuration lives at ~/.config/backupify_exporter/config.toml — an onboarding
wizard creates it on first run, or power users can write it by hand.

First-time setup — install Chromium browser driver (once only):
    uv run python -m playwright install chromium

Then just run:
    uv run backupify_export.py
"""

import asyncio
import copy
import html
import json
import logging
import re
import time
import tomllib
from datetime import datetime
from pathlib import Path

import httpx
from playwright.async_api import (
    async_playwright,
    BrowserContext,
    Error as PWError,
    TimeoutError as PWTimeout,
)

from rich.markup import escape as rich_escape

from textual import on, work
from textual.app import App, ComposeResult
from textual.binding import Binding
from textual.containers import Horizontal, Vertical
from textual.message import Message
from textual.screen import Screen
from textual.widgets import (
    Button,
    ContentSwitcher,
    DataTable,
    Footer,
    Header,
    Input,
    Label,
    RichLog,
    Select,
    SelectionList,
    Static,
    TabbedContent,
    TabPane,
)
from textual.widgets.selection_list import Selection


# ─── Config file ──────────────────────────────────────────────────────────────

CONFIG_DIR  = Path.home() / ".config" / "backupify_exporter"
CONFIG_FILE = CONFIG_DIR / "config.toml"

_DEFAULT_CONFIG: dict = {
    "account": {
        "email":        "",
        "totp_enabled": True,
    },
    "server": {
        "base_url":        "",
        "customer_id":     "",
        "ext_customer_id": "",
    },
    "defaults": {
        "concurrency": 8,
        "output_dir":  str(Path.home() / "backupify_exports"),
    },
}


def load_config() -> dict:
    cfg = copy.deepcopy(_DEFAULT_CONFIG)
    if not CONFIG_FILE.exists():
        return cfg
    try:
        with open(CONFIG_FILE, "rb") as fh:
            data = tomllib.load(fh)
        for section, values in data.items():
            if section in cfg and isinstance(cfg[section], dict):
                cfg[section].update(values)
            else:
                cfg[section] = values
    except Exception:
        pass
    return cfg


def save_config(cfg: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    lines = [
        "# Backupify Exporter configuration",
        "#",
        "# To find your server details, open Backupify in a browser.",
        "# The address bar will show something like:",
        "#   https://<base_url>/<customer_id>/o365?external_customer_id=<ext_customer_id>",
        "",
    ]
    for section, values in cfg.items():
        lines.append(f"[{section}]")
        for key, val in values.items():
            if isinstance(val, bool):
                lines.append(f"{key} = {'true' if val else 'false'}")
            elif isinstance(val, int):
                lines.append(f"{key} = {val}")
            else:
                escaped = str(val).replace("\\", "\\\\").replace('"', '\\"')
                lines.append(f'{key} = "{escaped}"')
        lines.append("")
    CONFIG_FILE.write_text("\n".join(lines), encoding="utf-8")


def config_is_complete(cfg: dict) -> bool:
    s = cfg.get("server", {})
    a = cfg.get("account", {})
    return bool(s.get("base_url") and s.get("customer_id") and a.get("email"))


# ─── Runtime globals (set by apply_config before any business logic runs) ─────

BASE_URL        = ""
CUSTOMER_ID     = ""
EXT_CUSTOMER_ID = ""
LOGIN_URL       = ""
DASHBOARD_URL   = ""
EXCHANGE_URL    = ""
EXPORT_ACTION   = ""
COOKIE_FILE: Path = Path.home() / ".backupify_session.json"
USERS_FILE:  Path = Path.home() / ".backupify_users_.json"


def apply_config(cfg: dict) -> None:
    """Overwrite module-level runtime globals from the loaded config dict."""
    global BASE_URL, CUSTOMER_ID, EXT_CUSTOMER_ID, LOGIN_URL
    global DASHBOARD_URL, EXCHANGE_URL, EXPORT_ACTION, COOKIE_FILE, USERS_FILE

    s = cfg["server"]
    a = cfg["account"]

    BASE_URL        = s.get("base_url", "").rstrip("/")
    CUSTOMER_ID     = s.get("customer_id", "")
    EXT_CUSTOMER_ID = s.get("ext_customer_id", "")
    LOGIN_URL       = f"https://auth.datto.com/login?login_hint={a.get('email', '')}"
    DASHBOARD_URL   = f"{BASE_URL}/{CUSTOMER_ID}/o365?external_customer_id={EXT_CUSTOMER_ID}"
    EXCHANGE_URL    = f"{BASE_URL}/{CUSTOMER_ID}/o365/exchange"
    EXPORT_ACTION   = f"{BASE_URL}/{CUSTOMER_ID}/restoreExportAction"
    COOKIE_FILE     = Path.home() / ".backupify_session.json"
    USERS_FILE      = Path.home() / f".backupify_users_{CUSTOMER_ID}.json"


# ─── Poll / download tunables ─────────────────────────────────────────────────

POLL_INTERVAL    = 30           # seconds between export-status scans
EXPORT_TIMEOUT   = 10_800       # max wait for one export job (3 h)
DOWNLOAD_TIMEOUT = 7_200        # max read time for one download (2 h)
CHUNK_SIZE       = 16 * 1024 * 1024

PAGE_LOAD_MS = 60_000
ELEMENT_MS   = 30_000
LOGIN_MS     = 300_000


# ─── Logging ──────────────────────────────────────────────────────────────────

def setup_logging(output_dir: Path) -> logging.Logger:
    """File-only logger — stdout belongs to the TUI."""
    log_path = output_dir / f"backupify_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    logger = logging.getLogger("backupify")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        fh = logging.FileHandler(log_path, encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        logger.addHandler(fh)
    return logger


def setup_playwright_debug_logging() -> logging.Logger:
    debug_log = logging.getLogger("pw_debug")
    debug_log.setLevel(logging.DEBUG)
    debug_log.propagate = False
    if not debug_log.handlers:
        handler = logging.FileHandler("playwright_debug.log", encoding="utf-8", mode="w")
        handler.setFormatter(logging.Formatter("%(asctime)s %(message)s", datefmt="%H:%M:%S"))
        debug_log.addHandler(handler)
    return debug_log


def attach_page_listeners(page, debug_log: logging.Logger) -> None:
    page.on("console", lambda msg: debug_log.debug(
        f"[CONSOLE:{msg.type.upper():7}] {msg.text}"
    ))
    page.on("pageerror", lambda exc: debug_log.debug(f"[JSERROR ] {exc}"))

    def _log_request(req):
        line = f"[REQUEST ] {req.method:<6} {req.url}"
        if req.method == "POST" and req.post_data and BASE_URL in req.url:
            line += f"\n           BODY: {req.post_data[:400]}"
        debug_log.debug(line)

    page.on("request", _log_request)
    page.on("response", lambda resp: debug_log.debug(
        f"[RESPONSE] {resp.status} {resp.url}"
    ))
    page.on("framenavigated", lambda frame: (
        debug_log.debug(f"[NAVIGATE] {frame.url}") if frame == page.main_frame else None
    ))


# ─── State / resume ───────────────────────────────────────────────────────────

class StateManager:
    def __init__(self, output_dir: Path):
        output_dir.mkdir(parents=True, exist_ok=True)
        self.path  = output_dir / "progress.json"
        self.state = self._load()

    def _load(self) -> dict:
        if self.path.exists():
            try:
                with open(self.path) as f:
                    data = json.load(f)
                    data.setdefault("in_progress", {})
                    return data
            except Exception:
                pass
        return {"completed": {}, "failed": {}, "in_progress": {}}

    def save(self):
        with open(self.path, "w") as f:
            json.dump(self.state, f, indent=2)

    def is_done(self, service_id: str) -> bool:
        return service_id in self.state["completed"]

    def is_failed(self, service_id: str) -> bool:
        return service_id in self.state["failed"]

    def get_in_progress(self, service_id: str) -> dict | None:
        return self.state["in_progress"].get(service_id)

    def mark_in_progress(self, service_id: str, job_id: str, snapshot_id: str):
        self.state["in_progress"][service_id] = {
            "job_id":       job_id,
            "snapshot_id":  snapshot_id,
            "triggered_at": datetime.now().isoformat(),
        }
        self.save()

    def clear_in_progress(self, service_id: str):
        self.state["in_progress"].pop(service_id, None)
        self.save()

    def mark_complete(self, service_id: str, filename: str):
        self.state["completed"][service_id] = {
            "filename":  filename,
            "timestamp": datetime.now().isoformat(),
        }
        self.state["in_progress"].pop(service_id, None)
        self.save()

    def mark_failed(self, service_id: str, reason: str):
        self.state["failed"][service_id] = {
            "reason":    reason,
            "timestamp": datetime.now().isoformat(),
        }
        self.state["in_progress"].pop(service_id, None)
        self.save()


# ─── Authentication ───────────────────────────────────────────────────────────

async def _save_cookies(context: BrowserContext) -> None:
    cookies = await context.cookies()
    COOKIE_FILE.write_text(json.dumps(cookies, indent=2), encoding="utf-8")


async def _try_restore_session(context: BrowserContext, logger: logging.Logger) -> bool:
    """
    Loads cookies from COOKIE_FILE into the browser context, navigates to
    the dashboard, and returns True if the session is still valid.
    """
    if not COOKIE_FILE.exists():
        return False
    try:
        cookies = json.loads(COOKIE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return False
    if not cookies:
        return False
    try:
        await context.add_cookies(cookies)
    except Exception as e:
        logger.warning(f"Could not restore saved cookies: {e}")
        return False

    logger.info("Saved session found — checking if it's still valid...")
    page = await context.new_page()
    try:
        await page.goto(DASHBOARD_URL, wait_until="domcontentloaded", timeout=30_000)
        if BASE_URL not in page.url:
            logger.info(f"Session expired (redirected to {page.url[:60]}...).")
            return False
        for sel in ["nav", "[class*='nav']", "[class*='sidebar']",
                    "a[href*='/o365']", ".app-container", "#app"]:
            try:
                await page.wait_for_selector(sel, timeout=5_000)
                logger.info("Saved session is valid — skipping interactive login.")
                return True
            except PWTimeout:
                continue
        logger.info("Dashboard element not found after session restore.")
        return False
    except PWTimeout:
        logger.warning("Timed out while validating saved session.")
        return False
    finally:
        await page.close()


def _make_listener(debug_log):
    return lambda p: attach_page_listeners(p, debug_log)


async def login(
    logger: logging.Logger,
    debug_log: logging.Logger,
    on_status=None,
) -> BrowserContext:
    """
    Returns an authenticated BrowserContext.  Tries a saved session first
    (fully headless).  Falls back to a visible browser for interactive login,
    then saves the cookies so the next run is headless.

    on_status: optional callable(str) — phase updates for a UI to display.
    """
    def _status(text: str) -> None:
        if on_status:
            on_status(text)

    _status("Starting browser engine…")
    pw = await async_playwright().start()
    _args = ["--no-first-run", "--no-default-browser-check", "--disable-extensions"]

    if COOKIE_FILE.exists():
        _status("Validating saved session…")
        browser = await pw.chromium.launch(headless=True, args=_args)
        context = await browser.new_context()
        context.on("page", _make_listener(debug_log))
        if await _try_restore_session(context, logger):
            return context
        logger.info("Saved session invalid — falling back to interactive login.")
        await browser.close()

    _status("Browser window opened — complete login there (password + TOTP)")
    logger.info("Opening browser for interactive login (password + TOTP if enabled).")
    browser = await pw.chromium.launch(headless=False, slow_mo=50, args=_args)
    context = await browser.new_context()
    context.on("page", _make_listener(debug_log))

    page = await context.new_page()
    page.on("framenavigated", lambda frame: (
        logger.info(f"  → {frame.url}") if frame == page.main_frame else None
    ))

    await page.goto(LOGIN_URL, wait_until="domcontentloaded")

    deadline = time.monotonic() + (LOGIN_MS / 1000)
    while time.monotonic() < deadline:
        await asyncio.sleep(1)
        if BASE_URL not in page.url:
            continue
        for sel in ["nav", "[class*='nav']", "[class*='sidebar']", "[class*='dashboard']",
                    "[class*='header']", "a[href*='/o365']", ".app-container", "#app"]:
            try:
                await page.wait_for_selector(sel, timeout=5_000)
                logger.info("Login complete.")
                await _save_cookies(context)
                logger.info(f"Session saved to {COOKIE_FILE}")
                return context
            except PWTimeout:
                continue
        await asyncio.sleep(2)

    await browser.close()
    raise RuntimeError(f"Login timed out after 5 minutes. Last URL: {page.url}")


async def extract_cookies(context: BrowserContext) -> dict:
    all_cookies = await context.cookies()
    wanted      = {"__backupify_session", "PHPSESSID"}
    cookies     = {c["name"]: c["value"] for c in all_cookies if c["name"] in wanted}
    missing     = wanted - set(cookies.keys())
    if missing:
        found = [c["name"] for c in all_cookies]
        raise RuntimeError(
            f"Missing session cookies: {missing}\nCookies present: {found}"
        )
    return cookies


# ─── User discovery ───────────────────────────────────────────────────────────

async def get_all_users(
    context: BrowserContext,
    cookies: dict,
    logger: logging.Logger,
) -> list[dict]:
    """
    Returns the full Exchange user list.  Loads from USERS_FILE cache on
    subsequent runs; delete the file (or press F5 in the TUI) to re-fetch.

    Returns [{"service_id", "name", "email", "snapshot_id"}, ...].
    """
    if USERS_FILE.exists():
        try:
            users = json.loads(USERS_FILE.read_text(encoding="utf-8"))
            if users:
                for u in users:  # older caches stored HTML entities (&#039; etc.)
                    u["name"]  = html.unescape(u.get("name")  or "")
                    u["email"] = html.unescape(u.get("email") or "")
                logger.info(f"Loaded {len(users)} users from cache ({USERS_FILE}).")
                return users
        except Exception as e:
            logger.warning(f"Could not read user cache: {e} — re-fetching.")

    logger.info("Loading Exchange user list via customerServices XHR...")
    page = await context.new_page()

    captured_req: dict = {}

    def _capture_request(req):
        if "/customerServices" in req.url and req.method == "POST" and not captured_req:
            captured_req["url"]       = req.url
            captured_req["post_data"] = req.post_data or ""
            captured_req["headers"]   = {
                k: v for k, v in req.headers.items()
                if k.lower() not in ("host", "content-length")
            }

    page.on("request", _capture_request)

    first_data: dict = {}
    try:
        async with page.expect_response(
            lambda r: "/customerServices" in r.url and r.status == 200,
            timeout=45_000,
        ) as resp_info:
            await page.goto(EXCHANGE_URL, wait_until="domcontentloaded")

        first_data    = await (await resp_info.value).json()
        records_total = first_data.get("recordsTotal", 0)
        logger.info(
            f"customerServices page-1: {len(first_data.get('data', []))} rows, "
            f"{records_total} total."
        )
    except PWTimeout:
        await page.close()
        raise RuntimeError("customerServices XHR not seen within 45 s.")

    await page.close()

    all_rows: list[dict] = list(first_data.get("data", []))
    records_total = first_data.get("recordsTotal", len(all_rows))

    if records_total > len(all_rows) and captured_req:
        logger.info(f"Paginating to fetch all {records_total} records...")
        headers = {k: v for k, v in captured_req["headers"].items() if k.lower() != "cookie"}
        headers["cookie"] = "; ".join(f"{k}={v}" for k, v in cookies.items())
        draw_num = 2
        offset   = len(all_rows)

        async with httpx.AsyncClient(follow_redirects=False, timeout=60.0) as client:
            while offset < records_total:
                body = captured_req["post_data"]
                body = re.sub(r"\bstart=\d+", f"start={offset}", body)
                body = re.sub(r"\bdraw=\d+",  f"draw={draw_num}",  body)
                draw_num += 1
                resp = await client.post(
                    captured_req["url"], content=body.encode(), headers=headers
                )
                if resp.status_code != 200:
                    logger.warning(f"Pagination at start={offset} returned {resp.status_code}.")
                    break
                batch = resp.json().get("data", [])
                if not batch:
                    break
                all_rows.extend(batch)
                offset += len(batch)
                logger.info(f"  {offset}/{records_total} users fetched...")

    users: list[dict] = []
    for row in all_rows:
        service_id = str(row.get("id") or "")
        if not service_id:
            continue
        name  = html.unescape(re.sub(r"<[^>]+>", "", str(row.get("name")  or ""))).strip()
        email = html.unescape(re.sub(r"<[^>]+>", "", str(row.get("email") or ""))).strip()
        backups     = row.get("perfectBackups") or []
        snapshot_id = (
            str(max(backups, key=lambda b: b.get("snapshotId", 0))["snapshotId"])
            if backups else None
        )
        users.append({
            "service_id":  service_id,
            "name":        name or service_id,
            "email":       email,
            "snapshot_id": snapshot_id,
        })

    if not users:
        raise RuntimeError("No users extracted from customerServices response.")

    logger.info(f"Found {len(users)} Exchange users.")
    USERS_FILE.write_text(json.dumps(users, indent=2), encoding="utf-8")
    logger.info(f"User list cached to {USERS_FILE}.")
    return users


# ─── Snapshot extraction ──────────────────────────────────────────────────────

async def get_latest_snapshot_id(
    context: BrowserContext,
    service_id: str,
    logger: logging.Logger,
) -> str | None:
    service_url = f"{BASE_URL}/{CUSTOMER_ID}/o365/exchange/service?serviceId={service_id}"
    page = await context.new_page()
    try:
        await page.goto(service_url, wait_until="domcontentloaded")
        await page.wait_for_load_state("networkidle", timeout=PAGE_LOAD_MS)

        snapshot_id = await page.evaluate("""() => {
            if (window.__INITIAL_STATE__) {
                const s = JSON.stringify(window.__INITIAL_STATE__);
                const m = s.match(/"snapshotId"\\s*:\\s*"?(\\d+)"?/);
                if (m) return m[1];
            }
            if (window.__APP_STATE__) {
                const s = JSON.stringify(window.__APP_STATE__);
                const m = s.match(/"snapshotId"\\s*:\\s*"?(\\d+)"?/);
                if (m) return m[1];
            }
            const el = document.querySelector('[data-snapshot-id]');
            if (el) return el.getAttribute('data-snapshot-id');
            const snap = document.querySelector('.snapshot-selector, #snapshot-selector, [class*="snapshot"]');
            if (snap) {
                const attr = snap.getAttribute('data-value') || snap.getAttribute('data-id') || snap.value;
                if (attr && /^\\d+$/.test(attr)) return attr;
            }
            return null;
        }""")

        if snapshot_id:
            return str(snapshot_id)

        html = await page.content()
        for pat in [
            r'"snapshotId"\s*:\s*"?(\d{10,})"?',
            r'data-snapshot-id=["\'](\d+)["\']',
            r'snapshot_id["\']?\s*[=:]\s*["\']?(\d{10,})',
            r'value=["\'](\d{13})["\']',
        ]:
            m = re.search(pat, html)
            if m:
                return m.group(1)

        logger.warning(f"[{service_id}] Could not find snapshot ID.")
        return None
    except PWTimeout as e:
        logger.error(f"[{service_id}] Timeout loading service page: {e}")
        return None
    finally:
        await page.close()


# ─── Export triggering ────────────────────────────────────────────────────────

async def trigger_export(
    context: BrowserContext,
    cookies: dict,
    service_id: str,
    snapshot_id: str,
    logger: logging.Logger,
) -> str | None:
    payload = {
        "actionType":         "export",
        "appType":            "office365_exchange",
        "snapshotId":         snapshot_id,
        "token":              "",
        "exportFormat":       "pst",
        "includePermissions": "false",
        "includeAttachments": "false",
        "services[]":         service_id,
    }
    cookie_header = "; ".join(f"{k}={v}" for k, v in cookies.items())
    async with httpx.AsyncClient(follow_redirects=True, timeout=60.0) as client:
        resp = await client.post(
            EXPORT_ACTION,
            data=payload,
            headers={
                "Cookie":           cookie_header,
                "Referer":          f"{BASE_URL}/{CUSTOMER_ID}/o365/exchange/service?serviceId={service_id}",
                "User-Agent":       "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36",
                "X-Requested-With": "XMLHttpRequest",
                "Accept":           "application/json, text/javascript, */*",
            },
        )
    if resp.status_code != 200:
        logger.error(f"[{service_id}] Export POST returned HTTP {resp.status_code}: {resp.text[:300]}")
        return None
    try:
        data   = resp.json()
        job_id = str(data["responseData"]["id"])
        logger.info(f"[{service_id}] Export job started — ID: {job_id}")
        return job_id
    except (KeyError, ValueError) as e:
        logger.error(f"[{service_id}] Unexpected export response: {resp.text[:300]} — {e}")
        return None


# ─── Export polling ───────────────────────────────────────────────────────────

_export_cache:      dict  = {}
_export_cache_time: float = 0.0
_export_scan_lock: asyncio.Lock | None = None


async def scan_export_page(context: BrowserContext, logger: logging.Logger) -> list[dict]:
    """
    Navigates to the export page and parses all server-rendered DataTable rows,
    paginating as needed.

    Each record: {"job_id", "source_name", "status", "download_url"}.
    """
    export_page_url = f"{BASE_URL}/{CUSTOMER_ID}/o365/exchange/export"
    page = await context.new_page()
    records: list[dict] = []

    try:
        await page.goto(export_page_url, wait_until="domcontentloaded", timeout=PAGE_LOAD_MS)
        try:
            await page.wait_for_selector("#exportListItems tbody tr", timeout=ELEMENT_MS)
        except PWTimeout:
            logger.info("Export table empty — no export jobs yet.")
            return []

        prev_first: str | None = None
        pages = 0
        while True:
            batch = await page.evaluate("""() => {
                const rows = document.querySelectorAll('#exportListItems tbody tr');
                return Array.from(rows).map(tr => {
                    const cells = tr.querySelectorAll('td');
                    const source = cells[0] ? cells[0].innerText.trim() : '';
                    const status = cells[7] ? cells[7].innerText.trim().toLowerCase() : '';
                    let job_id = null, download_href = null;
                    if (cells[8]) {
                        for (const a of cells[8].querySelectorAll('a[href]')) {
                            const href = a.getAttribute('href') || '';
                            const m = href.match(/[?&]id=(\\d+)/);
                            if (m) job_id = m[1];
                            if (href.includes('/download')) download_href = href;
                        }
                    }
                    return { source, status, job_id, download_href };
                });
            }""")

            first_id = batch[0]["job_id"] if batch else None
            if first_id is not None and first_id == prev_first:
                break  # pagination didn't advance — we're done
            prev_first = first_id

            for row in batch:
                href = row.pop("download_href") or ""
                if href and not href.startswith("http"):
                    href = BASE_URL + href
                row["download_url"] = href
                row["source_name"]  = row.pop("source")
                records.append(row)

            pages += 1
            if pages >= 40:
                logger.warning("Export page scan stopped after 40 pages.")
                break
            # The heartbeat XHR can redraw the table at any moment, detaching
            # the button between locate and click — retry with a fresh locator,
            # then settle for the rows scraped so far.
            clicked = False
            for _ in range(3):
                try:
                    next_btn = page.locator("#exportListItems_next:not(.disabled)")
                    if not await next_btn.count():
                        break
                    await next_btn.first.click(timeout=5_000)
                    clicked = True
                    break
                except PWError:
                    await page.wait_for_timeout(1_000)
            if not clicked:
                break
            # "networkidle" never settles on this page — give the AJAX redraw
            # a fixed moment to land instead.
            await page.wait_for_timeout(2_500)

    except PWError as e:
        logger.warning(f"Export page scan aborted, keeping {len(records)} row(s): {e}")
    finally:
        await page.close()

    return records


async def refresh_export_cache(context: BrowserContext, logger: logging.Logger) -> None:
    """Scans the export page once and updates the shared status cache."""
    global _export_cache, _export_cache_time
    records = await scan_export_page(context, logger)
    for rec in records:
        if rec["job_id"]:
            _export_cache[rec["job_id"]] = rec
    _export_cache_time = time.monotonic()


async def poll_for_download_url(
    context: BrowserContext,
    job_id: str,
    service_id: str,
    logger: logging.Logger,
    on_status=None,
    max_unseen: int = 4,
) -> tuple[str, str | None]:
    """
    Waits until the export job is completed and returns (outcome, url):
      ("ok", url)       — job completed, download URL available
      ("missing", None) — job absent from the export page for max_unseen scans
      ("failed", None)  — job ended in failed/error/cancelled
      ("timeout", None) — EXPORT_TIMEOUT exceeded

    All concurrent workers share one page-scan via _export_scan_lock.
    on_status: optional callable(str) for live UI updates each poll.
    """
    global _export_cache_time

    start          = time.monotonic()
    deadline       = start + EXPORT_TIMEOUT
    unseen_scans   = 0
    last_scan_time = _export_cache_time

    while time.monotonic() < deadline:
        cache_age = time.monotonic() - _export_cache_time
        if cache_age >= POLL_INTERVAL:
            async with _export_scan_lock:
                if time.monotonic() - _export_cache_time >= POLL_INTERVAL:
                    try:
                        await refresh_export_cache(context, logger)
                    except Exception as e:
                        # A flaky scan must not kill the worker — back off
                        # until the next interval and try again.
                        logger.warning(f"Export scan failed (will retry): {e}")
                        _export_cache_time = time.monotonic()

        record  = _export_cache.get(str(job_id))
        elapsed = int(time.monotonic() - start)

        if record:
            unseen_scans = 0
            status = record["status"]
            if on_status:
                on_status(f"job {job_id}: {status} · {elapsed // 60}m {elapsed % 60:02d}s")
            if status == "completed":
                url = record["download_url"]
                if url:
                    logger.info(f"[{service_id}] Export ready after {elapsed}s.")
                    return ("ok", url)
            elif status in ("failed", "error", "cancelled"):
                logger.error(f"[{service_id}] Export job {job_id} ended: {status}")
                return ("failed", None)
        else:
            # Count scans that completed while the job was still absent
            if _export_cache_time != last_scan_time:
                unseen_scans  += 1
                last_scan_time = _export_cache_time
            if on_status:
                on_status(f"job {job_id}: not listed yet (scan {unseen_scans}/{max_unseen})")
            if unseen_scans >= max_unseen:
                logger.warning(
                    f"[{service_id}] Job {job_id} absent after {unseen_scans} scans — treating as stale."
                )
                return ("missing", None)

        await asyncio.sleep(POLL_INTERVAL)

    logger.error(f"[{service_id}] Export timed out after {EXPORT_TIMEOUT}s.")
    return ("timeout", None)


# ─── File download ────────────────────────────────────────────────────────────

def pst_path(output_dir: Path, user_name: str, service_id: str) -> Path:
    """Canonical on-disk path for a mailbox export (server sends raw PST)."""
    safe_name = re.sub(r'[<>:"/\\|?*\s]', "_", user_name).strip("_")
    return output_dir / f"{safe_name}__{service_id}.pst"


async def download_file(
    download_url: str,
    cookies: dict,
    output_dir: Path,
    user_name: str,
    service_id: str,
    logger: logging.Logger,
    on_progress=None,
) -> str | None:
    if download_url.startswith("/"):
        download_url = BASE_URL + download_url

    cookie_header = "; ".join(f"{k}={v}" for k, v in cookies.items())
    filename      = pst_path(output_dir, user_name, service_id)
    part_file     = filename.with_name(filename.name + ".part")

    if filename.exists():
        size_mb = filename.stat().st_size / 1_048_576
        logger.info(f"[{service_id}] File already exists ({size_mb:.1f} MB), skipping.")
        return filename.name

    if part_file.exists():
        logger.info(f"[{service_id}] Removing stale partial file {part_file.name}")
        part_file.unlink()

    logger.info(f"[{service_id}] Downloading → {filename.name}")

    try:
        async with httpx.AsyncClient(
            follow_redirects=True,
            timeout=httpx.Timeout(connect=30.0, read=DOWNLOAD_TIMEOUT, write=30.0, pool=30.0),
        ) as client:
            async with client.stream(
                "GET",
                download_url,
                headers={
                    "Cookie":     cookie_header,
                    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36",
                },
            ) as resp:
                if resp.status_code != 200:
                    logger.error(f"[{service_id}] Download HTTP {resp.status_code}")
                    return None

                total      = int(resp.headers.get("content-length", 0))
                downloaded = 0

                with open(part_file, "wb") as f:
                    async for chunk in resp.aiter_bytes(CHUNK_SIZE):
                        f.write(chunk)
                        downloaded += len(chunk)
                        if on_progress:
                            on_progress(downloaded, total)

        part_file.rename(filename)
        size_mb = filename.stat().st_size / 1_048_576
        logger.info(f"[{service_id}] Saved {filename.name} ({size_mb:.1f} MB)")
        return filename.name

    except (httpx.RequestError, OSError) as e:
        logger.error(f"[{service_id}] Download failed: {e}")
        if part_file.exists():
            part_file.unlink()
        return None


# ─── TUI — sort helpers ───────────────────────────────────────────────────────

_SORT_OPTIONS = [
    ("Alphabetical (A → Z)",   "az"),
    ("Alphabetical (Z → A)",   "za"),
    ("Largest mailbox first",  "size_desc"),
    ("Smallest mailbox first", "size_asc"),
]

_STATUS_ICONS = {
    "queued":      "○",
    "triggering":  "⟳",
    "polling":     "⟳",
    "resuming":    "⟳",
    "downloading": "↓",
    "complete":    "✓",
    "failed":      "✗",
    "skipped":     "—",
}


def _sort_users(users: list[dict], key: str) -> list[dict]:
    if key == "az":
        return sorted(users, key=lambda u: u["name"].lower())
    if key == "za":
        return sorted(users, key=lambda u: u["name"].lower(), reverse=True)
    # snapshot_id is epoch-ms — used as a size proxy until real size data is
    # extracted (see ROADMAP.md: "Actual size-based sorting")
    if key == "size_desc":
        return sorted(users, key=lambda u: int(u.get("snapshot_id") or 0), reverse=True)
    if key == "size_asc":
        return sorted(users, key=lambda u: int(u.get("snapshot_id") or 0))
    return users


# ─── TUI — messages ───────────────────────────────────────────────────────────

class SessionReady(Message):
    pass


class StatusUpdate(Message):
    def __init__(self, text: str) -> None:
        super().__init__()
        self.text = text


class UsersLoaded(Message):
    def __init__(self, users: list[dict]) -> None:
        super().__init__()
        self.users = users


class LogEntry(Message):
    def __init__(self, level: str, text: str) -> None:
        super().__init__()
        self.level = level
        self.text  = text


class ExportUpdate(Message):
    def __init__(self, service_id: str, status: str, detail: str = "") -> None:
        super().__init__()
        self.service_id = service_id
        self.status     = status
        self.detail     = detail


class DownloadProgress(Message):
    def __init__(self, service_id: str, done: int, total: int) -> None:
        super().__init__()
        self.service_id = service_id
        self.done       = done
        self.total      = total


# ─── TUI — onboarding screen ──────────────────────────────────────────────────

class OnboardingScreen(Screen):
    CSS = """
    OnboardingScreen { align: center middle; }
    #ob-card {
        width: 76;
        height: auto;
        max-height: 100%;
        overflow-y: auto;
        border: double $accent;
        padding: 1 2;
        background: $surface;
    }
    #ob-title {
        text-align: center;
        text-style: bold;
        color: $accent;
        padding-bottom: 1;
    }
    #ob-indicator { text-align: center; color: $text-muted; margin-bottom: 1; }
    .ob-label { text-style: bold; margin-top: 1; }
    .ob-hint  { color: $text-muted; margin-bottom: 1; }
    ContentSwitcher { height: auto; }
    ContentSwitcher > Vertical { height: auto; }
    #ob-nav { height: 3; margin-top: 1; align: right middle; }
    #ob-nav Button { margin-left: 1; }
    """

    _STEP_TITLES = ["Welcome", "Server details", "Account", "Defaults"]

    def __init__(self, cfg: dict) -> None:
        super().__init__()
        self._cfg  = cfg
        self._step = 0

    def compose(self) -> ComposeResult:
        with Vertical(id="ob-card"):
            yield Static("⚡  Backupify Exporter — First Run Setup", id="ob-title")
            yield Static(
                f"Step 1 / {len(self._STEP_TITLES)}: {self._STEP_TITLES[0]}",
                id="ob-indicator",
            )
            with ContentSwitcher(initial="ob-step-0"):
                with Vertical(id="ob-step-0"):
                    yield Static(
                        "Welcome! This tool bulk-exports M365 mailboxes (and soon\n"
                        "OneDrive, SharePoint, and Teams) from Backupify/Datto.\n\n"
                        "You will need:\n"
                        "  • Your Backupify admin account email\n"
                        "  • The server URL and customer ID from the browser address bar\n\n"
                        "Power user tip: once created you can edit the config directly:\n"
                        f"  {CONFIG_FILE}",
                        classes="ob-hint",
                    )

                with Vertical(id="ob-step-1"):
                    yield Static(
                        "Open Backupify in your browser. The URL looks like:\n"
                        "  https://SERVER.backupify.com/CUSTOMER_ID/o365\n"
                        "      ?external_customer_id=EXT_ID\n"
                        "Copy each piece into the fields below.",
                        classes="ob-hint",
                    )
                    yield Static("Server URL", classes="ob-label")
                    yield Input(
                        value=self._cfg["server"]["base_url"],
                        placeholder="https://xxx-ext.backupify.com",
                        id="ob-base-url",
                    )
                    yield Static("Customer ID  (the number in the URL path)", classes="ob-label")
                    yield Input(
                        value=self._cfg["server"]["customer_id"],
                        placeholder="123456",
                        id="ob-customer-id",
                    )
                    yield Static("External Customer ID  (may be blank)", classes="ob-label")
                    yield Input(
                        value=self._cfg["server"]["ext_customer_id"],
                        placeholder="xxxxxxxx-xxxx-...",
                        id="ob-ext-id",
                    )

                with Vertical(id="ob-step-2"):
                    yield Static("Your Backupify admin account details.", classes="ob-hint")
                    yield Static("Email address", classes="ob-label")
                    yield Input(
                        value=self._cfg["account"]["email"],
                        placeholder="you@company.com",
                        id="ob-email",
                    )
                    yield Static("Two-factor authentication", classes="ob-label")
                    yield Select(
                        [
                            ("Yes — I enter a 6-digit code on login", "true"),
                            ("No — password only",                    "false"),
                        ],
                        value="true" if self._cfg["account"]["totp_enabled"] else "false",
                        allow_blank=False,
                        id="ob-totp",
                    )

                with Vertical(id="ob-step-3"):
                    yield Static(
                        "Where should exports be saved, and how many run at once?\n"
                        "Each service type gets its own subfolder (exchange/, onedrive/, ...).",
                        classes="ob-hint",
                    )
                    yield Static("Default output directory", classes="ob-label")
                    yield Input(
                        value=self._cfg["defaults"]["output_dir"],
                        placeholder="/path/to/exports",
                        id="ob-output",
                    )
                    yield Static("Concurrent exports  (2–16 recommended)", classes="ob-label")
                    yield Input(
                        value=str(self._cfg["defaults"]["concurrency"]),
                        placeholder="8",
                        id="ob-concurrency",
                    )

            with Horizontal(id="ob-nav"):
                yield Button("← Back", id="ob-back", variant="default", disabled=True)
                yield Button("Next →", id="ob-next", variant="primary")

    def _update_nav(self) -> None:
        self.query_one("#ob-indicator", Static).update(
            f"Step {self._step + 1} / {len(self._STEP_TITLES)}: "
            f"{self._STEP_TITLES[self._step]}"
        )
        self.query_one("#ob-back", Button).disabled = (self._step == 0)
        last = self._step == len(self._STEP_TITLES) - 1
        self.query_one("#ob-next", Button).label = "Finish ✓" if last else "Next →"

    def _collect(self) -> bool:
        if self._step == 1:
            base_url    = self.query_one("#ob-base-url",    Input).value.strip().rstrip("/")
            customer_id = self.query_one("#ob-customer-id", Input).value.strip()
            ext_id      = self.query_one("#ob-ext-id",      Input).value.strip()
            if not base_url or not customer_id:
                self.notify("Server URL and Customer ID are required.", severity="error")
                return False
            if not base_url.startswith("http"):
                base_url = "https://" + base_url
            self._cfg["server"]["base_url"]        = base_url
            self._cfg["server"]["customer_id"]     = customer_id
            self._cfg["server"]["ext_customer_id"] = ext_id

        elif self._step == 2:
            email    = self.query_one("#ob-email", Input).value.strip()
            totp_val = self.query_one("#ob-totp",  Select).value
            if not email or "@" not in email:
                self.notify("A valid email address is required.", severity="error")
                return False
            self._cfg["account"]["email"]        = email
            self._cfg["account"]["totp_enabled"] = (totp_val != "false")

        elif self._step == 3:
            out_dir  = self.query_one("#ob-output",      Input).value.strip()
            conc_str = self.query_one("#ob-concurrency", Input).value.strip()
            if not out_dir:
                self.notify("Output directory is required.", severity="error")
                return False
            try:
                conc = int(conc_str)
                if not 1 <= conc <= 32:
                    raise ValueError
            except ValueError:
                self.notify("Concurrency must be a whole number 1–32.", severity="error")
                return False
            self._cfg["defaults"]["output_dir"]  = out_dir
            self._cfg["defaults"]["concurrency"] = conc

        return True

    def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "ob-back":
            self._step = max(0, self._step - 1)
            self.query_one(ContentSwitcher).current = f"ob-step-{self._step}"
            self._update_nav()

        elif event.button.id == "ob-next":
            if not self._collect():
                return
            if self._step == len(self._STEP_TITLES) - 1:
                save_config(self._cfg)
                self.app.exit(self._cfg)
            else:
                self._step += 1
                self.query_one(ContentSwitcher).current = f"ob-step-{self._step}"
                self._update_nav()


class OnboardingApp(App):
    """Minimal host app for the onboarding wizard. exit() returns the config."""

    def __init__(self, cfg: dict) -> None:
        super().__init__()
        self._cfg = cfg

    def on_mount(self) -> None:
        self.push_screen(OnboardingScreen(self._cfg))


# ─── TUI — main app ───────────────────────────────────────────────────────────

class BackupifyApp(App):
    TITLE = "Backupify Exporter"

    CSS = """
    TabbedContent { height: 1fr; }
    /* Textual defaults TabPane + inner ContentSwitcher to height:auto, which
       collapses children using percentage heights to 0 — pin them to fill */
    TabbedContent > ContentSwitcher { height: 1fr; }
    TabPane { height: 100%; }

    #exchange-view { height: 100%; layout: vertical; }

    #settings-bar {
        height: 3;
        padding: 0 1;
        background: $boost;
        align: left middle;
    }
    #settings-bar Label { width: auto; padding: 0 1; color: $text-muted; }
    #output-dir-input  { width: 1fr; }
    #concurrency-input { width: 6; }
    #sort-select       { width: 30; }

    #main-panel { height: 1fr; }

    #user-panel {
        width: 40%;
        layout: vertical;
        border-right: solid $primary;
    }
    #search-input { dock: top; }
    #exchange-user-list { height: 1fr; }
    #user-stats {
        height: 1;
        padding: 0 1;
        color: $text-muted;
        background: $boost;
    }
    #user-actions {
        height: 3;
        layout: horizontal;
        background: $boost;
        align: left middle;
        padding: 0 1;
    }
    #user-actions Button { margin-right: 1; min-width: 8; }

    #export-panel { width: 60%; layout: vertical; }
    #progress-table { height: 1fr; }
    #log-pane { height: 10; border-top: solid $primary; }

    #control-bar {
        height: 3;
        layout: horizontal;
        background: $boost;
        align: left middle;
        padding: 0 1;
    }
    #control-bar Button { margin-right: 1; }
    #status-text { color: $text-muted; content-align: left middle; width: 1fr; height: 3; }

    .coming-soon {
        height: 100%;
        content-align: center middle;
        color: $text-muted;
    }
    """

    BINDINGS = [
        Binding("q",      "quit",          "Quit"),
        Binding("ctrl+s", "start_export",  "Start"),
        Binding("f5",     "refresh_users", "Refresh Users"),
        Binding("ctrl+r", "retry_failed",  "Retry Failed"),
    ]

    def __init__(self, cfg: dict) -> None:
        super().__init__()
        self._cfg = cfg

        self._pw_context: BrowserContext | None = None
        self._cookies: dict = {}

        # Exchange tab state — OneDrive/SharePoint/Teams get their own
        # parallel sets when implemented (different entity types per tab)
        self._exchange_users:    list[dict] = []
        self._exchange_filtered: list[dict] = []
        self._exchange_selected: set[str]   = set()
        self._exchange_state:    StateManager | None = None
        self._exchange_sort   = "az"
        self._exchange_search = ""

        self._bfy_logger:    logging.Logger | None = None
        self._debug_log: logging.Logger | None = None
        self._semaphore: asyncio.Semaphore | None = None
        self._export_running = False
        self._run_pending: set[str] = set()  # service_ids not yet finished this run

    # ── Layout ────────────────────────────────────────────────────────────

    def compose(self) -> ComposeResult:
        yield Header()
        with TabbedContent(initial="tab-exchange"):
            with TabPane("📧 Exchange", id="tab-exchange"):
                with Vertical(id="exchange-view"):
                    with Horizontal(id="settings-bar"):
                        yield Label("Output:")
                        yield Input(
                            value=self._cfg["defaults"]["output_dir"],
                            placeholder="/path/to/exports",
                            id="output-dir-input",
                        )
                        yield Label("Workers:")
                        yield Input(
                            value=str(self._cfg["defaults"]["concurrency"]),
                            id="concurrency-input",
                            restrict=r"[0-9]*",
                        )
                        yield Label("Sort:")
                        yield Select(
                            _SORT_OPTIONS,
                            value="az",
                            allow_blank=False,
                            id="sort-select",
                        )
                    with Horizontal(id="main-panel"):
                        with Vertical(id="user-panel"):
                            yield Input(placeholder="🔍 Filter users…", id="search-input")
                            yield SelectionList(id="exchange-user-list")
                            yield Static("Loading…", id="user-stats")
                            with Horizontal(id="user-actions"):
                                yield Button("All",  id="btn-all",  variant="default")
                                yield Button("None", id="btn-none", variant="default")
                        with Vertical(id="export-panel"):
                            yield DataTable(
                                id="progress-table",
                                zebra_stripes=True,
                                cursor_type="row",
                            )
                            yield RichLog(id="log-pane", highlight=True, markup=True)
                    with Horizontal(id="control-bar"):
                        yield Button("▶ Start",        id="btn-start", variant="success", disabled=True)
                        yield Button("↺ Retry Failed", id="btn-retry", variant="default")
                        yield Static("⟳ Starting up…", id="status-text")

            with TabPane("📁 OneDrive", id="tab-onedrive"):
                yield Static("OneDrive export — coming soon (see ROADMAP.md)", classes="coming-soon")
            with TabPane("📋 SharePoint", id="tab-sharepoint"):
                yield Static("SharePoint export — coming soon (see ROADMAP.md)", classes="coming-soon")
            with TabPane("💬 Teams", id="tab-teams"):
                yield Static("Teams export — coming soon (see ROADMAP.md)", classes="coming-soon")
        yield Footer()

    # ── Lifecycle ─────────────────────────────────────────────────────────

    def on_mount(self) -> None:
        global _export_scan_lock
        apply_config(self._cfg)

        exchange_dir = Path(self._cfg["defaults"]["output_dir"]) / "exchange"
        exchange_dir.mkdir(parents=True, exist_ok=True)

        self._exchange_state = StateManager(exchange_dir)
        self._bfy_logger         = setup_logging(exchange_dir)
        self._debug_log      = setup_playwright_debug_logging()
        self._semaphore      = asyncio.Semaphore(int(self._cfg["defaults"]["concurrency"]))
        _export_scan_lock    = asyncio.Lock()

        table = self.query_one("#progress-table", DataTable)
        table.add_column("Name",     key="name",   width=28)
        table.add_column("Status",   key="status", width=15)
        table.add_column("Progress", key="detail", width=40)

        # Spinner over the user list until UsersLoaded arrives
        self.query_one("#exchange-user-list", SelectionList).loading = True

        self.init_session()

    # ── Background workers ────────────────────────────────────────────────

    @work(exclusive=True, group="session")
    async def init_session(self) -> None:
        self.post_message(LogEntry("info", "Starting session — checking saved cookies…"))

        def _status(text: str) -> None:
            self.post_message(StatusUpdate(f"⟳ {text}"))
            self.post_message(LogEntry("info", text))

        try:
            context = await login(self._bfy_logger, self._debug_log, on_status=_status)
            self._pw_context = context
            self._cookies = await extract_cookies(context)
            self.post_message(SessionReady())
            self.post_message(StatusUpdate("⟳ Session ready — loading user list…"))
            self.post_message(LogEntry("info", "Session ready — loading users…"))
            self.load_users()
        except Exception as exc:
            self.post_message(StatusUpdate("✗ Session failed — see log pane"))
            self.post_message(LogEntry("error", f"Session error: {exc}"))
            self.query_one("#exchange-user-list", SelectionList).loading = False

    @work(exclusive=True, group="users")
    async def load_users(self) -> None:
        try:
            users = await get_all_users(self._pw_context, self._cookies, self._bfy_logger)
            self.post_message(UsersLoaded(users))
        except Exception as exc:
            self.post_message(LogEntry("error", f"User load failed: {exc}"))

    @work(group="preflight")
    async def run_preflight(self, users: list[dict], output_dir: Path) -> None:
        self.post_message(LogEntry("info", "Scanning existing export jobs…"))
        try:
            await refresh_export_cache(self._pw_context, self._bfy_logger)
            self.post_message(LogEntry("info", f"Found {len(_export_cache)} existing job(s)."))
        except Exception as exc:
            self.post_message(LogEntry("warning", f"Export scan failed: {exc}"))

        for user in users:
            sid = user["service_id"]
            if self._exchange_state.is_done(sid):
                self.post_message(ExportUpdate(sid, "skipped", "already complete"))
            else:
                self.export_one_user(user, output_dir)

    @work(group="exports")
    async def export_one_user(self, user: dict, output_dir: Path) -> None:
        sid  = user["service_id"]
        name = user.get("name") or sid

        # Already on disk (e.g. downloaded by an earlier version) — don't
        # trigger a fresh server-side export for it.
        existing_file = pst_path(output_dir, name, sid)
        if existing_file.exists():
            self._exchange_state.mark_complete(sid, existing_file.name)
            self.post_message(ExportUpdate(sid, "skipped", "already on disk"))
            return

        async with self._semaphore:
            try:
                snapshot_id = user.get("snapshot_id") or await get_latest_snapshot_id(
                    self._pw_context, sid, self._bfy_logger
                )
                if not snapshot_id:
                    self._exchange_state.mark_failed(sid, "No snapshot ID")
                    self.post_message(ExportUpdate(sid, "failed", "no snapshot"))
                    return

                existing = self._exchange_state.get_in_progress(sid)
                resumed  = bool(existing)
                if existing:
                    job_id = existing["job_id"]
                    self.post_message(ExportUpdate(sid, "resuming", f"job {job_id}"))
                else:
                    self.post_message(ExportUpdate(sid, "triggering"))
                    job_id = await trigger_export(
                        self._pw_context, self._cookies, sid, snapshot_id, self._bfy_logger
                    )
                    if not job_id:
                        self._exchange_state.mark_failed(sid, "Trigger failed")
                        self.post_message(ExportUpdate(sid, "failed", "trigger failed"))
                        return
                    self._exchange_state.mark_in_progress(sid, job_id, snapshot_id)

                def _poll_status(text: str) -> None:
                    self.post_message(ExportUpdate(sid, "polling", text))

                self.post_message(ExportUpdate(sid, "polling", f"job {job_id}"))
                # A resumed job that's absent is stale (4 scans); a freshly
                # triggered one definitely exists, so wait longer (10 scans)
                # before giving up on it.
                outcome, download_url = await poll_for_download_url(
                    self._pw_context, job_id, sid, self._bfy_logger,
                    on_status=_poll_status,
                    max_unseen=4 if resumed else 10,
                )

                if outcome == "missing" and resumed:
                    # Saved job id from an earlier session no longer exists on
                    # the export page — start a fresh export instead.
                    self.post_message(ExportUpdate(sid, "triggering", "stale job — re-exporting"))
                    self.post_message(LogEntry(
                        "warning", f"{name}: saved job {job_id} is gone — re-triggering export."
                    ))
                    self._exchange_state.clear_in_progress(sid)
                    job_id = await trigger_export(
                        self._pw_context, self._cookies, sid, snapshot_id, self._bfy_logger
                    )
                    if not job_id:
                        self._exchange_state.mark_failed(sid, "Re-trigger failed")
                        self.post_message(ExportUpdate(sid, "failed", "re-trigger failed"))
                        return
                    self._exchange_state.mark_in_progress(sid, job_id, snapshot_id)
                    self.post_message(ExportUpdate(sid, "polling", f"job {job_id}"))
                    outcome, download_url = await poll_for_download_url(
                        self._pw_context, job_id, sid, self._bfy_logger,
                        on_status=_poll_status,
                        max_unseen=10,
                    )

                if not download_url:
                    self._exchange_state.mark_failed(sid, f"Export {outcome}")
                    self.post_message(ExportUpdate(sid, "failed", f"export {outcome}"))
                    return

                self.post_message(ExportUpdate(sid, "downloading"))

                def _on_progress(done: int, total: int) -> None:
                    self.post_message(DownloadProgress(sid, done, total))

                filename = await download_file(
                    download_url, self._cookies, output_dir,
                    name, sid, self._bfy_logger, on_progress=_on_progress,
                )
                if filename:
                    self._exchange_state.mark_complete(sid, filename)
                    self.post_message(ExportUpdate(sid, "complete", filename[:38]))
                    self.post_message(LogEntry("success", f"✓ {name}  →  {filename}"))
                else:
                    self._exchange_state.mark_failed(sid, "Download failed")
                    self.post_message(ExportUpdate(sid, "failed", "download error"))
                    self.post_message(LogEntry("error", f"✗ {name}  download failed"))

            except Exception as exc:
                self._exchange_state.mark_failed(sid, str(exc))
                self.post_message(ExportUpdate(sid, "failed", str(exc)[:38]))
                self.post_message(LogEntry("error", f"✗ {name}: {exc}"))

    # ── Worker→UI message handlers ────────────────────────────────────────

    @on(SessionReady)
    def _handle_session_ready(self) -> None:
        self.query_one("#status-text", Static).update("⟳ Session active — loading users…")

    @on(StatusUpdate)
    def _handle_status_update(self, event: StatusUpdate) -> None:
        self.query_one("#status-text", Static).update(event.text)

    @on(UsersLoaded)
    def _handle_users_loaded(self, event: UsersLoaded) -> None:
        self._exchange_users    = event.users
        self._exchange_selected = {u["service_id"] for u in event.users}
        self.query_one("#exchange-user-list", SelectionList).loading = False
        self._refresh_user_list()
        self.query_one("#status-text", Static).update(
            f"✓ Ready — {len(event.users)} mailboxes loaded. Select users and press Start."
        )
        if not self._export_running:
            self.query_one("#btn-start", Button).disabled = False
        self.notify(f"Ready — {len(event.users)} mailboxes loaded.", title="Backupify")

    @on(LogEntry)
    def _handle_log_entry(self, event: LogEntry) -> None:
        colour = {
            "info": "white", "warning": "yellow",
            "error": "red", "success": "green",
        }.get(event.level, "white")
        ts = datetime.now().strftime("%H:%M:%S")
        self.query_one("#log-pane", RichLog).write(
            f"[{colour}]{ts}  {rich_escape(event.text)}[/{colour}]"
        )

    @on(ExportUpdate)
    def _handle_export_update(self, event: ExportUpdate) -> None:
        icon  = _STATUS_ICONS.get(event.status, "?")
        table = self.query_one("#progress-table", DataTable)
        try:
            table.update_cell(event.service_id, "status", f"{icon} {event.status.capitalize()}")
            if event.detail:
                table.update_cell(event.service_id, "detail", event.detail)
        except Exception:
            pass  # row may not exist (e.g. update for a user not in this run)
        if event.status in ("complete", "failed", "skipped"):
            self._update_stats()
            if self._export_running and event.service_id in self._run_pending:
                self._run_pending.discard(event.service_id)
                if not self._run_pending:
                    self._finish_run()

    @on(DownloadProgress)
    def _handle_download_progress(self, event: DownloadProgress) -> None:
        if event.total:
            pct      = min(100, int(event.done / event.total * 100))
            done_mb  = event.done  / 1_048_576
            total_mb = event.total / 1_048_576
            filled   = pct // 5
            bar      = "█" * filled + "░" * (20 - filled)
            detail   = f"{bar} {pct:3d}%  {done_mb:.0f}/{total_mb:.0f} MB"
        else:
            detail = f"{event.done / 1_048_576:.0f} MB"
        try:
            self.query_one("#progress-table", DataTable).update_cell(
                event.service_id, "detail", detail
            )
        except Exception:
            pass

    # ── UI event handlers ─────────────────────────────────────────────────

    @on(Input.Changed, "#search-input")
    def _on_search(self, event: Input.Changed) -> None:
        self._exchange_search = event.value
        self._refresh_user_list()

    @on(Select.Changed, "#sort-select")
    def _on_sort(self, event: Select.Changed) -> None:
        if event.value is not Select.BLANK:
            self._exchange_sort = str(event.value)
            self._refresh_user_list()

    @on(Input.Changed, "#output-dir-input")
    def _on_output_dir_changed(self, event: Input.Changed) -> None:
        if event.value.strip():
            self._cfg["defaults"]["output_dir"] = event.value.strip()

    @on(Input.Changed, "#concurrency-input")
    def _on_concurrency_changed(self, event: Input.Changed) -> None:
        try:
            n = int(event.value)
        except ValueError:
            return
        if 1 <= n <= 32:
            self._cfg["defaults"]["concurrency"] = n
            if not self._export_running:
                self._semaphore = asyncio.Semaphore(n)

    @on(SelectionList.SelectedChanged, "#exchange-user-list")
    def _on_user_selection(self, event: SelectionList.SelectedChanged) -> None:
        visible_ids    = {u["service_id"] for u in self._exchange_filtered}
        newly_selected = set(event.selection_list.selected)
        self._exchange_selected = (self._exchange_selected - visible_ids) | newly_selected
        self._update_stats()

    @on(Button.Pressed, "#btn-all")
    def _on_select_all(self) -> None:
        self._exchange_selected |= {u["service_id"] for u in self._exchange_filtered}
        self._refresh_user_list()

    @on(Button.Pressed, "#btn-none")
    def _on_select_none(self) -> None:
        self._exchange_selected -= {u["service_id"] for u in self._exchange_filtered}
        self._refresh_user_list()

    @on(Button.Pressed, "#btn-start")
    def _on_btn_start(self) -> None:
        self.action_start_export()

    @on(Button.Pressed, "#btn-retry")
    def _on_btn_retry(self) -> None:
        self.action_retry_failed()

    # ── Actions ───────────────────────────────────────────────────────────

    def action_start_export(self) -> None:
        if not self._pw_context:
            self.notify("Session not ready yet — please wait.", severity="warning")
            return
        if self._export_running:
            self.notify("Export already running.", severity="warning")
            return

        selected = [
            u for u in self._exchange_users
            if u["service_id"] in self._exchange_selected
        ]
        if not selected:
            self.notify("No users selected.", severity="warning")
            return

        selected   = _sort_users(selected, self._exchange_sort)
        output_dir = Path(self._cfg["defaults"]["output_dir"]) / "exchange"
        output_dir.mkdir(parents=True, exist_ok=True)
        save_config(self._cfg)

        self._export_running = True
        self._run_pending    = {u["service_id"] for u in selected}
        self.query_one("#btn-start", Button).disabled = True

        table = self.query_one("#progress-table", DataTable)
        table.clear()
        for user in selected:
            table.add_row(
                user["name"][:28], "○ Queued", "",
                key=user["service_id"],
            )

        self.post_message(StatusUpdate(f"⟳ Exporting {len(selected)} mailbox(es)…"))
        self.post_message(LogEntry("info", f"Starting export for {len(selected)} user(s)…"))
        self.run_preflight(selected, output_dir)

    def action_refresh_users(self) -> None:
        if not self._pw_context:
            self.notify("Session not ready yet.", severity="warning")
            return
        if USERS_FILE.exists():
            USERS_FILE.unlink()
        self.query_one("#exchange-user-list", SelectionList).loading = True
        self.post_message(StatusUpdate("⟳ Re-fetching user list from Backupify…"))
        self.post_message(LogEntry("info", "User cache cleared — re-fetching…"))
        self.load_users()

    def action_retry_failed(self) -> None:
        if not self._exchange_state:
            return
        n = len(self._exchange_state.state["failed"])
        if n == 0:
            self.notify("No failed exports to retry.")
            return
        self._exchange_state.state["failed"].clear()
        self._exchange_state.save()
        self._export_running = False
        self.query_one("#btn-start", Button).disabled = False
        self.notify(f"Cleared {n} failed entr{'y' if n == 1 else 'ies'} — press Start to retry.")
        self._refresh_user_list()

    def _finish_run(self) -> None:
        self._export_running = False
        self.query_one("#btn-start", Button).disabled = False
        done   = len(self._exchange_state.state["completed"])
        failed = len(self._exchange_state.state["failed"])
        self.query_one("#status-text", Static).update(
            f"✓ Run finished — {done} complete, {failed} failed."
        )
        self.post_message(LogEntry("info", f"Run finished — {done} complete, {failed} failed."))
        self.notify(f"Run finished — {done} complete, {failed} failed.", title="Backupify")
        self._refresh_user_list()

    def action_quit(self) -> None:
        save_config(self._cfg)
        self._close_browser_and_exit()

    @work(group="shutdown")
    async def _close_browser_and_exit(self) -> None:
        # Close the headless browser so no orphan Chromium lingers
        try:
            if self._pw_context and self._pw_context.browser:
                await self._pw_context.browser.close()
        except Exception:
            pass
        self.exit()

    # ── Helpers ───────────────────────────────────────────────────────────

    def _refresh_user_list(self) -> None:
        search = self._exchange_search.lower()
        filtered = [
            u for u in self._exchange_users
            if not search
            or search in u["name"].lower()
            or search in (u.get("email") or "").lower()
        ]
        self._exchange_filtered = _sort_users(filtered, self._exchange_sort)

        sl    = self.query_one("#exchange-user-list", SelectionList)
        state = self._exchange_state
        sl.clear_options()
        for user in self._exchange_filtered:
            sid = user["service_id"]
            if state and state.is_done(sid):
                label = f"✓ {user['name']}"
            elif state and state.get_in_progress(sid):
                label = f"⟳ {user['name']}"
            elif state and state.is_failed(sid):
                label = f"✗ {user['name']}"
            else:
                label = f"  {user['name']}"
            sl.add_option(Selection(label, sid, sid in self._exchange_selected))

        self._update_stats()

    def _update_stats(self) -> None:
        state = self._exchange_state
        total = len(self._exchange_users)
        if state:
            done    = len(state.state["completed"])
            in_prog = len(state.state["in_progress"])
            failed  = len(state.state["failed"])
            queued  = max(0, total - done - in_prog - failed)
        else:
            done = in_prog = failed = queued = 0
        try:
            self.query_one("#user-stats", Static).update(
                f"✓ {done}  ⟳ {in_prog}  ✗ {failed}  ○ {queued}"
                f"  │  {len(self._exchange_selected)}/{total} selected"
            )
        except Exception:
            pass


# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    _cfg = load_config()

    if not config_is_complete(_cfg):
        result = OnboardingApp(_cfg).run()
        if not (result and config_is_complete(result)):
            raise SystemExit("Setup not completed — exiting.")
        _cfg = result

    BackupifyApp(_cfg).run()
