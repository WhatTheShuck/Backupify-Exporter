# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "playwright>=1.44.0",
#   "httpx>=0.27.0",
#   "tqdm>=4.66.0",
# ]
# ///
"""
Backupify M365 PST Bulk Exporter
=================================
Automates exporting all Exchange mailbox PST files from Backupify.

Architecture:
  - Playwright handles everything that touches the SPA (login, user discovery,
    snapshot extraction, triggering exports, polling for completion)
  - httpx handles streaming large file downloads (unsuitable for Playwright)

Usage:
    uv run backupify_export.py --output /path/to/output/drive
    uv run backupify_export.py --output /path/to/output/drive --resume
    uv run backupify_export.py --output /path/to/output/drive --dry-run
    uv run backupify_export.py --output /path/to/output/drive --concurrency 6

First-time setup (once only — installs Chromium):
    uv run python -m playwright install chromium
"""

import argparse
import asyncio
import json
import logging
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import httpx
from tqdm import tqdm
from playwright.async_api import async_playwright, Page, BrowserContext, TimeoutError as PWTimeout

# ─── Configuration ────────────────────────────────────────────────────────────

BASE_URL        = "https://aue1-bfyii-2621-ext.backupify.com"
CUSTOMER_ID     = "292424"
EXT_CUSTOMER_ID = "24fd7c0d-2a29-11eb-ae0f-0cc47a57db64"
LOGIN_URL       = "https://auth.datto.com/login?login_hint=brandon.wiedman@ksb.com.au"
DASHBOARD_URL   = f"{BASE_URL}/{CUSTOMER_ID}/o365?external_customer_id={EXT_CUSTOMER_ID}"
EXCHANGE_URL    = f"{BASE_URL}/{CUSTOMER_ID}/o365/exchange"
EXPORT_ACTION   = f"{BASE_URL}/{CUSTOMER_ID}/restoreExportAction"

# Saved session file — stored in the user's home dir so it persists across runs
COOKIE_FILE     = Path.home() / ".backupify_session.json"
USERS_FILE      = Path.home() / f".backupify_users_{CUSTOMER_ID}.json"

# Seconds between polls when waiting for an export job to finish
POLL_INTERVAL    = 30
# Max seconds to wait for a single export job (3 hours)
EXPORT_TIMEOUT   = 10_800
# Max seconds allowed for a single file download (2 hours)
DOWNLOAD_TIMEOUT = 7_200
# Download chunk size
CHUNK_SIZE       = 16 * 1024 * 1024  # 16 MB

# Playwright timeouts (milliseconds)
PAGE_LOAD_MS     = 60_000   # 60 s for a page to load
ELEMENT_MS       = 30_000   # 30 s to find an element
LOGIN_MS         = 300_000  # 5 min for human to complete login + TOTP


# ─── Logging ──────────────────────────────────────────────────────────────────

def setup_logging(output_dir: Path) -> logging.Logger:
    log_path = output_dir / f"backupify_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    fmt = "%(asctime)s [%(levelname)s] %(message)s"
    logging.basicConfig(
        level=logging.INFO,
        format=fmt,
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return logging.getLogger("backupify")


def setup_playwright_debug_logging() -> logging.Logger:
    """
    Creates a dedicated logger that writes all Playwright page events
    (console, requests, responses, navigation, JS errors) to playwright_debug.log
    in the current working directory.  Does NOT propagate to the root logger.
    """
    debug_log = logging.getLogger("pw_debug")
    debug_log.setLevel(logging.DEBUG)
    debug_log.propagate = False

    handler = logging.FileHandler("playwright_debug.log", encoding="utf-8", mode="w")
    handler.setFormatter(logging.Formatter("%(asctime)s %(message)s", datefmt="%H:%M:%S"))
    debug_log.addHandler(handler)
    return debug_log


def attach_page_listeners(page, debug_log: logging.Logger) -> None:
    """Attaches debug event listeners to a single Playwright Page."""

    page.on("console", lambda msg: debug_log.debug(
        f"[CONSOLE:{msg.type.upper():7}] {msg.text}"
    ))
    page.on("pageerror", lambda exc: debug_log.debug(
        f"[JSERROR ] {exc}"
    ))
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
        self.path  = output_dir / "progress.json"
        self.state = self._load()

    def _load(self) -> dict:
        if self.path.exists():
            try:
                with open(self.path) as f:
                    data = json.load(f)
                    # Back-fill missing bucket for files written by older versions
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

    def summary(self) -> dict:
        return {
            "completed":   len(self.state["completed"]),
            "in_progress": len(self.state["in_progress"]),
            "failed":      len(self.state["failed"]),
        }


# ─── Authentication ───────────────────────────────────────────────────────────

async def _save_cookies(context: BrowserContext) -> None:
    cookies = await context.cookies()
    COOKIE_FILE.write_text(json.dumps(cookies, indent=2), encoding="utf-8")


async def _try_restore_session(
    context: BrowserContext,
    logger: logging.Logger,
) -> bool:
    """
    Loads cookies from COOKIE_FILE into the browser context, then navigates to
    the dashboard to verify the session is still valid.  Returns True if we're
    logged in, False if the cookies are missing or expired.
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
        # If we were redirected back to the auth domain the session has expired
        if BASE_URL not in page.url:
            logger.info(f"Session expired (redirected to {page.url[:60]}...).")
            return False

        # Confirm the dashboard is actually rendered
        dashboard_selectors = ["nav", "[class*='nav']", "[class*='sidebar']",
                               "a[href*='/o365']", ".app-container", "#app"]
        for sel in dashboard_selectors:
            try:
                await page.wait_for_selector(sel, timeout=5_000)
                logger.info("Saved session is valid — skipping interactive login.")
                return True
            except PWTimeout:
                continue

        logger.info("Dashboard element not found after session restore — treating as expired.")
        return False
    except PWTimeout:
        logger.warning("Timed out while validating saved session.")
        return False
    finally:
        await page.close()


def _make_context_listener(debug_log):
    return lambda p: attach_page_listeners(p, debug_log)


async def login(logger: logging.Logger, debug_log: logging.Logger) -> BrowserContext:
    """
    Returns an authenticated BrowserContext.

    If a valid saved session exists the browser runs headless — no window at
    all.  Only when a fresh login is needed does a visible window open; once
    the user completes login the cookies are saved so the next run is headless.
    """
    pw = await async_playwright().start()
    _quiet_args = ["--no-first-run", "--no-default-browser-check", "--disable-extensions"]

    # ── Try headless restore first ──
    if COOKIE_FILE.exists():
        browser = await pw.chromium.launch(headless=True, args=_quiet_args)
        context = await browser.new_context()
        context.on("page", _make_context_listener(debug_log))
        if await _try_restore_session(context, logger):
            return context
        logger.info("Saved session invalid — falling back to interactive login.")
        await browser.close()

    # ── Interactive login — visible browser required ──
    logger.info("Opening browser — please enter your password and TOTP code when prompted.")
    browser = await pw.chromium.launch(headless=False, slow_mo=50, args=_quiet_args)
    context = await browser.new_context()
    context.on("page", _make_context_listener(debug_log))

    page = await context.new_page()

    page.on("framenavigated", lambda frame: (
        logger.info(f"  → navigated to: {frame.url}") if frame == page.main_frame else None
    ))

    await page.goto(LOGIN_URL, wait_until="domcontentloaded")
    logger.info("Waiting for you to complete password + TOTP (5 minute limit)...")

    deadline = time.monotonic() + (LOGIN_MS / 1000)
    while time.monotonic() < deadline:
        await asyncio.sleep(1)
        current_url = page.url

        if BASE_URL not in current_url:
            continue

        logger.info(f"On backupify domain. Current URL: {current_url}")

        dashboard_selectors = [
            "nav", "[class*='nav']", "[class*='sidebar']", "[class*='dashboard']",
            "[class*='header']", "a[href*='/o365']", "a[href*='exchange']",
            ".app-container", "#app",
        ]
        for sel in dashboard_selectors:
            try:
                await page.wait_for_selector(sel, timeout=5_000)
                logger.info(f"Dashboard element found ({sel}). Login complete.")
                await _save_cookies(context)
                logger.info(f"Session cookies saved to {COOKIE_FILE}")
                return context
            except PWTimeout:
                continue

        logger.info("On backupify domain but dashboard not rendered yet, waiting...")
        await asyncio.sleep(2)

    await browser.close()
    raise RuntimeError(
        "Login timed out after 5 minutes.\n"
        f"Last URL was: {page.url}\n"
        "If you completed login successfully, the dashboard selector may need updating."
    )


# ─── User discovery ───────────────────────────────────────────────────────────

async def get_all_users(context: BrowserContext, cookies: dict, logger: logging.Logger) -> list[dict]:
    """
    Returns the full list of Exchange users.  On the first run, discovers them
    by intercepting the /customerServices XHR and paginates via httpx, then
    caches the result to USERS_FILE.  Subsequent runs load from cache
    (delete the file to force a re-fetch).

    Returns [{"name", "email", "service_id", "snapshot_id"}, ...].
    """
    if USERS_FILE.exists():
        try:
            users = json.loads(USERS_FILE.read_text(encoding="utf-8"))
            if users:
                logger.info(f"Loaded {len(users)} users from cache ({USERS_FILE}).")
                return users
        except Exception as e:
            logger.warning(f"Could not read user cache: {e} — re-fetching.")
    logger.info("Loading Exchange recovery page to discover users...")
    page = await context.new_page()

    # Capture the exact POST body/headers the browser sends to customerServices
    # so we can replay the request via httpx with a larger 'length'.
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

    # Intercept the first response (gives us recordsTotal and one page of data)
    first_data: dict = {}
    try:
        async with page.expect_response(
            lambda r: "/customerServices" in r.url and r.status == 200,
            timeout=45_000,
        ) as resp_info:
            await page.goto(EXCHANGE_URL, wait_until="domcontentloaded")

        first_data = await (await resp_info.value).json()
        records_total = first_data.get("recordsTotal", 0)
        first_batch   = first_data.get("data", [])
        logger.info(
            f"customerServices page-1: {len(first_batch)} rows returned, "
            f"{records_total} total."
        )

    except PWTimeout:
        logger.error("customerServices XHR not seen within 45 s.")
        await page.close()
        raise RuntimeError("customerServices XHR timed out — cannot discover users.")

    await page.close()

    # ── Fetch all records via httpx if the first page is incomplete ──────────
    all_rows: list[dict] = first_data.get("data", [])

    records_total = first_data.get("recordsTotal", len(all_rows))
    if records_total > len(all_rows) and captured_req:
        logger.info(f"Fetching all {records_total} records via httpx...")

        # Build headers from the captured request but override the Cookie header
        # with the session cookies extracted from the Playwright context — the
        # browser's internal cookie header isn't always forwarded correctly.
        headers = {
            k: v for k, v in captured_req["headers"].items()
            if k.lower() != "cookie"
        }
        headers["cookie"] = "; ".join(f"{k}={v}" for k, v in cookies.items())

        # Page through records with the server's own page size (changing only
        # 'start' and 'draw'), because requesting length=recordsTotal returns 500.
        page_size = len(first_data.get("data", [])) or 20
        draw_num  = 2
        offset    = len(all_rows)

        async with httpx.AsyncClient(follow_redirects=False, timeout=60.0) as client:
            while offset < records_total:
                paged_body = captured_req["post_data"]
                paged_body = re.sub(r'\bstart=\d+', f'start={offset}',   paged_body)
                paged_body = re.sub(r'\bdraw=\d+',  f'draw={draw_num}',  paged_body)
                draw_num  += 1

                resp = await client.post(
                    captured_req["url"],
                    content=paged_body.encode(),
                    headers=headers,
                )

                if resp.status_code != 200:
                    logger.warning(
                        f"Pagination page at start={offset} returned HTTP {resp.status_code}; "
                        "stopping early."
                    )
                    break

                batch = resp.json().get("data", [])
                if not batch:
                    break
                all_rows.extend(batch)
                offset += len(batch)
                logger.info(f"  Fetched {offset}/{records_total} users...")
    elif records_total > len(all_rows):
        logger.warning(
            "Could not capture customerServices request body — "
            f"only the first {len(all_rows)} of {records_total} users will be processed."
        )

    # ── Parse rows into user dicts ───────────────────────────────────────────
    if all_rows:
        logger.info(f"First row keys: {list(all_rows[0].keys())}")

    users: list[dict] = []
    for row in all_rows:
        service_id = str(row.get("id") or "")
        if not service_id:
            continue

        name  = re.sub(r"<[^>]+>", "", str(row.get("name")  or "")).strip()
        email = re.sub(r"<[^>]+>", "", str(row.get("email") or "")).strip()

        # Pull the latest snapshot ID from the perfectBackups list.
        # Each entry is {"snapshotId": <epoch_ms>}; we want the highest value.
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
        raise RuntimeError(
            "No users could be extracted from the customerServices response. "
            "Check the log for the 'First row keys' entry and update field-name mapping."
        )

    with_snap = sum(1 for u in users if u["snapshot_id"])
    logger.info(f"Found {len(users)} Exchange users ({with_snap} with snapshot IDs).")
    USERS_FILE.write_text(json.dumps(users, indent=2), encoding="utf-8")
    logger.info(f"User list cached to {USERS_FILE}.")
    return users


# ─── Snapshot extraction ──────────────────────────────────────────────────────

async def get_latest_snapshot_id(
    context: BrowserContext,
    service_id: str,
    logger: logging.Logger,
) -> str | None:
    """
    Navigates to the user's service page and reads the latest snapshot ID
    from the rendered DOM.
    """
    service_url = f"{BASE_URL}/{CUSTOMER_ID}/o365/exchange/service?serviceId={service_id}"
    page = await context.new_page()

    try:
        await page.goto(service_url, wait_until="domcontentloaded")
        await page.wait_for_load_state("networkidle", timeout=PAGE_LOAD_MS)

        # Strategy 1: intercept the XHR that returns snapshot data
        # We do this by reading window.__INITIAL_STATE__ or similar embedded JSON
        snapshot_id = await page.evaluate("""() => {
            // Try common SPA state patterns
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
            // Try data attributes on the page
            const el = document.querySelector('[data-snapshot-id]');
            if (el) return el.getAttribute('data-snapshot-id');
            // Try snapshot selector element
            const snap = document.querySelector('.snapshot-selector, #snapshot-selector, [class*="snapshot"]');
            if (snap) {
                const attr = snap.getAttribute('data-value') || snap.getAttribute('data-id') || snap.value;
                if (attr && /^\\d+$/.test(attr)) return attr;
            }
            return null;
        }""")

        if snapshot_id:
            return str(snapshot_id)

        # Strategy 2: look for snapshot ID in the rendered HTML source
        html = await page.content()

        patterns = [
            r'"snapshotId"\s*:\s*"?(\d{10,})"?',
            r'data-snapshot-id=["\'](\d+)["\']',
            r'snapshot_id["\']?\s*[=:]\s*["\']?(\d{10,})',
            r'value=["\'](\d{13})["\']',  # 13-digit epoch ms in a select/input
        ]
        for pat in patterns:
            m = re.search(pat, html)
            if m:
                return m.group(1)

        logger.warning(f"[{service_id}] Could not find snapshot ID — saving debug HTML.")
        Path(f"debug_service_{service_id}.html").write_text(html, encoding="utf-8")
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
    """
    POSTs to /restoreExportAction using the session cookies.
    We use httpx here since it's a simple JSON POST — no SPA rendering needed.
    Returns the job ID string, or None on failure.
    """
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

    # Build cookie header from Playwright context cookies
    cookie_header = "; ".join(f"{k}={v}" for k, v in cookies.items())

    async with httpx.AsyncClient(follow_redirects=True, timeout=60.0) as client:
        resp = await client.post(
            EXPORT_ACTION,
            data=payload,
            headers={
                "Cookie":       cookie_header,
                "Referer":      f"{BASE_URL}/{CUSTOMER_ID}/o365/exchange/service?serviceId={service_id}",
                "User-Agent":   "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36",
                "X-Requested-With": "XMLHttpRequest",
                "Accept":       "application/json, text/javascript, */*",
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

# Shared cache of export job statuses, refreshed by scanning the export page.
# Maps job_id (str) -> {"job_id", "source_name", "status", "download_url"}.
_export_cache: dict = {}
_export_cache_time: float = 0.0
_export_scan_lock: asyncio.Lock | None = None  # initialised in main()

# Captured getActivities XHR endpoint — populated on the first export-page load.
# Once set, all subsequent polls bypass Playwright and call it directly via httpx.
_activities_endpoint: dict = {}  # {"url": str, "headers": dict}


def _parse_activity_item(item: dict) -> dict:
    """
    Convert a single getActivities JSON row into the standard record dict.
    Uses multiple field-name fallbacks to handle Backupify API variations.
    """
    job_id = str(
        item.get("id") or item.get("jobId") or item.get("job_id") or ""
    ).strip() or None

    source_name = re.sub(r"<[^>]+>", "", str(
        item.get("sourceName") or item.get("source_name") or
        item.get("name") or item.get("source") or ""
    )).strip()

    details = item.get("details") or {}
    if isinstance(details, str):
        try:
            details = json.loads(details)
        except Exception:
            details = {}

    status = str(
        item.get("status") or details.get("status") or
        item.get("state") or details.get("state") or ""
    ).lower()

    download_url = str(
        item.get("downloadUrl") or item.get("download_url") or
        details.get("downloadUrl") or details.get("download_url") or
        item.get("exportUrl") or ""
    ).strip()
    if download_url and not download_url.startswith("http"):
        download_url = BASE_URL + download_url

    return {
        "job_id":       job_id,
        "source_name":  source_name,
        "status":       status,
        "download_url": download_url,
    }


async def _fetch_activities_httpx(cookies: dict, logger: logging.Logger) -> list[dict]:
    """
    Fetches all export activity rows directly from the getActivities endpoint
    via httpx, paginating until all records are retrieved.
    """
    url = _activities_endpoint["url"]
    headers = dict(_activities_endpoint["headers"])
    headers["cookie"] = "; ".join(f"{k}={v}" for k, v in cookies.items())

    records: list[dict] = []
    start = 0
    page_size = 100  # larger than the browser uses to reduce round trips
    draw = 1

    async with httpx.AsyncClient(follow_redirects=False, timeout=60.0) as client:
        while True:
            body = (
                f"start={start}&length={page_size}&draw={draw}"
                f"&shouldIncludeDetails=true&appType=office365_exchange&activityType=export"
            )
            try:
                resp = await client.post(url, content=body.encode(), headers=headers)
            except httpx.RequestError as e:
                logger.warning(f"getActivities request failed: {e}")
                break

            if resp.status_code != 200:
                logger.warning(f"getActivities returned HTTP {resp.status_code}")
                break

            try:
                data = resp.json()
            except Exception as e:
                logger.warning(f"getActivities returned non-JSON: {e}")
                break

            batch = data.get("data", [])
            if not batch:
                break

            if start == 0 and batch:
                logger.debug(f"getActivities first-item keys: {list(batch[0].keys())}")

            for item in batch:
                records.append(_parse_activity_item(item))

            start += len(batch)
            if start >= data.get("recordsTotal", start):
                break
            draw += 1

    return records


async def scan_export_page(
    context: BrowserContext,
    cookies: dict,
    logger: logging.Logger,
) -> list[dict]:
    """
    Returns export job records: [{"job_id", "source_name", "status", "download_url"}, ...].

    On the first call, loads the export page via Playwright solely to intercept the
    getActivities XHR endpoint URL and headers.  Every subsequent call skips Playwright
    entirely and calls that endpoint directly via httpx, avoiding the DataTable's
    continuous auto-poll loop that was hammering the service.
    """
    global _activities_endpoint

    if _activities_endpoint:
        return await _fetch_activities_httpx(cookies, logger)

    # ── First call: capture the endpoint from the browser, then switch to httpx ──
    export_page_url = f"{BASE_URL}/{CUSTOMER_ID}/o365/exchange/export"
    page = await context.new_page()

    try:
        async with page.expect_response(
            lambda r: "/getActivities" in r.url and r.status == 200,
            timeout=ELEMENT_MS,
        ) as resp_info:
            await page.goto(export_page_url, wait_until="domcontentloaded", timeout=PAGE_LOAD_MS)

        activities_resp = await resp_info.value
        req = activities_resp.request
        _activities_endpoint = {
            "url": req.url,
            "headers": {
                k: v for k, v in req.headers.items()
                if k.lower() not in ("host", "content-length", "cookie")
            },
        }
        logger.info(f"Captured getActivities endpoint: {_activities_endpoint['url']}")

    except PWTimeout:
        logger.warning(
            "getActivities XHR not observed within timeout — "
            "export page may be empty or the endpoint URL has changed."
        )
    except Exception as e:
        logger.warning(f"Error capturing getActivities endpoint: {e}")
    finally:
        await page.close()

    if _activities_endpoint:
        return await _fetch_activities_httpx(cookies, logger)

    logger.warning("Could not capture getActivities endpoint — returning empty list.")
    return []


async def discover_and_scan_exports(
    context: BrowserContext,
    cookies: dict,
    state: StateManager,
    logger: logging.Logger,
) -> None:
    """
    Scans the export page once at startup to populate the shared export cache
    and report the current status of any jobs already tracked in StateManager.
    """
    global _export_cache, _export_cache_time

    logger.info("Scanning export page for existing jobs...")
    records = await scan_export_page(context, cookies, logger)

    for rec in records:
        if rec["job_id"]:
            _export_cache[rec["job_id"]] = rec
    _export_cache_time = time.monotonic()

    if not records:
        logger.info("No existing export jobs found on the export page.")
        return

    # Report status for jobs we are already tracking
    job_to_service = {
        info["job_id"]: sid
        for sid, info in state.state["in_progress"].items()
        if info.get("job_id")
    }
    for rec in records:
        jid = rec["job_id"]
        if jid and jid in job_to_service:
            sid = job_to_service[jid]
            logger.info(
                f"  Job {jid} ({rec['source_name']}): {rec['status']}"
                + (f"  download_url={rec['download_url'][:80]}" if rec["download_url"] else "")
            )

    logger.info(f"Export page scan complete: {len(records)} record(s) found.")


async def poll_for_download_url(
    context: BrowserContext,
    cookies: dict,
    job_id: str,
    service_id: str,
    logger: logging.Logger,
) -> str | None:
    """
    Waits until the export job reaches "completed" status, then returns the
    download URL.  Uses a shared Playwright-based cache so that many concurrent
    workers share a single export-page scan rather than each loading the page
    independently.
    """
    global _export_cache, _export_cache_time, _export_scan_lock

    deadline = time.monotonic() + EXPORT_TIMEOUT
    attempt  = 0

    while time.monotonic() < deadline:
        # ── Refresh the cache if it's older than POLL_INTERVAL ──
        cache_age = time.monotonic() - _export_cache_time
        if cache_age >= POLL_INTERVAL:
            async with _export_scan_lock:
                # Re-check after acquiring the lock — another worker may have
                # already refreshed while we were waiting.
                if time.monotonic() - _export_cache_time >= POLL_INTERVAL:
                    logger.info(
                        f"[{service_id}] Refreshing export status "
                        f"(cache age {cache_age:.0f}s)..."
                    )
                    records = await scan_export_page(context, cookies, logger)
                    for rec in records:
                        if rec["job_id"]:
                            _export_cache[rec["job_id"]] = rec
                    _export_cache_time = time.monotonic()

        # ── Check this job's status in the cache ──
        record = _export_cache.get(str(job_id))
        if record:
            status = record["status"]
            if status == "completed":
                url = record["download_url"]
                if url:
                    logger.info(f"[{service_id}] Export ready after {attempt} poll(s).")
                    return url
                # Completed but no download URL yet — treat as still pending
            elif status in ("failed", "error", "cancelled"):
                logger.error(
                    f"[{service_id}] Export job {job_id} ended with status: {status}"
                )
                return None

        attempt += 1
        if attempt % 4 == 0:
            elapsed = int(time.monotonic() - (deadline - EXPORT_TIMEOUT))
            logger.info(
                f"[{service_id}] Still waiting... {elapsed}s elapsed "
                f"(job {job_id}, last status: {record['status'] if record else 'not seen'})"
            )

        await asyncio.sleep(POLL_INTERVAL)

    logger.error(f"[{service_id}] Export timed out after {EXPORT_TIMEOUT}s.")
    return None


# ─── File download ────────────────────────────────────────────────────────────

async def download_file(
    download_url: str,
    cookies: dict,
    output_dir: Path,
    user_name: str,
    service_id: str,
    logger: logging.Logger,
) -> str | None:
    """
    Streams the export file to disk. Returns the filename on success, None on failure.
    Uses httpx for streaming — Playwright is not suitable for 100 GB files.
    """
    if download_url.startswith("/"):
        download_url = BASE_URL + download_url

    cookie_header = "; ".join(f"{k}={v}" for k, v in cookies.items())
    safe_name     = re.sub(r'[<>:"/\\|?*\s]', "_", user_name).strip("_")
    filename      = output_dir / f"{safe_name}__{service_id}.pst.zip"
    part_file     = filename.with_suffix(".zip.part")

    # If the final file already exists (e.g. state was lost but file is intact),
    # skip the download — the caller's state.is_done() check should have caught
    # this, but be safe.
    if filename.exists():
        size_mb = filename.stat().st_size / 1024 / 1024
        logger.info(f"[{service_id}] Final file already exists ({size_mb:.1f} MB), skipping download.")
        return filename.name

    # Clean up any stale partial file from a previous interrupted download
    if part_file.exists():
        logger.info(f"[{service_id}] Removing stale partial file {part_file.name}")
        part_file.unlink()

    logger.info(f"[{service_id}] Downloading to {filename.name}")

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

                total = int(resp.headers.get("content-length", 0))
                desc  = f"{safe_name[:35]:<35}"

                with tqdm(
                    total=total or None,
                    unit="B",
                    unit_scale=True,
                    unit_divisor=1024,
                    desc=desc,
                    leave=False,
                ) as bar:
                    with open(part_file, "wb") as f:
                        async for chunk in resp.aiter_bytes(CHUNK_SIZE):
                            f.write(chunk)
                            bar.update(len(chunk))

        # Atomically promote the temp file to its final name
        part_file.rename(filename)
        size_mb = filename.stat().st_size / 1024 / 1024
        logger.info(f"[{service_id}] Saved {filename.name} ({size_mb:.1f} MB)")
        return filename.name

    except (httpx.RequestError, OSError) as e:
        logger.error(f"[{service_id}] Download failed: {e}")
        if part_file.exists():
            part_file.unlink()
        return None


# ─── Per-user worker ──────────────────────────────────────────────────────────

async def process_user(
    user: dict,
    context: BrowserContext,
    cookies: dict,
    state: StateManager,
    output_dir: Path,
    semaphore: asyncio.Semaphore,
    logger: logging.Logger,
    bar: tqdm,
):
    service_id = user["service_id"]
    name       = user.get("name") or user.get("email") or service_id

    async with semaphore:
        try:
            # 1. Snapshot ID — already in user dict from customerServices; fall
            #    back to per-service page only if somehow missing.
            snapshot_id = user.get("snapshot_id") or await get_latest_snapshot_id(context, service_id, logger)
            if not snapshot_id:
                state.mark_failed(service_id, "Could not determine snapshot ID")
                return

            # 2. Trigger the export — or resume an already-triggered job.
            existing = state.get_in_progress(service_id)
            if existing:
                job_id = existing["job_id"]
                logger.info(f"[{service_id}] Resuming existing export job {job_id} (triggered at {existing['triggered_at']})")
            else:
                job_id = await trigger_export(context, cookies, service_id, snapshot_id, logger)
                if not job_id:
                    state.mark_failed(service_id, "Export trigger failed")
                    return
                state.mark_in_progress(service_id, job_id, snapshot_id)

            # 3. Poll until the export is ready
            download_url = await poll_for_download_url(context, cookies, job_id, service_id, logger)
            if not download_url:
                state.mark_failed(service_id, "Export did not complete in time")
                return

            # 4. Stream the file to disk
            filename = await download_file(
                download_url, cookies, output_dir, name, service_id, logger
            )
            if filename:
                state.mark_complete(service_id, filename)
            else:
                state.mark_failed(service_id, "Download failed")

        except Exception as e:
            logger.error(f"[{service_id}] Unhandled error: {e}", exc_info=True)
            state.mark_failed(service_id, str(e))
        finally:
            bar.update(1)


# ─── Cookie extraction helper ─────────────────────────────────────────────────

async def extract_cookies(context: BrowserContext) -> dict:
    """
    Pulls the session cookies from the Playwright context and returns
    them as a plain dict for use in httpx requests.
    """
    all_cookies = await context.cookies()
    wanted      = {"__backupify_session", "PHPSESSID"}
    cookies     = {c["name"]: c["value"] for c in all_cookies if c["name"] in wanted}

    missing = wanted - set(cookies.keys())
    if missing:
        # Log what we did find — helps diagnose if cookie names have changed
        found = [c["name"] for c in all_cookies]
        raise RuntimeError(
            f"Missing expected session cookies: {missing}\n"
            f"Cookies present: {found}\n"
            f"Update the 'wanted' set in extract_cookies() if the names have changed."
        )
    return cookies


# ─── Main ─────────────────────────────────────────────────────────────────────

async def main(output_dir: Path, concurrency: int, resume: bool, dry_run: bool):
    global _export_scan_lock
    output_dir.mkdir(parents=True, exist_ok=True)

    logger    = setup_logging(output_dir)
    debug_log = setup_playwright_debug_logging()
    state     = StateManager(output_dir)
    _export_scan_lock = asyncio.Lock()

    logger.info("=" * 60)
    logger.info("Backupify PST Bulk Exporter")
    logger.info(f"Output dir   : {output_dir}")
    logger.info(f"Concurrency  : {concurrency}")
    logger.info(f"Resume mode  : {resume}")
    logger.info(f"Playwright debug log: playwright_debug.log")
    logger.info("=" * 60)

    # ── 1. Login ──
    context = await login(logger, debug_log)

    try:
        # ── 2. Extract session cookies for httpx calls ──
        cookies = await extract_cookies(context)
        logger.info(f"Session cookies captured: {list(cookies.keys())}")

        # ── 3. Discover all Exchange users ──
        users = await get_all_users(context, cookies, logger)
        logger.info(f"Total users: {len(users)}")

        # ── 4. Filter users ──
        # Completed users are always skipped (re-running the script is safe).
        # In-progress users (export triggered but not yet downloaded) skip
        # straight to polling — no second trigger.
        # Failed users are skipped when --resume is set; retried otherwise.
        pending = [u for u in users if not state.is_done(u["service_id"])]
        if resume:
            pending = [u for u in pending if not state.is_failed(u["service_id"])]

        n_done      = sum(1 for u in users if state.is_done(u["service_id"]))
        n_resuming  = sum(1 for u in pending if state.get_in_progress(u["service_id"]))
        n_failed    = sum(1 for u in users if state.is_failed(u["service_id"]))
        logger.info(
            f"Status: {n_done} completed, {n_resuming} resuming from poll, "
            f"{n_failed} failed ({'skipped' if resume else 'will retry'}), "
            f"{len(pending)} to process."
        )

        # Clean up any leftover partial download files from a previous run
        for part_file in output_dir.glob("*.part"):
            logger.info(f"Cleaning up stale partial file: {part_file.name}")
            part_file.unlink()

        # ── 5. Discover export API + preflight scan ──
        # Navigate to the export page once to find the real JSON endpoint URL
        # and pre-load any jobs that were triggered in a previous session.
        await discover_and_scan_exports(context, cookies, state, logger)

        # Recalculate n_resuming after the scan may have added entries
        n_resuming = sum(1 for u in pending if state.get_in_progress(u["service_id"]))
        if n_resuming:
            logger.info(f"  {n_resuming} will resume from polling (export already triggered).")

        # ── 6. Dry run ──
        if dry_run:
            logger.info("DRY RUN — users that would be exported:")
            for u in pending:
                logger.info(f"  {u['name']:<45} {u['email']:<45} serviceId={u['service_id']}  snapshotId={u.get('snapshot_id') or 'NONE'}")
            return

        if not pending:
            logger.info("Nothing to do.")
            return

        logger.info(f"Starting export of {len(pending)} mailboxes ({concurrency} concurrent)...")

        # ── 6. Run workers ──
        semaphore = asyncio.Semaphore(concurrency)
        bar = tqdm(total=len(pending), desc="Overall progress", unit="mailbox", position=0)

        tasks = [
            process_user(u, context, cookies, state, output_dir, semaphore, logger, bar)
            for u in pending
        ]
        await asyncio.gather(*tasks)
        bar.close()

    finally:
        # Always close the browser, even on error
        await context.browser.close()

    # ── 7. Summary ──
    summary = state.summary()
    logger.info("")
    logger.info("=" * 60)
    logger.info("COMPLETE")
    logger.info(f"  Completed    : {summary['completed']}")
    logger.info(f"  In-progress  : {summary['in_progress']}  (re-run to resume polling)")
    logger.info(f"  Failed       : {summary['failed']}")
    logger.info("=" * 60)

    if state.state["failed"]:
        logger.info("Failed mailboxes:")
        for sid, info in state.state["failed"].items():
            logger.info(f"  {sid} — {info['reason']}")


# ─── CLI ──────────────────────────────────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser(
        description="Bulk export Backupify M365 Exchange mailboxes as PST files.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  uv run backupify_export.py --output /mnt/usb_drive
  uv run backupify_export.py --output /mnt/usb_drive --resume
  uv run backupify_export.py --output /mnt/usb_drive --dry-run
  uv run backupify_export.py --output /mnt/usb_drive --concurrency 5
        """,
    )
    p.add_argument("--output",      "-o", required=True,      help="Directory for PST files and logs.")
    p.add_argument("--concurrency", "-c", type=int, default=8, help="Parallel exports (default: 8).")
    p.add_argument("--resume",      "-r", action="store_true", help="Skip already-completed mailboxes.")
    p.add_argument("--dry-run",           action="store_true", help="List users without exporting.")
    return p.parse_args()


if __name__ == "__main__":
    args = parse_args()
    asyncio.run(main(
        output_dir=Path(args.output),
        concurrency=args.concurrency,
        resume=args.resume,
        dry_run=args.dry_run,
    ))
