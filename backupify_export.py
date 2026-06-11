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
Supports Exchange PST, OneDrive, and SharePoint; Teams planned (see ROADMAP.md).

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
from dataclasses import dataclass
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
from textual.screen import ModalScreen, Screen
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


# ─── Service registry ─────────────────────────────────────────────────────────
#
# Every Backupify service section shares the same machinery (customerServices
# entity list, restoreExportAction trigger, /o365/<section>/export status
# table) — only these knobs differ.  Discovered per-service via the playbook
# in HANDOFF.md; don't guess values for new services.

@dataclass(frozen=True)
class ServiceDef:
    key:           str          # internal id, also output subfolder + widget id prefix
    label:         str          # tab title
    noun:          str          # what one entity is called in the UI
    noun_plural:   str
    section:       str          # URL path segment under /o365/
    app_type:      str          # appType in customerServices + restoreExportAction
    export_format: str | None   # exportFormat field, or None if the modal has no format radios
    file_ext:      str          # on-disk extension when the server names nothing better


SERVICES: dict[str, ServiceDef] = {
    "exchange": ServiceDef(
        key="exchange", label="📧 Exchange", noun="mailbox", noun_plural="mailboxes",
        section="exchange", app_type="office365_exchange",
        export_format="pst", file_ext=".pst",   # raw PST, no zip wrapper (verified)
    ),
    "onedrive": ServiceDef(
        key="onedrive", label="📁 OneDrive", noun="drive", noun_plural="drives",
        section="onedrive", app_type="office365_onedrive",
        export_format=None,                     # OneDrive export modal has no format choice
        file_ext=".zip",
    ),
    "sharepoint": ServiceDef(
        key="sharepoint", label="📋 SharePoint", noun="site", noun_plural="sites",
        section="sharepoint", app_type="office365_sharepoint",
        export_format=None,                     # SharePoint export modal has no format choice
        file_ext=".zip",
    ),
}


# ─── Runtime globals (set by apply_config before any business logic runs) ─────

BASE_URL        = ""
CUSTOMER_ID     = ""
EXT_CUSTOMER_ID = ""
LOGIN_URL       = ""
DASHBOARD_URL   = ""
EXPORT_ACTION   = ""
COOKIE_FILE: Path = Path.home() / ".backupify_session.json"


def apply_config(cfg: dict) -> None:
    """Overwrite module-level runtime globals from the loaded config dict."""
    global BASE_URL, CUSTOMER_ID, EXT_CUSTOMER_ID, LOGIN_URL
    global DASHBOARD_URL, EXPORT_ACTION, COOKIE_FILE

    s = cfg["server"]
    a = cfg["account"]

    BASE_URL        = s.get("base_url", "").rstrip("/")
    CUSTOMER_ID     = s.get("customer_id", "")
    EXT_CUSTOMER_ID = s.get("ext_customer_id", "")
    LOGIN_URL       = f"https://auth.datto.com/login?login_hint={a.get('email', '')}"
    DASHBOARD_URL   = f"{BASE_URL}/{CUSTOMER_ID}/o365?external_customer_id={EXT_CUSTOMER_ID}"
    EXPORT_ACTION   = f"{BASE_URL}/{CUSTOMER_ID}/restoreExportAction"
    COOKIE_FILE     = Path.home() / ".backupify_session.json"


def section_url(svc: ServiceDef) -> str:
    return f"{BASE_URL}/{CUSTOMER_ID}/o365/{svc.section}"


def export_page_url(svc: ServiceDef) -> str:
    return f"{BASE_URL}/{CUSTOMER_ID}/o365/{svc.section}/export"


def service_page_url(svc: ServiceDef, service_id: str) -> str:
    return f"{BASE_URL}/{CUSTOMER_ID}/o365/{svc.section}/service?serviceId={service_id}"


def entity_cache_path(svc: ServiceDef) -> Path:
    # Exchange keeps its pre-refactor cache filename so existing caches survive
    if svc.key == "exchange":
        return Path.home() / f".backupify_users_{CUSTOMER_ID}.json"
    return Path.home() / f".backupify_{svc.key}_{CUSTOMER_ID}.json"


# ─── Poll / download tunables ─────────────────────────────────────────────────

POLL_INTERVAL    = 30           # seconds between export-status scans
EXPORT_TIMEOUT   = 10_800       # max wait for one export job (3 h)

# How many consecutive scans a job may be absent from the export page before
# it's declared missing.  Under heavy load the server can take 45+ minutes to
# even list a freshly triggered job (an 8 MB OneDrive export once took 46 min
# end-to-end), so fresh triggers get a long leash; a resumed job id from an
# earlier session that stays absent is most likely genuinely gone, but still
# gets enough scans to ride out a slow page before we re-trigger.
MISSING_FRESH_SCANS   = 120     # fresh trigger:  ~60 min at POLL_INTERVAL=30s
MISSING_RESUMED_SCANS = 20      # resumed job id: ~10 min, then re-trigger
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


# ─── Entity discovery ─────────────────────────────────────────────────────────

async def get_entities(
    context: BrowserContext,
    cookies: dict,
    svc: ServiceDef,
    logger: logging.Logger,
) -> list[dict]:
    """
    Returns the full entity list for one service (Exchange users, OneDrive
    drives, …).  Loads from the per-service cache file on subsequent runs;
    delete the file (or press F5 in the TUI) to re-fetch.

    Returns [{"service_id", "name", "email", "snapshot_id",
              "size_bytes", "size_label"}, ...].
    """
    cache_file = entity_cache_path(svc)
    if cache_file.exists():
        try:
            users = json.loads(cache_file.read_text(encoding="utf-8"))
            if users:
                for u in users:  # older caches stored HTML entities (&#039; etc.)
                    u["name"]  = html.unescape(u.get("name")  or "")
                    u["email"] = html.unescape(u.get("email") or "")
                logger.info(f"Loaded {len(users)} {svc.key} entities from cache ({cache_file}).")
                return users
        except Exception as e:
            logger.warning(f"Could not read {svc.key} entity cache: {e} — re-fetching.")

    logger.info(f"Loading {svc.key} entity list via customerServices XHR...")
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
            await page.goto(section_url(svc), wait_until="domcontentloaded")

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
            "size_bytes":  row.get("ownSize"),
            "size_label":  row.get("usedBytes") or "",
        })

    if not users:
        raise RuntimeError("No entities extracted from customerServices response.")

    logger.info(f"Found {len(users)} {svc.key} entities.")
    cache_file.write_text(json.dumps(users, indent=2), encoding="utf-8")
    logger.info(f"Entity list cached to {cache_file}.")
    return users


# ─── Snapshot extraction ──────────────────────────────────────────────────────

async def get_latest_snapshot_id(
    context: BrowserContext,
    svc: ServiceDef,
    service_id: str,
    logger: logging.Logger,
) -> str | None:
    service_url = service_page_url(svc, service_id)
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
    svc: ServiceDef,
    service_id: str,
    snapshot_id: str,
    logger: logging.Logger,
) -> str | None:
    payload = {
        "actionType":         "export",
        "appType":            svc.app_type,
        "snapshotId":         snapshot_id,
        "token":              "",
        "includePermissions": "false",
        "includeAttachments": "false",
        "services[]":         service_id,
    }
    if svc.export_format:
        payload["exportFormat"] = svc.export_format
    cookie_header = "; ".join(f"{k}={v}" for k, v in cookies.items())
    async with httpx.AsyncClient(follow_redirects=True, timeout=60.0) as client:
        resp = await client.post(
            EXPORT_ACTION,
            data=payload,
            headers={
                "Cookie":           cookie_header,
                "Referer":          service_page_url(svc, service_id),
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
#
# One status cache per service — each service has its own export page, and
# job ids are only guaranteed unique within one.

_export_caches:      dict[str, dict]  = {key: {} for key in SERVICES}
_export_cache_times: dict[str, float] = {key: 0.0 for key in SERVICES}
_export_scan_locks:  dict[str, asyncio.Lock | None] = {key: None for key in SERVICES}
# Rows still queued/running carry NO action links on some services (SharePoint:
# no ?id= href until the job completes), so they can't go in the id-keyed cache.
# Kept per scan so the poll can match them by source name instead.
_export_pending:     dict[str, list]  = {key: [] for key in SERVICES}


async def scan_export_page(
    context: BrowserContext,
    svc: ServiceDef,
    logger: logging.Logger,
) -> list[dict]:
    """
    Navigates to the service's export page and parses all server-rendered
    DataTable rows, paginating as needed.  Exchange and OneDrive share the
    same #exportListItems layout (status col 7, action links col 8).

    Each record: {"job_id", "source_name", "status", "download_url"}.
    """
    page = await context.new_page()
    records: list[dict] = []

    try:
        await page.goto(export_page_url(svc), wait_until="domcontentloaded", timeout=PAGE_LOAD_MS)
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


async def refresh_export_cache(
    context: BrowserContext,
    svc: ServiceDef,
    logger: logging.Logger,
) -> None:
    """Scans the service's export page once and updates its status cache."""
    records = await scan_export_page(context, svc, logger)
    pending = []
    for rec in records:
        if rec["job_id"]:
            _export_caches[svc.key][rec["job_id"]] = rec
        elif rec["status"] not in ("completed", "failed", "error", "cancelled"):
            pending.append(rec)
    _export_pending[svc.key] = pending
    _export_cache_times[svc.key] = time.monotonic()


def find_server_export(svc: ServiceDef, source_name: str) -> dict | None:
    """
    Checks the last export-page scan for an export of this entity that some
    earlier run (or another machine) already started, so we don't trigger a
    duplicate.  Preference order:
      1. completed with a live download link  → caller can skip straight to it
      2. a row with an id and non-terminal status → adopt and poll by id
      3. an id-less queued/running row ("In Progress", no links yet) → adopt
         by name (poll with job_id=None)
    Expired exports (completed, no link) don't count.  Caller must ensure the
    name is unambiguous — the export page identifies rows by source name only.
    """
    records = [r for r in _export_caches[svc.key].values()
               if r["source_name"] == source_name]
    completed = [r for r in records if r["status"] == "completed" and r["download_url"]]
    if completed:
        return max(completed, key=lambda r: int(r["job_id"]))
    running = [r for r in records
               if r["status"] not in ("completed", "failed", "error", "cancelled")]
    if running:
        return max(running, key=lambda r: int(r["job_id"]))
    for r in _export_pending[svc.key]:
        if r["source_name"] == source_name:
            return r            # job_id is None — adopt by name
    return None


async def poll_for_download_url(
    context: BrowserContext,
    svc: ServiceDef,
    job_id: str | None,
    service_id: str,
    logger: logging.Logger,
    on_status=None,
    max_unseen: int = MISSING_RESUMED_SCANS,
    source_name: str | None = None,
    should_defer=None,
) -> tuple[str, str | None]:
    """
    Waits until the export job is completed and returns (outcome, url):
      ("ok", url)       — job completed, download URL available
      ("missing", None) — job absent from the export page for max_unseen scans
      ("failed", None)  — job ended in failed/error/cancelled
      ("timeout", None) — EXPORT_TIMEOUT exceeded
      ("deferred", None)— should_defer() returned True (wind-down); the job is
                          left running server-side for a later session to resume

    All concurrent workers of one service share one page-scan via its lock.
    on_status: optional callable(str) for live UI updates each poll.
    source_name: the entity's display name as it appears in the export table's
    Source column.  Queued/running rows carry no job id on some services, so a
    name match there means the job is alive — it doesn't count as unseen.
    """
    cache      = _export_caches[svc.key]
    scan_lock  = _export_scan_locks[svc.key]
    job_label  = job_id or f"unlisted export of '{source_name}'"

    # job_id None = adopting an id-less in-flight export found on the page:
    # completion is detected as the first NEW completed row with our name.
    known_ids = (
        {jid for jid, rec in cache.items() if rec["source_name"] == source_name}
        if job_id is None else set()
    )

    start          = time.monotonic()
    deadline       = start + EXPORT_TIMEOUT
    unseen_scans   = 0
    last_scan_time = _export_cache_times[svc.key]

    while time.monotonic() < deadline:
        if should_defer and should_defer():
            logger.info(f"[{service_id}] Poll deferred (wind-down) — {job_label} left running.")
            return ("deferred", None)
        cache_age = time.monotonic() - _export_cache_times[svc.key]
        if cache_age >= POLL_INTERVAL:
            async with scan_lock:
                if time.monotonic() - _export_cache_times[svc.key] >= POLL_INTERVAL:
                    try:
                        await refresh_export_cache(context, svc, logger)
                    except Exception as e:
                        # A flaky scan must not kill the worker — back off
                        # until the next interval and try again.
                        logger.warning(f"Export scan failed (will retry): {e}")
                        _export_cache_times[svc.key] = time.monotonic()

        if job_id is None:
            fresh  = [rec for jid, rec in cache.items()
                      if rec["source_name"] == source_name and jid not in known_ids]
            record = max(fresh, key=lambda r: int(r["job_id"])) if fresh else None
        else:
            record = cache.get(str(job_id))
        elapsed = int(time.monotonic() - start)

        if record:
            unseen_scans = 0
            status = record["status"]
            if on_status:
                on_status(f"job {job_label}: {status} · {elapsed // 60}m {elapsed % 60:02d}s")
            if status == "completed":
                url = record["download_url"]
                if url:
                    logger.info(f"[{service_id}] Export ready after {elapsed}s.")
                    return ("ok", url)
            elif status in ("failed", "error", "cancelled"):
                logger.error(f"[{service_id}] Export job {job_label} ended: {status}")
                return ("failed", None)
        elif source_name and any(
            r["source_name"] == source_name for r in _export_pending[svc.key]
        ):
            # The job has no id-bearing links yet, but a queued/running row
            # with our name is on the export page — it's alive, keep waiting.
            unseen_scans   = 0
            last_scan_time = _export_cache_times[svc.key]
            if on_status:
                on_status(f"job {job_label}: queued · {elapsed // 60}m {elapsed % 60:02d}s")
        else:
            # Count scans that completed while the job was still absent
            if _export_cache_times[svc.key] != last_scan_time:
                unseen_scans  += 1
                last_scan_time = _export_cache_times[svc.key]
            if on_status:
                on_status(f"job {job_label}: not listed yet (scan {unseen_scans}/{max_unseen})")
            if unseen_scans >= max_unseen:
                logger.warning(
                    f"[{service_id}] Job {job_label} absent after {unseen_scans} scans — treating as stale."
                )
                return ("missing", None)

        await asyncio.sleep(POLL_INTERVAL)

    logger.error(f"[{service_id}] Export timed out after {EXPORT_TIMEOUT}s.")
    return ("timeout", None)


# ─── File download ────────────────────────────────────────────────────────────

def _safe_name(user_name: str) -> str:
    return re.sub(r'[<>:"/\\|?*\s]', "_", user_name).strip("_")


def export_file_path(output_dir: Path, user_name: str, service_id: str, ext: str) -> Path:
    """Canonical on-disk path for one entity's export."""
    return output_dir / f"{_safe_name(user_name)}__{service_id}{ext}"


def find_existing_export(output_dir: Path, user_name: str, service_id: str) -> Path | None:
    """
    Finds a previously downloaded export for this entity regardless of
    extension (the server's Content-Disposition decides it at download time).
    """
    prefix = f"{_safe_name(user_name)}__{service_id}."
    if not output_dir.is_dir():
        return None
    for p in output_dir.iterdir():
        if p.name.startswith(prefix) and not p.name.endswith(".part"):
            return p
    return None


async def download_file(
    download_url: str,
    cookies: dict,
    output_dir: Path,
    user_name: str,
    service_id: str,
    svc: ServiceDef,
    logger: logging.Logger,
    on_progress=None,
) -> str | None:
    if download_url.startswith("/"):
        download_url = BASE_URL + download_url

    cookie_header = "; ".join(f"{k}={v}" for k, v in cookies.items())

    existing = find_existing_export(output_dir, user_name, service_id)
    if existing:
        size_mb = existing.stat().st_size / 1_048_576
        logger.info(f"[{service_id}] File already exists ({size_mb:.1f} MB), skipping.")
        return existing.name

    part_file = export_file_path(output_dir, user_name, service_id, ".part")
    if part_file.exists():
        logger.info(f"[{service_id}] Removing stale partial file {part_file.name}")
        part_file.unlink()

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

                # The server's filename (if any) decides the real extension;
                # fall back to the service default.
                ext = svc.file_ext
                disposition = resp.headers.get("content-disposition", "")
                m = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^";]+)', disposition)
                if m:
                    suffix = Path(m.group(1).strip()).suffix
                    if suffix:
                        ext = suffix
                logger.info(
                    f"[{service_id}] Downloading (content-type "
                    f"{resp.headers.get('content-type', '?')}, ext {ext})"
                )

                total      = int(resp.headers.get("content-length", 0))
                downloaded = 0

                with open(part_file, "wb") as f:
                    async for chunk in resp.aiter_bytes(CHUNK_SIZE):
                        f.write(chunk)
                        downloaded += len(chunk)
                        if on_progress:
                            on_progress(downloaded, total)

        filename = export_file_path(output_dir, user_name, service_id, ext)
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
    ("Alphabetical (A → Z)", "az"),
    ("Alphabetical (Z → A)", "za"),
    ("Largest first",        "size_desc"),
    ("Smallest first",       "size_asc"),
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
    "deferred":    "⏸",
}


def _size_sort_key(u: dict) -> int:
    # Real byte size when the entity cache has it; pre-refactor Exchange
    # caches lack it, so fall back to the old snapshot_id (epoch-ms) proxy
    # until the user refreshes the list (F5).
    size = u.get("size_bytes")
    if size is not None:
        return int(size)
    return int(u.get("snapshot_id") or 0)


def _sort_users(users: list[dict], key: str) -> list[dict]:
    if key == "az":
        return sorted(users, key=lambda u: u["name"].lower())
    if key == "za":
        return sorted(users, key=lambda u: u["name"].lower(), reverse=True)
    if key == "size_desc":
        return sorted(users, key=_size_sort_key, reverse=True)
    if key == "size_asc":
        return sorted(users, key=_size_sort_key)
    return users


# ─── TUI — messages ───────────────────────────────────────────────────────────
# `service` is a SERVICES key; None on StatusUpdate/LogEntry = broadcast to
# every tab (session-level events).

class SessionReady(Message):
    pass


class StatusUpdate(Message):
    def __init__(self, text: str, service: str | None = None) -> None:
        super().__init__()
        self.text    = text
        self.service = service


class UsersLoaded(Message):
    def __init__(self, service: str, users: list[dict]) -> None:
        super().__init__()
        self.service = service
        self.users   = users


class UsersLoadFailed(Message):
    def __init__(self, service: str) -> None:
        super().__init__()
        self.service = service


class LogEntry(Message):
    def __init__(self, level: str, text: str, service: str | None = None) -> None:
        super().__init__()
        self.level   = level
        self.text    = text
        self.service = service


class ExportUpdate(Message):
    def __init__(self, service: str, service_id: str, status: str, detail: str = "") -> None:
        super().__init__()
        self.service    = service
        self.service_id = service_id
        self.status     = status
        self.detail     = detail


class DownloadProgress(Message):
    def __init__(self, service: str, service_id: str, done: int, total: int) -> None:
        super().__init__()
        self.service    = service
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
                        "Welcome! This tool bulk-exports M365 mailboxes, OneDrive\n"
                        "drives, and SharePoint sites from Backupify/Datto.\n\n"
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


# ─── Wind-down modes ──────────────────────────────────────────────────────────
#
# App-wide switch for winding work down gracefully (swap drives, shut the
# machine off, …) without killing anything mid-flight.  Deferred entities
# keep their progress.json in_progress record, so the next run resumes the
# same server-side job instead of triggering a new one.

WINDDOWN_MODES: list[tuple[str, str]] = [
    ("off",              "off — normal operation"),
    ("no-new-jobs",      "finish in-flight, start nothing new"),
    ("no-new-downloads", "keep queueing exports, defer downloads"),
    ("drain",            "finish current downloads only"),
]


class ExtendPollScreen(ModalScreen[bool]):
    """'Export job still hasn't appeared — keep waiting?' dialog."""

    def __init__(self, entity_name: str, job_id: str, waited_min: int) -> None:
        super().__init__()
        self._entity_name = entity_name
        self._job_id      = job_id
        self._waited_min  = waited_min

    def compose(self) -> ComposeResult:
        with Vertical(id="extend-dialog"):
            yield Static("⚠  Export job not listed yet", id="extend-title")
            yield Static(
                f"{self._entity_name}: job {self._job_id} has not appeared on the "
                f"export page after ~{self._waited_min} min.\n\n"
                "Under heavy load Backupify can take a long time to even list a "
                "job — it may well still be queued server-side.",
                id="extend-body",
            )
            with Horizontal(id="extend-buttons"):
                yield Button(f"Keep waiting (+{self._waited_min} min)",
                             variant="primary", id="extend-wait")
                yield Button("Give up (mark failed)", variant="error", id="extend-fail")

    @on(Button.Pressed, "#extend-wait")
    def _handle_extend_wait(self) -> None:
        self.dismiss(True)

    @on(Button.Pressed, "#extend-fail")
    def _handle_extend_fail(self) -> None:
        self.dismiss(False)


# ─── TUI — main app ───────────────────────────────────────────────────────────

class TabState:
    """Per-service UI state — one instance per service tab."""

    def __init__(self, svc: ServiceDef) -> None:
        self.svc = svc
        self.users:    list[dict] = []
        self.filtered: list[dict] = []
        self.selected: set[str]   = set()
        self.state: StateManager | None = None
        self.sort   = "az"
        self.search = ""
        self.users_loaded   = False
        self.users_loading  = False
        self.export_running = False
        self.run_pending: set[str] = set()


class BackupifyApp(App):
    TITLE = "Backupify Exporter"

    CSS = """
    TabbedContent { height: 1fr; }
    /* Textual defaults TabPane + inner ContentSwitcher to height:auto, which
       collapses children using percentage heights to 0 — pin them to fill */
    TabbedContent > ContentSwitcher { height: 1fr; }
    TabPane { height: 100%; }

    .service-view { height: 100%; layout: vertical; }

    .settings-bar {
        height: 3;
        padding: 0 1;
        background: $boost;
        align: left middle;
    }
    .settings-bar Label { width: auto; padding: 0 1; color: $text-muted; }
    .output-dir-input  { width: 1fr; }
    .concurrency-input { width: 6; }
    .sort-select       { width: 30; }

    .main-panel { height: 1fr; }

    .user-panel {
        width: 40%;
        layout: vertical;
        border-right: solid $primary;
    }
    .search-input { dock: top; }
    .user-list { height: 1fr; }
    .user-stats {
        height: 1;
        padding: 0 1;
        color: $text-muted;
        background: $boost;
    }
    .user-actions {
        height: 3;
        layout: horizontal;
        background: $boost;
        align: left middle;
        padding: 0 1;
    }
    .user-actions Button { margin-right: 1; min-width: 8; }

    .export-panel { width: 60%; layout: vertical; }
    .progress-table { height: 1fr; }
    .log-pane { height: 10; border-top: solid $primary; }

    .control-bar {
        height: 3;
        layout: horizontal;
        background: $boost;
        align: left middle;
        padding: 0 1;
    }
    .control-bar Button { margin-right: 1; }
    .status-text { color: $text-muted; content-align: left middle; width: 1fr; height: 3; }

    .coming-soon {
        height: 100%;
        content-align: center middle;
        color: $text-muted;
    }

    ExtendPollScreen { align: center middle; }
    #extend-dialog {
        width: 64;
        height: auto;
        padding: 1 2;
        background: $surface;
        border: thick $warning;
    }
    #extend-title   { text-style: bold; }
    #extend-body    { padding: 1 0; }
    #extend-buttons { height: auto; align: center middle; }
    #extend-buttons Button { margin: 0 1; }
    """

    BINDINGS = [
        Binding("q",      "quit",           "Quit"),
        Binding("ctrl+s", "start_export",   "Start"),
        Binding("f5",     "refresh_users",  "Refresh Users"),
        Binding("ctrl+r", "retry_failed",   "Retry Failed"),
        Binding("ctrl+w", "cycle_winddown", "Wind-down"),
    ]

    def __init__(self, cfg: dict) -> None:
        super().__init__()
        self._cfg = cfg

        self._pw_context: BrowserContext | None = None
        self._cookies: dict = {}

        self._tabs: dict[str, TabState] = {
            key: TabState(svc) for key, svc in SERVICES.items()
        }

        self._bfy_logger: logging.Logger | None = None
        self._debug_log:  logging.Logger | None = None
        self._semaphore:  asyncio.Semaphore | None = None

        self.winddown = "off"                     # see WINDDOWN_MODES
        self._prompt_lock = asyncio.Lock()        # one missing-job dialog at a time

    # ── Layout ────────────────────────────────────────────────────────────

    def compose(self) -> ComposeResult:
        yield Header()
        with TabbedContent(initial="tab-exchange"):
            for key, svc in SERVICES.items():
                with TabPane(svc.label, id=f"tab-{key}"):
                    yield from self._compose_service_tab(svc)
            with TabPane("💬 Teams", id="tab-teams"):
                yield Static("Teams export — coming soon (see ROADMAP.md)", classes="coming-soon")
        yield Footer()

    def _compose_service_tab(self, svc: ServiceDef) -> ComposeResult:
        k = svc.key
        with Vertical(classes="service-view"):
            with Horizontal(classes="settings-bar"):
                yield Label("Output:")
                yield Input(
                    value=self._cfg["defaults"]["output_dir"],
                    placeholder="/path/to/exports",
                    id=f"{k}-output-dir-input",
                    classes="output-dir-input",
                )
                yield Label("Workers:")
                yield Input(
                    value=str(self._cfg["defaults"]["concurrency"]),
                    id=f"{k}-concurrency-input",
                    classes="concurrency-input",
                    restrict=r"[0-9]*",
                )
                yield Label("Sort:")
                yield Select(
                    _SORT_OPTIONS,
                    value="az",
                    allow_blank=False,
                    id=f"{k}-sort-select",
                    classes="sort-select",
                )
            with Horizontal(classes="main-panel"):
                with Vertical(classes="user-panel"):
                    yield Input(
                        placeholder=f"🔍 Filter {svc.noun_plural}…",
                        id=f"{k}-search-input",
                        classes="search-input",
                    )
                    yield SelectionList(id=f"{k}-user-list", classes="user-list")
                    yield Static("Loading…", id=f"{k}-user-stats", classes="user-stats")
                    with Horizontal(classes="user-actions"):
                        yield Button("All",  id=f"{k}-btn-all",  classes="btn-all",  variant="default")
                        yield Button("None", id=f"{k}-btn-none", classes="btn-none", variant="default")
                with Vertical(classes="export-panel"):
                    yield DataTable(
                        id=f"{k}-progress-table",
                        classes="progress-table",
                        zebra_stripes=True,
                        cursor_type="row",
                    )
                    yield RichLog(id=f"{k}-log-pane", classes="log-pane", highlight=True, markup=True)
            with Horizontal(classes="control-bar"):
                yield Button("▶ Start", id=f"{k}-btn-start", classes="btn-start",
                             variant="success", disabled=True)
                yield Button("↺ Retry Failed", id=f"{k}-btn-retry", classes="btn-retry",
                             variant="default")
                yield Static("⟳ Starting up…", id=f"{k}-status-text", classes="status-text")

    # ── Lifecycle ─────────────────────────────────────────────────────────

    def on_mount(self) -> None:
        apply_config(self._cfg)

        output_root = Path(self._cfg["defaults"]["output_dir"])
        output_root.mkdir(parents=True, exist_ok=True)

        self._bfy_logger = setup_logging(output_root)
        self._debug_log  = setup_playwright_debug_logging()
        self._semaphore  = asyncio.Semaphore(int(self._cfg["defaults"]["concurrency"]))

        for key, tab in self._tabs.items():
            service_dir = output_root / key
            service_dir.mkdir(parents=True, exist_ok=True)
            tab.state = StateManager(service_dir)
            _export_scan_locks[key] = asyncio.Lock()

            table = self.query_one(f"#{key}-progress-table", DataTable)
            table.add_column("Name",     key="name",   width=28)
            table.add_column("Status",   key="status", width=15)
            table.add_column("Progress", key="detail", width=40)

        # Spinner over the initial tab's list until UsersLoaded arrives
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
            self.post_message(StatusUpdate("⟳ Session ready — loading entity list…"))
            self.post_message(LogEntry("info", "Session ready — loading entities…"))
            active = self._active_service_key() or "exchange"
            self._request_users_load(active)
        except Exception as exc:
            self.post_message(StatusUpdate("✗ Session failed — see log pane"))
            self.post_message(LogEntry("error", f"Session error: {exc}"))
            for key in self._tabs:
                self.query_one(f"#{key}-user-list", SelectionList).loading = False

    def _request_users_load(self, key: str) -> None:
        """Loads a tab's entity list once, the first time it's needed."""
        tab = self._tabs[key]
        if not self._pw_context or tab.users_loaded or tab.users_loading:
            return
        tab.users_loading = True
        self.query_one(f"#{key}-user-list", SelectionList).loading = True
        self.load_users(key)

    @work(group="users")
    async def load_users(self, key: str) -> None:
        tab = self._tabs[key]
        try:
            users = await get_entities(
                self._pw_context, self._cookies, tab.svc, self._bfy_logger
            )
            self.post_message(UsersLoaded(key, users))
        except Exception as exc:
            self.post_message(LogEntry("error", f"Entity load failed: {exc}", service=key))
            self.post_message(UsersLoadFailed(key))

    @work(group="preflight")
    async def run_preflight(self, key: str, users: list[dict], output_dir: Path) -> None:
        tab = self._tabs[key]
        self.post_message(LogEntry("info", "Scanning existing export jobs…", service=key))
        try:
            await refresh_export_cache(self._pw_context, tab.svc, self._bfy_logger)
            self.post_message(LogEntry(
                "info", f"Found {len(_export_caches[key])} existing job(s).", service=key
            ))
        except Exception as exc:
            self.post_message(LogEntry("warning", f"Export scan failed: {exc}", service=key))

        for user in users:
            sid = user["service_id"]
            if tab.state.is_done(sid):
                self.post_message(ExportUpdate(key, sid, "skipped", "already complete"))
            else:
                self.export_one_user(key, user, output_dir)

    @work(group="exports")
    async def export_one_user(self, key: str, user: dict, output_dir: Path) -> None:
        tab   = self._tabs[key]
        svc   = tab.svc
        state = tab.state
        sid   = user["service_id"]
        name  = user.get("name") or sid

        # Already on disk (e.g. downloaded by an earlier version) — don't
        # trigger a fresh server-side export for it.
        existing_file = find_existing_export(output_dir, name, sid)
        if existing_file:
            state.mark_complete(sid, existing_file.name)
            self.post_message(ExportUpdate(key, sid, "skipped", "already on disk"))
            return

        if self.winddown in ("no-new-jobs", "drain"):
            self.post_message(ExportUpdate(key, sid, "deferred", "wind-down — not started"))
            return

        async with self._semaphore:
            # Mode may have flipped while this worker queued on the semaphore
            if self.winddown in ("no-new-jobs", "drain"):
                self.post_message(ExportUpdate(key, sid, "deferred", "wind-down — not started"))
                return
            try:
                snapshot_id = user.get("snapshot_id") or await get_latest_snapshot_id(
                    self._pw_context, svc, sid, self._bfy_logger
                )
                if not snapshot_id:
                    state.mark_failed(sid, "No snapshot ID")
                    self.post_message(ExportUpdate(key, sid, "failed", "no snapshot"))
                    return

                existing = state.get_in_progress(sid)
                resumed  = bool(existing)
                adopted: dict | None = None
                if existing:
                    job_id = existing["job_id"]
                    self.post_message(ExportUpdate(key, sid, "resuming", f"job {job_id}"))
                else:
                    # An earlier run (or another machine) may already have an
                    # export of this entity on the server — adopt it instead of
                    # triggering a duplicate.  Only safe when the name is
                    # unique: the export page identifies rows by name alone.
                    if sum(1 for u in tab.users if u.get("name") == name) == 1:
                        adopted = find_server_export(svc, name)
                    if adopted is not None:
                        job_id  = adopted.get("job_id")
                        resumed = True
                        if job_id:
                            state.mark_in_progress(sid, job_id, snapshot_id)
                        self.post_message(ExportUpdate(
                            key, sid, "resuming",
                            f"found on server ({adopted['status']})",
                        ))
                        self.post_message(LogEntry(
                            "info",
                            f"{name}: export already on server "
                            f"({adopted['status']}, job {job_id or 'unlisted'}) — adopting it.",
                            service=key,
                        ))
                    else:
                        self.post_message(ExportUpdate(key, sid, "triggering"))
                        job_id = await trigger_export(
                            self._pw_context, self._cookies, svc, sid, snapshot_id, self._bfy_logger
                        )
                        if not job_id:
                            state.mark_failed(sid, "Trigger failed")
                            self.post_message(ExportUpdate(key, sid, "failed", "trigger failed"))
                            return
                        state.mark_in_progress(sid, job_id, snapshot_id)

                def _poll_status(text: str) -> None:
                    self.post_message(ExportUpdate(key, sid, "polling", text))

                def _should_defer() -> bool:
                    return self.winddown == "drain"

                if adopted is not None and adopted["status"] == "completed":
                    # Server already has the finished export — skip the poll.
                    outcome, download_url = "ok", adopted["download_url"]
                else:
                    self.post_message(ExportUpdate(key, sid, "polling", f"job {job_id or 'unlisted'}"))
                    # A resumed job that stays absent is stale; a freshly
                    # triggered one definitely exists, so wait much longer
                    # before giving up on it.
                    outcome, download_url = await poll_for_download_url(
                        self._pw_context, svc, job_id, sid, self._bfy_logger,
                        on_status=_poll_status,
                        max_unseen=MISSING_RESUMED_SCANS if resumed else MISSING_FRESH_SCANS,
                        source_name=name,
                        should_defer=_should_defer,
                    )

                if outcome == "missing" and resumed:
                    # Saved job id from an earlier session no longer exists on
                    # the export page — start a fresh export instead.
                    self.post_message(ExportUpdate(key, sid, "triggering", "stale job — re-exporting"))
                    self.post_message(LogEntry(
                        "warning",
                        f"{name}: job {job_id or 'unlisted'} is gone — re-triggering export.",
                        service=key,
                    ))
                    state.clear_in_progress(sid)
                    job_id = await trigger_export(
                        self._pw_context, self._cookies, svc, sid, snapshot_id, self._bfy_logger
                    )
                    if not job_id:
                        state.mark_failed(sid, "Re-trigger failed")
                        self.post_message(ExportUpdate(key, sid, "failed", "re-trigger failed"))
                        return
                    state.mark_in_progress(sid, job_id, snapshot_id)
                    self.post_message(ExportUpdate(key, sid, "polling", f"job {job_id}"))
                    outcome, download_url = await poll_for_download_url(
                        self._pw_context, svc, job_id, sid, self._bfy_logger,
                        on_status=_poll_status,
                        max_unseen=MISSING_FRESH_SCANS,
                        source_name=name,
                        should_defer=_should_defer,
                    )

                # A freshly triggered job that's still unlisted is usually the
                # server lagging under load, not a lost job — ask before failing.
                while outcome == "missing" and self.winddown != "drain":
                    self.post_message(ExportUpdate(key, sid, "polling", "missing — awaiting your call"))
                    if not await self._ask_extend_poll(name, job_id):
                        break
                    self.post_message(LogEntry(
                        "info", f"{name}: extending wait for job {job_id}.", service=key
                    ))
                    outcome, download_url = await poll_for_download_url(
                        self._pw_context, svc, job_id, sid, self._bfy_logger,
                        on_status=_poll_status,
                        max_unseen=MISSING_FRESH_SCANS,
                        source_name=name,
                        should_defer=_should_defer,
                    )

                if outcome == "deferred":
                    # in_progress stays in progress.json → next run resumes it
                    self.post_message(ExportUpdate(key, sid, "deferred", "wind-down — job left running"))
                    self.post_message(LogEntry(
                        "info", f"{name}: poll deferred — job {job_id} resumes next run.", service=key
                    ))
                    return

                if not download_url:
                    state.mark_failed(sid, f"Export {outcome}")
                    self.post_message(ExportUpdate(key, sid, "failed", f"export {outcome}"))
                    return

                if self.winddown in ("no-new-downloads", "drain"):
                    # in_progress stays → next run resumes the completed job
                    # and goes straight to the download.
                    self.post_message(ExportUpdate(key, sid, "deferred", "export ready — download deferred"))
                    self.post_message(LogEntry(
                        "info", f"{name}: export ready — download deferred (wind-down).", service=key
                    ))
                    return

                self.post_message(ExportUpdate(key, sid, "downloading"))

                def _on_progress(done: int, total: int) -> None:
                    self.post_message(DownloadProgress(key, sid, done, total))

                filename = await download_file(
                    download_url, self._cookies, output_dir,
                    name, sid, svc, self._bfy_logger, on_progress=_on_progress,
                )
                if filename:
                    state.mark_complete(sid, filename)
                    self.post_message(ExportUpdate(key, sid, "complete", filename[:38]))
                    self.post_message(LogEntry("success", f"✓ {name}  →  {filename}", service=key))
                else:
                    state.mark_failed(sid, "Download failed")
                    self.post_message(ExportUpdate(key, sid, "failed", "download error"))
                    self.post_message(LogEntry("error", f"✗ {name}  download failed", service=key))

            except Exception as exc:
                state.mark_failed(sid, str(exc))
                self.post_message(ExportUpdate(key, sid, "failed", str(exc)[:38]))
                self.post_message(LogEntry("error", f"✗ {name}: {exc}", service=key))

    # ── Worker→UI message handlers ────────────────────────────────────────

    def _set_status(self, text: str, service: str | None) -> None:
        keys = [service] if service else list(self._tabs)
        for key in keys:
            self.query_one(f"#{key}-status-text", Static).update(text)

    @on(SessionReady)
    def _handle_session_ready(self) -> None:
        self._set_status("⟳ Session active — loading entities…", None)

    @on(StatusUpdate)
    def _handle_status_update(self, event: StatusUpdate) -> None:
        self._set_status(event.text, event.service)

    @on(UsersLoaded)
    def _handle_users_loaded(self, event: UsersLoaded) -> None:
        tab = self._tabs[event.service]
        tab.users         = event.users
        tab.selected      = {u["service_id"] for u in event.users}
        tab.users_loaded  = True
        tab.users_loading = False
        self.query_one(f"#{event.service}-user-list", SelectionList).loading = False
        self._refresh_user_list(event.service)
        plural = tab.svc.noun_plural
        self._set_status(
            f"✓ Ready — {len(event.users)} {plural} loaded. Select {plural} and press Start.",
            event.service,
        )
        if not tab.export_running:
            self.query_one(f"#{event.service}-btn-start", Button).disabled = False
        self.notify(f"Ready — {len(event.users)} {plural} loaded.", title="Backupify")

    @on(UsersLoadFailed)
    def _handle_users_load_failed(self, event: UsersLoadFailed) -> None:
        tab = self._tabs[event.service]
        tab.users_loading = False
        self.query_one(f"#{event.service}-user-list", SelectionList).loading = False
        self._set_status("✗ Entity load failed — see log pane", event.service)

    @on(LogEntry)
    def _handle_log_entry(self, event: LogEntry) -> None:
        colour = {
            "info": "white", "warning": "yellow",
            "error": "red", "success": "green",
        }.get(event.level, "white")
        ts   = datetime.now().strftime("%H:%M:%S")
        line = f"[{colour}]{ts}  {rich_escape(event.text)}[/{colour}]"
        keys = [event.service] if event.service else list(self._tabs)
        for key in keys:
            self.query_one(f"#{key}-log-pane", RichLog).write(line)

    @on(ExportUpdate)
    def _handle_export_update(self, event: ExportUpdate) -> None:
        tab   = self._tabs[event.service]
        icon  = _STATUS_ICONS.get(event.status, "?")
        table = self.query_one(f"#{event.service}-progress-table", DataTable)
        try:
            table.update_cell(event.service_id, "status", f"{icon} {event.status.capitalize()}")
            if event.detail:
                table.update_cell(event.service_id, "detail", event.detail)
        except Exception:
            pass  # row may not exist (e.g. update for a user not in this run)
        if event.status in ("complete", "failed", "skipped", "deferred"):
            self._update_stats(event.service)
            if tab.export_running and event.service_id in tab.run_pending:
                tab.run_pending.discard(event.service_id)
                if not tab.run_pending:
                    self._finish_run(event.service)

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
            self.query_one(f"#{event.service}-progress-table", DataTable).update_cell(
                event.service_id, "detail", detail
            )
        except Exception:
            pass

    # ── UI event handlers ─────────────────────────────────────────────────

    def _tab_for(self, widget_id: str | None) -> TabState | None:
        """Maps a per-tab widget id ('onedrive-btn-start') to its TabState."""
        if not widget_id:
            return None
        return self._tabs.get(widget_id.split("-", 1)[0])

    @on(TabbedContent.TabActivated)
    def _handle_tab_activated(self, event: TabbedContent.TabActivated) -> None:
        key = (event.pane.id or "").removeprefix("tab-")
        if key in self._tabs:
            self._request_users_load(key)

    @on(Input.Changed, ".search-input")
    def _handle_search(self, event: Input.Changed) -> None:
        tab = self._tab_for(event.input.id)
        if tab:
            tab.search = event.value
            self._refresh_user_list(tab.svc.key)

    @on(Select.Changed, ".sort-select")
    def _handle_sort(self, event: Select.Changed) -> None:
        tab = self._tab_for(event.select.id)
        if tab and event.value is not Select.BLANK:
            tab.sort = str(event.value)
            self._refresh_user_list(tab.svc.key)

    @on(Input.Changed, ".output-dir-input")
    def _handle_output_dir_changed(self, event: Input.Changed) -> None:
        if event.value.strip():
            self._cfg["defaults"]["output_dir"] = event.value.strip()

    @on(Input.Changed, ".concurrency-input")
    def _handle_concurrency_changed(self, event: Input.Changed) -> None:
        try:
            n = int(event.value)
        except ValueError:
            return
        if 1 <= n <= 32:
            self._cfg["defaults"]["concurrency"] = n
            if not any(t.export_running for t in self._tabs.values()):
                self._semaphore = asyncio.Semaphore(n)

    @on(SelectionList.SelectedChanged, ".user-list")
    def _handle_user_selection(self, event: SelectionList.SelectedChanged) -> None:
        tab = self._tab_for(event.selection_list.id)
        if not tab:
            return
        visible_ids    = {u["service_id"] for u in tab.filtered}
        newly_selected = set(event.selection_list.selected)
        tab.selected   = (tab.selected - visible_ids) | newly_selected
        self._update_stats(tab.svc.key)

    @on(Button.Pressed, ".btn-all")
    def _handle_select_all(self, event: Button.Pressed) -> None:
        tab = self._tab_for(event.button.id)
        if tab:
            tab.selected |= {u["service_id"] for u in tab.filtered}
            self._refresh_user_list(tab.svc.key)

    @on(Button.Pressed, ".btn-none")
    def _handle_select_none(self, event: Button.Pressed) -> None:
        tab = self._tab_for(event.button.id)
        if tab:
            tab.selected -= {u["service_id"] for u in tab.filtered}
            self._refresh_user_list(tab.svc.key)

    @on(Button.Pressed, ".btn-start")
    def _handle_btn_start(self, event: Button.Pressed) -> None:
        tab = self._tab_for(event.button.id)
        if tab:
            self._start_export_for(tab.svc.key)

    @on(Button.Pressed, ".btn-retry")
    def _handle_btn_retry(self, event: Button.Pressed) -> None:
        tab = self._tab_for(event.button.id)
        if tab:
            self._retry_failed_for(tab.svc.key)

    # ── Actions ───────────────────────────────────────────────────────────

    def _active_service_key(self) -> str | None:
        active = self.query_one(TabbedContent).active.removeprefix("tab-")
        return active if active in self._tabs else None

    def action_start_export(self) -> None:
        key = self._active_service_key()
        if key:
            self._start_export_for(key)

    def _start_export_for(self, key: str) -> None:
        tab = self._tabs[key]
        if not self._pw_context:
            self.notify("Session not ready yet — please wait.", severity="warning")
            return
        if tab.export_running:
            self.notify("Export already running.", severity="warning")
            return
        if self.winddown != "off":
            self.notify(f"Wind-down active ({self.winddown}) — work will be deferred accordingly.",
                        severity="warning")

        selected = [u for u in tab.users if u["service_id"] in tab.selected]
        if not selected:
            self.notify(f"No {tab.svc.noun_plural} selected.", severity="warning")
            return

        selected   = _sort_users(selected, tab.sort)
        output_dir = Path(self._cfg["defaults"]["output_dir"]) / key
        output_dir.mkdir(parents=True, exist_ok=True)
        save_config(self._cfg)

        tab.export_running = True
        tab.run_pending    = {u["service_id"] for u in selected}
        self.query_one(f"#{key}-btn-start", Button).disabled = True

        table = self.query_one(f"#{key}-progress-table", DataTable)
        table.clear()
        for user in selected:
            table.add_row(
                user["name"][:28], "○ Queued", "",
                key=user["service_id"],
            )

        self.post_message(StatusUpdate(
            f"⟳ Exporting {len(selected)} {tab.svc.noun}(s)…", service=key
        ))
        self.post_message(LogEntry(
            "info", f"Starting export for {len(selected)} {tab.svc.noun}(s)…", service=key
        ))
        self.run_preflight(key, selected, output_dir)

    def action_refresh_users(self) -> None:
        key = self._active_service_key()
        if not key:
            return
        if not self._pw_context:
            self.notify("Session not ready yet.", severity="warning")
            return
        tab        = self._tabs[key]
        cache_file = entity_cache_path(tab.svc)
        if cache_file.exists():
            cache_file.unlink()
        tab.users_loaded = False
        self.post_message(StatusUpdate("⟳ Re-fetching entity list from Backupify…", service=key))
        self.post_message(LogEntry("info", "Entity cache cleared — re-fetching…", service=key))
        self._request_users_load(key)

    def action_retry_failed(self) -> None:
        key = self._active_service_key()
        if key:
            self._retry_failed_for(key)

    def _retry_failed_for(self, key: str) -> None:
        tab = self._tabs[key]
        if not tab.state:
            return
        n = len(tab.state.state["failed"])
        if n == 0:
            self.notify("No failed exports to retry.")
            return
        tab.state.state["failed"].clear()
        tab.state.save()
        tab.export_running = False
        self.query_one(f"#{key}-btn-start", Button).disabled = False
        self.notify(f"Cleared {n} failed entr{'y' if n == 1 else 'ies'} — press Start to retry.")
        self._refresh_user_list(key)

    def _finish_run(self, key: str) -> None:
        tab = self._tabs[key]
        tab.export_running = False
        self.query_one(f"#{key}-btn-start", Button).disabled = False
        done   = len(tab.state.state["completed"])
        failed = len(tab.state.state["failed"])
        self._set_status(f"✓ Run finished — {done} complete, {failed} failed.", key)
        self.post_message(LogEntry(
            "info", f"Run finished — {done} complete, {failed} failed.", service=key
        ))
        self.notify(f"Run finished — {done} complete, {failed} failed.", title="Backupify")
        self._refresh_user_list(key)

    def action_cycle_winddown(self) -> None:
        keys = [k for k, _ in WINDDOWN_MODES]
        idx  = (keys.index(self.winddown) + 1) % len(keys)
        self.winddown, label = WINDDOWN_MODES[idx]
        self.sub_title = "" if self.winddown == "off" else f"⏸ Wind-down: {label}"
        self.notify(f"Wind-down: {label}", title="Backupify")
        self.post_message(LogEntry("warning" if self.winddown != "off" else "info",
                                   f"Wind-down mode: {label}"))

    async def _ask_extend_poll(self, name: str, job_id: str) -> bool:
        """Missing-job dialog (serialized — one at a time). True = keep waiting."""
        waited_min = int(MISSING_FRESH_SCANS * POLL_INTERVAL / 60)
        async with self._prompt_lock:
            return bool(await self.push_screen_wait(
                ExtendPollScreen(name, job_id, waited_min)
            ))

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

    def _refresh_user_list(self, key: str) -> None:
        tab    = self._tabs[key]
        search = tab.search.lower()
        filtered = [
            u for u in tab.users
            if not search
            or search in u["name"].lower()
            or search in (u.get("email") or "").lower()
        ]
        tab.filtered = _sort_users(filtered, tab.sort)

        sl    = self.query_one(f"#{key}-user-list", SelectionList)
        state = tab.state
        sl.clear_options()
        for user in tab.filtered:
            sid = user["service_id"]
            if state and state.is_done(sid):
                label = f"✓ {user['name']}"
            elif state and state.get_in_progress(sid):
                label = f"⟳ {user['name']}"
            elif state and state.is_failed(sid):
                label = f"✗ {user['name']}"
            else:
                label = f"  {user['name']}"
            sl.add_option(Selection(label, sid, sid in tab.selected))

        self._update_stats(key)

    def _update_stats(self, key: str) -> None:
        tab   = self._tabs[key]
        state = tab.state
        total = len(tab.users)
        if state:
            done    = len(state.state["completed"])
            in_prog = len(state.state["in_progress"])
            failed  = len(state.state["failed"])
            queued  = max(0, total - done - in_prog - failed)
        else:
            done = in_prog = failed = queued = 0
        try:
            self.query_one(f"#{key}-user-stats", Static).update(
                f"✓ {done}  ⟳ {in_prog}  ✗ {failed}  ○ {queued}"
                f"  │  {len(tab.selected)}/{total} selected"
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
