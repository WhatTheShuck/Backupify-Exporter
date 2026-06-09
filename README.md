# Backupify PST Bulk Exporter

Automates exporting all 260 Exchange mailbox PSTs from Backupify.
Works on both Linux and Windows from the same codebase.

---

## How it works

1. Opens a **visible Chromium window** for you to log in (password + TOTP)
2. Captures session cookies from the authenticated browser context
3. Uses Playwright to render the SPA and discover all Exchange users
4. For each user, renders their service page to get the latest snapshot ID
5. POSTs the export job via httpx using the session cookies
6. Polls the exports endpoint until each job completes
7. Streams each file to disk via httpx (suitable for 100 GB files)
8. Writes `progress.json` so you can **resume** if anything goes wrong

---

## Setup

### Install uv (once per machine)

**Linux:**
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

**Windows (PowerShell):**
```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

### Install Playwright's Chromium browser (once per machine)

Python dependencies are handled automatically by `uv run`.
Chromium itself needs a one-off install:

```bash
uv run python -m playwright install chromium
```

That's it. No venv, no pip, no separate Python install.

---

## Usage

### Dry run first — lists users without exporting anything
```bash
# Linux
uv run backupify_export.py --output /mnt/usb_drive --dry-run

# Windows
uv run backupify_export.py --output D:\backupify_exports --dry-run
```

### Full export run
```bash
uv run backupify_export.py --output /mnt/usb_drive
```

### Resume after interruption
```bash
uv run backupify_export.py --output /mnt/usb_drive --resume
```

### Adjust concurrency (lower if you see errors)
```bash
uv run backupify_export.py --output /mnt/usb_drive --concurrency 5
```

### Check progress mid-run or after
```bash
uv run progress_check.py /mnt/usb_drive
```

---

## Login flow

A Chromium window opens at the Datto login page with your email pre-filled.

1. Enter your **password**
2. Enter your **TOTP code** when prompted
3. Wait to land on the Backupify dashboard

The script detects the redirect and takes over automatically.
You have **5 minutes** to complete the login before it times out.

---

## Output files

All files land in the `--output` directory:

| File | Description |
|---|---|
| `DisplayName__serviceId.pst.zip` | PST archive — unzip to get the .pst |
| `progress.json` | Resume state — completed/failed |
| `backupify_export_YYYYMMDD_HHMMSS.log` | Full run log |
| `debug_exchange_page.html` | Only created if user discovery fails |
| `debug_service_<id>.html` | Only created if snapshot ID extraction fails |

---

## Splitting across two machines

Run both machines simultaneously pointing at different output drives.
Each machine authenticates independently with its own browser session.
Use `--concurrency 4` on each to avoid overwhelming Backupify's export queue.

```bash
# Machine 1 (Linux)
uv run backupify_export.py --output /mnt/drive_a --concurrency 4

# Machine 2 (Windows)
uv run backupify_export.py --output D:\drive_b --concurrency 4
```

Both will discover the full ~260 user list and export all of them.
This means some mailboxes will be exported twice — deduplicate afterwards,
or manually split the work by editing `progress.json` on one machine to
pre-populate the `completed` block with half the service IDs.

---

## Timeouts & tuning

| Setting | Default | Notes |
|---|---|---|
| `POLL_INTERVAL` | 30s | How often to check if an export job is ready |
| `EXPORT_TIMEOUT` | 3 hours | Max wait per export job |
| `DOWNLOAD_TIMEOUT` | 2 hours | Max time for a single file download |
| `CHUNK_SIZE` | 16 MB | Download chunk size |
| `--concurrency` | 8 | Parallel exports; reduce to 4–5 if errors increase |

---

## Troubleshooting

**"Could not find the user list on the Exchange recovery page"**
The script saves `debug_exchange_page.html` in your working directory.
Open it in a browser (or text editor), find what element contains the
user rows, and share it so the selector can be updated.

**"Missing expected session cookies"**
After logging in manually, open DevTools → Application → Cookies on the
backupify domain. Check what cookies are present and update the `wanted`
set in `extract_cookies()` accordingly.

**"Could not determine snapshot ID"**
The script saves `debug_service_<serviceId>.html`. Find the element that
shows the snapshot/date selector and report it for a fix.

**Lots of failed exports**
Reduce `--concurrency` to 3. Backupify may throttle simultaneous export
jobs per session.

**Script crashes mid-run**
Re-run with `--resume` — it reads `progress.json` and skips completed jobs.

**Download links expire**
Backupify download links expire after 14 days. For failed entries in
`progress.json`, remove them from the `"failed"` block and re-run with
`--resume` to re-trigger just those exports.
