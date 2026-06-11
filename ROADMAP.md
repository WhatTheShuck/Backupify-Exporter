# Backupify Exporter — Roadmap

> Implementing the OneDrive / SharePoint / Teams tabs? Start with **HANDOFF.md** —
> it has the architecture, the discovery playbook, and the list of landmines.

## Planned features

### Snapshot selection
Currently the exporter always uses the most recent snapshot for each user/service.
A future version should let you browse a user's snapshot history and choose which
point-in-time to export from (e.g. "give me the mailbox as it was 3 months ago").
The `perfectBackups` array returned by the `customerServices` API already contains
all available snapshot IDs — the UI just needs to expose them per-user.

### Teams export
Teams entities are *teams* (and optionally channels). Like SharePoint these are
not user mailboxes. Exports include chat history and files. Reserved in the TUI.

### TOTP auto-detection
Currently the login prompt always tells the user to enter their TOTP code. If a
Backupify account has 2FA disabled, the browser skips straight to the dashboard
after the password. Auto-detect this by watching whether the TOTP input field
appears, and adjust the on-screen message accordingly.

### Pause / resume in-flight downloads
The TUI has a Pause button that is not yet functional. Implementing it requires
a cooperative cancellation mechanism (e.g. an `asyncio.Event` checked between
download chunks) and storing the partial byte offset so a resumed download can
send a `Range:` header to avoid re-downloading from the start.
(Partially superseded by wind-down mode, which stops *new* work gracefully —
this item is about interrupting a download that is already streaming.)

### Multi-machine coordination
When running two machines concurrently against the same Backupify account,
they may attempt to export the same users. A simple coordination file (a shared
JSON on a network mount) listing claimed service IDs would prevent collisions
without requiring a network service.

### Export format selection
Backupify supports multiple export formats for some service types (e.g. PST vs
EML for Exchange). Expose a per-tab format selector in the TUI settings bar and
pass the chosen format to `trigger_export`.

### Per-item retry granularity
Failed exports are currently retried at the whole-mailbox level. For large
mailboxes that fail partway through it would be useful to retry only the failed
items within the export job.

## Done

### Adopt exports already on the server
Before triggering, each entity now checks the export-page scan (already done
once per run in preflight, so no extra requests) for an existing export with
its name: a completed one with a live Download link is downloaded directly; an
in-progress/queued one ("In Progress" instead of Download/Delete) is adopted
and polled instead of starting a duplicate job. Id-less in-progress rows are
adopted by name and completion is detected as the first *new* completed row
with that name. Adoption is skipped when two entities share a display name
(the export page identifies rows by name only). Expired exports don't count.

### Wind-down mode
`Ctrl+W` cycles an app-wide mode shown in the header: **off** → **no new jobs**
(in-flight entities finish completely, everything not yet started defers) →
**no new downloads** (exports still trigger and get polled so they're ready
server-side, but downloads defer) → **drain** (only currently-streaming
downloads finish; polling workers checkpoint and stop). Deferred entities keep
their `progress.json` `in_progress` record, so the next run resumes the same
server-side job — nothing is lost by quitting after the wind-down settles.

### Missing-job extension prompt
When a freshly triggered job still hasn't appeared on the export page after
`MISSING_FRESH_SCANS` (~60 min — heavy load can delay listing that long), the
app now asks per job whether to keep waiting another round or mark it failed,
instead of failing silently. Prompts are serialized (one dialog at a time).

### SharePoint export
Third `SERVICES` registry entry. Despite the old assumption that sites would need
a different endpoint, SharePoint speaks the exact same protocol as Exchange and
OneDrive: `customerServices` entity list (`appType=office365_sharepoint`, entities
are sites, `email` carries the site URL), the same `restoreExportAction` payload
(no `exportFormat` — verified by capturing a real trigger), and the same
`#exportListItems` export-status table (status col 7, actions col 8). Download is
a ZIP (`…&ext=zip`). One twist: queued/running rows show no `?id=` links and jobs
can queue for 10+ minutes, so the poll now also matches in-flight rows by Source
name (`poll_for_download_url(source_name=…)`) instead of failing them as missing.
Verified against the archived tenant on 2026-06-11 with the smallest site
(`Forms for Office 365 by Virto`, 36 MB).

### OneDrive export
Implemented as a second entry in the `SERVICES` registry (`ServiceDef` dataclass) —
same `customerServices` entity list and `restoreExportAction` trigger as Exchange,
with `appType=office365_onedrive` and no `exportFormat` field (the OneDrive modal
has no format radios). Each service tab has its own entity cache, export-status
cache, and `progress.json` under `<output_dir>/<service>/`. The download extension
comes from the server's `Content-Disposition` header, falling back to `.zip`.

### Actual size-based sorting
The `customerServices` response includes `ownSize` (bytes) and `usedBytes` (label);
both are now stored in the entity cache (`size_bytes` / `size_label`) and drive the
"Largest / Smallest first" sorts. Entity caches written by older versions lack the
field and fall back to the old `snapshot_id` proxy until refreshed with F5.
