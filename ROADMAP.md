# Backupify Exporter — Roadmap

## Planned features

### Snapshot selection
Currently the exporter always uses the most recent snapshot for each user/service.
A future version should let you browse a user's snapshot history and choose which
point-in-time to export from (e.g. "give me the mailbox as it was 3 months ago").
The `perfectBackups` array returned by the `customerServices` API already contains
all available snapshot IDs — the UI just needs to expose them per-user.

### OneDrive export
Each M365 user also has a OneDrive drive. Backupify exposes these under a separate
service type. The user list for OneDrive is the same set of people as Exchange but
the export format and endpoint differ. Implementation should reuse the session,
cookie, and user-loading machinery already in place.

### SharePoint export
SharePoint entities are *sites*, not users — the item list comes from a different
API endpoint. The export format is a ZIP archive of document libraries. The tab
structure in the TUI already reserves a slot for this.

### Teams export
Teams entities are *teams* (and optionally channels). Like SharePoint these are
not user mailboxes. Exports include chat history and files. Reserved in the TUI.

### Actual size-based sorting
The current "Largest / Smallest first" sort uses `snapshot_id` (an epoch-ms
timestamp) as a proxy, which is not the real mailbox size. The `customerServices`
API response likely includes an item count or size field — extract it and store it
in the user cache so that size sorting is accurate.

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
