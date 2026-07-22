# AP Automation Portfolio — Received Emails Dashboard

Google Apps Script web app that ingests every Gmail message addressed to
`ap-invoices@coreweave.com` / `apinvoices@coreweave.com` (or mentioning
those addresses anywhere in the headers) from **February 1, 2026** onward,
builds a searchable database in your Google Drive, and serves a dashboard
with filters, search, charts, and exports.

> Lives in a Google Apps Script project — nothing to run on a server. The
> data lives in a Google Sheet that the app creates for you on first run.

---

## What it does

- **Auto-creates a Google Spreadsheet** in your Drive with three sheets:
  - `Emails` — one row per matching message, with **every label and every
    header Gmail exposes** (From, To, CC, BCC, Reply-To, Subject, snippet,
    internal date, all `labelIds`, all human-readable labels, body plain
    text, body HTML, raw headers, attachment summary, plus flags for
    Unread / Starred / Important / Inbox / Trash / Spam / Draft / Sent).
  - `Documents` — one row per "billable" document. Zips are expanded so
    every inner file becomes its own row. Nested email attachments
    (`.eml` / `.msg`) are recursed into so the documents *inside* the
    nested email become the rows (the email itself is **not** counted).
    Images and videos are excluded as per spec.
  - `Settings` — runtime-overridable config (start date, target
    addresses, page size, time budget, etc.).
- **Resumable ingestion** that survives the Apps Script 6-minute
  execution limit. Each run processes a chunk, saves the cursor, and
  re-schedules itself via a time-based trigger.
- **Fallbacks everywhere**:
  - Gmail API transient failures (429 / 5xx) → exponential backoff.
  - 50,000-char Sheet cell limit → safe truncation with marker.
  - Oversize blobs (≥ 1 MB) → stashed in a `Received-Emails-Overflow-Blobs`
    Drive folder, with a Drive URL in the row.
  - Per-message parse errors → stub row so the loop never hangs.
- **Dashboard**:
  - KPI cards: total documents, total emails, ingestion state.
  - **Bar chart** — X = month, Y = count, **gray bars = documents, blue
    bars = emails**.
  - **Date range** filter.
  - **Search bar** that turns into filter chips. Press **Enter** to add
    a chip, **Backspace** on an empty input to remove the last chip.
    Tokens supported:
    - `Subject: <text>` · `From: <text>` · `To: <text>` · `CC: <text>` ·
      `BCC: <text>` · `Label: <text>` · `Body: <text>` · `Address: <text>`
    - `PDF: <name>` / `Doc: <name>` / `Filename: <name>` /
      `Attachment: <name>` — searches the Documents sheet
    - `Has: zip | nested | attachment | unread | important | starred |
      inbox | spam | trash | draft | sent`
    - `Month: YYYY-MM` · `Date: YYYY-MM-DD` or `YYYY-MM`
    - bare text → free-text across subject, from, to, snippet, labels,
      attachment names
  - **Download CSV** and **Download XLSX** of the currently visible
    (filtered) data only.
- **Safe console logs**: every meaningful step is logged. All messages
  pass through a `%`-escaper so a stray `%` in a Gmail header or error
  message cannot crash the V8 engine's `util.format`. `Logger.log` is
  used as a secondary sink that doesn't do format expansion.

---

## What was fixed in this version

1. **Stack-overflow recursion** — previous version had
   `loadConfigFromSheet_` → `getOrCreateSpreadsheet_` →
   `ensureSettingsSheet_` → `loadConfigFromSheet_` ... → stack overflow
   after ~10 min, reported as a misleading "Invalid regular expression"
   error. New design: `getOrCreateSpreadsheet_` uses hard-coded `CFG`
   only; a separate `loadConfigOverridesFromSheet_` is called once at
   the start of each ingestion / API call and never calls back into
   spreadsheet creation.
2. **V8 format-string crash** — `console.log` in Apps Script V8 does
   printf-style expansion (`%s`, `%d`, `%j`). A stray `%` in an error
   message or a Gmail header threw a secondary `SyntaxError: Invalid
   regular expression` that masked the real bug. `log_` now escapes
   every `%` to `%%` before printing, and also writes to `Logger.log`
   as a non-format-expanding sink.
3. **Running flag self-healing** — if the trigger somehow gets stuck,
   `apiGetStatus` auto-clears the flag after 30 min. If the user clicks
   Reset while a run is in progress, the next iteration of the loop
   sees the flag cleared and bails out cleanly.

---

## Setup (one-time, ~5 minutes)

1. **Create the Apps Script project**
   - Go to https://script.google.com → **New project**.
   - Delete the default `Code.gs` content.
   - Create files matching this repo and paste contents in:
     - `Code.gs` · `Config.gs` · `Utils.gs` · `SheetManager.gs` ·
       `GmailIngestion.gs` · `DashboardAPI.gs` · `Dashboard.html` ·
       `Stylesheet.html` · `JavaScript.html` · `appsscript.json`
   - For `appsscript.json`: ⚙️ Project settings →
     ☑ "Show appsscript.json manifest file in editor", then paste.
2. **Authorize scopes**
   - Run the `install` function once. Accept the OAuth prompts
     (Gmail read/modify, Drive, Sheets, script management).
   - The first run creates your spreadsheet.
3. **Deploy as a Web App**
   - **Deploy → New deployment → Web app**.
   - **Execute as**: *Me*. **Who has access**: *Only myself*.
   - Copy the **Web app URL** — that's your dashboard.
4. **(Optional) Tweak the `Settings` sheet**
   - Open the auto-created spreadsheet. The `Settings` tab lets you
     change start date, target addresses, page size, time budget, etc.
     without redeploying.

---

## File map

| File | Role |
|------|------|
| `appsscript.json` | Manifest — scopes + time zone + web app config |
| `Code.gs` | `doGet` entry, install/menu helpers |
| `Config.gs` | All configurable defaults + `loadConfigOverridesFromSheet_()` override |
| `Utils.gs` | Logging, truncation, hashing, MIME / extension classifiers, time budget |
| `SheetManager.gs` | Auto-create spreadsheet + sheets, batched writes, dedupe check |
| `GmailIngestion.gs` | The actual engine — pagination, parsers, zip / .eml expanders, oversize Drive stashing |
| `DashboardAPI.gs` | Server-side functions called by the HTML dashboard: `apiGetStatus`, `apiStartIngestion`, `apiSearch`, `apiExportCsv`, `apiExportXlsx` |
| `Dashboard.html` | The web app shell (cards, chart, search, tables) |
| `Stylesheet.html` | All CSS (dark theme) |
| `JavaScript.html` | All client-side logic (filters, chart rendering, exports) |

---

## Known limits & gotchas

- **50,000 chars / cell**: bodies and headers are truncated with a
  marker so nothing is silently dropped. For truly huge bodies the
  full content is moved to the overflow folder and only a pointer
  remains in the sheet.
- **500 messages / page**: hard Gmail limit; we default to 100.
- **6-minute execution ceiling**: the time budget is 300s, the resumable
  trigger fires after `RESCHEDULE_MINUTES` (default 1 min) and picks up
  where the previous run stopped.
- **Labels**: human-readable names of custom labels (`Label_123`) are
  fetched lazily on first encounter and cached in-memory for the
  duration of the run.
- **Zips nested more than 2 levels deep** are not expanded. Spec was
  silent on this; 2 levels is plenty for invoice workflows.
- **Export** uses the standard "temp sheet → blob → trash" trick so
  you get a real `.xlsx` (not a CSV with an XLSX extension).

---

## License

MIT — do whatever you want, but no warranty.
