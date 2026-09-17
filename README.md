# Roof CRM — Lightweight Local CRM for Roof Coating Sales

A fast, single-file-database CRM you run on your own computer and use from any
device on your Wi-Fi (desktop, phone, tablet). Built with Flask + SQLite +
Bootstrap. No cloud, no accounts, no monthly fee — your data lives in `crm.db`
next to the app.

## Quick Start

```bash
# 1. Install dependencies (one time)
pip install -r requirements.txt

# 2. Run the app
python3 app.py
```

The database (`crm.db`) is created automatically on first run.

## Access From Your Phone / Tablet

The server binds to `0.0.0.0:8000`, so any device on the **same Wi-Fi network**
can reach it. When the app starts it prints your computer's local IP and the
exact URL to type into your phone's browser, e.g.:

```
On your phone/tablet:  http://192.168.1.23:8000
```

If you need to find the IP yourself:

- **Windows:** open Command Prompt → `ipconfig` → look for "IPv4 Address"
- **Mac:** System Settings → Wi-Fi → Details, or `ipconfig getifaddr en0`
- **Linux:** `hostname -I`

If the page won't load from your phone, allow inbound port 8000 through your
computer's firewall (Windows will usually prompt you the first time — choose
"Allow" for private networks).

## Features

- **Daily Dashboard** — due cadence reminders (pinned until logged or checked
  off), active-prospect stats, and recent activity.
- **Accounts** — searchable list with filters for Prospecting Status and
  Pipeline Milestone; tap-to-call / tap-to-text / tap-to-email links on mobile.
- **Account Detail** — edit everything in place, log interactions, see a
  chronological history timeline and live cadence progress.
- **Import** — upload `.xlsx` or `.csv` prospecting lists (e.g.
  `HTX_Office_5kto10k.xlsx`). Column headers are matched automatically and
  duplicates (same company name) are skipped, so re-uploads are safe.
  Imported accounts default to **Prospecting / None / In Cadence** and enter
  the cadence immediately.

## Cadence Logic

Standard 10-day cadence, computed from each account's cadence start date:

| Day | Step          |
|-----|---------------|
| 1   | Email 1       |
| 3   | Call & Text   |
| 6   | Call 2        |
| 8   | Email 2       |
| 10  | Breakup Email |

Rules:

- Cadence applies **only** while Prospecting Status = "Prospecting" **and**
  Pipeline Milestone = "None / In Cadence".
- Moving an account to any other status or milestone stops the cadence and
  clears its reminders **instantly** (reminders are computed live, never stored
  stale).
- A due reminder stays pinned to the dashboard until you either log that
  interaction (✓ Log) or check it off manually (✕ Skip).
- "Restart Cadence" on an account's detail page resets the clock to today —
  useful for re-engaging a cold prospect.

## Backup

Your entire CRM is the single file `crm.db`. Copy it anywhere (USB drive,
cloud folder) to back up; restore by copying it back.
