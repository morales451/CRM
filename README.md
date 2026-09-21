# Roof CRM — Lightweight Local CRM for Roof Coating Sales

A fast, single-file-database CRM you run on your own computer and use from any
device on your Wi-Fi (desktop, phone, tablet). Built with Flask + SQLite +
Bootstrap. No cloud, no accounts, no monthly fee — your data lives in `crm.db`
next to the app.

## Quick Start

Double-click **`start.bat`** (Windows) or run **`./start.sh`** (Mac/Linux) —
it installs dependencies on first run and starts the app. Or manually:

```bash
pip install -r requirements.txt
python3 app.py
```

The database (`crm.db`) is created automatically on first run, and a dated
backup copy is saved to `backups/` once per day on startup (newest 14 kept).
The app runs on the production-grade `waitress` server.

**Phone tip:** open the app in your phone's browser and use "Add to Home
Screen" — it installs like an app with its own icon and opens full-screen.

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
  off), due follow-ups with snooze buttons, stale-deal alerts, active-prospect
  stats, and recent activity.
- **Work the Queue** — one-task-at-a-time focus mode: each due cadence step or
  follow-up appears with the account's info and the personalized script on one
  screen; log it and the next task loads. Turns a call block into a flow.
- **Follow-Ups** — set a "next follow-up" date + note on any account (quick
  buttons: tomorrow / +3 days / +1 week). Due follow-ups pin to the dashboard
  and queue until done or snoozed — so warm prospects outside the cold cadence
  never slip. Accounts in the pipeline with no activity for 14+ days and no
  follow-up set are flagged as **going stale**.
- **Pipeline board** — a column per milestone with each deal as a card; move
  deals between milestones right from the board.
- **Projects & Invoicing** — life after Closed Won. Each won deal gets a
  project with a standard job checklist (contract → deposit → materials →
  crew → completion → final payment → warranty; add your own steps),
  contract amount, start/completion dates, and invoices tracked
  Draft → Sent → Paid. "Mark Sent" stamps the date and defaults the due
  date to 30 days out; sent invoices pin to the Dashboard until paid, with
  overdue ones flagged red. The Projects page and Insights show the money
  rollup: contracted, invoiced, collected, outstanding.
- **Accounts** — searchable list with filters for Prospecting Status,
  Pipeline Milestone and minimum matching buildings (🎯), sorted by
  opportunity size by default; tap-to-call / tap-to-text / tap-to-email
  links on mobile.
- **Priority** — the criteria-matching building count drives what you work
  first: the dashboard, the queue and the accounts list all lead with the
  biggest portfolios, and a 🎯 Priority / 📅 Due date toggle switches the
  ordering.
- **Account Detail** — edit everything in place, log interactions, see a
  chronological history timeline and live cadence progress.
- **Contacts** — each account holds a primary contact plus unlimited
  additional people (from ZoomInfo research). Add them on the account detail
  page, promote anyone to primary with ★, and use the 🔎 ZoomInfo / LinkedIn
  buttons to jump straight to research for that company. You can also
  bulk-upload a ZoomInfo contact export on the Import page — people are
  matched to accounts by company name (LLC/Inc suffixes ignored), the first
  person on an empty account becomes its primary contact, and duplicates are
  skipped.
- **Outreach Templates & Scripts** — your cold call script, voicemail,
  text message, and three cadence emails live in the app (Templates page)
  and are fully editable, with placeholders like `{first_name}`, `{company}`,
  and `{matching_properties}` (the buildings that fit your criteria). Every dashboard reminder has a **📄 Script** button
  that opens the right script personalized for that account: one-tap
  **Open in Email app** (subject and body pre-filled), **Dial**, or
  **Open in Messages**, plus Copy buttons and a **✓ Log** button to record
  the step when done. Missing info (no first name yet, etc.) is flagged in
  [brackets] so nothing goes out half-baked. Set your own name, company, and
  phone once on the Templates page and they fill into every script.
- **Warranty calculator & pricing** — application rates ported from the
  Warranty Roofing Calculator: Silicone / Acrylic (Standard + Reinforced) /
  Aluminum across capsheet, single-ply, sprayfoam and metal at 10/15/20
  years. Only combinations that actually exist can be selected. Suggested
  pricing is $4.50/sq ft for capsheet and $4.00 elsewhere on a 10-year
  system, +$0.15 at 15 years and +$0.10 more at 20 — all editable on the
  Templates page.
- **Roof Reports / Bid Generator** — create a full branded restoration
  proposal from any account page: cover, CEO letter (personalized to the
  contact), process overview, roof facts, site assessment with an uploaded
  photo survey (photos auto-resize and auto-rotate; captions print beside
  them), the coating spec, and a materials quote computed by the built-in
  warranty calculator — real Henry Prograde / Enduraroof application rates
  for the chosen system, roof type and warranty length, plus primers,
  mastic and fabric, rounded to 5-gallon pails. Also quotes a suggested
  price per square foot, then prints the figure in words, the warranties,
  and a signature page. Print or save as PDF from the browser.
- **Import** — upload `.xlsx` or `.csv` prospecting lists (e.g.
  `HTX_Office_5kto10k.xlsx`). Column headers are matched automatically and
  duplicates (same company name) are skipped, so re-uploads are safe.
  Imported accounts default to **Prospecting / None / In Cadence** and enter
  the cadence immediately — or use the **pacing option** to stagger starts
  (e.g. 25 accounts per business day) so a big list becomes a steady daily
  routine instead of hundreds of Day-1 tasks at once. Forgot to pace? The
  **Re-Pace Cadence** tool on the Import page re-staggers all untouched
  in-cadence accounts after the fact.
- **Archive** — remove a company from your working list without losing it:
  archived accounts keep all their history, disappear from every view, and
  are **skipped by future imports**, so re-uploading the same CoStar list
  can't resurrect them. Restore any time from Accounts → Archived.
  (Permanent delete also exists, but it forgets the company entirely, so a
  later import can add it back.)
- **Export** — download accounts and interaction history as CSV from the
  Import page any time.

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

Your entire CRM is the single file `crm.db`. A dated copy lands in
`backups/` automatically each day (newest 14 kept). Set an **off-machine
backup folder** on the Import page — point it at OneDrive/Google
Drive/Dropbox and each daily backup is mirrored there, so a dead computer
can't take your data with it. A **Download Full Backup** button (also on the
Import page) grabs a consistent snapshot from any device. Restore by copying
a backup file back to `crm.db`. CSV exports are on the Import page too.

Note: bid photos live in `uploads/` next to the app — include that folder
when copying the app to a new machine.

## Tests

The end-to-end test suite lives in `tests/test_crm.py` and covers imports,
cadence rules, follow-ups, the queue, templates, insights, exports, and
backups against a throwaway database (your `crm.db` is never touched):

```bash
python3 tests/test_crm.py
```
