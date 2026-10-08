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

**Your data lives outside this folder** — in `Documents/RoofCRM` (the startup
window prints the path), so you can replace or delete the app folder when
updating without losing anything. Data left in the app folder by older
versions is moved there automatically on first run. To use a different
location, set the `ROOF_CRM_DATA` environment variable.

The database (`crm.db`) is created automatically on first run, and a dated
backup copy is saved once per day on startup (newest 14 kept). The app runs
on the production-grade `waitress` server.

### Updating

**It updates itself.** `start.bat` / `./start.sh` runs `updater.py` before
starting. That script asks GitHub for the newest commit on `main`, and if this
folder isn't on it, downloads the branch zip (about 300 KB) and copies over
only the files that changed. It needs no Git. Offline or a failed download
means the app starts the version you have, unchanged. Every file is checked
before anything is written, and files that aren't part of the app are left
alone. `update.bat` / `./update.sh` does the same without starting. A folder
that is a `git clone` updates with `git pull --ff-only` instead. Set
`ROOF_CRM_NO_AUTO_UPDATE=1` to skip the check on start. Your data in
`Documents/RoofCRM` is never touched. Hit **Download Full Backup** on the
Import page first if you want a guaranteed rollback point.

The start and update scripts wrap their commands in a single `( )` / `{ }`
block, because cmd and bash read a script while running it and an update
that rewrites the running script would otherwise derail it.

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

- **Daily Dashboard** — a "16 of 24 done today" progress bar, due cadence
  reminders (pinned until logged or checked off), due follow-ups with snooze
  buttons, stale-deal alerts, active-prospect stats, and recent activity.
- **Work the Queue** — one-task-at-a-time focus mode: each due cadence step or
  follow-up appears with the account's info and the personalized script on one
  screen; log it and the next task loads. Turns a call block into a flow.
  Keyboard shortcuts (`L` log, `S` skip, `N`/`P` next and previous, `?` help)
  keep a calling block moving without the mouse. On a call step the outcome
  buttons sit right under the account, numbered `1`–`7`.
- **Call outcomes** — logging a call records how it went in one tap: No
  answer, Voicemail, Gatekeeper, Spoke, Meeting booked, Not interested, Bad
  number. *Meeting booked* moves the deal to Accepted Meeting and *Not
  interested* sets the status (both end the cadence); *Bad number* sends the
  account back to Research and leaves that call to do again once the number
  is fixed. Every one-tap log is undoable. Insights turns the outcomes into
  calls made, connect rate and meetings per connect.
- **Research list** — an account only starts its cadence once it has someone
  to contact: a named person with an email or phone. Until then it waits in
  Research (no clock, no overdue tasks), sorted biggest-portfolio-first as the
  list of who to look up in ZoomInfo next. Adding a contact — by hand, by
  pasting a ZoomInfo page or roster, or by a ZoomInfo upload — starts the
  cadence that day. *Start the cadence anyway* is there for cold-calling a
  switchboard before you have a name.
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
  links on mobile. Search covers the company name, **every contact on the
  account**, titles, management level, notes and phone numbers in any
  formatting — and says which person a result matched on.
- **Bulk actions** — tick any number of accounts (or the whole filtered
  list) and set a status or milestone, set/clear a follow-up, restart the
  cadence, or archive them in one click. Restore and permanent delete on
  the Archived view — only archived accounts can be deleted. Every bulk action is undoable.
- **Undo** — deletes and bulk changes put an ↶ Undo bar at the top of the
  page for 15 minutes. Undoing an account delete brings back its contacts, history, bids,
  projects, invoices **and photo files** (deleted photos wait in a trash
  folder for 7 days). Logged interactions can also be edited — type, notes
  and date — or deleted; correcting a date corrects the cadence with it.
- **Priority** — the criteria-matching building count drives what you work
  first: the dashboard, the queue and the accounts list all lead with the
  biggest portfolios, and a 🎯 Priority / 📅 Due date toggle switches the
  ordering.
- **Account Detail** — edit everything in place, log interactions, see a
  chronological history timeline and live cadence progress.
- **Contacts & ZoomInfo** — each account holds a primary contact plus
  unlimited additional people. Add them on the account detail page, promote
  anyone to primary with ★, and use the 🔎 ZoomInfo / LinkedIn buttons to
  jump straight to research for that company.
  - **Paste a whole ZoomInfo page** — on the contact's page press Ctrl+A,
    Ctrl+C, paste into the box and press **Read the paste**. Name, title,
    company, email, direct line, mobile and LinkedIn fill the form so you can
    check them before saving. Menus, buttons and the colleagues listed further
    down (Similar Contacts, Org Chart) are discarded, and the contact is found
    by anchoring on the company link in the page header — so other people on
    the same page can't get mixed in, and it works on pages with no revealed
    email. Markdown links from the browser copy are unpacked. The **Contact Details**
    panel's `(B)` / `(HQ)` / `(D)` / `(M)` tags decide which number is the
    mobile and which is the work line. Smaller selections, "Label: value"
    lists and spreadsheet rows work too; a paste with no contact in it is
    refused rather than guessed at.
  - **Several people in one paste** — on ZoomInfo's company *Employees* tab,
    expand the few rows you want (the ▾ on each person reveals their email and
    phones inline), then Ctrl+A, Ctrl+C the page once and press **Find
    people**. Everyone is listed with title, management level and — for the
    rows you expanded — email, work phone and mobile. Nothing is pre-selected;
    tick your three or four, optionally mark one primary, press Add Selected.
    They're created complete without opening a single profile. *Tick
    decision-makers* selects every C-Level, VP-Level and Director at once, a
    direct dial beats the company switchboard, and people already on the
    account are greyed out.
  - **Batch loop (the main way)** — the Research list's **⬇ Download for
    ZoomInfo** button saves those companies (name, website, address; current
    search and 🎯 filter applied) as a CSV. Upload it to ZoomInfo as a
    company list, search contacts at those companies filtered to your
    management levels, export them, and upload the export on the Import page.
    People are matched to accounts by company name (legal suffixes like LLC,
    Inc, LP, Corp, P.C. and punctuation ignored), then by **web domain**:
    Email Domain/Website against each account's website (never the emails
    of people on it, since a management firm's domain would pull in its own
    staff). Every domain match is listed back, a domain shared by two
    accounts matches neither, and free mail domains never count. The first
    person on an empty account becomes primary, duplicates are skipped, and
    archived companies are left alone. Accounts leaving Research start their
    cadence today, or **N per business day** when you set a pace. The whole
    upload is **undoable**. Tick *"Open a new account for any company I don't
    have yet"* and unmatched companies become new prospects.
  - A stock ZoomInfo export maps with no editing, including **LinkedIn
    Contact Profile URL**, **Management Level**, Website, Email Domain and
    Company HQ Phone (the dial line for people with no direct number).
    Management level is guessed from the job title when the file doesn't
    carry it. Extensions like `ext. 32` dial as a pause plus the extension.
    The fixture `tests/fixtures/zoominfo_contacts_export.csv` is a real
    export.
  - Accounts have a **Website** field. Account imports fill it from a
    Website column, and older databases lift it out of the "Website:" line
    in Notes.
- **Which app the buttons open** — Templates → Your Info picks the call app
  (the phone's own dialer, or **Google Voice**, which opens voice.google.com
  ready to call or text from your Google Voice number) and the email app
  (default mail app, **Gmail** or **Outlook** web compose with the subject and
  body filled in). Nothing is ever sent automatically. Accounts with no email
  say so instead of hiding the button.
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
  Company names are matched on their meaning, not their spelling — "Boxer
  Property Corp" and "Boxer Property, Corp." are one company — so a re-pull
  from CoStar can't create a second account or un-archive one you removed.
  Rows with a contact go into the cadence; rows without one go to Research.
- **Merge duplicates** — duplicates left by older versions are listed at the
  bottom of the Import page. Pick the one to keep (the most-worked is
  pre-selected) and press Merge: people, history, roof reports, projects and
  checked-off steps all move onto it, gaps in its details are filled, the
  duplicate's notes are appended, and the duplicates are removed. Fully
  undoable — photos and invoices included. An archived duplicate never
  blocks imports for the live company.
  Imported accounts default to **Prospecting / None / In Cadence** and enter
  the cadence immediately — or use the **pacing option** to stagger starts
  (about a fifth of your daily touch capacity, e.g. 10 a business day — each
  account brings five touches over ten business days, so starts stack up) so a big list becomes a steady daily
  routine instead of hundreds of Day-1 tasks at once. Forgot to pace? The
  **Re-Pace Cadence** tool on the Import page re-staggers all untouched
  in-cadence accounts after the fact.
- **Archive** — remove a company from your working list without losing it:
  archived accounts keep all their history, disappear from every view, and
  are **skipped by future imports** — however the next list spells the name —
  so re-uploading the same CoStar list can't resurrect them. Restore any time from Accounts → Archived.
  (Permanent delete also exists, but it forgets the company entirely, so a
  later import can add it back.)
- **Insights** — conversion, pipeline and money, plus a **daily activity
  chart**: the last 14 days against an editable daily goal, with the bars that
  hit it in green. Tells you whether a slow month is an activity problem or a
  conversion problem.
- **Export Everything (Excel)** — one workbook with a summary tab plus
  accounts, contacts, full interaction history, roof reports (with computed
  application rates and suggested prices), projects, invoices and today's
  tasks. Every sheet is frozen and filterable. Also plain CSV exports of
  accounts and interaction history.

## Speed

Pages are bounded so the app stays quick however big the list gets: the
dashboard draws the top 25 due reminders (with the true total and a link to
see them all), the accounts list pages at 100, and an account's history shows
the 25 most recent entries. With 400 accounts and 773 due reminders this takes
the dashboard from 1.3 MB and ~950 ms of browser load down to 52 KB and ~50 ms.

## Cadence Logic

Five steps over ten **business** days, computed from each account's cadence
start date. Day 1 is the start date itself; every later step counts working
days, so nothing ever falls on a weekend and Monday doesn't inherit Saturday's
and Sunday's work:

| Day | Step          | e.g. started Friday |
|-----|---------------|---------------------|
| 1   | Email 1       | Friday              |
| 3   | Call & Text   | Tuesday             |
| 6   | Call 2        | Friday              |
| 8   | Email 2       | Tuesday             |
| 10  | Breakup Email | Thursday            |

Rules:

- Cadence applies **only** while Prospecting Status = "Prospecting" **and**
  Pipeline Milestone = "None / In Cadence" — and only once the clock has
  started. An account with nobody to contact waits in **Research** with no
  clock; adding a contact starts it that day. When this version first runs it
  moves untouched accounts with nobody to contact into Research (anything you
  have already worked keeps its dates) and says how many in the startup
  window.
- Moving an account to any other status or milestone stops the cadence and
  clears its reminders **instantly** (reminders are computed live, never stored
  stale).
- A due reminder stays pinned to the dashboard until you either log that
  interaction (✓ Log) or check it off manually (✕ Skip).
- **One step at a time.** The cadence is a sequence, so an account appears
  once, at the step it is actually waiting on — never five times. If it has
  fallen behind, the row says "+N more steps behind"; logging or skipping the
  current step brings the next one forward.
- Importing a big list without pacing puts every account on the same clock,
  so they all come due together. Use **Re-Pace Cadence** on the Import page
  to stagger untouched accounts — pick about a fifth of the touches you can do in a day — each account brings five touches over ten business days, so at 10 a day you settle at exactly 50 tasks a day. It never moves
  anyone you have already contacted.
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

Note: bid photos live in `Documents/RoofCRM/uploads/` — include that folder
when copying your data to a new machine. Deleted photos wait in
`uploads/trash/` for 7 days so an Undo can put them back, then they're
cleared automatically on startup.

## When a page errors

A crash shows the actual error on the page, with a **Copy details** button,
and appends the full traceback to `Documents/RoofCRM/error.log` (the startup
window prints the path). The rest of the app keeps working and your data is
untouched. Send that text on and it says exactly which line failed.

## Tests

The end-to-end test suite lives in `tests/test_crm.py` and covers imports,
cadence rules, follow-ups, the queue, templates, insights, exports, and
backups against a throwaway database (your `crm.db` is never touched):

```bash
python3 tests/test_crm.py
```

650 checks, about two seconds.
