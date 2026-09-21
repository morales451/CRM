"""SQLite database layer for the Roof Coating CRM.

All timestamps are strict ISO-8601 with local timezone offset,
e.g. 2026-09-17T14:03:22-05:00.
"""

import shutil
import sqlite3
from datetime import datetime, date
from pathlib import Path

DB_PATH = Path(__file__).parent / "crm.db"
BACKUP_DIR = Path(__file__).parent / "backups"

PREFERRED_CONTACT_METHODS = ["Email", "Call", "Text", "Unknown"]

PROSPECTING_STATUSES = [
    "Prospecting",
    "Interested - Continue Conversation",
    "Might be Interested",
    "Not Interested",
]

PIPELINE_MILESTONES = [
    "None / In Cadence",
    "Accepted Meeting",
    "Walked Roof",
    "Created Report/Bid",
    "Presented Report/Bid",
    "Resolving Objections",
    "Closed Won",
    "Closed Lost",
    "On Hold",
]

INTERACTION_TYPES = [
    "Email 1",
    "Call & Text",
    "Call 2",
    "Email 2",
    "Breakup Email",
    "General Note",
    "Meeting",
]

PROJECT_STATUSES = [
    "Not Started",
    "Scheduled",
    "In Progress",
    "Completed",
    "Closed",
]

INVOICE_STATUSES = ["Draft", "Sent", "Paid"]

# Checklist seeded onto every new project — the standard steps of a
# roof-coating job from contract to warranty. Fully editable per project.
DEFAULT_PROJECT_TASKS = [
    "Contract signed",
    "Deposit invoiced",
    "Deposit received",
    "Materials ordered",
    "Crew scheduled",
    "Job started",
    "Job completed",
    "Final walkthrough with owner",
    "Final invoice sent",
    "Final payment received",
    "Warranty documents delivered",
]

SCHEMA = """
CREATE TABLE IF NOT EXISTS accounts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    company_name TEXT NOT NULL,
    first_name TEXT DEFAULT '',
    last_name TEXT DEFAULT '',
    title TEXT DEFAULT '',
    num_properties INTEGER,               -- total properties owned
    matching_properties INTEGER,          -- how many fit the sales criteria (pre-1980s)
    email TEXT DEFAULT '',
    work_phone TEXT DEFAULT '',
    mobile_phone TEXT DEFAULT '',
    preferred_contact TEXT NOT NULL DEFAULT 'Unknown',
    notes TEXT DEFAULT '',
    prospecting_status TEXT NOT NULL DEFAULT 'Prospecting',
    pipeline_milestone TEXT NOT NULL DEFAULT 'None / In Cadence',
    cadence_start TEXT NOT NULL,          -- ISO date the cadence clock starts from
    next_follow_up TEXT DEFAULT '',       -- ISO date; pins to dashboard when due
    follow_up_note TEXT DEFAULT '',
    created_at TEXT NOT NULL,             -- ISO timestamp
    updated_at TEXT NOT NULL              -- ISO timestamp
);

-- Additional people at an account. The account's own first/last/email/phone
-- fields hold the PRIMARY contact; these rows are the rest of the org chart.
CREATE TABLE IF NOT EXISTS contacts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    account_id INTEGER NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
    first_name TEXT DEFAULT '',
    last_name TEXT DEFAULT '',
    title TEXT DEFAULT '',
    email TEXT DEFAULT '',
    work_phone TEXT DEFAULT '',
    mobile_phone TEXT DEFAULT '',
    created_at TEXT NOT NULL              -- ISO timestamp
);

CREATE TABLE IF NOT EXISTS interactions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    account_id INTEGER NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
    interaction_type TEXT NOT NULL,
    notes TEXT DEFAULT '',
    created_at TEXT NOT NULL              -- ISO timestamp
);

-- Editable outreach scripts (seeded from the user's cold-call/email doc).
-- steps = comma-separated cadence step names this template applies to.
CREATE TABLE IF NOT EXISTS templates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    kind TEXT NOT NULL DEFAULT 'email',   -- email | call | text
    steps TEXT DEFAULT '',
    subject TEXT DEFAULT '',
    body TEXT NOT NULL DEFAULT '',
    sort_order INTEGER DEFAULT 0,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL DEFAULT ''
);

-- Roof report / bid per account: feeds the printable proposal generator.
CREATE TABLE IF NOT EXISTS bids (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    account_id INTEGER NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
    roof_address TEXT DEFAULT '',
    roof_size_sqft INTEGER,
    deduction_sqft INTEGER DEFAULT 0,     -- non-coated area (e.g. skylights)
    surface_type TEXT DEFAULT '',
    candidate TEXT DEFAULT 'Yes',
    warranty_years INTEGER DEFAULT 10,
    coating_system TEXT DEFAULT 'Silicone',      -- Silicone | Acrylic | Aluminum
    acrylic_system_type TEXT DEFAULT 'Standard', -- Standard | Reinforced
    roof_type TEXT DEFAULT 'Capsheet',           -- calculator roof category
    linear_feet INTEGER DEFAULT 0,               -- seams/penetrations for mastic
    waste_pct REAL DEFAULT 5,
    stretch_pct REAL DEFAULT 0,
    passed_adhesion INTEGER DEFAULT 1,
    has_rust INTEGER DEFAULT 0,
    rust_prime_method TEXT DEFAULT 'field',
    price REAL,
    assessment_date TEXT DEFAULT '',      -- ISO date
    assessment_notes TEXT DEFAULT '',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS bid_photos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bid_id INTEGER NOT NULL REFERENCES bids(id) ON DELETE CASCADE,
    filename TEXT NOT NULL,
    caption TEXT DEFAULT '',
    sort_order INTEGER DEFAULT 0,
    created_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_bids_account ON bids(account_id);
CREATE INDEX IF NOT EXISTS idx_bid_photos_bid ON bid_photos(bid_id);

-- Post-sale: one project per won deal, with a task checklist and invoices.
CREATE TABLE IF NOT EXISTS projects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    account_id INTEGER NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'Not Started',
    contract_amount REAL,
    start_date TEXT DEFAULT '',           -- ISO date
    completion_date TEXT DEFAULT '',      -- ISO date
    notes TEXT DEFAULT '',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS project_tasks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
    title TEXT NOT NULL,
    done INTEGER NOT NULL DEFAULT 0,
    done_at TEXT DEFAULT '',              -- ISO timestamp when checked
    sort_order INTEGER DEFAULT 0,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS invoices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
    invoice_number TEXT DEFAULT '',
    amount REAL NOT NULL DEFAULT 0,
    status TEXT NOT NULL DEFAULT 'Draft', -- Draft | Sent | Paid
    sent_date TEXT DEFAULT '',            -- ISO date
    due_date TEXT DEFAULT '',             -- ISO date; Sent + past due = overdue
    paid_date TEXT DEFAULT '',            -- ISO date
    notes TEXT DEFAULT '',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_projects_account ON projects(account_id);
CREATE INDEX IF NOT EXISTS idx_tasks_project ON project_tasks(project_id);
CREATE INDEX IF NOT EXISTS idx_invoices_project ON invoices(project_id);

-- A manually checked-off cadence step ("clear the task without logging it").
CREATE TABLE IF NOT EXISTS cadence_dismissals (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    account_id INTEGER NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
    step_type TEXT NOT NULL,
    dismissed_at TEXT NOT NULL,           -- ISO timestamp
    UNIQUE (account_id, step_type)
);

CREATE INDEX IF NOT EXISTS idx_interactions_account ON interactions(account_id);
CREATE INDEX IF NOT EXISTS idx_contacts_account ON contacts(account_id);
CREATE INDEX IF NOT EXISTS idx_accounts_status ON accounts(prospecting_status, pipeline_milestone);
"""


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def today_iso() -> str:
    return date.today().isoformat()


def get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    # WAL lets reads and writes overlap (phone + laptop at once) without
    # "database is locked" errors; busy_timeout covers the rest.
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA busy_timeout = 5000")
    return conn


# Default settings and outreach templates, seeded on first run.
# Placeholders: {first_name} {last_name} {title} {company} {num_properties}
#               {matching_properties} {my_name} {my_company} {my_phone}
DEFAULT_SETTINGS = {
    "my_name": "Alexis Morales",
    "my_title": "President & CEO",
    "my_company": "Silicone Roof Pros, Inc.",
    "my_phone": "(832) 303-3183",
    "my_email": "sales@siliconeroofpros.com",
    "my_website": "siliconeroofpros.com",
    "my_address": "",
    # Shown at the bottom of printed invoices; drawn from the standard
    # SRP contract payment terms. Editable on the Templates page.
    "invoice_terms": (
        "Payment accepted by check, bank transfer, or credit/debit card "
        "(a 3% processing fee applies to credit and debit card transactions). "
        "Invoices not paid within the specified term are subject to a late fee "
        "of 1.5% per month (18% per annum) on the outstanding balance. "
        "Upon receipt of payment, a conditional lien waiver will be provided; "
        "a final unconditional lien waiver follows final payment."),
}

SEED_TEMPLATES = [
    {
        "name": "Cold Call Script", "kind": "call",
        "steps": "Call & Text,Call 2", "subject": "", "sort_order": 1,
        "body": """{first_name}?

This is {my_name} with {my_company}.

(Pause. Wait for response)

I'll be honest, this is a cold call, but it's a well-researched one. Would you be open to giving me 30 seconds to explain why I'm reaching out?

(Pause. Wait for response)

I noticed you're the {title} at {company}, and that you manage several properties built before the 1980s. When commercial buildings hit that age, owners are usually staring down a massive, highly disruptive full roof replacement.

We help businesses bypass that entirely. We restore aging roofs with a silicone coating system that cuts the cost of a full replacement by 50% and keeps the building fully operational while we work.

Would you be open to learning how this might work for your properties?""",
    },
    {
        "name": "Voicemail Script", "kind": "call",
        "steps": "Call & Text,Call 2", "subject": "", "sort_order": 2,
        "body": """Hi {first_name}, this is {my_name} with {my_company}.

I'm calling because I noticed {company} manages several properties built before the 1980s. Usually, that means you're bracing for a massive, disruptive roof replacement.

We help businesses bypass that entirely with a silicone system that cuts costs by 50% and keeps the building operational.

I'll email you my contact info so you can easily reply, but if you want to chat, my number is {my_phone}.

Again, {my_name} at {my_phone}. Thanks.""",
    },
    {
        "name": "Text Message", "kind": "text",
        "steps": "Call & Text", "subject": "", "sort_order": 3,
        "body": "Hi {first_name}, this is {my_name} with {my_company} — just left you a "
                "voicemail. We restore aging commercial roofs for about 50% less than a "
                "full replacement, with no tear-off. Worth a quick chat about {company}'s "
                "properties?",
    },
    {
        "name": "Email 1", "kind": "email",
        "steps": "Email 1", "sort_order": 4,
        "subject": "{company}'s {matching_properties} older properties",
        "body": """Hi {first_name},

I noticed you're the {title} at {company}, and it looks like you manage roughly {matching_properties} properties built before the 1980s.

When commercial buildings hit that age, owners are usually staring down a massive, highly disruptive full roof replacement.

We help businesses bypass that process entirely. We restore aging roofs using a commercial silicone system that:

- Cuts the cost of a full replacement by up to 50%
- Eliminates the need for an expensive tear-off
- Keeps the building fully operational while we work

Would you be open to learning how this might work for your portfolio?

Best,
{my_name}""",
    },
    {
        "name": "Email 2", "kind": "email",
        "steps": "Email 2", "sort_order": 5,
        "subject": "Extending the life of {company}'s roofs",
        "body": """Hi {first_name},

I know how busy things get, so I'll keep this brief.

When managers at companies like {company} evaluate older roofs, the biggest headache often isn't just the replacement cost, it's the operational downtime and tenant disruption of a full tear-off.

Because our fluid-applied silicone systems skip the tear-off entirely, we can safely extend the life of those aging roofs by 10, 15, or even 20 years with zero disruption to the businesses inside.

Is extending the life of these older assets a priority for you this year, or is this currently on the back burner?

Best,
{my_name}""",
    },
    {
        "name": "Breakup Email", "kind": "email",
        "steps": "Breakup Email", "sort_order": 6,
        "subject": "Closing the loop",
        "body": """Hi {first_name},

I haven't heard back, so I'm going to assume that restoring those pre-1980s properties isn't a priority right now, or you already have a trusted maintenance plan in place.

I'll stop reaching out here, but I want to leave my contact info below in case you ever need a second opinion before pulling the trigger on a massive replacement project.

If things change and you want to see how skipping the tear-off can cut your capital expenditures by 50%, my door is always open.

Best,
{my_name}
{my_company}
{my_phone}""",
    },
]


def _migrate(conn) -> None:
    """Add columns introduced after the first release to existing databases."""
    cols = {row[1] for row in conn.execute("PRAGMA table_info(accounts)")}
    if "next_follow_up" not in cols:
        conn.execute("ALTER TABLE accounts ADD COLUMN next_follow_up TEXT DEFAULT ''")
    if "follow_up_note" not in cols:
        conn.execute("ALTER TABLE accounts ADD COLUMN follow_up_note TEXT DEFAULT ''")
    bid_cols = {row[1] for row in conn.execute("PRAGMA table_info(bids)")}
    if bid_cols:  # table exists
        for name, ddl in (
                ("coating_system", "TEXT DEFAULT 'Silicone'"),
                ("acrylic_system_type", "TEXT DEFAULT 'Standard'"),
                ("roof_type", "TEXT DEFAULT 'Capsheet'"),
                ("linear_feet", "INTEGER DEFAULT 0"),
                ("waste_pct", "REAL DEFAULT 5"),
                ("stretch_pct", "REAL DEFAULT 0"),
                ("passed_adhesion", "INTEGER DEFAULT 1"),
                ("has_rust", "INTEGER DEFAULT 0"),
                ("rust_prime_method", "TEXT DEFAULT 'field'")):
            if name not in bid_cols:
                conn.execute(f"ALTER TABLE bids ADD COLUMN {name} {ddl}")

    if "matching_properties" not in cols:
        conn.execute("ALTER TABLE accounts ADD COLUMN matching_properties INTEGER")
        # Templates seeded before this version used {num_properties} where the
        # criteria-matching count was meant ("...properties built before the
        # 1980s"); swap in the new placeholder.
        conn.execute("UPDATE templates SET "
                     "body = REPLACE(body, '{num_properties}', '{matching_properties}'), "
                     "subject = REPLACE(subject, '{num_properties}', '{matching_properties}')")


def _get_backup_dir_setting() -> str:
    """The user's off-machine backup folder (e.g. a OneDrive/Dropbox path)."""
    try:
        with sqlite3.connect(DB_PATH) as conn:
            row = conn.execute(
                "SELECT value FROM settings WHERE key='backup_dir'").fetchone()
            return (row[0] or "").strip() if row else ""
    except sqlite3.Error:
        return ""


def backup_db(keep: int = 14):
    """Copy crm.db into backups/ (at most once per day), keep the newest
    `keep` copies. Also mirrors the backup into the user's configured
    off-machine folder (backup_dir setting) when one is set — pointing that
    at a synced folder (OneDrive, Google Drive, Dropbox) protects the data
    if this computer dies. Returns the new backup path, or None if skipped."""
    if not DB_PATH.exists():
        return None
    BACKUP_DIR.mkdir(exist_ok=True)
    today_tag = date.today().strftime("%Y%m%d")
    existing = sorted(BACKUP_DIR.glob("crm-*.db"))
    made = None
    if not any(f.name.startswith(f"crm-{today_tag}-") for f in existing):
        made = BACKUP_DIR / f"crm-{today_tag}-{datetime.now().strftime('%H%M%S')}.db"
        shutil.copy2(DB_PATH, made)
        existing.append(made)
    for old in existing[:-keep]:
        old.unlink(missing_ok=True)

    mirror = _get_backup_dir_setting()
    latest = existing[-1] if existing else None
    if mirror and latest:
        try:
            mdir = Path(mirror).expanduser()
            mdir.mkdir(parents=True, exist_ok=True)
            if not (mdir / latest.name).exists():
                shutil.copy2(latest, mdir / latest.name)
            for old in sorted(mdir.glob("crm-*.db"))[:-30]:
                old.unlink(missing_ok=True)
        except OSError:
            pass  # never let a bad mirror path break startup
    return made


def init_db() -> None:
    with get_db() as conn:
        conn.executescript(SCHEMA)
        _migrate(conn)
        for key, value in DEFAULT_SETTINGS.items():
            conn.execute("INSERT OR IGNORE INTO settings (key, value) VALUES (?,?)",
                         (key, value))
        if conn.execute("SELECT COUNT(*) FROM templates").fetchone()[0] == 0:
            for t in SEED_TEMPLATES:
                conn.execute(
                    "INSERT INTO templates (name, kind, steps, subject, body, "
                    "sort_order, updated_at) VALUES (?,?,?,?,?,?,?)",
                    (t["name"], t["kind"], t["steps"], t["subject"], t["body"],
                     t["sort_order"], now_iso()))


if __name__ == "__main__":
    init_db()
    print(f"Database initialized at {DB_PATH}")
