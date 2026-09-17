"""SQLite database layer for the Roof Coating CRM.

All timestamps are strict ISO-8601 with local timezone offset,
e.g. 2026-09-17T14:03:22-05:00.
"""

import sqlite3
from datetime import datetime, date
from pathlib import Path

DB_PATH = Path(__file__).parent / "crm.db"

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

SCHEMA = """
CREATE TABLE IF NOT EXISTS accounts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    company_name TEXT NOT NULL,
    first_name TEXT DEFAULT '',
    last_name TEXT DEFAULT '',
    title TEXT DEFAULT '',
    num_properties INTEGER,
    email TEXT DEFAULT '',
    work_phone TEXT DEFAULT '',
    mobile_phone TEXT DEFAULT '',
    preferred_contact TEXT NOT NULL DEFAULT 'Unknown',
    notes TEXT DEFAULT '',
    prospecting_status TEXT NOT NULL DEFAULT 'Prospecting',
    pipeline_milestone TEXT NOT NULL DEFAULT 'None / In Cadence',
    cadence_start TEXT NOT NULL,          -- ISO date the cadence clock starts from
    created_at TEXT NOT NULL,             -- ISO timestamp
    updated_at TEXT NOT NULL              -- ISO timestamp
);

CREATE TABLE IF NOT EXISTS interactions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    account_id INTEGER NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
    interaction_type TEXT NOT NULL,
    notes TEXT DEFAULT '',
    created_at TEXT NOT NULL              -- ISO timestamp
);

-- A manually checked-off cadence step ("clear the task without logging it").
CREATE TABLE IF NOT EXISTS cadence_dismissals (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    account_id INTEGER NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
    step_type TEXT NOT NULL,
    dismissed_at TEXT NOT NULL,           -- ISO timestamp
    UNIQUE (account_id, step_type)
);

CREATE INDEX IF NOT EXISTS idx_interactions_account ON interactions(account_id);
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
    return conn


def init_db() -> None:
    with get_db() as conn:
        conn.executescript(SCHEMA)


if __name__ == "__main__":
    init_db()
    print(f"Database initialized at {DB_PATH}")
