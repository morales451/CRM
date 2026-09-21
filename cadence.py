"""Cadence & reminder engine.

Reminders are COMPUTED, not stored. A cadence step is "due" for an account when:
  1. The account is active in cadence:
       prospecting_status == 'Prospecting' AND pipeline_milestone == 'None / In Cadence'
  2. cadence_start + (step_day - 1) days <= today
  3. No interaction of that step's type has been logged for the account
  4. The step has not been manually checked off (cadence_dismissals)

Because reminders are derived on the fly, they persist on the dashboard until
logged or dismissed, and they vanish the instant an account leaves the active
status/milestone — no stored reminder rows to sync or clean up.
"""

from datetime import date, timedelta

# (day number, interaction type) — Day 1 is the cadence start date itself.
CADENCE_STEPS = [
    (1, "Email 1"),
    (3, "Call & Text"),
    (6, "Call 2"),
    (8, "Email 2"),
    (10, "Breakup Email"),
]

ACTIVE_STATUS = "Prospecting"
ACTIVE_MILESTONE = "None / In Cadence"


def _parse_date(iso_str: str) -> date:
    return date.fromisoformat(iso_str[:10])


def get_due_reminders(conn, account_id: int | None = None,
                      order: str = "due") -> list[dict]:
    """Return due cadence reminders.

    order="due"      -> oldest due date first (classic cadence order)
    order="priority" -> accounts with the most criteria-matching buildings
                        first, then oldest due date. Biggest portfolios are
                        the highest-potential deals, so they get worked first.

    Each reminder: {account_id, company_name, first_name, last_name,
                    work_phone, mobile_phone, email, preferred_contact,
                    matching_properties, num_properties,
                    day, step_type, due_date, days_overdue}
    """
    today = date.today()

    sql = """
        SELECT id, company_name, first_name, last_name, work_phone,
               mobile_phone, email, preferred_contact, cadence_start,
               matching_properties, num_properties
        FROM accounts
        WHERE prospecting_status = ? AND pipeline_milestone = ?
          AND COALESCE(archived_at, '') = ''
    """
    params: list = [ACTIVE_STATUS, ACTIVE_MILESTONE]
    if account_id is not None:
        sql += " AND id = ?"
        params.append(account_id)
    accounts = conn.execute(sql, params).fetchall()
    if not accounts:
        return []

    ids = [a["id"] for a in accounts]
    ph = ",".join("?" * len(ids))

    logged: set[tuple[int, str]] = set()
    for row in conn.execute(
        f"SELECT DISTINCT account_id, interaction_type FROM interactions "
        f"WHERE account_id IN ({ph})", ids):
        logged.add((row["account_id"], row["interaction_type"]))

    dismissed: set[tuple[int, str]] = set()
    for row in conn.execute(
        f"SELECT account_id, step_type FROM cadence_dismissals "
        f"WHERE account_id IN ({ph})", ids):
        dismissed.add((row["account_id"], row["step_type"]))

    reminders = []
    for acct in accounts:
        start = _parse_date(acct["cadence_start"])
        for day, step_type in CADENCE_STEPS:
            due = start + timedelta(days=day - 1)
            if due > today:
                continue
            key = (acct["id"], step_type)
            if key in logged or key in dismissed:
                continue
            reminders.append({
                "account_id": acct["id"],
                "company_name": acct["company_name"],
                "first_name": acct["first_name"],
                "last_name": acct["last_name"],
                "work_phone": acct["work_phone"],
                "mobile_phone": acct["mobile_phone"],
                "email": acct["email"],
                "preferred_contact": acct["preferred_contact"],
                "matching_properties": acct["matching_properties"],
                "num_properties": acct["num_properties"],
                "day": day,
                "step_type": step_type,
                "due_date": due.isoformat(),
                "days_overdue": (today - due).days,
            })

    if order == "priority":
        reminders.sort(key=lambda r: (-(r["matching_properties"] or 0),
                                      r["due_date"], r["company_name"].lower(),
                                      r["day"]))
    else:
        reminders.sort(key=lambda r: (r["due_date"], r["company_name"].lower(),
                                      r["day"]))
    return reminders


def get_cadence_progress(conn, account_id: int) -> list[dict]:
    """Full cadence step status for one account's detail page."""
    acct = conn.execute(
        "SELECT prospecting_status, pipeline_milestone, cadence_start "
        "FROM accounts WHERE id = ?", (account_id,)).fetchone()
    if acct is None:
        return []
    active = (acct["prospecting_status"] == ACTIVE_STATUS
              and acct["pipeline_milestone"] == ACTIVE_MILESTONE)
    start = _parse_date(acct["cadence_start"])
    today = date.today()

    logged = {r["interaction_type"] for r in conn.execute(
        "SELECT DISTINCT interaction_type FROM interactions WHERE account_id = ?",
        (account_id,))}
    dismissed = {r["step_type"] for r in conn.execute(
        "SELECT step_type FROM cadence_dismissals WHERE account_id = ?",
        (account_id,))}

    steps = []
    for day, step_type in CADENCE_STEPS:
        due = start + timedelta(days=day - 1)
        if step_type in logged:
            state = "done"
        elif step_type in dismissed:
            state = "skipped"
        elif not active:
            state = "inactive"
        elif due <= today:
            state = "due"
        else:
            state = "upcoming"
        steps.append({"day": day, "step_type": step_type,
                      "due_date": due.isoformat(), "state": state})
    return steps
