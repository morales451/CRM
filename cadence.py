"""Cadence & reminder engine.

Reminders are COMPUTED, not stored. A cadence step is "due" for an account when:
  1. The account is active in cadence:
       prospecting_status == 'Prospecting' AND pipeline_milestone == 'None / In Cadence'
     and its clock has started (cadence_start is set — an empty cadence_start
     means it's still waiting in Research for someone to contact)
  2. Its due date has arrived. Step 1 is due on the start date itself. Each
     later step is due the cadence's gap in BUSINESS days after the step
     before it: after the day that step was done if it's done, otherwise after
     its due date. So the clock resets from every action you take. Send Email 1
     a week late and the call comes two business days after you sent it, not
     overdue the moment you log it.
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


def today() -> date:
    """The date the cadence is measured against. A function rather than a
    direct date.today() call so tests can pin it to a known weekday."""
    return date.today()


def add_business_days(start: date, n: int) -> date:
    """start moved forward n working days, skipping Saturdays and Sundays."""
    d = start
    while n > 0:
        d += timedelta(days=1)
        if d.weekday() < 5:
            n -= 1
    return d


def step_due(start: date, day: int) -> date:
    """When cadence step `day` falls due for a clock started on `start`.

    Day 1 is the start date whatever day that is — import on a Saturday and
    the first email is due that Saturday. Every later step counts business
    days, so Day 3 of a Friday start is Tuesday, not Sunday."""
    return start if day <= 1 else add_business_days(start, day - 1)


def schedule(start: date, done_on: dict[str, date]) -> dict[str, date]:
    """Due date of every step, given when the finished steps were done.

    Each step's due date is counted from the step before it: from the day
    that step was done if it's done, otherwise from its own due date. With
    nothing done this is the plain Day 1/3/6/8/10 business-day schedule.
    """
    dues: dict[str, date] = {}
    anchor, prev_day = start, None
    for day, step_type in CADENCE_STEPS:
        due = start if prev_day is None else add_business_days(anchor, day - prev_day)
        dues[step_type] = due
        anchor = done_on.get(step_type, due)
        prev_day = day
    return dues


def _done_dates(conn, ids: list[int]) -> dict[int, dict[str, date]]:
    """{account_id: {step: date first done}} from logged steps and check-offs.
    A call that hit a bad number didn't complete its step."""
    out: dict[int, dict[str, date]] = {}
    if not ids:
        return out
    ph = ",".join("?" * len(ids))
    steps = [st for _, st in CADENCE_STEPS]
    sph = ",".join("?" * len(steps))
    rows = list(conn.execute(
        f"SELECT account_id, interaction_type AS step, MIN(created_at) AS at "
        f"FROM interactions WHERE account_id IN ({ph}) AND interaction_type IN ({sph}) "
        f"AND COALESCE(outcome, '') != 'Bad number' GROUP BY account_id, interaction_type",
        (*ids, *steps)))
    rows += list(conn.execute(
        f"SELECT account_id, step_type AS step, dismissed_at AS at "
        f"FROM cadence_dismissals WHERE account_id IN ({ph})", ids))
    for r in rows:
        try:
            d = _parse_date(r["at"])
        except (TypeError, ValueError):
            continue
        steps_done = out.setdefault(r["account_id"], {})
        if r["step"] not in steps_done or d < steps_done[r["step"]]:
            steps_done[r["step"]] = d
    return out


def has_contact(acct) -> bool:
    """Is there someone to reach? A named person plus an email or a phone.

    The cadence opens with an email, so starting the clock on a company with
    nothing but a switchboard number just manufactures overdue tasks nobody
    can do. Accounts without a contact wait in Research until one is added.
    """
    def val(key):
        try:
            return (acct[key] or "").strip()
        except (KeyError, IndexError):
            return ""
    named = bool(val("first_name") or val("last_name"))
    reachable = bool(val("email") or val("work_phone") or val("mobile_phone"))
    return named and reachable


def start_cadence_if_ready(conn, account_id, start: str | None = None) -> bool:
    """Start an account's clock today if it was waiting in Research and now
    has someone to contact. Call after anything that adds contact details.
    `start` (ISO date) overrides today — a bulk contact upload uses it to
    stagger a big batch instead of starting them all at once.

    Only ever starts a clock — never stops or moves one that's running.
    Returns True when it started one. The caller commits.
    """
    acct = conn.execute(
        "SELECT * FROM accounts WHERE id = ?", (account_id,)).fetchone()
    if acct is None or (acct["cadence_start"] or "").strip():
        return False
    if (acct["archived_at"] or "") or acct["prospecting_status"] != ACTIVE_STATUS \
            or acct["pipeline_milestone"] != ACTIVE_MILESTONE:
        return False
    if not has_contact(acct):
        return False
    conn.execute("UPDATE accounts SET cadence_start = ? WHERE id = ?",
                 (start or today().isoformat(), account_id))
    return True


def _stage_sql(clock_running: bool, alias: str = "") -> str:
    a = f"{alias}." if alias else ""
    return (f"COALESCE({a}archived_at, '') = '' "
            f"AND {a}prospecting_status = '{ACTIVE_STATUS}' "
            f"AND {a}pipeline_milestone = '{ACTIVE_MILESTONE}' "
            f"AND COALESCE({a}cadence_start, '') {'!=' if clock_running else '='} ''")


def research_where(alias: str = "") -> str:
    """SQL condition for "waiting in Research": active, not archived, no clock
    yet. Shared so every count and list agrees on what Research means."""
    return _stage_sql(False, alias)


def in_cadence_where(alias: str = "") -> str:
    """SQL condition for "in the cadence": the same, with the clock running."""
    return _stage_sql(True, alias)


RESEARCH_SQL = research_where()
IN_CADENCE_SQL = in_cadence_where()


def get_due_reminders(conn, account_id: int | None = None,
                      order: str = "due", collapse: bool = True) -> list[dict]:
    """Return due cadence reminders — by default, the NEXT one per account.

    The cadence is a sequence: Email 1, then the call, then Email 2. You
    cannot do all five to the same person on the same day, so an account that
    has sat untouched for two weeks should appear once, at the step it is
    actually waiting on — not five times. Logging that step brings the next
    one forward.

    collapse=False returns every outstanding step, which is what the
    "steps behind" count on a reminder is worked out from.

    order="due"      -> oldest due date first (classic cadence order)
    order="priority" -> accounts with the most criteria-matching buildings
                        first, then oldest due date. Biggest portfolios are
                        the highest-potential deals, so they get worked first.

    Each reminder: {account_id, company_name, first_name, last_name,
                    work_phone, mobile_phone, email, preferred_contact,
                    matching_properties, num_properties,
                    day, step_type, due_date, days_overdue, steps_behind}
    """
    now = today()

    sql = """
        SELECT id, company_name, first_name, last_name, work_phone,
               mobile_phone, email, preferred_contact, cadence_start,
               matching_properties, num_properties
        FROM accounts
        WHERE prospecting_status = ? AND pipeline_milestone = ?
          AND COALESCE(archived_at, '') = ''
          AND COALESCE(cadence_start, '') != ''
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
    # A call that hit a bad number didn't complete its step — the step stays
    # outstanding so it comes round again once the number is fixed.
    for row in conn.execute(
        f"SELECT DISTINCT account_id, interaction_type FROM interactions "
        f"WHERE account_id IN ({ph}) AND COALESCE(outcome, '') != 'Bad number'", ids):
        logged.add((row["account_id"], row["interaction_type"]))

    dismissed: set[tuple[int, str]] = set()
    for row in conn.execute(
        f"SELECT account_id, step_type FROM cadence_dismissals "
        f"WHERE account_id IN ({ph})", ids):
        dismissed.add((row["account_id"], row["step_type"]))

    done = _done_dates(conn, ids)
    reminders = []
    for acct in accounts:
        start = _parse_date(acct["cadence_start"])
        dues = schedule(start, done.get(acct["id"], {}))
        for day, step_type in CADENCE_STEPS:
            due = dues[step_type]
            if due > now:
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
                "days_overdue": (now - due).days,
                "steps_behind": 1,
            })

    if collapse:
        # Keep only the earliest outstanding step for each account, and note
        # how many others are stacked up behind it.
        nxt: dict[int, dict] = {}
        for r in reminders:
            current = nxt.get(r["account_id"])
            if current is None:
                nxt[r["account_id"]] = r
            elif r["day"] < current["day"]:
                r["steps_behind"] = current["steps_behind"] + 1
                nxt[r["account_id"]] = r
            else:
                current["steps_behind"] += 1
        reminders = list(nxt.values())

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
    started = bool((acct["cadence_start"] or "").strip())
    start = _parse_date(acct["cadence_start"]) if started else None
    now = today()

    logged = {r["interaction_type"] for r in conn.execute(
        "SELECT DISTINCT interaction_type FROM interactions WHERE account_id = ? "
        "AND COALESCE(outcome, '') != 'Bad number'", (account_id,))}
    dismissed = {r["step_type"] for r in conn.execute(
        "SELECT step_type FROM cadence_dismissals WHERE account_id = ?",
        (account_id,))}

    done_on = _done_dates(conn, [account_id]).get(account_id, {})
    dues = schedule(start, done_on) if started else {}
    steps = []
    for day, step_type in CADENCE_STEPS:
        due = dues.get(step_type)
        if step_type in logged:
            state = "done"
        elif step_type in dismissed:
            state = "skipped"
        elif not active:
            state = "inactive"
        elif not started:
            state = "waiting"          # in Research: no clock yet
        elif due <= now:
            state = "due"
        else:
            state = "upcoming"
        done = done_on.get(step_type)
        steps.append({"day": day, "step_type": step_type,
                      "due_date": due.isoformat() if due else "", "state": state,
                      "done_on": done.isoformat() if done else "",
                      "days_late": (now - due).days if due and state == "due" else 0})
    return steps
