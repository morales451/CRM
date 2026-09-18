"""Roof Coating CRM — lightweight local Flask app.

Run:  python3 app.py
Then open http://<your-local-ip>:8000 from any device on your Wi-Fi.
"""

import csv
import io
import re
import socket
from datetime import datetime, timedelta
from urllib.parse import quote

from flask import (Flask, flash, redirect, render_template, request,
                   send_file, url_for)

import cadence
import importer
from db import (INTERACTION_TYPES, PIPELINE_MILESTONES,
                PREFERRED_CONTACT_METHODS, PROSPECTING_STATUSES,
                backup_db, get_db, init_db, now_iso, today_iso)

app = Flask(__name__)
app.secret_key = "local-crm-flash-messages"  # local single-user app; used only for flash()
PORT = 8000


@app.template_filter("dt")
def format_datetime(iso_str):
    """Render an ISO timestamp as 'Sep 17, 2026 2:03 PM'."""
    try:
        return datetime.fromisoformat(iso_str).strftime("%b %d, %Y %I:%M %p")
    except (ValueError, TypeError):
        return iso_str or ""


@app.template_filter("pref")
def pref_badge(method):
    """Preferred contact method as a compact badge, empty for Unknown."""
    icons = {"Email": "✉️ prefers email", "Call": "📞 prefers call",
             "Text": "💬 prefers text"}
    return icons.get(method, "")


@app.template_filter("d")
def format_date(iso_str):
    try:
        return datetime.fromisoformat(iso_str[:10]).strftime("%b %d, %Y")
    except (ValueError, TypeError):
        return iso_str or ""


@app.context_processor
def inject_constants():
    return {
        "STATUSES": PROSPECTING_STATUSES,
        "MILESTONES": PIPELINE_MILESTONES,
        "CONTACT_METHODS": PREFERRED_CONTACT_METHODS,
        "INTERACTION_TYPES": INTERACTION_TYPES,
        "today": today_iso(),
    }


def _account_or_404(conn, account_id):
    acct = conn.execute("SELECT * FROM accounts WHERE id = ?", (account_id,)).fetchone()
    if acct is None:
        from flask import abort
        abort(404)
    return acct


def _account_fields_from_form(form):
    def int_or_none(name):
        raw = form.get(name, "").strip()
        return int(raw) if raw.isdigit() else None

    return {
        "company_name": form.get("company_name", "").strip(),
        "first_name": form.get("first_name", "").strip(),
        "last_name": form.get("last_name", "").strip(),
        "title": form.get("title", "").strip(),
        "num_properties": int_or_none("num_properties"),
        "matching_properties": int_or_none("matching_properties"),
        "email": form.get("email", "").strip(),
        "work_phone": form.get("work_phone", "").strip(),
        "mobile_phone": form.get("mobile_phone", "").strip(),
        "preferred_contact": form.get("preferred_contact", "Unknown"),
        "notes": form.get("notes", "").strip(),
        "prospecting_status": form.get("prospecting_status", "Prospecting"),
        "pipeline_milestone": form.get("pipeline_milestone", "None / In Cadence"),
    }


# ---------------------------------------------------------------- Dashboard

STALE_DAYS = 14
STALE_EXCLUDED_MILESTONES = ("None / In Cadence", "Closed Won", "Closed Lost", "On Hold")


def _due_followups(conn):
    return conn.execute(
        "SELECT * FROM accounts WHERE next_follow_up != '' AND next_follow_up <= ? "
        "ORDER BY next_follow_up, company_name COLLATE NOCASE",
        (today_iso(),)).fetchall()


def _stale_deals(conn):
    """Pipeline accounts with no touch in STALE_DAYS days and no follow-up set."""
    ph = ",".join("?" * len(STALE_EXCLUDED_MILESTONES))
    rows = conn.execute(
        f"""SELECT a.*, COALESCE(
                (SELECT MAX(created_at) FROM interactions i WHERE i.account_id = a.id),
                a.updated_at) AS last_touch
            FROM accounts a
            WHERE a.pipeline_milestone NOT IN ({ph})
              AND a.prospecting_status != 'Not Interested'
              AND (a.next_follow_up IS NULL OR a.next_follow_up = '')""",
        STALE_EXCLUDED_MILESTONES).fetchall()
    cutoff = (datetime.now() - timedelta(days=STALE_DAYS)).date().isoformat()
    return [r for r in rows if (r["last_touch"] or "")[:10] <= cutoff]


@app.route("/")
def dashboard():
    conn = get_db()
    try:
        reminders = cadence.get_due_reminders(conn)
        followups = _due_followups(conn)
        stale = _stale_deals(conn)
        stats = {
            "total_accounts": conn.execute(
                "SELECT COUNT(*) c FROM accounts").fetchone()["c"],
            "active_prospects": conn.execute(
                "SELECT COUNT(*) c FROM accounts WHERE prospecting_status "
                "NOT IN ('Not Interested') AND pipeline_milestone "
                "NOT IN ('Closed Won','Closed Lost')").fetchone()["c"],
            "in_cadence": conn.execute(
                "SELECT COUNT(*) c FROM accounts WHERE prospecting_status = ? "
                "AND pipeline_milestone = ?",
                (cadence.ACTIVE_STATUS, cadence.ACTIVE_MILESTONE)).fetchone()["c"],
            "closed_won": conn.execute(
                "SELECT COUNT(*) c FROM accounts "
                "WHERE pipeline_milestone = 'Closed Won'").fetchone()["c"],
        }
        recent = conn.execute(
            """SELECT i.*, a.company_name FROM interactions i
               JOIN accounts a ON a.id = i.account_id
               ORDER BY i.created_at DESC, i.id DESC LIMIT 10""").fetchall()
        return render_template("dashboard.html", reminders=reminders,
                               followups=followups, stale=stale,
                               stats=stats, recent=recent, today=today_iso())
    finally:
        conn.close()


@app.route("/reminders/dismiss", methods=["POST"])
def dismiss_reminder():
    account_id = request.form["account_id"]
    step_type = request.form["step_type"]
    conn = get_db()
    try:
        conn.execute(
            "INSERT OR IGNORE INTO cadence_dismissals "
            "(account_id, step_type, dismissed_at) VALUES (?,?,?)",
            (account_id, step_type, now_iso()))
        conn.commit()
    finally:
        conn.close()
    flash(f"Checked off “{step_type}”.", "success")
    return redirect(request.form.get("next") or url_for("dashboard"))


@app.route("/reminders/quicklog", methods=["POST"])
def quick_log():
    """Log a cadence step straight from the dashboard (clears its reminder)."""
    account_id = request.form["account_id"]
    step_type = request.form["step_type"]
    notes = request.form.get("notes", "").strip()
    conn = get_db()
    try:
        conn.execute(
            "INSERT INTO interactions (account_id, interaction_type, notes, created_at) "
            "VALUES (?,?,?,?)", (account_id, step_type, notes, now_iso()))
        conn.commit()
    finally:
        conn.close()
    flash(f"Logged “{step_type}”.", "success")
    return redirect(request.form.get("next") or url_for("dashboard"))


# ------------------------------------------------------- Follow-ups & queue

@app.route("/accounts/<int:account_id>/followup", methods=["POST"])
def set_followup(account_id):
    """Set, snooze, or clear an account's follow-up reminder."""
    conn = get_db()
    try:
        _account_or_404(conn, account_id)
        if request.form.get("clear"):
            conn.execute(
                "UPDATE accounts SET next_follow_up='', follow_up_note='', "
                "updated_at=? WHERE id=?", (now_iso(), account_id))
            flash("Follow-up cleared.", "success")
        else:
            days = request.form.get("days", "")
            if days.lstrip("-").isdigit():
                due = (datetime.now() + timedelta(days=int(days))).date().isoformat()
            else:
                due = request.form.get("date", "").strip()
            if not due:
                flash("Pick a follow-up date.", "danger")
                return redirect(request.form.get("next")
                                or url_for("account_detail", account_id=account_id))
            if "note" in request.form:
                conn.execute(
                    "UPDATE accounts SET next_follow_up=?, follow_up_note=?, "
                    "updated_at=? WHERE id=?",
                    (due, request.form.get("note", "").strip(), now_iso(), account_id))
            else:
                conn.execute(
                    "UPDATE accounts SET next_follow_up=?, updated_at=? WHERE id=?",
                    (due, now_iso(), account_id))
            flash(f"Follow-up set for {due}.", "success")
        conn.commit()
    finally:
        conn.close()
    return redirect(request.form.get("next")
                    or url_for("account_detail", account_id=account_id))


def _build_queue(conn):
    """Today's work: due cadence steps + due follow-ups, oldest first."""
    tasks = [{"kind": "cadence", "due": r["due_date"], "account_id": r["account_id"],
              "reminder": r} for r in cadence.get_due_reminders(conn)]
    tasks += [{"kind": "followup", "due": a["next_follow_up"], "account_id": a["id"],
               "note": a["follow_up_note"]} for a in _due_followups(conn)]
    tasks.sort(key=lambda t: t["due"])
    return tasks


@app.route("/queue")
def queue():
    """One-task-at-a-time focus mode: script on screen, log and move on."""
    conn = get_db()
    try:
        tasks = _build_queue(conn)
        total = len(tasks)
        if total == 0:
            return render_template("queue.html", task=None, total=0, pos=0)
        pos = max(0, min(int(request.args.get("pos", 0) or 0), total - 1))
        task = tasks[pos]
        acct = conn.execute("SELECT * FROM accounts WHERE id=?",
                            (task["account_id"],)).fetchone()
        scripts = []
        if task["kind"] == "cadence":
            settings = _get_settings(conn)
            rows = conn.execute(
                "SELECT * FROM templates ORDER BY sort_order, id").fetchall()
            scripts = _build_scripts(acct, settings, rows,
                                     task["reminder"]["step_type"])
    finally:
        conn.close()
    return render_template("queue.html", task=task, account=acct,
                           scripts=scripts, total=total, pos=pos)


# ----------------------------------------------------------------- Insights

def _bar_items(pairs, total=None):
    """[(label, count)] -> render-ready bars with widths scaled to the max."""
    peak = max((c for _, c in pairs), default=0) or 1
    return [{"label": l, "count": c, "pct": round(100 * c / peak),
             "share": (round(100 * c / total) if total else None)}
            for l, c in pairs]


@app.route("/insights")
def insights():
    """Critical numbers: activity trend, funnel, cadence completion, breakdowns."""
    conn = get_db()
    try:
        today = datetime.now().date()

        def count(sql, *params):
            return conn.execute(sql, params).fetchone()[0]

        total_accounts = count("SELECT COUNT(*) FROM accounts")
        won = count("SELECT COUNT(*) FROM accounts WHERE pipeline_milestone='Closed Won'")
        lost = count("SELECT COUNT(*) FROM accounts WHERE pipeline_milestone='Closed Lost'")
        tiles = {
            "tasks_due": len(cadence.get_due_reminders(conn)) + len(_due_followups(conn)),
            "total_accounts": total_accounts,
            "active_prospects": count(
                "SELECT COUNT(*) FROM accounts WHERE prospecting_status != 'Not Interested' "
                "AND pipeline_milestone NOT IN ('Closed Won','Closed Lost')"),
            "in_cadence": count(
                "SELECT COUNT(*) FROM accounts WHERE prospecting_status=? AND pipeline_milestone=?",
                cadence.ACTIVE_STATUS, cadence.ACTIVE_MILESTONE),
            "won": won,
            "win_rate": round(100 * won / (won + lost)) if (won + lost) else None,
        }

        # Interactions per week, last 8 weeks (Mondays as bucket starts)
        monday = today - timedelta(days=today.weekday())
        week_starts = [monday - timedelta(weeks=i) for i in range(7, -1, -1)]
        buckets = {w.isoformat(): 0 for w in week_starts}
        since = week_starts[0].isoformat()
        for (created,) in conn.execute(
                "SELECT created_at FROM interactions WHERE created_at >= ?", (since,)):
            d = datetime.fromisoformat(created).date()
            key = (d - timedelta(days=d.weekday())).isoformat()
            if key in buckets:
                buckets[key] += 1
        weekly_counts = [buckets[w.isoformat()] for w in week_starts]
        peak = max(weekly_counts) or 1
        weekly = [{"label": f"{w.strftime('%b')} {w.day}", "count": c,
                   "pct": round(100 * c / peak)}
                  for w, c in zip(week_starts, weekly_counts)]
        tiles["this_week"] = weekly_counts[-1]
        tiles["week_delta"] = weekly_counts[-1] - weekly_counts[-2]

        # Pipeline funnel (the cadence pool would dwarf it, so it's a tile instead)
        pipeline_bars = _bar_items([
            (m, count("SELECT COUNT(*) FROM accounts WHERE pipeline_milestone=?", m))
            for m in PIPELINE_MILESTONES if m != "None / In Cadence"])

        # How far accounts get through the cadence (distinct accounts per step)
        cadence_bars = _bar_items([
            (f"Day {day}: {step}",
             count("SELECT COUNT(DISTINCT account_id) FROM interactions "
                   "WHERE interaction_type=?", step))
            for day, step in cadence.CADENCE_STEPS], total=total_accounts)

        status_bars = _bar_items([
            (s, count("SELECT COUNT(*) FROM accounts WHERE prospecting_status=?", s))
            for s in PROSPECTING_STATUSES], total=total_accounts)

        month_ago = (today - timedelta(days=30)).isoformat()
        type_bars = _bar_items(conn.execute(
            "SELECT interaction_type, COUNT(*) FROM interactions "
            "WHERE created_at >= ? GROUP BY interaction_type "
            "ORDER BY COUNT(*) DESC", (month_ago,)).fetchall())
    finally:
        conn.close()
    return render_template("insights.html", tiles=tiles, weekly=weekly,
                           pipeline_bars=pipeline_bars, cadence_bars=cadence_bars,
                           status_bars=status_bars, type_bars=type_bars)


# ----------------------------------------------------------------- Pipeline

@app.route("/pipeline")
def pipeline():
    """Board view: one column per milestone (cadence pool shown as a count)."""
    conn = get_db()
    try:
        columns = []
        for m in PIPELINE_MILESTONES:
            rows = conn.execute(
                """SELECT a.*, (SELECT MAX(created_at) FROM interactions i
                                WHERE i.account_id = a.id) AS last_activity
                   FROM accounts a WHERE a.pipeline_milestone = ?
                   ORDER BY a.updated_at DESC""", (m,)).fetchall()
            columns.append({"milestone": m, "count": len(rows), "accounts": rows[:20]})
    finally:
        conn.close()
    return render_template("pipeline.html", columns=columns)


@app.route("/accounts/<int:account_id>/milestone", methods=["POST"])
def move_milestone(account_id):
    milestone = request.form.get("pipeline_milestone", "")
    if milestone in PIPELINE_MILESTONES:
        conn = get_db()
        try:
            _account_or_404(conn, account_id)
            conn.execute("UPDATE accounts SET pipeline_milestone=?, updated_at=? "
                         "WHERE id=?", (milestone, now_iso(), account_id))
            conn.commit()
        finally:
            conn.close()
        flash(f"Moved to “{milestone}”.", "success")
    return redirect(request.form.get("next") or url_for("pipeline"))


# ------------------------------------------------------------------- Export

def _csv_response(rows, headers, filename):
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(headers)
    writer.writerows(rows)
    return send_file(io.BytesIO(buf.getvalue().encode("utf-8-sig")),
                     mimetype="text/csv", as_attachment=True,
                     download_name=filename)


@app.route("/export/accounts.csv")
def export_accounts():
    conn = get_db()
    try:
        rows = conn.execute(
            "SELECT * FROM accounts ORDER BY company_name COLLATE NOCASE").fetchall()
    finally:
        conn.close()
    if not rows:
        headers = ["company_name"]
        data = []
    else:
        headers = rows[0].keys()
        data = [tuple(r) for r in rows]
    return _csv_response(data, headers, "crm_accounts.csv")


@app.route("/export/interactions.csv")
def export_interactions():
    conn = get_db()
    try:
        rows = conn.execute(
            """SELECT i.id, a.company_name, i.interaction_type, i.notes, i.created_at
               FROM interactions i JOIN accounts a ON a.id = i.account_id
               ORDER BY i.created_at""").fetchall()
    finally:
        conn.close()
    return _csv_response([tuple(r) for r in rows],
                         ["id", "company_name", "interaction_type", "notes",
                          "created_at"], "crm_interactions.csv")


# ----------------------------------------------------------------- Accounts

@app.route("/accounts")
def accounts():
    status = request.args.get("status", "")
    milestone = request.args.get("milestone", "")
    q = request.args.get("q", "").strip()

    sql = """SELECT a.*,
                    (SELECT MAX(created_at) FROM interactions i
                     WHERE i.account_id = a.id) AS last_activity
             FROM accounts a WHERE 1=1"""
    params = []
    if status:
        sql += " AND a.prospecting_status = ?"
        params.append(status)
    if milestone:
        sql += " AND a.pipeline_milestone = ?"
        params.append(milestone)
    if q:
        sql += (" AND (a.company_name LIKE ? OR a.first_name LIKE ? "
                "OR a.last_name LIKE ? OR a.email LIKE ?)")
        params += [f"%{q}%"] * 4
    sql += " ORDER BY a.company_name COLLATE NOCASE"

    conn = get_db()
    try:
        rows = conn.execute(sql, params).fetchall()
    finally:
        conn.close()
    return render_template("accounts.html", accounts=rows,
                           status=status, milestone=milestone, q=q)


@app.route("/accounts/new", methods=["GET", "POST"])
def new_account():
    if request.method == "POST":
        fields = _account_fields_from_form(request.form)
        if not fields["company_name"]:
            flash("Company name is required.", "danger")
            return render_template("account_form.html", account=request.form)
        ts = now_iso()
        conn = get_db()
        try:
            cur = conn.execute(
                """INSERT INTO accounts
                   (company_name, first_name, last_name, title, num_properties,
                    matching_properties, email, work_phone, mobile_phone,
                    preferred_contact, notes, prospecting_status,
                    pipeline_milestone, cadence_start, created_at, updated_at)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (*fields.values(), today_iso(), ts, ts))
            conn.commit()
            new_id = cur.lastrowid
        finally:
            conn.close()
        flash(f"Account “{fields['company_name']}” created.", "success")
        return redirect(url_for("account_detail", account_id=new_id))
    return render_template("account_form.html", account=None)


@app.route("/accounts/<int:account_id>")
def account_detail(account_id):
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        contacts = conn.execute(
            "SELECT * FROM contacts WHERE account_id = ? "
            "ORDER BY last_name COLLATE NOCASE, first_name COLLATE NOCASE",
            (account_id,)).fetchall()
        interactions = conn.execute(
            "SELECT * FROM interactions WHERE account_id = ? "
            "ORDER BY created_at DESC, id DESC", (account_id,)).fetchall()
        steps = cadence.get_cadence_progress(conn, account_id)
        in_cadence = (acct["prospecting_status"] == cadence.ACTIVE_STATUS
                      and acct["pipeline_milestone"] == cadence.ACTIVE_MILESTONE)
    finally:
        conn.close()
    return render_template("account_detail.html", account=acct,
                           contacts=contacts, interactions=interactions,
                           steps=steps, in_cadence=in_cadence)


@app.route("/accounts/<int:account_id>/edit", methods=["POST"])
def edit_account(account_id):
    fields = _account_fields_from_form(request.form)
    if not fields["company_name"]:
        flash("Company name is required.", "danger")
        return redirect(url_for("account_detail", account_id=account_id))
    conn = get_db()
    try:
        _account_or_404(conn, account_id)
        conn.execute(
            """UPDATE accounts SET company_name=?, first_name=?, last_name=?,
               title=?, num_properties=?, matching_properties=?, email=?,
               work_phone=?, mobile_phone=?, preferred_contact=?, notes=?,
               prospecting_status=?, pipeline_milestone=?, updated_at=?
               WHERE id=?""",
            (*fields.values(), now_iso(), account_id))
        conn.commit()
    finally:
        conn.close()
    flash("Account updated.", "success")
    return redirect(url_for("account_detail", account_id=account_id))


@app.route("/accounts/<int:account_id>/restart-cadence", methods=["POST"])
def restart_cadence(account_id):
    """Reset the cadence clock to today and clear prior step history."""
    conn = get_db()
    try:
        _account_or_404(conn, account_id)
        conn.execute(
            "UPDATE accounts SET cadence_start=?, prospecting_status=?, "
            "pipeline_milestone=?, updated_at=? WHERE id=?",
            (today_iso(), cadence.ACTIVE_STATUS, cadence.ACTIVE_MILESTONE,
             now_iso(), account_id))
        conn.execute("DELETE FROM cadence_dismissals WHERE account_id=?",
                     (account_id,))
        # Old logged cadence steps would suppress the new cycle's reminders,
        # so retag them as day-0 history under General Note.
        conn.execute(
            "UPDATE interactions SET interaction_type='General Note', "
            "notes='[' || interaction_type || ' — previous cadence] ' || notes "
            "WHERE account_id=? AND interaction_type IN "
            "('Email 1','Call & Text','Call 2','Email 2','Breakup Email')",
            (account_id,))
        conn.commit()
    finally:
        conn.close()
    flash("Cadence restarted from today.", "success")
    return redirect(url_for("account_detail", account_id=account_id))


@app.route("/accounts/<int:account_id>/delete", methods=["POST"])
def delete_account(account_id):
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        conn.execute("DELETE FROM accounts WHERE id=?", (account_id,))
        conn.commit()
    finally:
        conn.close()
    flash(f"Account “{acct['company_name']}” deleted.", "warning")
    return redirect(url_for("accounts"))


@app.route("/accounts/<int:account_id>/log", methods=["POST"])
def log_interaction(account_id):
    itype = request.form.get("interaction_type", "General Note")
    notes = request.form.get("notes", "").strip()
    conn = get_db()
    try:
        _account_or_404(conn, account_id)
        conn.execute(
            "INSERT INTO interactions (account_id, interaction_type, notes, created_at) "
            "VALUES (?,?,?,?)", (account_id, itype, notes, now_iso()))
        conn.commit()
    finally:
        conn.close()
    flash(f"Logged “{itype}”.", "success")
    return redirect(request.form.get("next")
                    or url_for("account_detail", account_id=account_id))


# ----------------------------------------------------------------- Contacts

CONTACT_COLS = ("first_name", "last_name", "title", "email", "work_phone", "mobile_phone")


def _has_primary_contact(acct) -> bool:
    return any(acct[c] for c in ("first_name", "last_name", "email"))


def _demote_primary_to_contact(conn, acct):
    """Move the account's current primary-contact fields into a contacts row."""
    if _has_primary_contact(acct):
        conn.execute(
            "INSERT INTO contacts (account_id, first_name, last_name, title, "
            "email, work_phone, mobile_phone, created_at) VALUES (?,?,?,?,?,?,?,?)",
            (acct["id"], acct["first_name"], acct["last_name"], acct["title"],
             acct["email"], acct["work_phone"], acct["mobile_phone"], now_iso()))


def _set_primary_contact(conn, account_id, person: dict):
    conn.execute(
        "UPDATE accounts SET first_name=?, last_name=?, title=?, email=?, "
        "work_phone=?, mobile_phone=?, updated_at=? WHERE id=?",
        (person["first_name"], person["last_name"], person["title"],
         person["email"], person["work_phone"], person["mobile_phone"],
         now_iso(), account_id))


@app.route("/accounts/<int:account_id>/contacts/add", methods=["POST"])
def add_contact(account_id):
    person = {c: request.form.get(c, "").strip() for c in CONTACT_COLS}
    if not (person["first_name"] or person["last_name"]):
        flash("Contact needs at least a first or last name.", "danger")
        return redirect(url_for("account_detail", account_id=account_id))
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        make_primary = bool(request.form.get("set_primary")) or not _has_primary_contact(acct)
        if make_primary:
            _demote_primary_to_contact(conn, acct)
            _set_primary_contact(conn, account_id, person)
        else:
            conn.execute(
                "INSERT INTO contacts (account_id, first_name, last_name, title, "
                "email, work_phone, mobile_phone, created_at) VALUES (?,?,?,?,?,?,?,?)",
                (account_id, *person.values(), now_iso()))
        conn.commit()
    finally:
        conn.close()
    flash(f"Contact {person['first_name']} {person['last_name']} added"
          + (" as primary." if make_primary else "."), "success")
    return redirect(url_for("account_detail", account_id=account_id))


@app.route("/accounts/<int:account_id>/contacts/<int:contact_id>/promote", methods=["POST"])
def promote_contact(account_id, contact_id):
    """Swap a contact with the account's primary-contact fields."""
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        row = conn.execute("SELECT * FROM contacts WHERE id=? AND account_id=?",
                           (contact_id, account_id)).fetchone()
        if row:
            _demote_primary_to_contact(conn, acct)
            _set_primary_contact(conn, account_id, dict(row))
            conn.execute("DELETE FROM contacts WHERE id=?", (contact_id,))
            conn.commit()
            flash(f"{row['first_name']} {row['last_name']} is now the primary contact.",
                  "success")
    finally:
        conn.close()
    return redirect(url_for("account_detail", account_id=account_id))


@app.route("/accounts/<int:account_id>/contacts/<int:contact_id>/delete", methods=["POST"])
def delete_contact(account_id, contact_id):
    conn = get_db()
    try:
        conn.execute("DELETE FROM contacts WHERE id=? AND account_id=?",
                     (contact_id, account_id))
        conn.commit()
    finally:
        conn.close()
    flash("Contact removed.", "warning")
    return redirect(url_for("account_detail", account_id=account_id))


# --------------------------------------------------- Templates & scripts

PLACEHOLDER_LABELS = {
    "first_name": "first name", "last_name": "last name", "title": "title",
    "company": "company", "num_properties": "total properties",
    "matching_properties": "matching properties",
    "email": "email", "my_name": "your name", "my_company": "your company",
    "my_phone": "your phone number",
}


def _get_settings(conn) -> dict:
    return {r["key"]: r["value"] for r in conn.execute("SELECT key, value FROM settings")}


def render_script(text: str, acct, settings: dict):
    """Fill {placeholders} with account + settings values.

    Returns (rendered_text, missing_labels). Missing values stay visible as
    bracketed hints like [first name] so nothing gets sent half-baked
    unnoticed.
    """
    values = {
        "first_name": acct["first_name"],
        "last_name": acct["last_name"],
        "title": acct["title"],
        "company": acct["company_name"],
        "num_properties": str(acct["num_properties"]) if acct["num_properties"] else "",
        "matching_properties": (str(acct["matching_properties"])
                                if acct["matching_properties"] else ""),
        "email": acct["email"],
        "my_name": settings.get("my_name", ""),
        "my_company": settings.get("my_company", ""),
        "my_phone": settings.get("my_phone", ""),
    }
    missing = []

    def sub(match):
        key = match.group(1)
        if key not in values:
            return match.group(0)
        if values[key]:
            return values[key]
        label = PLACEHOLDER_LABELS.get(key, key)
        if label not in missing:
            missing.append(label)
        return f"[{label}]"

    return re.sub(r"\{(\w+)\}", sub, text or ""), missing


@app.route("/templates", methods=["GET"])
def templates_page():
    conn = get_db()
    try:
        templates = conn.execute(
            "SELECT * FROM templates ORDER BY sort_order, id").fetchall()
        settings = _get_settings(conn)
    finally:
        conn.close()
    return render_template("templates.html", templates=templates,
                           settings=settings,
                           cadence_steps=[s for _, s in cadence.CADENCE_STEPS])


@app.route("/templates/<int:template_id>", methods=["POST"])
def save_template(template_id):
    steps = ",".join(request.form.getlist("steps"))
    conn = get_db()
    try:
        conn.execute(
            "UPDATE templates SET name=?, kind=?, steps=?, subject=?, body=?, "
            "updated_at=? WHERE id=?",
            (request.form.get("name", "").strip() or "Untitled",
             request.form.get("kind", "email"), steps,
             request.form.get("subject", "").strip(),
             request.form.get("body", ""), now_iso(), template_id))
        conn.commit()
    finally:
        conn.close()
    flash("Template saved.", "success")
    return redirect(url_for("templates_page"))


@app.route("/settings", methods=["POST"])
def save_settings():
    conn = get_db()
    try:
        for key in ("my_name", "my_company", "my_phone"):
            conn.execute(
                "INSERT INTO settings (key, value) VALUES (?,?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                (key, request.form.get(key, "").strip()))
        conn.commit()
    finally:
        conn.close()
    flash("Your info saved — it now fills into every script.", "success")
    return redirect(url_for("templates_page"))


def _build_scripts(acct, settings, rows, step=""):
    """Render templates for an account, optionally filtered to a cadence step."""
    scripts = []
    for t in rows:
        t_steps = [s.strip() for s in (t["steps"] or "").split(",") if s.strip()]
        if step and step not in t_steps:
            continue
        subject, missing_s = render_script(t["subject"], acct, settings)
        body, missing_b = render_script(t["body"], acct, settings)
        action_url = ""
        if t["kind"] == "email" and acct["email"]:
            action_url = (f"mailto:{acct['email']}?subject={quote(subject)}"
                          f"&body={quote(body)}")
        elif t["kind"] == "call" and (acct["work_phone"] or acct["mobile_phone"]):
            action_url = f"tel:{acct['work_phone'] or acct['mobile_phone']}"
        elif t["kind"] == "text" and acct["mobile_phone"]:
            action_url = f"sms:{acct['mobile_phone']}?body={quote(body)}"
        scripts.append({
            "template": t, "subject": subject, "body": body,
            "missing": list(dict.fromkeys(missing_s + missing_b)),
            "action_url": action_url, "log_steps": t_steps,
        })
    return scripts


@app.route("/accounts/<int:account_id>/scripts")
def account_scripts(account_id):
    step = request.args.get("step", "")
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        settings = _get_settings(conn)
        rows = conn.execute(
            "SELECT * FROM templates ORDER BY sort_order, id").fetchall()
    finally:
        conn.close()
    scripts = _build_scripts(acct, settings, rows, step)
    return render_template("scripts.html", account=acct, scripts=scripts, step=step)


# ------------------------------------------------------------------- Import

@app.route("/import", methods=["GET", "POST"])
def import_page():
    result = None
    if request.method == "POST":
        file = request.files.get("file")
        if not file or not file.filename:
            flash("Choose a .xlsx or .csv file first.", "danger")
        else:
            per_day_raw = request.form.get("per_day", "").strip()
            per_day = int(per_day_raw) if per_day_raw.isdigit() and int(per_day_raw) > 0 else None
            conn = get_db()
            try:
                result = importer.import_accounts(conn, file, per_day=per_day)
                msg = f"Imported {result['imported']} account(s)."
                if per_day:
                    msg += (f" Cadence starts staggered at {per_day}/business day "
                            f"through {result['last_start_date']}.")
                flash(msg, "success")
            except ValueError as e:
                flash(str(e), "danger")
            except Exception as e:
                flash(f"Import failed: {e}", "danger")
            finally:
                conn.close()
    return render_template("import.html", result=result)


@app.route("/repace", methods=["POST"])
def repace():
    """Re-stagger cadence starts for untouched in-cadence accounts.

    Only accounts that are active in cadence AND have no logged interactions
    or checked-off steps are re-paced — anything already being worked keeps
    its dates.
    """
    per_day_raw = request.form.get("per_day", "").strip()
    if not (per_day_raw.isdigit() and int(per_day_raw) > 0):
        flash("Enter how many accounts should start per business day.", "danger")
        return redirect(url_for("import_page"))
    per_day = int(per_day_raw)
    conn = get_db()
    try:
        rows = conn.execute(
            """SELECT id FROM accounts a
               WHERE prospecting_status = ? AND pipeline_milestone = ?
                 AND NOT EXISTS (SELECT 1 FROM interactions i WHERE i.account_id = a.id)
                 AND NOT EXISTS (SELECT 1 FROM cadence_dismissals d WHERE d.account_id = a.id)
               ORDER BY id""",
            (cadence.ACTIVE_STATUS, cadence.ACTIVE_MILESTONE)).fetchall()
        start = importer._next_business_day(datetime.now().date())
        ts = now_iso()
        for i, row in enumerate(rows):
            if i > 0 and i % per_day == 0:
                start = importer._next_business_day(start + timedelta(days=1))
            conn.execute("UPDATE accounts SET cadence_start=?, updated_at=? WHERE id=?",
                         (start.isoformat(), ts, row["id"]))
        conn.commit()
    finally:
        conn.close()
    if rows:
        flash(f"Re-paced {len(rows)} untouched account(s) at {per_day}/business day, "
              f"{importer._next_business_day(datetime.now().date()).isoformat()} "
              f"through {start.isoformat()}.", "success")
    else:
        flash("No untouched in-cadence accounts to re-pace.", "warning")
    return redirect(url_for("import_page"))


@app.route("/import/contacts", methods=["POST"])
def import_contacts_route():
    """Bulk-attach a ZoomInfo (or similar) contact export to existing accounts."""
    file = request.files.get("file")
    if not file or not file.filename:
        flash("Choose a .xlsx or .csv contact file first.", "danger")
        return render_template("import.html", result=None)
    conn = get_db()
    try:
        contact_result = importer.import_contacts(conn, file)
        flash(f"Attached {contact_result['attached']} contact(s) to "
              f"{contact_result['companies_matched']} account(s).", "success")
    except ValueError as e:
        flash(str(e), "danger")
        contact_result = None
    except Exception as e:
        flash(f"Contact import failed: {e}", "danger")
        contact_result = None
    finally:
        conn.close()
    return render_template("import.html", result=None, contact_result=contact_result)


@app.route("/import/template")
def import_template():
    csv = ("Company Name,First Name,Last Name,Title,Number of Properties,"
           "Email,Work Phone,Mobile Phone,Notes\r\n")
    return send_file(io.BytesIO(csv.encode()), mimetype="text/csv",
                     as_attachment=True, download_name="crm_import_template.csv")


# ------------------------------------------------------------------ Startup

def _local_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))  # no traffic is sent; just picks the LAN interface
        ip = s.getsockname()[0]
        s.close()
        return ip
    except OSError:
        return "127.0.0.1"


def print_network_instructions():
    ip = _local_ip()
    print(f"""
============================================================
  Roof Coating CRM is starting
============================================================
  On this computer:      http://localhost:{PORT}
  On your phone/tablet:  http://{ip}:{PORT}
                         (must be on the same Wi-Fi network)

  If that address doesn't work, find your computer's local
  IP address manually:
    - Windows:  open Command Prompt, run  ipconfig
                (look for "IPv4 Address", e.g. 192.168.1.23)
    - Mac:      System Settings > Wi-Fi > Details,
                or run  ipconfig getifaddr en0
    - Linux:    run  hostname -I

  Then browse to  http://<that-ip>:{PORT}  from your phone.
  Tip: your firewall may need to allow inbound port {PORT}.
============================================================
""")


if __name__ == "__main__":
    init_db()
    backup = backup_db()
    if backup:
        print(f"  Daily backup saved: {backup}")
    print_network_instructions()
    try:
        from waitress import serve
        print("  Server: waitress (production WSGI)\n")
        serve(app, host="0.0.0.0", port=PORT, threads=8)
    except ImportError:
        print("  Server: Flask dev server (pip install waitress for the "
              "production server)\n")
        app.run(host="0.0.0.0", port=PORT, debug=False)
