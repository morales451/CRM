"""Roof Coating CRM — lightweight local Flask app.

Run:  python3 app.py
Then open http://<your-local-ip>:8000 from any device on your Wi-Fi.
"""

import csv
import io
import re
import socket
import uuid
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import quote

from flask import (Flask, flash, redirect, render_template, request,
                   send_file, session, url_for)

import cadence
import db as db_module
import excel_export
import importer
import undo as undo_module
import warranty_calc
from db import (ARCHIVE_REASONS, DEFAULT_PROJECT_TASKS, INTERACTION_TYPES,
                INVOICE_STATUSES,
                PIPELINE_MILESTONES, PREFERRED_CONTACT_METHODS,
                PROJECT_STATUSES, PROSPECTING_STATUSES,
                backup_db, get_db, init_db, now_iso, today_iso)

app = Flask(__name__)
# Local single-user app: the session cookie only carries flash messages and
# the id of the last undoable action, never credentials.
app.secret_key = "local-crm-flash-messages"
PORT = 8000
PHOTO_MAX_DIM = 1600  # uploaded photos are resized to fit the report/PDF


def photo_dir():
    """Where bid photos live. Read through db_module every time rather than
    copied into a constant, so this and undo.py can never disagree about
    which folder holds the images."""
    return db_module.UPLOAD_DIR


@app.template_filter("dt")
def format_datetime(iso_str):
    """Render an ISO timestamp as 'Sep 17, 2026 2:03 PM'."""
    try:
        return datetime.fromisoformat(iso_str).strftime("%b %d, %Y %I:%M %p")
    except (ValueError, TypeError):
        return iso_str or ""


@app.template_filter("tel")
def clean_tel(phone):
    """Phone number as dialable digits for tel:/sms: links."""
    return re.sub(r"[^\d+]", "", phone or "")


_last_backup_date = None


@app.before_request
def _daily_backup():
    """Keep daily backups flowing even when the app runs for weeks
    without a restart (backup_db itself is a no-op after the first
    call each day)."""
    global _last_backup_date
    t = today_iso()
    if _last_backup_date != t:
        _last_backup_date = t
        try:
            backup_db()
        except OSError:
            pass


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
        "PROJECT_STATUSES": PROJECT_STATUSES,
        "ARCHIVE_REASONS": ARCHIVE_REASONS,
        "INVOICE_STATUSES": INVOICE_STATUSES,
        "today": today_iso(),
    }


@app.template_filter("money")
def format_money(value):
    if value is None:
        return "—"
    return f"${value:,.0f}" if value == int(value) else f"${value:,.2f}"


# ---------------------------------------------------------------------- Undo
#
# Every destructive action snapshots what it changed into the undo_log table,
# then drops the record's id into the session. base.html renders an Undo bar
# for as long as that record is unused, so a mis-tap is a click away from
# being put back instead of a restore-from-backup.

UNDO_SESSION_KEY = "undo_id"


def _offer_undo(conn, label: str, ops: list[dict]) -> None:
    """Record an undo for the action about to be committed.

    Call BEFORE conn.commit() so the snapshot and the change land together —
    a crash between them would otherwise leave an undo pointing at nothing.
    """
    undo_id = undo_module.record(conn, label, ops)
    if undo_id is not None:
        # The label rides along in the session so drawing the Undo bar costs
        # nothing: this runs on every page render, and opening a second
        # database connection here (while the page's own is still open) made
        # every click pay for a lock it didn't need.
        session[UNDO_SESSION_KEY] = {"id": undo_id, "label": label}


@app.context_processor
def inject_undo():
    """The pending undo, if the last action left one. Session-only — the
    record itself is re-checked (and may be refused as used or expired) when
    the button is actually pressed."""
    pending = session.get(UNDO_SESSION_KEY)
    if isinstance(pending, dict) and pending.get("id"):
        return {"pending_undo": pending}
    return {"pending_undo": None}


@app.route("/undo/<int:undo_id>", methods=["POST"])
def undo_action(undo_id):
    conn = get_db()
    try:
        label = undo_module.restore(conn, undo_id)
    finally:
        conn.close()
    session.pop(UNDO_SESSION_KEY, None)
    if label:
        flash(f"Undone: {label}.", "success")
    else:
        flash("That action can no longer be undone.", "warning")
    return redirect(request.form.get("next") or url_for("dashboard"))


def _account_or_404(conn, account_id):
    acct = conn.execute("SELECT * FROM accounts WHERE id = ?", (account_id,)).fetchone()
    if acct is None:
        from flask import abort
        abort(404)
    return acct


def _account_fields_from_form(form, previous=None):
    def int_or_none(name):
        prior = previous[name] if previous is not None else None
        if name not in form:
            return prior
        value = _parse_int(form.get(name), on_error=_MISSING)
        return prior if value is _MISSING else value

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
        "linkedin_url": form.get("linkedin_url", "").strip(),
        "seniority": form.get("seniority", "").strip(),
        "preferred_contact": form.get("preferred_contact", "Unknown"),
        "notes": form.get("notes", "").strip(),
        "prospecting_status": form.get("prospecting_status", "Prospecting"),
        "pipeline_milestone": form.get("pipeline_milestone", "None / In Cadence"),
    }


# ---------------------------------------------------------------- Dashboard

STALE_DAYS = 14
STALE_EXCLUDED_MILESTONES = ("None / In Cadence", "Closed Won", "Closed Lost", "On Hold")

TASK_ORDERS = ("priority", "due")


def _task_order(conn) -> str:
    """How the dashboard and queue order today's work.

    'priority' (default) works accounts with the most criteria-matching
    buildings first — biggest portfolios are the biggest deals.
    'due' keeps the classic oldest-due-first cadence order.
    """
    row = conn.execute(
        "SELECT value FROM settings WHERE key='task_order'").fetchone()
    value = (row["value"] if row else "") or "priority"
    return value if value in TASK_ORDERS else "priority"


@app.route("/settings/task-order", methods=["POST"])
def set_task_order():
    order = request.form.get("order", "priority")
    if order in TASK_ORDERS:
        conn = get_db()
        try:
            conn.execute(
                "INSERT INTO settings (key, value) VALUES ('task_order', ?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (order,))
            conn.commit()
        finally:
            conn.close()
    return redirect(request.form.get("next") or url_for("dashboard"))


def _due_followups(conn, order: str = "due"):
    """Follow-ups whose date has arrived. In priority order, accounts with
    more matching buildings come first."""
    sort = ("COALESCE(matching_properties, 0) DESC, next_follow_up"
            if order == "priority" else "next_follow_up")
    return conn.execute(
        f"SELECT * FROM accounts WHERE COALESCE(archived_at, '') = '' AND next_follow_up != '' "
        f"AND next_follow_up <= ? "
        f"ORDER BY {sort}, company_name COLLATE NOCASE",
        (today_iso(),)).fetchall()


def _top_priority_accounts(conn, limit: int = 8):
    """Highest-potential accounts still in play, most matching buildings first."""
    return conn.execute(
        """SELECT a.*,
                  (SELECT MAX(created_at) FROM interactions i
                   WHERE i.account_id = a.id) AS last_activity
           FROM accounts a
           WHERE COALESCE(a.archived_at, '') = ''
             AND COALESCE(a.matching_properties, 0) > 0
             AND a.prospecting_status != 'Not Interested'
             AND a.pipeline_milestone NOT IN ('Closed Won', 'Closed Lost')
           ORDER BY a.matching_properties DESC, a.company_name COLLATE NOCASE
           LIMIT ?""", (limit,)).fetchall()


def _stale_deals(conn):
    """Pipeline accounts with no touch in STALE_DAYS days and no follow-up set."""
    ph = ",".join("?" * len(STALE_EXCLUDED_MILESTONES))
    rows = conn.execute(
        f"""SELECT a.*, COALESCE(
                (SELECT MAX(created_at) FROM interactions i WHERE i.account_id = a.id),
                a.updated_at) AS last_touch
            FROM accounts a
            WHERE COALESCE(a.archived_at, '') = ''
              AND a.pipeline_milestone NOT IN ({ph})
              AND a.prospecting_status != 'Not Interested'
              AND (a.next_follow_up IS NULL OR a.next_follow_up = '')""",
        STALE_EXCLUDED_MILESTONES).fetchall()
    cutoff = (datetime.now() - timedelta(days=STALE_DAYS)).date().isoformat()
    return [r for r in rows if (r["last_touch"] or "")[:10] <= cutoff]


def _today_progress(conn, remaining: int) -> dict:
    """"9 of 14 done today" — the single number that says whether the day's
    outreach actually happened. Done = anything logged or checked off today;
    remaining = reminders and follow-ups still sitting on the dashboard."""
    t = today_iso()
    done = conn.execute(
        "SELECT COUNT(*) c FROM interactions i JOIN accounts a ON a.id = i.account_id "
        "WHERE COALESCE(a.archived_at, '') = '' AND i.created_at LIKE ?",
        (t + "%",)).fetchone()["c"]
    done += conn.execute(
        "SELECT COUNT(*) c FROM cadence_dismissals d "
        "JOIN accounts a ON a.id = d.account_id "
        "WHERE COALESCE(a.archived_at, '') = '' AND d.dismissed_at LIKE ?",
        (t + "%",)).fetchone()["c"]
    total = done + remaining
    return {"done": done, "remaining": remaining, "total": total,
            "pct": round(100 * done / total) if total else 0,
            # NOT "clear": Jinja would resolve progress.clear to dict.clear,
            # a bound method that is always truthy.
            "all_done": total > 0 and remaining == 0}


# How many rows of each list the dashboard draws. Past a few hundred accounts
# the page was rendering every due reminder — 1,500+ forms and over a megabyte
# of HTML on the page you land on after EVERY action, which is what made the
# whole app feel slow. The Queue is the right tool for working a long list;
# the dashboard only has to show you the top of it. "?all=1" still renders
# everything for anyone who wants it.
DASHBOARD_ROWS = 25


@app.route("/")
def dashboard():
    conn = get_db()
    try:
        order = _task_order(conn)
        show_all = request.args.get("all") == "1"
        cut = (lambda rows: rows) if show_all else (lambda rows: rows[:DASHBOARD_ROWS])
        all_reminders = cadence.get_due_reminders(conn, order=order)
        all_followups = _due_followups(conn, order)
        all_stale = _stale_deals(conn)
        totals = {"reminders": len(all_reminders), "followups": len(all_followups),
                  "stale": len(all_stale)}
        reminders, followups, stale = (cut(all_reminders), cut(all_followups),
                                       cut(all_stale))
        owed = _outstanding_invoices(conn)
        top_priority = _top_priority_accounts(conn)
        stats = {
            "total_accounts": conn.execute(
                "SELECT COUNT(*) c FROM accounts WHERE COALESCE(archived_at, '') = ''").fetchone()["c"],
            "active_prospects": conn.execute(
                "SELECT COUNT(*) c FROM accounts WHERE COALESCE(archived_at, '') = '' "
                "AND prospecting_status NOT IN ('Not Interested') AND pipeline_milestone "
                "NOT IN ('Closed Won','Closed Lost')").fetchone()["c"],
            "in_cadence": conn.execute(
                "SELECT COUNT(*) c FROM accounts WHERE COALESCE(archived_at, '') = '' "
                "AND prospecting_status = ? AND pipeline_milestone = ?",
                (cadence.ACTIVE_STATUS, cadence.ACTIVE_MILESTONE)).fetchone()["c"],
            "closed_won": conn.execute(
                "SELECT COUNT(*) c FROM accounts "
                f"WHERE COALESCE(archived_at, '') = '' AND pipeline_milestone = 'Closed Won'").fetchone()["c"],
            "matching_due": sum(r["matching_properties"] or 0 for r in all_reminders),
        }
        recent = conn.execute(
            """SELECT i.*, a.company_name FROM interactions i
               JOIN accounts a ON a.id = i.account_id
               WHERE COALESCE(a.archived_at, '') = ''
               ORDER BY i.created_at DESC, i.id DESC LIMIT 10""").fetchall()
        progress = _today_progress(conn, totals["reminders"] + totals["followups"])
        return render_template("dashboard.html", reminders=reminders,
                               followups=followups, stale=stale, owed=owed,
                               stats=stats, recent=recent, today=today_iso(),
                               order=order, top_priority=top_priority,
                               progress=progress, totals=totals,
                               show_all=show_all)
    finally:
        conn.close()


def _account_exists(conn, account_id) -> bool:
    return conn.execute("SELECT 1 FROM accounts WHERE id = ?",
                        (account_id,)).fetchone() is not None


@app.route("/reminders/dismiss", methods=["POST"])
def dismiss_reminder():
    account_id = request.form["account_id"]
    step_type = request.form["step_type"]
    conn = get_db()
    try:
        if not _account_exists(conn, account_id):
            flash("That account no longer exists.", "danger")
            return redirect(request.form.get("next") or url_for("dashboard"))
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
        if not _account_exists(conn, account_id):
            flash("That account no longer exists.", "danger")
            return redirect(request.form.get("next") or url_for("dashboard"))
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


def _build_queue(conn, order: str | None = None):
    """Today's work: due cadence steps + due follow-ups.

    Ordered by the saved task order — biggest matching portfolios first by
    default, so the highest-potential accounts get worked while you're fresh.
    """
    order = order or _task_order(conn)
    tasks = [{"kind": "cadence", "due": r["due_date"], "account_id": r["account_id"],
              "matching": r["matching_properties"] or 0, "reminder": r}
             for r in cadence.get_due_reminders(conn, order=order)]
    tasks += [{"kind": "followup", "due": a["next_follow_up"], "account_id": a["id"],
               "matching": a["matching_properties"] or 0,
               "note": a["follow_up_note"]} for a in _due_followups(conn, order)]
    if order == "priority":
        tasks.sort(key=lambda t: (-t["matching"], t["due"]))
    else:
        tasks.sort(key=lambda t: (t["due"], -t["matching"]))
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
        try:
            pos = max(0, min(int(request.args.get("pos", 0)), total - 1))
        except (TypeError, ValueError):
            pos = 0
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


DEFAULT_DAILY_GOAL = 20


def _daily_goal(conn) -> int:
    """How many logged touches count as a full day's outreach. 0 turns the
    goal line off."""
    row = conn.execute(
        "SELECT value FROM settings WHERE key='daily_goal'").fetchone()
    try:
        return max(0, int((row["value"] if row else "").strip()))
    except (TypeError, ValueError, AttributeError):
        return DEFAULT_DAILY_GOAL


@app.route("/insights")
def insights():
    """Critical numbers: activity trend, funnel, cadence completion, breakdowns."""
    conn = get_db()
    try:
        today = datetime.now().date()

        def count(sql, *params):
            return conn.execute(sql, params).fetchone()[0]

        total_accounts = count(f"SELECT COUNT(*) FROM accounts WHERE COALESCE(archived_at, '') = ''")
        won = count(f"SELECT COUNT(*) FROM accounts WHERE COALESCE(archived_at, '') = '' "
                    "AND pipeline_milestone='Closed Won'")
        lost = count(f"SELECT COUNT(*) FROM accounts WHERE COALESCE(archived_at, '') = '' "
                     "AND pipeline_milestone='Closed Lost'")
        tiles = {
            "tasks_due": len(cadence.get_due_reminders(conn)) + len(_due_followups(conn)),
            "total_accounts": total_accounts,
            "active_prospects": count(
                f"SELECT COUNT(*) FROM accounts WHERE COALESCE(archived_at, '') = '' "
                "AND prospecting_status != 'Not Interested' "
                "AND pipeline_milestone NOT IN ('Closed Won','Closed Lost')"),
            "in_cadence": count(
                f"SELECT COUNT(*) FROM accounts WHERE COALESCE(archived_at, '') = '' "
                "AND prospecting_status=? AND pipeline_milestone=?",
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
                "SELECT i.created_at FROM interactions i JOIN accounts a "
                "ON a.id = i.account_id WHERE COALESCE(a.archived_at, '') = '' "
                "AND i.created_at >= ?", (since,)):
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

        # Interactions per DAY, last 14 days, against your daily goal. The
        # weekly chart shows the trend; this one answers "am I actually
        # making the calls?" while there's still time to fix it.
        goal = _daily_goal(conn)
        days = [today - timedelta(days=i) for i in range(13, -1, -1)]
        day_counts = {d.isoformat(): 0 for d in days}
        for (created,) in conn.execute(
                "SELECT i.created_at FROM interactions i JOIN accounts a "
                "ON a.id = i.account_id WHERE COALESCE(a.archived_at, '') = '' "
                "AND i.created_at >= ?", (days[0].isoformat(),)):
            key = created[:10]
            if key in day_counts:
                day_counts[key] += 1
        day_peak = max(max(day_counts.values()), goal, 1)
        daily = [{"label": d.strftime("%a"), "date": d.isoformat(),
                  "day": d.day, "count": day_counts[d.isoformat()],
                  "pct": round(100 * day_counts[d.isoformat()] / day_peak),
                  "hit": goal > 0 and day_counts[d.isoformat()] >= goal,
                  "weekend": d.weekday() >= 5, "is_today": d == today}
                 for d in days]
        workdays = [d for d in daily if not d["weekend"]]
        activity = {
            "goal": goal,
            "days": daily,
            "goal_pct": round(100 * goal / day_peak) if goal else 0,
            "today_count": daily[-1]["count"],
            "hit_days": sum(1 for d in workdays if d["hit"]),
            "workdays": len(workdays),
            "avg": round(sum(d["count"] for d in workdays) / len(workdays), 1)
                   if workdays else 0,
        }

        # Pipeline funnel (the cadence pool would dwarf it, so it's a tile instead)
        pipeline_bars = _bar_items([
            (m, count(f"SELECT COUNT(*) FROM accounts WHERE COALESCE(archived_at, '') = '' "
                      "AND pipeline_milestone=?", m))
            for m in PIPELINE_MILESTONES if m != "None / In Cadence"])

        # How far accounts get through the cadence (distinct accounts per step)
        cadence_bars = _bar_items([
            (f"Day {day}: {step}",
             count("SELECT COUNT(DISTINCT i.account_id) FROM interactions i "
                   "JOIN accounts a ON a.id = i.account_id "
                   "WHERE COALESCE(a.archived_at, '') = '' AND i.interaction_type=?", step))
            for day, step in cadence.CADENCE_STEPS], total=total_accounts)

        status_bars = _bar_items([
            (s, count(f"SELECT COUNT(*) FROM accounts WHERE COALESCE(archived_at, '') = '' "
                      "AND prospecting_status=?", s))
            for s in PROSPECTING_STATUSES], total=total_accounts)

        month_ago = (today - timedelta(days=30)).isoformat()
        type_bars = _bar_items(conn.execute(
            "SELECT i.interaction_type, COUNT(*) FROM interactions i "
            "JOIN accounts a ON a.id = i.account_id "
            "WHERE COALESCE(a.archived_at, '') = '' AND i.created_at >= ? "
            "GROUP BY i.interaction_type ORDER BY COUNT(*) DESC", (month_ago,)).fetchall())

        money = None
        if count("SELECT COUNT(*) FROM projects"):
            inv = conn.execute(
                """SELECT COALESCE(SUM(CASE WHEN status != 'Draft' THEN amount END), 0) AS invoiced,
                          COALESCE(SUM(CASE WHEN status = 'Paid' THEN amount END), 0) AS paid,
                          COALESCE(SUM(CASE WHEN status = 'Sent' THEN amount END), 0) AS outstanding
                   FROM invoices""").fetchone()
            money = {
                "contracted": conn.execute(
                    "SELECT COALESCE(SUM(contract_amount), 0) FROM projects").fetchone()[0],
                "invoiced": inv["invoiced"], "paid": inv["paid"],
                "outstanding": inv["outstanding"],
                "projects": count("SELECT COUNT(*) FROM projects"),
                "active_projects": count(
                    "SELECT COUNT(*) FROM projects WHERE status NOT IN ('Closed')"),
            }
    finally:
        conn.close()
    return render_template("insights.html", tiles=tiles, weekly=weekly,
                           activity=activity,
                           pipeline_bars=pipeline_bars, cadence_bars=cadence_bars,
                           status_bars=status_bars, type_bars=type_bars,
                           money=money)


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
                   FROM accounts a WHERE COALESCE(a.archived_at, '') = '' AND a.pipeline_milestone = ?
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


# ------------------------------------------------------- Bids & roof reports


_ONES = ["", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT",
         "NINE", "TEN", "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN",
         "SIXTEEN", "SEVENTEEN", "EIGHTEEN", "NINETEEN"]
_TENS = ["", "", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY",
         "EIGHTY", "NINETY"]


def dollars_in_words(amount) -> str:
    """15000 -> 'FIFTEEN THOUSAND DOLLARS' (contract-style quotation line)."""
    def words(n):
        if n < 20:
            return _ONES[n]
        if n < 100:
            return (_TENS[n // 10] + ("-" + _ONES[n % 10] if n % 10 else "")).strip()
        if n < 1000:
            return (_ONES[n // 100] + " HUNDRED"
                    + (" " + words(n % 100) if n % 100 else ""))
        for div, name in ((1_000_000, "MILLION"), (1_000, "THOUSAND")):
            if n >= div:
                return (words(n // div) + " " + name
                        + (" " + words(n % div) if n % div else ""))
        return ""
    n = int(round(amount or 0))
    if n <= 0:
        return ""
    return words(n) + " DOLLARS"



def _support_map():
    """{system: {acrylic_type: {roof_type: [warranty years]}}} for the editor,
    so only combinations that exist can be picked."""
    out = {}
    for system in warranty_calc.COATING_SYSTEMS:
        types = warranty_calc.ACRYLIC_TYPES if system == "Acrylic" else ["Standard"]
        out[system] = {
            t: {rt: warranty_calc.supported_warranties(system, rt, t)
                for rt in warranty_calc.supported_roof_types(system, t)}
            for t in types}
    return out


def _pricing_settings(conn):
    """Per-sq-ft pricing rules, editable on the Templates page."""
    rows = {r["key"]: r["value"] for r in conn.execute(
        "SELECT key, value FROM settings WHERE key LIKE 'price_%'")}
    pricing = dict(warranty_calc.DEFAULT_PRICING)
    for field, key in (("capsheet_base", "price_capsheet_base"),
                       ("other_base", "price_other_base"),
                       ("add_15", "price_add_15"), ("add_20", "price_add_20")):
        try:
            if rows.get(key, "").strip():
                pricing[field] = float(rows[key])
        except ValueError:
            pass
    return pricing


def _bid_suggested_price(conn, bid, plan):
    """Suggested sell price for a bid, based on its coated area."""
    net = plan["net_sqft"] if plan else max(
        0, (bid["roof_size_sqft"] or 0) - (bid["deduction_sqft"] or 0))
    if not net:
        return None
    return warranty_calc.suggested_price(
        net, bid["roof_type"] or "Capsheet", bid["warranty_years"] or 10,
        _pricing_settings(conn))


def _bid_plan(bid, warranty_years=None):
    """Materials plan for a bid, using the warranty calculator's rates."""
    return warranty_calc.calculate(
        bid["roof_size_sqft"] or 0,
        coating_system=bid["coating_system"] or "Silicone",
        roof_type=bid["roof_type"] or "Capsheet",
        warranty_years=warranty_years or bid["warranty_years"] or 10,
        acrylic_system_type=bid["acrylic_system_type"] or "Standard",
        deduction_sqft=bid["deduction_sqft"] or 0,
        linear_feet=bid["linear_feet"] or 0,
        waste_pct=bid["waste_pct"] or 0,
        stretch_pct=bid["stretch_pct"] or 0,
        passed_adhesion=bool(bid["passed_adhesion"]),
        has_rust=bool(bid["has_rust"]),
        rust_prime_method=bid["rust_prime_method"] or "field",
        topcoat=bid["selected_topcoat"] or "",
        basecoat=bid["selected_basecoat"] or "",
        butter_grade=bid["selected_butter_grade"] or "")


def _bid_warranty_options(bid):
    """The same roof at 10/15/20 years, for the comparison table."""
    out = []
    for years in warranty_calc.WARRANTY_YEARS:
        plan = _bid_plan(bid, warranty_years=years)
        if plan:
            out.append(plan)
    return out


def _bid_or_404(conn, bid_id):
    bid = conn.execute(
        """SELECT b.*, a.company_name, a.first_name, a.last_name, a.title,
                  a.email, a.work_phone, a.mobile_phone, a.notes AS account_notes
           FROM bids b JOIN accounts a ON a.id = b.account_id
           WHERE b.id = ?""", (bid_id,)).fetchone()
    if bid is None:
        from flask import abort
        abort(404)
    return bid


@app.route("/accounts/<int:account_id>/bids/new", methods=["POST"])
def new_bid(account_id):
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        ts = now_iso()
        address = (request.form.get("roof_address", "").strip()
                   or _address_from_notes(acct["notes"]))
        sqft_raw = request.form.get("roof_size_sqft", "").strip()
        surface = request.form.get("surface_type", "").strip()
        cur = conn.execute(
            """INSERT INTO bids (account_id, roof_address, roof_size_sqft,
               surface_type, roof_type, assessment_date, created_at, updated_at)
               VALUES (?,?,?,?,?,?,?,?)""",
            (account_id, address, int(sqft_raw) if sqft_raw.isdigit() else None,
             surface, warranty_calc.guess_roof_type(surface), today_iso(), ts, ts))
        conn.commit()
        bid_id = cur.lastrowid
    finally:
        conn.close()
    flash("Roof report created — fill in the details, add photos, then open "
          "the report.", "success")
    return redirect(url_for("bid_edit", bid_id=bid_id))


@app.route("/bids/<int:bid_id>")
def bid_edit(bid_id):
    conn = get_db()
    try:
        bid = _bid_or_404(conn, bid_id)
        photos = conn.execute(
            "SELECT * FROM bid_photos WHERE bid_id = ? ORDER BY sort_order, id",
            (bid_id,)).fetchall()
        plan = _bid_plan(bid)
        suggested = _bid_suggested_price(conn, bid, plan)
        pricing = _pricing_settings(conn)
    finally:
        conn.close()
    return render_template("bid_edit.html", bid=bid, photos=photos,
                           plan=plan, suggested=suggested, pricing=pricing,
                           options=_bid_warranty_options(bid),
                           COATING_SYSTEMS=warranty_calc.COATING_SYSTEMS,
                           ACRYLIC_TYPES=warranty_calc.ACRYLIC_TYPES,
                           ROOF_TYPES=warranty_calc.ROOF_TYPES,
                           WARRANTY_YEARS=warranty_calc.WARRANTY_YEARS,
                           CATALOG=warranty_calc.PRODUCT_CATALOG,
                           SUPPORT=_support_map())


@app.route("/bids/<int:bid_id>/edit", methods=["POST"])
def save_bid(bid_id):
    conn = get_db()
    try:
        bid = _bid_or_404(conn, bid_id)
        rejected = []

        def num(name, previous=None):
            """Parse a number, keeping the stored value if it's unreadable."""
            if name not in request.form:
                return previous
            value = _parse_int(request.form.get(name), on_error=_MISSING)
            if value is _MISSING:
                rejected.append(name.replace("_", " "))
                return previous
            return value
        def pct(name, previous=0.0):
            if name not in request.form:
                return previous
            raw = (request.form.get(name) or "").strip()
            if not raw:
                return 0.0
            try:
                return max(0.0, float(raw))
            except ValueError:
                rejected.append(name.replace("_pct", " %").replace("_", " "))
                return previous

        system = request.form.get("coating_system", "Silicone")
        if system not in warranty_calc.COATING_SYSTEMS:
            system = "Silicone"
        roof_type = request.form.get("roof_type", "Capsheet")
        if roof_type not in warranty_calc.ROOF_TYPES:
            roof_type = "Capsheet"
        acrylic_type = request.form.get("acrylic_system_type", "Standard")
        if acrylic_type not in warranty_calc.ACRYLIC_TYPES:
            acrylic_type = "Standard"

        # Only combinations the manufacturers actually publish can be saved.
        snapped = []
        roofs = warranty_calc.supported_roof_types(system, acrylic_type)
        if roof_type not in roofs:
            snapped.append(f"{system}"
                           + (f" ({acrylic_type})" if system == "Acrylic" else "")
                           + f" isn't offered over {roof_type}")
            roof_type = roofs[0] if roofs else "Capsheet"
        years = num("warranty_years", bid["warranty_years"]) or 10
        supported_years = warranty_calc.supported_warranties(system, roof_type, acrylic_type)
        if supported_years and years not in supported_years:
            snapped.append(f"{system} on {roof_type} only comes in "
                           + " / ".join(f"{y}-year" for y in supported_years))
            years = supported_years[0]
        products = warranty_calc.resolve_products(
            system, request.form.get("selected_topcoat", ""),
            request.form.get("selected_basecoat", ""),
            request.form.get("selected_butter_grade", ""))
        conn.execute(
            """UPDATE bids SET roof_address=?, roof_size_sqft=?, deduction_sqft=?,
               surface_type=?, candidate=?, warranty_years=?, price=?,
               assessment_date=?, assessment_notes=?, coating_system=?,
               acrylic_system_type=?, roof_type=?, linear_feet=?, waste_pct=?,
               stretch_pct=?, passed_adhesion=?, has_rust=?, rust_prime_method=?,
               selected_topcoat=?, selected_basecoat=?, selected_butter_grade=?,
               updated_at=? WHERE id=?""",
            (request.form.get("roof_address", "").strip(),
             num("roof_size_sqft", bid["roof_size_sqft"]),
             num("deduction_sqft", bid["deduction_sqft"]) or 0,
             request.form.get("surface_type", "").strip(),
             request.form.get("candidate", "Yes"),
             years,
             _parse_amount(request.form.get("price"), on_error=bid["price"]),
             request.form.get("assessment_date", "").strip(),
             request.form.get("assessment_notes", "").strip(),
             system, acrylic_type, roof_type,
             num("linear_feet", bid["linear_feet"]) or 0,
             pct("waste_pct", bid["waste_pct"]), pct("stretch_pct", bid["stretch_pct"]),
             1 if request.form.get("passed_adhesion") else 0,
             1 if request.form.get("has_rust") else 0,
             request.form.get("rust_prime_method", "field"),
             products["top"], products["base"], products["mastic"],
             now_iso(), bid_id))
        conn.commit()
    finally:
        conn.close()
    notes = list(snapped)
    if rejected:
        notes.append("couldn't read " + ", ".join(sorted(set(rejected)))
                     + " so the previous value was kept")
    if notes:
        flash("Saved, with adjustments: " + "; ".join(notes) + ".", "warning")
    else:
        flash("Roof report saved.", "success")
    return redirect(url_for("bid_edit", bid_id=bid_id))


@app.route("/bids/<int:bid_id>/use-suggested-price", methods=["POST"])
def use_suggested_price(bid_id):
    conn = get_db()
    try:
        bid = _bid_or_404(conn, bid_id)
        suggested = _bid_suggested_price(conn, bid, _bid_plan(bid))
        if not suggested:
            flash("Enter a roof size first.", "danger")
        else:
            conn.execute("UPDATE bids SET price=?, updated_at=? WHERE id=?",
                         (suggested["total"], now_iso(), bid_id))
            conn.commit()
            flash(f"Price set to {format_money(suggested['total'])} "
                  f"({suggested['rate']:.2f}/sq ft × "
                  f"{suggested['sqft']:,} sq ft).", "success")
    finally:
        conn.close()
    return redirect(url_for("bid_edit", bid_id=bid_id))


@app.route("/bids/<int:bid_id>/delete", methods=["POST"])
def delete_bid(bid_id):
    conn = get_db()
    try:
        bid = _bid_or_404(conn, bid_id)
        ops = [undo_module.insert_op("bids", [dict(bid)]),
               undo_module.insert_op("bid_photos", undo_module.capture(
                   conn, "bid_photos", "bid_id=?", (bid_id,)))]
        trashed = [row["filename"] for row in conn.execute(
            "SELECT filename FROM bid_photos WHERE bid_id=?", (bid_id,))
            if undo_module.trash_photo(row["filename"])]
        ops.append(undo_module.files_op(trashed))
        conn.execute("DELETE FROM bids WHERE id=?", (bid_id,))
        _offer_undo(conn, "the deleted roof report", ops)
        conn.commit()
    finally:
        conn.close()
    flash("Roof report deleted.", "warning")
    return redirect(url_for("account_detail", account_id=bid["account_id"]))


@app.route("/bids/<int:bid_id>/photos", methods=["POST"])
def add_bid_photos(bid_id):
    """Upload assessment photos; each is auto-resized to fit the report."""
    from PIL import Image, ImageOps
    files = request.files.getlist("photos")
    added = failed = 0
    conn = get_db()
    try:
        _bid_or_404(conn, bid_id)
        photo_dir().mkdir(parents=True, exist_ok=True)
        last = conn.execute(
            "SELECT COALESCE(MAX(sort_order), -1) FROM bid_photos WHERE bid_id=?",
            (bid_id,)).fetchone()[0]
        for f in files:
            if not f or not f.filename:
                continue
            try:
                img = Image.open(f.stream)
                img = ImageOps.exif_transpose(img)  # honor phone orientation
                img.thumbnail((PHOTO_MAX_DIM, PHOTO_MAX_DIM))
                if img.mode not in ("RGB", "L"):
                    img = img.convert("RGB")
                name = f"bid{bid_id}_{uuid.uuid4().hex[:10]}.jpg"
                img.save(photo_dir() / name, "JPEG", quality=85, optimize=True)
            except Exception:
                failed += 1
                continue
            last += 1
            conn.execute(
                "INSERT INTO bid_photos (bid_id, filename, caption, sort_order, "
                "created_at) VALUES (?,?,?,?,?)",
                (bid_id, name, "", last, now_iso()))
            added += 1
        conn.commit()
    finally:
        conn.close()
    msg = f"Added {added} photo(s)."
    if failed:
        msg += (f" {failed} file(s) could not be read — use JPEG or PNG "
                "(on iPhone: Settings > Camera > Formats > Most Compatible).")
    flash(msg, "success" if added else "danger")
    return redirect(url_for("bid_edit", bid_id=bid_id))


@app.route("/bids/photos/<int:photo_id>/caption", methods=["POST"])
def caption_bid_photo(photo_id):
    conn = get_db()
    try:
        row = conn.execute("SELECT bid_id FROM bid_photos WHERE id=?",
                           (photo_id,)).fetchone()
        if row:
            conn.execute("UPDATE bid_photos SET caption=? WHERE id=?",
                         (request.form.get("caption", "").strip(), photo_id))
            conn.commit()
    finally:
        conn.close()
    return redirect(url_for("bid_edit", bid_id=row["bid_id"]) if row
                    else url_for("accounts"))


@app.route("/bids/photos/<int:photo_id>/delete", methods=["POST"])
def delete_bid_photo(photo_id):
    conn = get_db()
    try:
        row = conn.execute("SELECT * FROM bid_photos WHERE id=?",
                           (photo_id,)).fetchone()
        if row:
            trashed = [row["filename"]] if undo_module.trash_photo(row["filename"]) else []
            conn.execute("DELETE FROM bid_photos WHERE id=?", (photo_id,))
            _offer_undo(conn, "the removed photo",
                        [undo_module.insert_op("bid_photos", [dict(row)]),
                         undo_module.files_op(trashed)])
            conn.commit()
    finally:
        conn.close()
    return redirect(url_for("bid_edit", bid_id=row["bid_id"]) if row
                    else url_for("accounts"))


@app.route("/uploads/bid_photos/<path:filename>")
def bid_photo_file(filename):
    from flask import send_from_directory
    return send_from_directory(photo_dir(), filename)


@app.route("/bids/<int:bid_id>/report")
def bid_report(bid_id):
    """The full branded, printable roof report / proposal."""
    conn = get_db()
    try:
        bid = _bid_or_404(conn, bid_id)
        photos = conn.execute(
            "SELECT * FROM bid_photos WHERE bid_id = ? ORDER BY sort_order, id",
            (bid_id,)).fetchall()
        settings = _get_settings(conn)
    finally:
        conn.close()
    net_sqft = (bid["roof_size_sqft"] or 0) - (bid["deduction_sqft"] or 0)
    plan = _bid_plan(bid) if net_sqft > 0 else None
    return render_template("bid_report.html", bid=bid, photos=photos,
                           settings=settings, plan=plan, net_sqft=net_sqft,
                           price_words=dollars_in_words(bid["price"]))


# ----------------------------------------------------- Projects & invoicing

_MISSING = object()


def _parse_amount(raw, on_error=None):
    """Money from a form field. Accepts $, commas and spaces. An empty field
    clears the value; unparseable text returns `on_error` so a typo like
    "15,00O" can keep the previous number instead of wiping it."""
    cleaned = (raw or "").replace("$", "").replace(",", "").replace(" ", "").strip()
    if not cleaned:
        return None
    try:
        return round(float(cleaned), 2)
    except ValueError:
        return on_error


def _parse_int(raw, on_error=None):
    """Whole number from a form field, tolerant of commas and spaces."""
    cleaned = (raw or "").replace(",", "").replace(" ", "").strip()
    if not cleaned:
        return None
    try:
        value = int(float(cleaned))
    except ValueError:
        return on_error
    return value if value >= 0 else on_error


def _project_money(conn, project_id):
    row = conn.execute(
        """SELECT COALESCE(SUM(CASE WHEN status != 'Draft' THEN amount END), 0) AS invoiced,
                  COALESCE(SUM(CASE WHEN status = 'Paid' THEN amount END), 0) AS paid,
                  COALESCE(SUM(CASE WHEN status = 'Sent' THEN amount END), 0) AS outstanding
           FROM invoices WHERE project_id = ?""", (project_id,)).fetchone()
    return dict(row)


def _project_or_404(conn, project_id):
    proj = conn.execute(
        """SELECT p.*, a.company_name, a.first_name, a.last_name, a.email,
                  a.work_phone, a.mobile_phone
           FROM projects p JOIN accounts a ON a.id = p.account_id
           WHERE p.id = ?""", (project_id,)).fetchone()
    if proj is None:
        from flask import abort
        abort(404)
    return proj


@app.route("/projects")
def projects():
    status = request.args.get("status", "")
    sql = """SELECT p.*, a.company_name,
                    (SELECT COUNT(*) FROM project_tasks t
                     WHERE t.project_id = p.id AND t.done = 1) AS tasks_done,
                    (SELECT COUNT(*) FROM project_tasks t
                     WHERE t.project_id = p.id) AS tasks_total
             FROM projects p JOIN accounts a ON a.id = p.account_id"""
    params = []
    if status:
        sql += " WHERE p.status = ?"
        params.append(status)
    sql += " ORDER BY p.updated_at DESC"
    conn = get_db()
    try:
        rows = conn.execute(sql, params).fetchall()
        items = [{"project": r, "money": _project_money(conn, r["id"])} for r in rows]
        # accounts eligible for a new project: Closed Won without one
        eligible = conn.execute(
            f"""SELECT id, company_name FROM accounts
               WHERE COALESCE(archived_at, '') = '' AND pipeline_milestone = 'Closed Won'
                 AND id NOT IN (SELECT account_id FROM projects)
               ORDER BY company_name COLLATE NOCASE""").fetchall()
        totals = conn.execute(
            """SELECT COALESCE(SUM(CASE WHEN status != 'Draft' THEN amount END), 0) AS invoiced,
                      COALESCE(SUM(CASE WHEN status = 'Paid' THEN amount END), 0) AS paid,
                      COALESCE(SUM(CASE WHEN status = 'Sent' THEN amount END), 0) AS outstanding
               FROM invoices""").fetchone()
        contracted = conn.execute(
            "SELECT COALESCE(SUM(contract_amount), 0) FROM projects").fetchone()[0]
    finally:
        conn.close()
    return render_template("projects.html", items=items, status=status,
                           eligible=eligible, totals=totals, contracted=contracted)


@app.route("/projects/new", methods=["POST"])
def new_project():
    account_id = request.form.get("account_id", "")
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        existing = conn.execute("SELECT id FROM projects WHERE account_id = ?",
                                (account_id,)).fetchone()
        if existing:
            flash("This account already has a project.", "warning")
            return redirect(url_for("project_detail", project_id=existing["id"]))
        ts = now_iso()
        name = (request.form.get("name", "").strip()
                or f"{acct['company_name']} — Roof Restoration")
        cur = conn.execute(
            "INSERT INTO projects (account_id, name, contract_amount, created_at, "
            "updated_at) VALUES (?,?,?,?,?)",
            (account_id, name, _parse_amount(request.form.get("contract_amount")),
             ts, ts))
        project_id = cur.lastrowid
        for i, title in enumerate(DEFAULT_PROJECT_TASKS):
            conn.execute(
                "INSERT INTO project_tasks (project_id, title, sort_order, created_at) "
                "VALUES (?,?,?,?)", (project_id, title, i, ts))
        conn.commit()
    finally:
        conn.close()
    flash(f"Project created for {acct['company_name']} with the standard job checklist.",
          "success")
    return redirect(url_for("project_detail", project_id=project_id))


@app.route("/projects/<int:project_id>")
def project_detail(project_id):
    conn = get_db()
    try:
        proj = _project_or_404(conn, project_id)
        tasks = conn.execute(
            "SELECT * FROM project_tasks WHERE project_id = ? "
            "ORDER BY sort_order, id", (project_id,)).fetchall()
        invs = conn.execute(
            "SELECT * FROM invoices WHERE project_id = ? ORDER BY created_at, id",
            (project_id,)).fetchall()
        money = _project_money(conn, project_id)
    finally:
        conn.close()
    return render_template("project_detail.html", project=proj, tasks=tasks,
                           invoices=invs, money=money)


@app.route("/projects/<int:project_id>/edit", methods=["POST"])
def edit_project(project_id):
    conn = get_db()
    try:
        _project_or_404(conn, project_id)
        conn.execute(
            """UPDATE projects SET name=?, status=?, contract_amount=?,
               start_date=?, completion_date=?, notes=?, updated_at=? WHERE id=?""",
            (request.form.get("name", "").strip() or "Untitled Project",
             request.form.get("status", "Not Started"),
             _parse_amount(request.form.get("contract_amount")),
             request.form.get("start_date", "").strip(),
             request.form.get("completion_date", "").strip(),
             request.form.get("notes", "").strip(), now_iso(), project_id))
        conn.commit()
    finally:
        conn.close()
    flash("Project updated.", "success")
    return redirect(url_for("project_detail", project_id=project_id))


@app.route("/projects/<int:project_id>/delete", methods=["POST"])
def delete_project(project_id):
    conn = get_db()
    try:
        proj = _project_or_404(conn, project_id)
        ops = [undo_module.insert_op("projects", [dict(proj)])]
        for table in ("project_tasks", "invoices"):
            ops.append(undo_module.insert_op(table, undo_module.capture(
                conn, table, "project_id=?", (project_id,))))
        conn.execute("DELETE FROM projects WHERE id=?", (project_id,))
        _offer_undo(conn, f"delete of project “{proj['name']}”", ops)
        conn.commit()
    finally:
        conn.close()
    flash(f"Project “{proj['name']}” deleted.", "warning")
    return redirect(url_for("projects"))


@app.route("/projects/<int:project_id>/tasks/add", methods=["POST"])
def add_project_task(project_id):
    title = request.form.get("title", "").strip()
    if title:
        conn = get_db()
        try:
            _project_or_404(conn, project_id)
            last = conn.execute(
                "SELECT COALESCE(MAX(sort_order), -1) FROM project_tasks "
                "WHERE project_id=?", (project_id,)).fetchone()[0]
            conn.execute(
                "INSERT INTO project_tasks (project_id, title, sort_order, created_at) "
                "VALUES (?,?,?,?)", (project_id, title, last + 1, now_iso()))
            conn.commit()
        finally:
            conn.close()
    return redirect(url_for("project_detail", project_id=project_id))


@app.route("/projects/<int:project_id>/tasks/<int:task_id>/toggle", methods=["POST"])
def toggle_project_task(project_id, task_id):
    conn = get_db()
    try:
        conn.execute(
            "UPDATE project_tasks SET done = 1 - done, "
            "done_at = CASE WHEN done = 0 THEN ? ELSE '' END "
            "WHERE id = ? AND project_id = ?", (now_iso(), task_id, project_id))
        conn.execute("UPDATE projects SET updated_at=? WHERE id=?",
                     (now_iso(), project_id))
        conn.commit()
    finally:
        conn.close()
    return redirect(url_for("project_detail", project_id=project_id))


@app.route("/projects/<int:project_id>/tasks/<int:task_id>/delete", methods=["POST"])
def delete_project_task(project_id, task_id):
    conn = get_db()
    try:
        rows = undo_module.capture(conn, "project_tasks", "id=? AND project_id=?",
                                   (task_id, project_id))
        conn.execute("DELETE FROM project_tasks WHERE id=? AND project_id=?",
                     (task_id, project_id))
        if rows:
            _offer_undo(conn, f"removal of “{rows[0]['title']}”",
                        [undo_module.insert_op("project_tasks", rows)])
        conn.commit()
    finally:
        conn.close()
    return redirect(url_for("project_detail", project_id=project_id))


@app.route("/projects/<int:project_id>/invoices/add", methods=["POST"])
def add_invoice(project_id):
    amount = _parse_amount(request.form.get("amount"))
    if amount is None or amount <= 0:
        flash("Enter an invoice amount.", "danger")
        return redirect(url_for("project_detail", project_id=project_id))
    conn = get_db()
    try:
        _project_or_404(conn, project_id)
        ts = now_iso()
        conn.execute(
            """INSERT INTO invoices (project_id, invoice_number, amount, status,
               due_date, notes, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?)""",
            (project_id, request.form.get("invoice_number", "").strip(), amount,
             "Draft", request.form.get("due_date", "").strip(),
             request.form.get("notes", "").strip(), ts, ts))
        conn.commit()
    finally:
        conn.close()
    flash(f"Invoice for {format_money(amount)} added as Draft.", "success")
    return redirect(url_for("project_detail", project_id=project_id))


@app.route("/invoices/<int:invoice_id>/status", methods=["POST"])
def invoice_status(invoice_id):
    """Advance an invoice: mark Sent (stamps sent/due dates) or Paid."""
    new_status = request.form.get("status", "")
    if new_status not in INVOICE_STATUSES:
        return redirect(request.form.get("next") or url_for("projects"))
    conn = get_db()
    try:
        inv = conn.execute("SELECT * FROM invoices WHERE id=?",
                           (invoice_id,)).fetchone()
        if inv is None:
            flash("Invoice not found.", "danger")
            return redirect(request.form.get("next") or url_for("projects"))
        if inv:
            sent = inv["sent_date"] or (today_iso() if new_status in ("Sent", "Paid") else "")
            due = inv["due_date"]
            if new_status == "Sent" and not due:
                due = (datetime.now() + timedelta(days=30)).date().isoformat()
            paid = today_iso() if new_status == "Paid" else ""
            conn.execute(
                "UPDATE invoices SET status=?, sent_date=?, due_date=?, paid_date=?, "
                "updated_at=? WHERE id=?",
                (new_status, sent, due, paid, now_iso(), invoice_id))
            conn.execute("UPDATE projects SET updated_at=? WHERE id=?",
                         (now_iso(), inv["project_id"]))
            conn.commit()
            flash(f"Invoice marked {new_status}.", "success")
    finally:
        conn.close()
    return redirect(request.form.get("next")
                    or url_for("project_detail", project_id=inv["project_id"]))


def _address_from_notes(notes):
    """Pull the 'Address: ...' line that imports write into account notes."""
    for line in (notes or "").splitlines():
        if line.strip().lower().startswith("address:"):
            return line.split(":", 1)[1].strip()
    return ""


@app.route("/invoices/<int:invoice_id>/print")
def print_invoice(invoice_id):
    """Branded, printable invoice (print or save as PDF from the browser)."""
    conn = get_db()
    try:
        inv = conn.execute(
            """SELECT i.*, p.name AS project_name, p.account_id
               FROM invoices i JOIN projects p ON p.id = i.project_id
               WHERE i.id = ?""", (invoice_id,)).fetchone()
        if inv is None:
            from flask import abort
            abort(404)
        acct = conn.execute("SELECT * FROM accounts WHERE id = ?",
                            (inv["account_id"],)).fetchone()
        settings = _get_settings(conn)
    finally:
        conn.close()
    overdue = (inv["status"] == "Sent" and inv["due_date"]
               and inv["due_date"] < today_iso())
    return render_template("invoice_print.html", inv=inv, account=acct,
                           settings=settings, overdue=overdue,
                           bill_address=_address_from_notes(acct["notes"]),
                           number=inv["invoice_number"] or f"INV-{inv['id']:04d}")


@app.route("/invoices/<int:invoice_id>/delete", methods=["POST"])
def delete_invoice(invoice_id):
    conn = get_db()
    try:
        rows = undo_module.capture(conn, "invoices", "id=?", (invoice_id,))
        inv = rows[0] if rows else None
        conn.execute("DELETE FROM invoices WHERE id=?", (invoice_id,))
        if rows:
            _offer_undo(conn, "the deleted invoice",
                        [undo_module.insert_op("invoices", rows)])
        conn.commit()
    finally:
        conn.close()
    flash("Invoice deleted.", "warning")
    return redirect(url_for("project_detail", project_id=inv["project_id"])
                    if inv else url_for("projects"))


def _outstanding_invoices(conn):
    """Sent, unpaid invoices for the dashboard — overdue first."""
    rows = conn.execute(
        """SELECT i.*, p.name AS project_name, p.id AS pid, a.company_name
           FROM invoices i JOIN projects p ON p.id = i.project_id
           JOIN accounts a ON a.id = p.account_id
           WHERE i.status = 'Sent'
           ORDER BY CASE WHEN i.due_date = '' THEN 1 ELSE 0 END, i.due_date""").fetchall()
    t = today_iso()
    return [dict(r, overdue=bool(r["due_date"] and r["due_date"] < t)) for r in rows]


# ------------------------------------------------------------------- Export

def _csv_response(rows, headers, filename):
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(headers)
    writer.writerows(rows)
    return send_file(io.BytesIO(buf.getvalue().encode("utf-8-sig")),
                     mimetype="text/csv", as_attachment=True,
                     download_name=filename)


@app.route("/backup/download")
def download_backup():
    """Consistent snapshot of the whole database, sent as a file — an easy
    off-machine copy from any device (including the phone)."""
    import sqlite3 as _sq
    import tempfile
    from db import DB_PATH
    with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as tf:
        tmp_path = tf.name
    src = _sq.connect(DB_PATH)
    dst = _sq.connect(tmp_path)
    try:
        src.backup(dst)  # sqlite's online backup: safe while the app is in use
    finally:
        dst.close()
        src.close()
    data = Path(tmp_path).read_bytes()
    Path(tmp_path).unlink(missing_ok=True)
    return send_file(io.BytesIO(data), mimetype="application/octet-stream",
                     as_attachment=True,
                     download_name=f"crm-backup-{today_iso()}.db")


@app.route("/export/workbook.xlsx")
def export_workbook():
    """Everything in one Excel file: summary, accounts, contacts, history,
    roof reports, projects, invoices and today's tasks."""
    conn = get_db()
    try:
        buf = excel_export.build_workbook(conn)
    finally:
        conn.close()
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True, download_name=f"roof-crm-{today_iso()}.xlsx")


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

ACCOUNT_SORTS = {
    "priority": ("COALESCE(a.matching_properties, 0) DESC, "
                 "a.company_name COLLATE NOCASE"),
    "name": "a.company_name COLLATE NOCASE",
    "recent": ("COALESCE((SELECT MAX(created_at) FROM interactions i "
               "WHERE i.account_id = a.id), '') DESC, "
               "a.company_name COLLATE NOCASE"),
}


# Phone numbers get typed with every punctuation style there is, so a digit
# search compares against the number with its formatting stripped out.
_PHONE_SQL = ("REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE({c},"
              "'-',''),' ',''),'(',''),')',''),'.',''),'+','')")


def _search_clause(q: str):
    """SQL for the accounts search box.

    Looks at everything you might remember about a company: its name, the
    primary contact, the notes, AND the other people on the account — so a
    name you only ever added as a second contact is still findable.
    """
    like = f"%{q}%"
    cols = ["a.company_name", "a.first_name", "a.last_name", "a.email",
            "a.title", "a.notes", "a.seniority"]
    parts = [f"{c} LIKE ?" for c in cols]
    params = [like] * len(cols)

    digits = re.sub(r"\D", "", q)
    if len(digits) >= 4:
        for col in ("a.work_phone", "a.mobile_phone"):
            parts.append(f"{_PHONE_SQL.format(c=col)} LIKE ?")
            params.append(f"%{digits}%")

    contact_cols = ["c.first_name", "c.last_name", "c.email", "c.title",
                    "c.seniority"]
    contact_parts = [f"{c} LIKE ?" for c in contact_cols]
    contact_params = [like] * len(contact_cols)
    if len(digits) >= 4:
        for col in ("c.work_phone", "c.mobile_phone"):
            contact_parts.append(f"{_PHONE_SQL.format(c=col)} LIKE ?")
            contact_params.append(f"%{digits}%")
    parts.append("EXISTS (SELECT 1 FROM contacts c WHERE c.account_id = a.id "
                 "AND (" + " OR ".join(contact_parts) + "))")
    params += contact_params
    return "(" + " OR ".join(parts) + ")", params


# A page of accounts. Drawing every row turned a 2,000-account list into a
# multi-megabyte page; paging keeps each click the same speed whatever the
# size of the list.
ACCOUNTS_PER_PAGE = 100

# Timeline entries drawn on an account page before the "show all" link.
TIMELINE_ROWS = 25


def _contact_match_hints(conn, q: str, rows) -> dict:
    """{account_id: "Jane Doe (VP)"} for rows whose match came from a contact
    rather than from the account's own fields."""
    if not rows:
        return {}
    like = f"%{q}%"
    ids = [r["id"] for r in rows]
    ph = ",".join("?" * len(ids))
    hits = conn.execute(
        f"""SELECT account_id, first_name, last_name, title FROM contacts
            WHERE account_id IN ({ph})
              AND (first_name LIKE ? OR last_name LIKE ? OR email LIKE ?
                   OR title LIKE ? OR seniority LIKE ?)
            ORDER BY id""",
        ids + [like] * 5).fetchall()
    hints = {}
    for h in hits:
        if h["account_id"] in hints:
            continue
        name = f"{h['first_name']} {h['last_name']}".strip()
        hints[h["account_id"]] = f"{name} — {h['title']}" if h["title"] else name
    low = q.lower()
    # Drop the hint when the account itself matched; it would just be noise.
    for r in rows:
        if r["id"] in hints and any(
                low in (r[c] or "").lower()
                for c in ("company_name", "first_name", "last_name", "email")):
            hints.pop(r["id"])
    return hints


@app.route("/accounts")
def accounts():
    status = request.args.get("status", "")
    milestone = request.args.get("milestone", "")
    q = request.args.get("q", "").strip()
    min_matching_raw = request.args.get("min_matching", "").strip()
    min_matching = int(min_matching_raw) if min_matching_raw.isdigit() else None
    sort = request.args.get("sort", "priority")
    if sort not in ACCOUNT_SORTS:
        sort = "priority"
    view = "archived" if request.args.get("view") == "archived" else "active"

    page = _parse_int(request.args.get("page", "1"), on_error=1) or 1
    page = max(1, page)

    where = (" AND COALESCE(a.archived_at, '') != ''" if view == "archived"
             else " AND COALESCE(a.archived_at, '') = ''")
    params = []
    if status:
        where += " AND a.prospecting_status = ?"
        params.append(status)
    if milestone:
        where += " AND a.pipeline_milestone = ?"
        params.append(milestone)
    if min_matching is not None:
        where += " AND COALESCE(a.matching_properties, 0) >= ?"
        params.append(min_matching)
    if q:
        clause, search_params = _search_clause(q)
        where += " AND " + clause
        params += search_params

    conn = get_db()
    try:
        # Totals come from a COUNT/SUM over the whole filter, so the header
        # still describes every match even though only one page is drawn.
        summary = conn.execute(
            f"SELECT COUNT(*) c, COALESCE(SUM(a.matching_properties), 0) m "
            f"FROM accounts a WHERE 1=1{where}", params).fetchone()
        total, total_matching = summary["c"], summary["m"]
        pages = max(1, -(-total // ACCOUNTS_PER_PAGE))
        page = min(page, pages)
        rows = conn.execute(
            f"""SELECT a.*,
                       (SELECT MAX(created_at) FROM interactions i
                        WHERE i.account_id = a.id) AS last_activity
                FROM accounts a WHERE 1=1{where}
                ORDER BY {ACCOUNT_SORTS[sort]}
                LIMIT ? OFFSET ?""",
            params + [ACCOUNTS_PER_PAGE, (page - 1) * ACCOUNTS_PER_PAGE]).fetchall()
        archived_count = conn.execute(
            "SELECT COUNT(*) c FROM accounts "
            "WHERE COALESCE(archived_at, '') != ''").fetchone()["c"]
        # When a row only matched because of someone in its contact list, say
        # who — otherwise the result looks like a mystery.
        match_hints = _contact_match_hints(conn, q, rows) if q else {}
    finally:
        conn.close()
    return render_template("accounts.html", accounts=rows,
                           status=status, milestone=milestone, q=q,
                           min_matching=min_matching_raw, sort=sort,
                           total_matching=total_matching, view=view,
                           archived_count=archived_count,
                           match_hints=match_hints,
                           total=total, page=page, pages=pages,
                           per_page=ACCOUNTS_PER_PAGE)


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
            cols = list(fields) + ["cadence_start", "created_at", "updated_at"]
            cur = conn.execute(
                f"INSERT INTO accounts ({','.join(cols)}) "
                f"VALUES ({','.join('?' * len(cols))})",
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
    return _render_account_detail(account_id)


def _render_account_detail(account_id, **extra):
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        contacts = conn.execute(
            "SELECT * FROM contacts WHERE account_id = ? "
            "ORDER BY last_name COLLATE NOCASE, first_name COLLATE NOCASE",
            (account_id,)).fetchall()
        # Each timeline entry carries an inline edit form, so a long history
        # is the heaviest thing on this page. Draw the recent ones and offer
        # the rest behind a link.
        history_total = conn.execute(
            "SELECT COUNT(*) c FROM interactions WHERE account_id = ?",
            (account_id,)).fetchone()["c"]
        show_all_history = request.args.get("history") == "all"
        interactions = conn.execute(
            "SELECT * FROM interactions WHERE account_id = ? "
            "ORDER BY created_at DESC, id DESC" +
            ("" if show_all_history else f" LIMIT {TIMELINE_ROWS}"),
            (account_id,)).fetchall()
        steps = cadence.get_cadence_progress(conn, account_id)
        in_cadence = (acct["prospecting_status"] == cadence.ACTIVE_STATUS
                      and acct["pipeline_milestone"] == cadence.ACTIVE_MILESTONE)
        project = conn.execute("SELECT * FROM projects WHERE account_id = ?",
                               (account_id,)).fetchone()
        bids = conn.execute(
            "SELECT * FROM bids WHERE account_id = ? ORDER BY created_at DESC",
            (account_id,)).fetchall()
    finally:
        conn.close()
    return render_template("account_detail.html", account=acct,
                           contacts=contacts, interactions=interactions,
                           steps=steps, in_cadence=in_cadence, project=project,
                           bids=bids, history_total=history_total,
                           show_all_history=show_all_history, **extra)


@app.route("/accounts/<int:account_id>/edit", methods=["POST"])
def edit_account(account_id):
    conn = get_db()
    try:
        existing = _account_or_404(conn, account_id)
        fields = _account_fields_from_form(request.form, existing)
        if not fields["company_name"]:
            flash("Company name is required.", "danger")
            return redirect(url_for("account_detail", account_id=account_id))
        conn.execute(
            f"UPDATE accounts SET {','.join(c + '=?' for c in fields)}, "
            f"updated_at=? WHERE id=?",
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


@app.route("/accounts/<int:account_id>/archive", methods=["POST"])
def archive_account(account_id):
    """Take an account off the working list without losing it. Archived
    accounts keep every interaction, bid and project, disappear from the
    dashboard, queue, lists and stats, and are skipped by future imports."""
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        reason = request.form.get("archive_reason", "").strip()
        conn.execute("UPDATE accounts SET archived_at=?, archive_reason=?, "
                     "next_follow_up='', follow_up_note='', updated_at=? WHERE id=?",
                     (now_iso(), reason, now_iso(), account_id))
        conn.commit()
    finally:
        conn.close()
    flash(f"Archived “{acct['company_name']}”. It won't come back on future "
          f"imports — find it under Accounts → Archived to restore it.", "warning")
    return redirect(request.form.get("next") or url_for("accounts"))


@app.route("/accounts/<int:account_id>/restore", methods=["POST"])
def restore_account(account_id):
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        conn.execute("UPDATE accounts SET archived_at='', archive_reason='', "
                     "updated_at=? WHERE id=?", (now_iso(), account_id))
        conn.commit()
    finally:
        conn.close()
    flash(f"Restored “{acct['company_name']}” to your working list.", "success")
    return redirect(request.form.get("next")
                    or url_for("account_detail", account_id=account_id))


# ------------------------------------------------------------ Bulk actions

# Columns a bulk action may write, and the form field each reads from.
BULK_FIELD_ACTIONS = {
    "status": ("prospecting_status", PROSPECTING_STATUSES),
    "milestone": ("pipeline_milestone", PIPELINE_MILESTONES),
}


@app.route("/accounts/bulk", methods=["POST"])
def bulk_accounts():
    """Apply one change to every ticked account.

    Cleaning up a 400-row import one account at a time is the kind of chore
    that stops people using a CRM, so every bulk action is undoable: the
    previous values are snapshotted before anything is written.
    """
    action = request.form.get("action", "")
    ids = [int(i) for i in request.form.getlist("account_ids") if i.isdigit()]
    back = request.form.get("next") or url_for("accounts")
    if not ids:
        flash("Tick at least one account first.", "warning")
        return redirect(back)

    ph = ",".join("?" * len(ids))
    ts = now_iso()
    conn = get_db()
    try:
        rows = conn.execute(
            f"SELECT * FROM accounts WHERE id IN ({ph})", ids).fetchall()
        ids = [r["id"] for r in rows]          # drop ids that no longer exist
        if not ids:
            flash("Those accounts no longer exist.", "danger")
            return redirect(back)
        ph = ",".join("?" * len(ids))
        n = len(ids)

        def snapshot(*cols):
            return [undo_module.update_op("accounts", [
                {"id": r["id"], **{c: r[c] for c in cols}} for r in rows])]

        if action in BULK_FIELD_ACTIONS:
            column, allowed = BULK_FIELD_ACTIONS[action]
            value = request.form.get("value", "")
            if value not in allowed:
                flash("Pick a value to set first.", "warning")
                return redirect(back)
            _offer_undo(conn, f"the change to {n} account(s)", snapshot(column))
            conn.execute(
                f"UPDATE accounts SET {column}=?, updated_at=? WHERE id IN ({ph})",
                [value, ts] + ids)
            flash(f"Set {column.replace('_', ' ')} to “{value}” on {n} account(s).",
                  "success")

        elif action == "archive":
            reason = request.form.get("value", "").strip()
            _offer_undo(conn, f"archiving {n} account(s)",
                        snapshot("archived_at", "archive_reason",
                                 "next_follow_up", "follow_up_note"))
            conn.execute(
                f"UPDATE accounts SET archived_at=?, archive_reason=?, "
                f"next_follow_up='', follow_up_note='', updated_at=? "
                f"WHERE id IN ({ph}) AND COALESCE(archived_at,'') = ''",
                [ts, reason, ts] + ids)
            flash(f"Archived {n} account(s). Future imports will skip them.",
                  "warning")

        elif action == "restore":
            _offer_undo(conn, f"restoring {n} account(s)",
                        snapshot("archived_at", "archive_reason"))
            conn.execute(
                f"UPDATE accounts SET archived_at='', archive_reason='', "
                f"updated_at=? WHERE id IN ({ph})", [ts] + ids)
            flash(f"Restored {n} account(s) to your working list.", "success")

        elif action == "followup":
            when = request.form.get("value", "").strip()
            note = request.form.get("follow_up_note", "").strip()
            if when:
                try:
                    datetime.fromisoformat(when)
                except ValueError:
                    flash(f"Couldn't read the date “{when}”.", "danger")
                    return redirect(back)
            _offer_undo(conn, f"the follow-up change on {n} account(s)",
                        snapshot("next_follow_up", "follow_up_note"))
            conn.execute(
                f"UPDATE accounts SET next_follow_up=?, follow_up_note=?, "
                f"updated_at=? WHERE id IN ({ph})", [when, note, ts] + ids)
            flash((f"Follow-up set for {when} on {n} account(s)." if when
                   else f"Cleared the follow-up on {n} account(s)."), "success")

        elif action == "restart_cadence":
            _offer_undo(conn, f"restarting the cadence on {n} account(s)",
                        snapshot("cadence_start", "prospecting_status",
                                 "pipeline_milestone"))
            conn.execute(
                f"UPDATE accounts SET cadence_start=?, prospecting_status=?, "
                f"pipeline_milestone=?, updated_at=? WHERE id IN ({ph})",
                [today_iso(), cadence.ACTIVE_STATUS, cadence.ACTIVE_MILESTONE,
                 ts] + ids)
            conn.execute(
                f"DELETE FROM cadence_dismissals WHERE account_id IN ({ph})", ids)
            flash(f"Restarted the cadence on {n} account(s) from today.", "success")

        elif action == "delete":
            # Only offered on the Archived view — you have to archive a
            # company before you can erase it, which makes this hard to do
            # by accident. Still undoable, photos included.
            ops = [undo_module.insert_op("accounts", [dict(r) for r in rows])]
            for table in ("contacts", "interactions", "cadence_dismissals",
                          "bids", "projects"):
                ops.append(undo_module.insert_op(table, undo_module.capture(
                    conn, table, f"account_id IN ({ph})", ids)))
            bid_ids = [r["id"] for r in conn.execute(
                f"SELECT id FROM bids WHERE account_id IN ({ph})", ids)]
            project_ids = [r["id"] for r in conn.execute(
                f"SELECT id FROM projects WHERE account_id IN ({ph})", ids)]
            trashed = []
            if bid_ids:
                bph = ",".join("?" * len(bid_ids))
                ops.append(undo_module.insert_op("bid_photos", undo_module.capture(
                    conn, "bid_photos", f"bid_id IN ({bph})", bid_ids)))
                trashed = [r["filename"] for r in conn.execute(
                    f"SELECT filename FROM bid_photos WHERE bid_id IN ({bph})",
                    bid_ids) if undo_module.trash_photo(r["filename"])]
            if project_ids:
                pph = ",".join("?" * len(project_ids))
                for table in ("project_tasks", "invoices"):
                    ops.append(undo_module.insert_op(table, undo_module.capture(
                        conn, table, f"project_id IN ({pph})", project_ids)))
            ops.append(undo_module.files_op(trashed))
            conn.execute(f"DELETE FROM accounts WHERE id IN ({ph})", ids)
            _offer_undo(conn, f"deletion of {n} account(s)", ops)
            flash(f"Permanently deleted {n} account(s). A future import can "
                  f"add these companies again.", "warning")

        else:
            flash("Pick an action first.", "warning")
            return redirect(back)
        conn.commit()
    finally:
        conn.close()
    return redirect(back)


@app.route("/accounts/<int:account_id>/delete", methods=["POST"])
def delete_account(account_id):
    """Erase an account for good. Unlike archiving, this forgets the company
    entirely, so a future import can bring it back."""
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        # Snapshot the whole subtree, parents first, so Undo can rebuild it.
        bid_ids = [r["id"] for r in conn.execute(
            "SELECT id FROM bids WHERE account_id=?", (account_id,))]
        project_ids = [r["id"] for r in conn.execute(
            "SELECT id FROM projects WHERE account_id=?", (account_id,))]
        ops = [undo_module.insert_op("accounts", [dict(acct)])]
        for table in ("contacts", "interactions", "cadence_dismissals"):
            ops.append(undo_module.insert_op(
                table, undo_module.capture(conn, table, "account_id=?", (account_id,))))
        ops.append(undo_module.insert_op(
            "bids", undo_module.capture(conn, "bids", "account_id=?", (account_id,))))
        ops.append(undo_module.insert_op(
            "projects", undo_module.capture(conn, "projects", "account_id=?", (account_id,))))
        if bid_ids:
            ph = ",".join("?" * len(bid_ids))
            ops.append(undo_module.insert_op("bid_photos", undo_module.capture(
                conn, "bid_photos", f"bid_id IN ({ph})", bid_ids)))
        if project_ids:
            ph = ",".join("?" * len(project_ids))
            for table in ("project_tasks", "invoices"):
                ops.append(undo_module.insert_op(table, undo_module.capture(
                    conn, table, f"project_id IN ({ph})", project_ids)))
        # Photos go to the trash rather than straight to the bin, so Undo can
        # bring the images back and not just their database rows.
        trashed = [row["filename"] for row in conn.execute(
            "SELECT p.filename FROM bid_photos p JOIN bids b ON b.id = p.bid_id "
            "WHERE b.account_id = ?", (account_id,))
            if undo_module.trash_photo(row["filename"])]
        ops.append(undo_module.files_op(trashed))
        conn.execute("DELETE FROM accounts WHERE id=?", (account_id,))
        _offer_undo(conn, f"delete of “{acct['company_name']}”", ops)
        conn.commit()
    finally:
        conn.close()
    flash(f"Permanently deleted “{acct['company_name']}” and all of its history. "
          f"A future import can add this company again.", "warning")
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


@app.route("/interactions/<int:interaction_id>/edit", methods=["POST"])
def edit_interaction(interaction_id):
    """Fix a logged interaction — wrong type, typo in the notes, or a call
    logged on the wrong day. The date matters: cadence reminders are computed
    from what's logged, so correcting it corrects the reminders too."""
    conn = get_db()
    try:
        row = conn.execute("SELECT * FROM interactions WHERE id=?",
                           (interaction_id,)).fetchone()
        if row is None:
            flash("That entry no longer exists.", "danger")
            return redirect(request.form.get("next") or url_for("dashboard"))
        itype = request.form.get("interaction_type", row["interaction_type"])
        if itype not in INTERACTION_TYPES:
            itype = row["interaction_type"]
        notes = request.form.get("notes", "").strip()
        created_at = row["created_at"]
        new_date = request.form.get("created_date", "").strip()
        if new_date and new_date != created_at[:10]:
            try:
                datetime.fromisoformat(new_date)
                created_at = new_date + created_at[10:]
            except ValueError:
                flash(f"Couldn't read the date “{new_date}” — kept the original.",
                      "warning")
        conn.execute(
            "UPDATE interactions SET interaction_type=?, notes=?, created_at=? "
            "WHERE id=?", (itype, notes, created_at, interaction_id))
        _offer_undo(conn, "that edit", [undo_module.update_op("interactions", [
            {"id": row["id"], "interaction_type": row["interaction_type"],
             "notes": row["notes"], "created_at": row["created_at"]}])])
        conn.commit()
        account_id = row["account_id"]
    finally:
        conn.close()
    flash("Entry updated.", "success")
    return redirect(request.form.get("next")
                    or url_for("account_detail", account_id=account_id))


@app.route("/interactions/<int:interaction_id>/delete", methods=["POST"])
def delete_interaction(interaction_id):
    """Remove a logged interaction. If it was a cadence step, its reminder
    comes back — the reminders are derived from what's logged."""
    conn = get_db()
    try:
        rows = undo_module.capture(conn, "interactions", "id=?", (interaction_id,))
        if not rows:
            flash("That entry no longer exists.", "danger")
            return redirect(request.form.get("next") or url_for("dashboard"))
        conn.execute("DELETE FROM interactions WHERE id=?", (interaction_id,))
        _offer_undo(conn, f"deletion of “{rows[0]['interaction_type']}”",
                    [undo_module.insert_op("interactions", rows)])
        conn.commit()
        account_id = rows[0]["account_id"]
    finally:
        conn.close()
    flash("Entry deleted.", "warning")
    return redirect(request.form.get("next")
                    or url_for("account_detail", account_id=account_id))


# ----------------------------------------------------------------- Contacts

CONTACT_COLS = ("first_name", "last_name", "title", "email", "work_phone",
                "mobile_phone", "linkedin_url", "seniority")


def _has_primary_contact(acct) -> bool:
    return any(acct[c] for c in ("first_name", "last_name", "email"))


def _demote_primary_to_contact(conn, acct):
    """Move the account's current primary-contact fields into a contacts row."""
    if _has_primary_contact(acct):
        conn.execute(
            f"INSERT INTO contacts (account_id, {','.join(CONTACT_COLS)}, created_at) "
            f"VALUES ({','.join('?' * (len(CONTACT_COLS) + 2))})",
            (acct["id"], *(acct[c] for c in CONTACT_COLS), now_iso()))


def _set_primary_contact(conn, account_id, person: dict):
    conn.execute(
        f"UPDATE accounts SET {','.join(c + '=?' for c in CONTACT_COLS)}, "
        f"updated_at=? WHERE id=?",
        (*(person.get(c, "") for c in CONTACT_COLS), now_iso(), account_id))


@app.route("/accounts/<int:account_id>/contacts/add", methods=["POST"])
def add_contact(account_id):
    person = {c: request.form.get(c, "").strip() for c in CONTACT_COLS}
    parsed_note = ""
    paste = request.form.get("paste", "").strip()
    if paste and not request.form.get("already_parsed"):
        # "Paste from ZoomInfo": the blob fills whatever the form left blank,
        # so a copied profile becomes a contact without retyping six fields.
        parsed = importer.parse_contact_blob(paste)
        filled = [c for c in CONTACT_COLS if not person[c] and parsed.get(c)]
        for c in filled:
            person[c] = parsed[c]
        if filled:
            parsed_note = " Read from the paste: " + ", ".join(
                c.replace("_", " ") for c in filled) + "."
    if not person["seniority"] and person["title"]:
        person["seniority"] = importer.seniority_from_title(person["title"])
    if not (person["first_name"] or person["last_name"]):
        flash("Contact needs at least a first or last name — the paste didn't "
              "contain one, so nothing was added." if paste else
              "Contact needs at least a first or last name.", "danger")
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
                f"INSERT INTO contacts (account_id, {','.join(CONTACT_COLS)}, created_at) "
                f"VALUES ({','.join('?' * (len(CONTACT_COLS) + 2))})",
                (account_id, *(person[c] for c in CONTACT_COLS), now_iso()))
        conn.commit()
    finally:
        conn.close()
    flash(f"Contact {person['first_name']} {person['last_name']} added"
          + (" as primary." if make_primary else ".") + parsed_note, "success")
    return redirect(url_for("account_detail", account_id=account_id))


@app.route("/accounts/<int:account_id>/contacts/parse", methods=["POST"])
def parse_contact(account_id):
    """Read a pasted profile and show what came out of it, without saving.

    The paste box on its own asked people to trust it blind. This fills the
    form in front of them so they can see what it got, fix anything it missed,
    and only then press Add.
    """
    paste = request.form.get("paste", "").strip()
    if not paste:
        flash("Paste something into the box first.", "warning")
        return redirect(url_for("account_detail", account_id=account_id))
    parsed = importer.parse_contact_blob(paste)
    found = [c for c in CONTACT_COLS if parsed.get(c)]
    if not parsed["first_name"] and not parsed["last_name"]:
        flash("Couldn't find a person's name in that paste. Make sure the "
              "contact's name is in what you copied, or just type it in below.",
              "warning")
    else:
        missing = [c.replace("_", " ") for c in ("email", "work_phone", "title")
                   if not parsed.get(c)]
        msg = "Read: " + ", ".join(c.replace("_", " ") for c in found) + "."
        if missing:
            msg += " Didn't find: " + ", ".join(missing) + " — add below if you have it."
        msg += " Check it over, then press Add Contact."
        flash(msg, "info")
    return _render_account_detail(account_id, contact_prefill=parsed,
                                  contact_paste=paste)


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
        rows = undo_module.capture(conn, "contacts", "id=? AND account_id=?",
                                   (contact_id, account_id))
        conn.execute("DELETE FROM contacts WHERE id=? AND account_id=?",
                     (contact_id, account_id))
        if rows:
            who = f"{rows[0]['first_name']} {rows[0]['last_name']}".strip() or "contact"
            _offer_undo(conn, f"removal of {who}",
                        [undo_module.insert_op("contacts", rows)])
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
        pricing = _pricing_settings(conn)
    finally:
        conn.close()
    return render_template("templates.html", templates=templates,
                           settings=settings, pricing=pricing,
                           ROOF_TYPES=warranty_calc.ROOF_TYPES,
                           WARRANTY_YEARS=warranty_calc.WARRANTY_YEARS,
                           price_per_sqft=warranty_calc.price_per_sqft,
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
        bad_prices = []
        for key in ("my_name", "my_title", "my_company", "my_phone", "my_email",
                    "my_website", "my_address", "invoice_terms", "backup_dir",
                    "price_capsheet_base", "price_other_base", "price_add_15",
                    "price_add_20", "daily_goal"):
            if key not in request.form:  # only touch submitted fields
                continue
            value = request.form.get(key, "").strip()
            if key == "daily_goal":
                count = _parse_int(value, on_error=_MISSING)
                if count is _MISSING or count is None or count < 0:
                    bad_prices.append("daily activity goal")
                    continue
                value = str(count)
            elif key.startswith("price_"):
                amount = _parse_amount(value, on_error=_MISSING)
                if amount is _MISSING or amount is None:
                    bad_prices.append(key.replace("price_", "").replace("_", " "))
                    continue
                value = f"{amount:.2f}"
            conn.execute(
                "INSERT INTO settings (key, value) VALUES (?,?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                (key, value))
        conn.commit()
    finally:
        conn.close()
    if bad_prices:
        flash("Saved, but these weren't numbers and were left unchanged: "
              + ", ".join(bad_prices) + ".", "warning")
    else:
        flash("Settings saved.", "success")
    return redirect(request.form.get("next") or url_for("templates_page"))


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
            action_url = f"tel:{clean_tel(acct['work_phone'] or acct['mobile_phone'])}"
        elif t["kind"] == "text" and acct["mobile_phone"]:
            action_url = f"sms:{clean_tel(acct['mobile_phone'])}?body={quote(body)}"
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

def _backup_dir_setting():
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT value FROM settings WHERE key='backup_dir'").fetchone()
        return row["value"] if row else ""
    finally:
        conn.close()


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
    conn = get_db()
    try:
        duplicates = importer.find_duplicate_groups(conn)
    finally:
        conn.close()
    return render_template("import.html", result=result, duplicates=duplicates,
                           backup_dir=_backup_dir_setting())


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
            f"""SELECT id FROM accounts a
               WHERE COALESCE(a.archived_at, '') = '' AND prospecting_status = ? AND pipeline_milestone = ?
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
    create_missing = bool(request.form.get("create_missing"))
    contact_result = None
    duplicates = []
    conn = get_db()
    try:
        if not file or not file.filename:
            flash("Choose a .xlsx or .csv contact file first.", "danger")
        else:
            try:
                contact_result = importer.import_contacts(
                    conn, file, create_missing=create_missing)
                msg = (f"Attached {contact_result['attached']} contact(s) to "
                       f"{contact_result['companies_matched']} account(s).")
                if contact_result["accounts_created"]:
                    msg += (f" Opened {contact_result['accounts_created']} new "
                            f"account(s) for companies you didn't have.")
                flash(msg, "success")
            except ValueError as e:
                flash(str(e), "danger")
            except Exception as e:
                flash(f"Contact import failed: {e}", "danger")
        duplicates = importer.find_duplicate_groups(conn)
    finally:
        conn.close()
    return render_template("import.html", result=None, duplicates=duplicates,
                           contact_result=contact_result,
                           backup_dir=_backup_dir_setting())


@app.route("/import/template")
def import_template():
    csv = ("Company Name,First Name,Last Name,Job Title,Number of Properties,"
           "# Properties (in search),Email Address,Direct Phone Number,"
           "Mobile phone,LinkedIn Contact Profile URL,Management Level,Notes\r\n")
    return send_file(io.BytesIO(csv.encode()), mimetype="text/csv",
                     as_attachment=True, download_name="crm_import_template.csv")


# -------------------------------------------------------------------- Guide

@app.route("/guide")
def guide():
    return render_template("guide.html")


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
    moved = db_module.migrate_legacy_data()
    init_db()
    print(f"\n  Your data:  {db_module.DATA_DIR}")
    print("  (kept outside this folder, so updating the app never touches it)")
    if moved:
        print("  Moved from the old location:")
        for line in moved:
            print(f"    • {line}")
    backup = backup_db()
    if backup:
        print(f"  Daily backup saved: {backup}")
    # Drop undo records (and the photos they were holding) past their shelf life.
    conn = get_db()
    try:
        undo_module.purge(conn)
    finally:
        conn.close()
    print_network_instructions()
    try:
        from waitress import serve
        print("  Server: waitress (production WSGI)\n")
        serve(app, host="0.0.0.0", port=PORT, threads=8)
    except ImportError:
        print("  Server: Flask dev server (pip install waitress for the "
              "production server)\n")
        app.run(host="0.0.0.0", port=PORT, debug=False)
