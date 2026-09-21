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
                   send_file, url_for)

import cadence
import importer
import warranty_calc
from db import (DEFAULT_PROJECT_TASKS, INTERACTION_TYPES, INVOICE_STATUSES,
                PIPELINE_MILESTONES, PREFERRED_CONTACT_METHODS,
                PROJECT_STATUSES, PROSPECTING_STATUSES,
                backup_db, get_db, init_db, now_iso, today_iso)

app = Flask(__name__)
app.secret_key = "local-crm-flash-messages"  # local single-user app; used only for flash()
PORT = 8000
PHOTO_DIR = Path(__file__).parent / "uploads" / "bid_photos"
PHOTO_MAX_DIM = 1600  # uploaded photos are resized to fit the report/PDF


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
        "INVOICE_STATUSES": INVOICE_STATUSES,
        "today": today_iso(),
    }


@app.template_filter("money")
def format_money(value):
    if value is None:
        return "—"
    return f"${value:,.0f}" if value == int(value) else f"${value:,.2f}"


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
        f"SELECT * FROM accounts WHERE next_follow_up != '' AND next_follow_up <= ? "
        f"ORDER BY {sort}, company_name COLLATE NOCASE",
        (today_iso(),)).fetchall()


def _top_priority_accounts(conn, limit: int = 8):
    """Highest-potential accounts still in play, most matching buildings first."""
    return conn.execute(
        """SELECT a.*,
                  (SELECT MAX(created_at) FROM interactions i
                   WHERE i.account_id = a.id) AS last_activity
           FROM accounts a
           WHERE COALESCE(a.matching_properties, 0) > 0
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
        order = _task_order(conn)
        reminders = cadence.get_due_reminders(conn, order=order)
        followups = _due_followups(conn, order)
        stale = _stale_deals(conn)
        owed = _outstanding_invoices(conn)
        top_priority = _top_priority_accounts(conn)
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
            "matching_due": sum(r["matching_properties"] or 0 for r in reminders),
        }
        recent = conn.execute(
            """SELECT i.*, a.company_name FROM interactions i
               JOIN accounts a ON a.id = i.account_id
               ORDER BY i.created_at DESC, i.id DESC LIMIT 10""").fetchall()
        return render_template("dashboard.html", reminders=reminders,
                               followups=followups, stale=stale, owed=owed,
                               stats=stats, recent=recent, today=today_iso(),
                               order=order, top_priority=top_priority)
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
        for row in conn.execute("SELECT filename FROM bid_photos WHERE bid_id=?",
                                (bid_id,)):
            (PHOTO_DIR / row["filename"]).unlink(missing_ok=True)
        conn.execute("DELETE FROM bids WHERE id=?", (bid_id,))
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
        PHOTO_DIR.mkdir(parents=True, exist_ok=True)
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
                img.save(PHOTO_DIR / name, "JPEG", quality=85, optimize=True)
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
            (PHOTO_DIR / row["filename"]).unlink(missing_ok=True)
            conn.execute("DELETE FROM bid_photos WHERE id=?", (photo_id,))
            conn.commit()
    finally:
        conn.close()
    return redirect(url_for("bid_edit", bid_id=row["bid_id"]) if row
                    else url_for("accounts"))


@app.route("/uploads/bid_photos/<path:filename>")
def bid_photo_file(filename):
    from flask import send_from_directory
    return send_from_directory(PHOTO_DIR, filename)


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
            """SELECT id, company_name FROM accounts
               WHERE pipeline_milestone = 'Closed Won'
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
        conn.execute("DELETE FROM projects WHERE id=?", (project_id,))
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
        conn.execute("DELETE FROM project_tasks WHERE id=? AND project_id=?",
                     (task_id, project_id))
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
        inv = conn.execute("SELECT project_id FROM invoices WHERE id=?",
                           (invoice_id,)).fetchone()
        conn.execute("DELETE FROM invoices WHERE id=?", (invoice_id,))
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
    if min_matching is not None:
        sql += " AND COALESCE(a.matching_properties, 0) >= ?"
        params.append(min_matching)
    if q:
        sql += (" AND (a.company_name LIKE ? OR a.first_name LIKE ? "
                "OR a.last_name LIKE ? OR a.email LIKE ?)")
        params += [f"%{q}%"] * 4
    sql += " ORDER BY " + ACCOUNT_SORTS[sort]

    conn = get_db()
    try:
        rows = conn.execute(sql, params).fetchall()
        total_matching = sum(r["matching_properties"] or 0 for r in rows)
    finally:
        conn.close()
    return render_template("accounts.html", accounts=rows,
                           status=status, milestone=milestone, q=q,
                           min_matching=min_matching_raw, sort=sort,
                           total_matching=total_matching)


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
                           bids=bids)


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
                    "price_add_20"):
            if key not in request.form:  # only touch submitted fields
                continue
            value = request.form.get(key, "").strip()
            if key.startswith("price_"):
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
        flash("Saved, but these prices weren't numbers and were left unchanged: "
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
    return render_template("import.html", result=result,
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
        return render_template("import.html", result=None,
                               backup_dir=_backup_dir_setting())
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
    return render_template("import.html", result=None,
                           contact_result=contact_result,
                           backup_dir=_backup_dir_setting())


@app.route("/import/template")
def import_template():
    csv = ("Company Name,First Name,Last Name,Title,Number of Properties,"
           "Email,Work Phone,Mobile Phone,Notes\r\n")
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
