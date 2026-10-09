"""Roof Coating CRM — lightweight local Flask app.

Run:  python3 app.py
Then open http://<your-local-ip>:8000 from any device on your Wi-Fi.
"""

import csv
import io
import re
import socket
import uuid
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import quote

from flask import (Flask, flash, g, redirect, render_template, request,
                   send_file, session, url_for)
from markupsafe import Markup, escape

import cadence
import db as db_module
import excel_export
import importer
import undo as undo_module
import warranty_calc
from db import (ALL_OUTCOMES, ARCHIVE_REASONS, CALL_OUTCOMES, CALL_STEPS,
                CONNECT_OUTCOMES, DOOR_KNOCK, DOOR_OUTCOMES,
                PERSON_DONE_OUTCOMES, RECYCLE_DAYS, TRIED_STATUSES,
                DEFAULT_PROJECT_TASKS, INTERACTION_TYPES,
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
    """Phone number as dialable digits for tel:/sms: links.

    An extension ("(713) 555-0100 ext. 32", as ZoomInfo writes direct lines)
    becomes a pause — "7135550100,32" — so the phone dials the main number,
    waits for the pickup, then keys the extension. Run together, the digits
    would dial a number that doesn't exist.
    """
    m = re.match(r"(?i)(.*?\d.*?)\s*(?:ext\.?|extension|x|#)\s*(\d+)\s*$",
                 phone or "")
    number, ext = (m.group(1), m.group(2)) if m else (phone or "", "")
    digits = re.sub(r"[^\d+]", "", number)
    return f"{digits},{ext}" if digits and ext else digits


@app.template_filter("sms")
def clean_sms(phone):
    """Phone number for sms: links — a text can't dial an extension."""
    return clean_tel(phone).split(",")[0]


# Which app the Call / Text / Email buttons open (Templates page → Your Info).
CALL_APPS = {"phone": "Phone's own dialer", "google_voice": "Google Voice"}
EMAIL_APPS = {"default": "Default mail app", "gmail": "Gmail (in the browser)",
              "outlook": "Outlook.com / Microsoft 365 (in the browser)"}
# Phones can follow a different choice: e.g. Gmail on the computer, but the
# iPhone's Mail app (holding only the CRM account) on the phone.
PHONE_EMAIL_APPS = {"same": "Same as the computer",
                    "default": "Phone's default mail app (e.g. Apple Mail)",
                    "gmail": "Gmail app", "outlook": "Outlook app"}


def _link_prefs() -> dict:
    """The call/email app settings, read once per request."""
    if "link_prefs" not in g:
        conn = get_db()
        try:
            rows = dict(conn.execute(
                "SELECT key, value FROM settings WHERE key IN "
                "('call_app', 'email_app', 'email_app_phone', 'send_from', 'my_email')"
            ).fetchall())
        finally:
            conn.close()
        g.link_prefs = {"call_app": rows.get("call_app") or "phone",
                        "email_app": rows.get("email_app") or "default",
                        "email_app_phone": rows.get("email_app_phone") or "same",
                        "send_from": (rows.get("send_from") or rows.get("my_email") or "").strip()}
    return g.link_prefs


def _device() -> str:
    """'ios', 'android' or 'desktop', from the browser's user agent."""
    ua = request.headers.get("User-Agent", "") if request else ""
    if any(k in ua for k in ("iPhone", "iPad", "iPod")):
        return "ios"
    return "android" if "Android" in ua else "desktop"


def _e164(phone: str) -> str:
    """+17135550100 from any US-style number; the extension is dropped."""
    digits = re.sub(r"\D", "", clean_tel(phone).split(",")[0])
    if len(digits) == 10:
        digits = "1" + digits
    return "+" + digits if digits else ""


def _attrs(url: str, external: bool) -> Markup:
    extra = ' target="_blank" rel="noopener"' if external else ""
    return Markup(f'href="{escape(url)}"{extra}')


def phone_link(phone: str, kind: str = "call", body: str = "") -> tuple[str, bool]:
    """(url, opens_in_new_tab) for calling or texting `phone`.

    Google Voice has no call link a phone or browser hands to it, so with
    Google Voice chosen the buttons open Google Voice's own web pages: a new
    call to the number, or the message thread with it. Those can't dial an
    extension or pre-fill a text; the scripts page has Copy buttons for that.
    """
    if not phone:
        return "", False
    if _link_prefs()["call_app"] == "google_voice":
        num = quote(_e164(phone))
        if kind == "text":
            return f"https://voice.google.com/u/0/messages?itemId=t.{num}", True
        return f"https://voice.google.com/u/0/calls?a=nc,{num}", True
    if kind == "text":
        return f"sms:{clean_sms(phone)}" + (f"?body={quote(body)}" if body else ""), False
    return f"tel:{clean_tel(phone)}", False


def email_link(email: str, subject: str = "", body: str = "") -> tuple[str, bool]:
    """(url, opens_in_new_tab) for writing to `email` in the chosen mail app."""
    if not email:
        return "", False
    prefs = _link_prefs()
    device = _device()
    app_ = prefs["email_app"]
    if device != "desktop" and prefs["email_app_phone"] != "same":
        app_ = prefs["email_app_phone"]
    fields = "subject=" + quote(subject) + "&body=" + quote(body)
    # On a phone the web compose page is handed to the Gmail/Outlook APP,
    # which opens on the inbox and drops the draft. Those apps have their
    # own compose links that keep everything, so phones get those instead.
    if app_ == "gmail" and device == "ios":
        return f"googlegmail:///co?to={quote(email)}&{fields}", False
    if app_ == "outlook" and device == "ios":
        return f"ms-outlook://compose?to={quote(email)}&{fields}", False
    if app_ in ("gmail", "outlook") and device == "android":
        # Android offers the installed mail apps for a mailto, draft intact.
        return f"mailto:{email}?{fields}", False
    if app_ == "gmail":
        # authuser picks which signed-in Google account writes the email, so
        # a browser also signed in to a personal Gmail never sends from it.
        sender = _link_prefs()["send_from"]
        who = f"authuser={quote(sender)}&" if sender else ""
        return (f"https://mail.google.com/mail/?{who}view=cm&fs=1&to=" + quote(email)
                + "&su=" + quote(subject) + "&body=" + quote(body)), True
    if app_ == "outlook":
        return ("https://outlook.office.com/mail/deeplink/compose?to=" + quote(email)
                + "&subject=" + quote(subject) + "&body=" + quote(body)), True
    query = "&".join(f"{k}={quote(v)}" for k, v in (("subject", subject), ("body", body)) if v)
    return f"mailto:{email}" + (f"?{query}" if query else ""), False


@app.template_global()
def zoominfo_link(acct):
    """The account's own ZoomInfo company page once it's known (saved from a
    ZoomInfo export or paste); until then, a search for the company."""
    url = (acct["zoominfo_url"] if "zoominfo_url" in acct.keys() else "") or ""
    if url:
        return _attrs(url, True)
    return _attrs("https://www.google.com/search?q="
                  + quote("site:zoominfo.com " + acct["company_name"]), True)


def _save_zoominfo_url(conn, account_id, paste: str) -> None:
    """Remember the company's ZoomInfo page from a pasted ZoomInfo page."""
    url = importer.zoominfo_company_url(paste or "")
    if url:
        conn.execute("UPDATE accounts SET zoominfo_url=? WHERE id=? "
                     "AND COALESCE(zoominfo_url, '') = ''", (url, account_id))


@app.template_global()
def dial(phone):
    """Attributes for a Call link: {{ dial(phone) }} inside <a ...>."""
    return _attrs(*phone_link(phone, "call"))


@app.template_global()
def text_to(phone, body=""):
    return _attrs(*phone_link(phone, "text", body))


@app.template_global()
def mail_to(email, subject="", body=""):
    return _attrs(*email_link(email, subject, body))


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
        # Undo records and trashed photos expire after a week. Doing this only
        # at startup meant an app left running for a month never cleared them.
        try:
            conn = get_db()
            try:
                undo_module.purge(conn)
            finally:
                conn.close()
        except Exception:
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
        "CALL_OUTCOMES": CALL_OUTCOMES,
        "DOOR_OUTCOMES": DOOR_OUTCOMES,
        "CALL_STEPS": CALL_STEPS,
        "today": today_iso(),
    }


@app.template_filter("money")
def format_money(value):
    if value is None:
        return "—"
    return f"${value:,.0f}" if value == int(value) else f"${value:,.2f}"


# ------------------------------------------------------------ Error reporting

MAX_ERROR_LOG_BYTES = 512 * 1024


@app.errorhandler(Exception)
def handle_error(err):
    """Show what actually went wrong, and write it down.

    This runs on one computer for one person, so hiding the error behind
    "Internal Server Error" helps nobody — it just means the terminal window
    has to be hunted down before anything can be fixed. The page names the
    error and points at the log; the log holds the full traceback.
    """
    from werkzeug.exceptions import HTTPException
    if isinstance(err, HTTPException):
        return err                      # 404s and the like are not crashes

    import traceback
    detail = traceback.format_exc()
    try:
        log = db_module.ERROR_LOG
        log.parent.mkdir(parents=True, exist_ok=True)
        if log.exists() and log.stat().st_size > MAX_ERROR_LOG_BYTES:
            log.rename(log.with_suffix(".log.old"))
        with log.open("a", encoding="utf-8") as fh:
            fh.write(f"\n{'=' * 70}\n{now_iso()}  {request.method} "
                     f"{request.full_path}\n{'=' * 70}\n{detail}\n")
    except OSError:
        pass
    app.logger.error("Unhandled error on %s\n%s", request.full_path, detail)
    try:
        return render_template("error.html", error=err,
                               error_type=type(err).__name__,
                               detail=detail,
                               log_path=str(db_module.ERROR_LOG)), 500
    except Exception:
        # error.html extends base.html. If the fault is in the shared page
        # chrome, rendering this page fails too — and Flask would fall back to
        # the bare "Internal Server Error" this handler exists to replace.
        # Plain text always works.
        body = (f"Roof CRM hit an error on {request.path}\n\n"
                f"{type(err).__name__}: {err}\n\n"
                f"Your saved data is intact. Full details are in "
                f"{db_module.ERROR_LOG}\n\n{detail}")
        return body, 500, {"Content-Type": "text/plain; charset=utf-8"}


# ---------------------------------------------------------------------- Undo
#
# Every destructive action snapshots what it changed into the undo_log table,
# then drops the record's id into the session. base.html renders an Undo bar
# for as long as that record is unused, so a mis-tap is a click away from
# being put back instead of a restore-from-backup.

UNDO_SESSION_KEY = "undo_id"

# How long the Undo bar keeps showing after an action. It used to stay until
# the next undoable action, which meant "Just did: removal of Jane Doe" could
# still be on screen days and dozens of unrelated actions later.
UNDO_BAR_SECONDS = 15 * 60


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
        session[UNDO_SESSION_KEY] = {"id": undo_id, "label": label,
                                     "at": datetime.now().timestamp()}


@app.context_processor
def inject_undo():
    """The pending undo, if the last action left one. Session-only — the
    record itself is re-checked (and may be refused as used or expired) when
    the button is actually pressed."""
    pending = session.get(UNDO_SESSION_KEY)
    if not (isinstance(pending, dict) and pending.get("id")):
        return {"pending_undo": None}
    age = datetime.now().timestamp() - float(pending.get("at") or 0)
    if age > UNDO_BAR_SECONDS:
        session.pop(UNDO_SESSION_KEY, None)
        return {"pending_undo": None}
    return {"pending_undo": pending}


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
        "website": form.get("website", "").strip(),
        "street": form.get("street", "").strip(),
        "city": form.get("city", "").strip(),
        "state": form.get("state", "").strip(),
        "zip": form.get("zip", "").strip(),
        "zoominfo_url": (importer.zoominfo_company_url(form.get("zoominfo_url", ""))
                         or form.get("zoominfo_url", "").strip()),
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
        finished_cadences = _finished_cadences(conn)
        stats = {
            "total_accounts": conn.execute(
                "SELECT COUNT(*) c FROM accounts WHERE COALESCE(archived_at, '') = ''").fetchone()["c"],
            "active_prospects": conn.execute(
                "SELECT COUNT(*) c FROM accounts WHERE COALESCE(archived_at, '') = '' "
                "AND prospecting_status NOT IN ('Not Interested') AND pipeline_milestone "
                "NOT IN ('Closed Won','Closed Lost')").fetchone()["c"],
            "in_cadence": conn.execute(
                f"SELECT COUNT(*) c FROM accounts WHERE {cadence.IN_CADENCE_SQL}"
            ).fetchone()["c"],
            "research": conn.execute(
                f"SELECT COUNT(*) c FROM accounts WHERE {cadence.RESEARCH_SQL}"
            ).fetchone()["c"],
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
        # Who to look up next: research accounts, biggest portfolios first.
        research_top = conn.execute(
            f"SELECT id, company_name, matching_properties, work_phone, zoominfo_url "
            f"FROM accounts WHERE {cadence.RESEARCH_SQL} "
            f"ORDER BY COALESCE(matching_properties, 0) DESC, "
            f"company_name COLLATE NOCASE LIMIT 6").fetchall()
        progress = _today_progress(conn, totals["reminders"] + totals["followups"])
        return render_template("dashboard.html", reminders=reminders,
                               followups=followups, stale=stale, owed=owed,
                               stats=stats, recent=recent, today=today_iso(),
                               order=order, top_priority=top_priority,
                               progress=progress, totals=totals,
                               show_all=show_all, research_top=research_top,
                               finished_cadences=finished_cadences,
                               RECYCLE_DAYS=RECYCLE_DAYS)
    finally:
        conn.close()


def _account_exists(conn, account_id) -> bool:
    return conn.execute("SELECT 1 FROM accounts WHERE id = ?",
                        (account_id,)).fetchone() is not None


# Who to try next at an account: the people who own the roof problem first,
# the C-suite last as the escalation.
NEXT_CONTACT_ORDER = {"Director": 0, "Manager": 1, "VP-Level": 2, "C-Level": 3}


def _person_name(row) -> str:
    return " ".join(x for x in (row["first_name"], row["last_name"]) if x).strip()


def _next_contacts(conn, account_id) -> list:
    """People at the account who haven't had their turn, best first:
    reachable before not, then by NEXT_CONTACT_ORDER, then oldest first."""
    rows = conn.execute(
        "SELECT * FROM contacts WHERE account_id=? AND COALESCE(tried_status, '') = ''",
        (account_id,)).fetchall()
    return sorted(rows, key=lambda r: (not cadence.has_contact(r),
                                       NEXT_CONTACT_ORDER.get(r["seniority"] or "", 4),
                                       r["id"]))


def _rotate_contact(conn, account_id, status: str, contact_id=None):
    """End the current primary's turn (marked `status`) and start the next
    person's cadence today. Returns (their name, undo ops) or None when
    nobody untried is left. The caller records the undo and commits."""
    acct = conn.execute("SELECT * FROM accounts WHERE id=?", (account_id,)).fetchone()
    nxt = _next_contacts(conn, account_id)
    if contact_id is not None:
        nxt = [r for r in nxt if r["id"] == int(contact_id)]
    if not acct or not nxt:
        return None
    pick = nxt[0]
    acct_before = undo_module.capture(conn, "accounts", "id=?", (account_id,))
    contacts_before = undo_module.capture(conn, "contacts", "account_id=?", (account_id,))
    created = []
    if _has_primary_contact(acct):
        cur = conn.execute(
            f"INSERT INTO contacts (account_id, {','.join(CONTACT_COLS)}, tried_status, "
            f"tried_at, created_at) VALUES ({','.join('?' * (len(CONTACT_COLS) + 4))})",
            (account_id, *(acct[c] for c in CONTACT_COLS), status, today_iso(), now_iso()))
        created.append(cur.lastrowid)
    _set_primary_contact(conn, account_id, dict(pick))
    conn.execute("DELETE FROM contacts WHERE id=?", (pick["id"],))
    conn.execute("UPDATE accounts SET previous_contact=? WHERE id=?",
                 (acct["first_name"] or acct["last_name"] or "", account_id))
    restart_ops = _restart_cadence(conn, [account_id])
    ops = ([undo_module.delete_op("contacts", created),
            undo_module.insert_op("contacts", contacts_before),
            undo_module.update_op("accounts", acct_before)] + restart_ops)
    return _person_name(pick) or "the next contact", ops


def _rest_account(conn, account_id, why: str) -> str:
    """Everyone at the company has been tried: rest it for RECYCLE_DAYS,
    then a follow-up brings it back to start again with the most senior."""
    back = (date.today() + timedelta(days=RECYCLE_DAYS)).isoformat()
    note = f"Everyone tried ({why.lower()}). Start again with the most senior person."
    if why == "Not interested":
        conn.execute("UPDATE accounts SET prospecting_status='Not Interested', "
                     "next_follow_up=?, follow_up_note=?, updated_at=? WHERE id=?",
                     (back, note, now_iso(), account_id))
        return (f" Nobody else is on file, so the company is marked Not Interested and "
                f"comes back as a follow-up on {back}.")
    conn.execute("UPDATE accounts SET next_follow_up=?, follow_up_note=?, updated_at=? "
                 "WHERE id=?", (back, note, now_iso(), account_id))
    return f" Resting it until {back}, when it comes back as a follow-up."


def _finished_cadences(conn, account_id=None) -> list[dict]:
    """In-cadence accounts whose every step is done or skipped: the person
    had their full run without a reply (a reply moves the deal out of
    the cadence). Each comes with who's next, if anyone."""
    sql = (f"SELECT a.* FROM accounts a WHERE {cadence.in_cadence_where('a')}"
           + (" AND a.id = ?" if account_id else ""))
    rows = conn.execute(sql, (account_id,) if account_id else ()).fetchall()
    done = cadence._done_dates(conn, [r["id"] for r in rows])
    n_steps = len(cadence.CADENCE_STEPS)
    out = []
    for r in rows:
        finished_steps = set(done.get(r["id"], {}))
        if cadence.skips_email(r):
            finished_steps |= set(cadence.EMAIL_STEPS)
        if len(finished_steps) < n_steps or not done.get(r["id"]):
            continue
        if r["next_follow_up"] and r["next_follow_up"] > today_iso():
            continue                     # already resting
        nxt = _next_contacts(conn, r["id"])
        out.append({"account": r, "next": nxt,
                    "finished": max(done[r["id"]].values()).isoformat()})
    return out


def _log_interaction(conn, account_id, itype: str, notes: str,
                     outcome: str = "") -> str:
    """Record an interaction and carry out what its outcome means.

    One place for the dashboard, the queue and the account page, so a call
    marked "Meeting booked" moves the deal wherever it was logged. Records an
    undo covering both the log and anything the outcome changed — these are
    one-tap buttons on a calling list, and a mis-tap on "Not interested"
    shouldn't quietly end a cadence.

    Returns a sentence describing any knock-on change, for the flash message.
    """
    outcome = outcome if outcome in ALL_OUTCOMES else ""
    before = conn.execute(
        "SELECT id, prospecting_status, pipeline_milestone, cadence_start, notes, "
        "next_follow_up, follow_up_note FROM accounts WHERE id=?", (account_id,)).fetchone()
    cur = conn.execute(
        "INSERT INTO interactions (account_id, interaction_type, notes, outcome, "
        "created_at) VALUES (?,?,?,?,?)",
        (account_id, itype, notes, outcome, now_iso()))
    effect = ""
    if outcome in PERSON_DONE_OUTCOMES and itype != DOOR_KNOCK:
        # One person's no isn't the company's no: hand over to the next
        # contact, or rest the company once everyone has been tried.
        who = _person_name(conn.execute(
            "SELECT * FROM accounts WHERE id=?", (account_id,)).fetchone())
        rotated = _rotate_contact(conn, account_id, outcome)
        if rotated:
            name, ops = rotated
            _offer_undo(conn, f"logging “{itype}” ({outcome})",
                        [undo_module.delete_op("interactions", [cur.lastrowid])] + ops)
            return (f" {who or 'They'} marked {outcome.lower()}. Moved on to {name}: "
                    f"their cadence starts today, with an email that mentions {who or 'the first contact'}.")
        if outcome == "Wrong person":
            conn.execute("UPDATE accounts SET cadence_start='', updated_at=? WHERE id=?",
                         (now_iso(), account_id))
            effect = (" Nobody else is on file, so the account is back in Research. Add the "
                      "right person and tick Make primary; their cadence starts that day.")
        else:
            effect = _rest_account(conn, account_id, "Not interested")
        _offer_undo(conn, f"logging “{itype}” ({outcome})",
                    [undo_module.delete_op("interactions", [cur.lastrowid]),
                     undo_module.update_op("accounts", [dict(before)])])
        return effect
    if outcome == "Meeting booked":
        conn.execute("UPDATE accounts SET pipeline_milestone='Accepted Meeting', "
                     "updated_at=? WHERE id=?", (now_iso(), account_id))
        effect = " Moved to Accepted Meeting — the cadence stops here."
    elif outcome == "Not interested":
        conn.execute("UPDATE accounts SET prospecting_status='Not Interested', "
                     "updated_at=? WHERE id=?", (now_iso(), account_id))
        effect = " Marked Not Interested — the cadence stops here."
    elif outcome == "Bad number":
        stamp = datetime.now().strftime("%b %d, %Y")
        note = f"Bad number reported on {stamp} ({itype})."
        conn.execute(
            "UPDATE accounts SET cadence_start='', notes=?, updated_at=? WHERE id=?",
            (((before["notes"] + "\n") if before["notes"] else "") + note,
             now_iso(), account_id))
        effect = (" Moved back to Research — fix the number (or add another "
                  "contact) and the cadence picks up again, with this call still to do.")
    _offer_undo(conn, f"logging “{itype}”" + (f" ({outcome})" if outcome else ""),
                [undo_module.delete_op("interactions", [cur.lastrowid]),
                 undo_module.update_op("accounts", [dict(before)])])
    return effect


@app.route("/reminders/dismiss", methods=["POST"])
def dismiss_reminder():
    account_id = request.form["account_id"]
    step_type = request.form["step_type"]
    conn = get_db()
    try:
        if not _account_exists(conn, account_id):
            flash("That account no longer exists.", "danger")
            return redirect(request.form.get("next") or url_for("dashboard"))
        cur = conn.execute(
            "INSERT OR IGNORE INTO cadence_dismissals "
            "(account_id, step_type, dismissed_at) VALUES (?,?,?)",
            (account_id, step_type, now_iso()))
        if cur.rowcount:
            # One tap sits right next to Log, so a mis-tap must be reversible.
            _offer_undo(conn, f"skipping “{step_type}”",
                        [undo_module.delete_op("cadence_dismissals", [cur.lastrowid])])
        conn.commit()
    finally:
        conn.close()
    flash(f"Skipped “{step_type}”.", "success")
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
        outcome = request.form.get("outcome", "")
        effect = _log_interaction(conn, account_id, step_type, notes, outcome)
        conn.commit()
    finally:
        conn.close()
    flash(f"Logged “{step_type}”" + (f" — {outcome}" if outcome in ALL_OUTCOMES else "")
          + "." + effect, "success")
    return redirect(request.form.get("next") or url_for("dashboard"))


# ------------------------------------------------------- Follow-ups & queue

@app.route("/accounts/<int:account_id>/followup", methods=["POST"])
def set_followup(account_id):
    """Set, snooze, or clear an account's follow-up reminder."""
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        undo_ops = [undo_module.update_op("accounts", [{
            "id": acct["id"], "next_follow_up": acct["next_follow_up"],
            "follow_up_note": acct["follow_up_note"]}])]
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
        _offer_undo(conn, f"changing the follow-up for {acct['company_name']}", undo_ops)
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
                f"SELECT COUNT(*) FROM accounts WHERE {cadence.IN_CADENCE_SQL}"),
            "research": count(
                f"SELECT COUNT(*) FROM accounts WHERE {cadence.RESEARCH_SQL}"),
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

        # Calls, last 30 days: how many, how often you reach someone, how often
        # that turns into a meeting. Rates use only calls with an outcome
        # recorded, so older calls logged before outcomes existed can't drag
        # them down.
        month_back = (today - timedelta(days=30)).isoformat()
        ph = ",".join("?" * len(CALL_STEPS))
        call_rows = conn.execute(
            f"""SELECT COALESCE(i.outcome, '') AS outcome FROM interactions i
                JOIN accounts a ON a.id = i.account_id
                WHERE COALESCE(a.archived_at, '') = '' AND i.created_at >= ?
                  AND (i.interaction_type IN ({ph}) OR COALESCE(i.outcome, '') != '')""",
            (month_back, *CALL_STEPS)).fetchall()
        tagged = [r["outcome"] for r in call_rows if r["outcome"]]
        connects = sum(1 for o in tagged if o in CONNECT_OUTCOMES)
        meetings = sum(1 for o in tagged if o == "Meeting booked")
        calls = {
            "total": len(call_rows),
            "tagged": len(tagged),
            "connects": connects,
            "meetings": meetings,
            "connect_rate": round(100 * connects / len(tagged)) if tagged else None,
            "meeting_rate": round(100 * meetings / connects) if connects else None,
            "bars": _bar_items([(o, tagged.count(o)) for o in CALL_OUTCOMES],
                               total=len(tagged) or None),
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
                           activity=activity, calls=calls,
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
                   or account_address(acct))
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


def account_address(acct) -> str:
    """One-line office address: the address fields, else the Notes line."""
    keys = acct.keys()
    parts = [acct[k] for k in ("street", "city") if k in keys and acct[k]]
    state_zip = " ".join(acct[k] for k in ("state", "zip") if k in keys and acct[k])
    if state_zip:
        parts.append(state_zip)
    return ", ".join(parts) or _address_from_notes(acct["notes"] if "notes" in keys else "")


app.add_template_global(account_address)


@app.template_global()
def map_link(acct):
    """Google Maps for the account's office (directions are one tap away)."""
    addr = account_address(acct)
    if not addr:
        return ""
    return _attrs("https://www.google.com/maps/search/?api=1&query="
                  + quote(acct["company_name"] + ", " + addr), True)


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
                           bill_address=account_address(acct),
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
            "a.title", "a.notes", "a.seniority", "a.website", "a.street",
            "a.city", "a.zip"]
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


def _account_filters(args) -> dict:
    """The Accounts list's filters, read from the query string, as a SQL
    WHERE fragment over alias `a` plus the values the template echoes back.
    Shared with the ZoomInfo download so it exports exactly what's listed."""
    status = args.get("status", "")
    milestone = args.get("milestone", "")
    q = args.get("q", "").strip()
    min_matching_raw = args.get("min_matching", "").strip()
    min_matching = int(min_matching_raw) if min_matching_raw.isdigit() else None
    sort = args.get("sort", "priority")
    if sort not in ACCOUNT_SORTS:
        sort = "priority"
    view = args.get("view", "")
    if view not in ("archived", "research"):
        view = "active"

    if view == "archived":
        where = " AND COALESCE(a.archived_at, '') != ''"
    elif view == "research":
        # Waiting for someone to contact. Sorted biggest-portfolio-first by
        # default, this is the list of who to look up in ZoomInfo next.
        where = " AND " + cadence.research_where("a")
    else:
        where = " AND COALESCE(a.archived_at, '') = ''"
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
    visited = args.get("visited", "")
    if visited in ("no", "yes"):
        where += (" AND " + ("NOT " if visited == "no" else "")
                  + f"EXISTS (SELECT 1 FROM interactions dk WHERE dk.account_id = a.id "
                    f"AND dk.interaction_type = '{DOOR_KNOCK}')")
    else:
        visited = ""
    return {"status": status, "visited": visited, "milestone": milestone, "q": q,
            "min_matching_raw": min_matching_raw, "sort": sort, "view": view,
            "where": where, "params": params}


@app.route("/accounts")
def accounts():
    f = _account_filters(request.args)
    status, milestone, q = f["status"], f["milestone"], f["q"]
    min_matching_raw, sort, view = f["min_matching_raw"], f["sort"], f["view"]
    where, params = f["where"], f["params"]
    visited = f["visited"]

    page = _parse_int(request.args.get("page", "1"), on_error=1) or 1
    page = max(1, page)

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
                        WHERE i.account_id = a.id) AS last_activity,
                       (SELECT MAX(created_at) FROM interactions dk
                        WHERE dk.account_id = a.id
                          AND dk.interaction_type = '{DOOR_KNOCK}') AS last_visit
                FROM accounts a WHERE 1=1{where}
                ORDER BY {ACCOUNT_SORTS[sort]}
                LIMIT ? OFFSET ?""",
            params + [ACCOUNTS_PER_PAGE, (page - 1) * ACCOUNTS_PER_PAGE]).fetchall()
        archived_count = conn.execute(
            "SELECT COUNT(*) c FROM accounts "
            "WHERE COALESCE(archived_at, '') != ''").fetchone()["c"]
        research_count = conn.execute(
            f"SELECT COUNT(*) c FROM accounts WHERE {cadence.RESEARCH_SQL}"
        ).fetchone()["c"]
        # When a row only matched because of someone in its contact list, say
        # who — otherwise the result looks like a mystery.
        match_hints = _contact_match_hints(conn, q, rows) if q else {}
    finally:
        conn.close()
    return render_template("accounts.html", accounts=rows,
                           status=status, milestone=milestone, q=q, visited=visited,
                           min_matching=min_matching_raw, sort=sort,
                           total_matching=total_matching, view=view,
                           archived_count=archived_count,
                           research_count=research_count,
                           match_hints=match_hints,
                           total=total, page=page, pages=pages,
                           per_page=ACCOUNTS_PER_PAGE)


@app.route("/accounts/door-knock.csv")
def door_knock_list():
    """The accounts on screen with their office address, sorted by zip code
    then street, so neighbouring offices sit together for planning a route.
    Defaults to the active accounts you haven't visited yet."""
    args = request.args.to_dict()
    args.setdefault("visited", "no")
    f = _account_filters(args)
    conn = get_db()
    try:
        rows = conn.execute(
            f"""SELECT a.*, (SELECT MAX(created_at) FROM interactions dk
                    WHERE dk.account_id = a.id AND dk.interaction_type = '{DOOR_KNOCK}')
                    AS last_visit
                FROM accounts a WHERE 1=1{f['where']}""", f["params"]).fetchall()
    finally:
        conn.close()
    data = []
    for r in rows:
        if not account_address(r):
            continue
        street, city, state, zip_ = r["street"], r["city"], r["state"], r["zip"]
        if not (street or city):
            a = importer.address_from_notes(r["notes"])
            street, city, state, zip_ = a["street"], a["city"], a["state"], a["zip"]
        data.append([zip_ or "", street or "", city or "", state or "", r["company_name"],
                     " ".join(x for x in (r["first_name"], r["last_name"]) if x),
                     r["title"] or "", r["work_phone"] or r["mobile_phone"] or "",
                     r["matching_properties"] or "", (r["last_visit"] or "")[:10],
                     r["prospecting_status"], r["pipeline_milestone"]])
    data.sort(key=lambda d: (d[0] == "", d[0], d[2].lower(), d[1].lower()))
    return _csv_response(data, ["Zip", "Street", "City", "State", "Company", "Contact",
                                "Title", "Phone", "Matching buildings", "Last visit",
                                "Status", "Milestone"],
                         f"door-knock-list-{today_iso()}.csv")


# Columns ZoomInfo's company-list upload recognizes on its own.
ZOOMINFO_LIST_HEADERS = ["Company Name", "Website", "Street", "City", "State",
                         "Zip Code", "Country"]


@app.route("/accounts/zoominfo.csv")
def zoominfo_list():
    """The accounts on screen (Research by default) as a company list to
    upload to ZoomInfo: find the people there in one search, export them,
    and upload that file back on the Import page."""
    args = request.args.to_dict()
    args.setdefault("view", "research")
    f = _account_filters(args)
    conn = get_db()
    try:
        rows = conn.execute(
            f"SELECT a.company_name, a.website, a.notes, a.street, a.city, a.state, a.zip "
            f"FROM accounts a "
            f"WHERE 1=1{f['where']} ORDER BY {ACCOUNT_SORTS[f['sort']]}",
            f["params"]).fetchall()
    finally:
        conn.close()
    data = []
    for r in rows:
        addr = ({"street": r["street"], "city": r["city"], "state": r["state"],
                 "zip": r["zip"]} if (r["street"] or r["city"])
                else importer.address_from_notes(r["notes"]))
        data.append([r["company_name"], r["website"] or "", addr["street"],
                     addr["city"], addr["state"], addr["zip"], "United States"])
    return _csv_response(data, ZOOMINFO_LIST_HEADERS,
                         f"zoominfo-companies-{today_iso()}.csv")


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
            ready = cadence.has_contact(fields)
            cur = conn.execute(
                f"INSERT INTO accounts ({','.join(cols)}) "
                f"VALUES ({','.join('?' * len(cols))})",
                (*fields.values(), today_iso() if ready else "", ts, ts))
            conn.commit()
            new_id = cur.lastrowid
        finally:
            conn.close()
        flash(f"Account “{fields['company_name']}” created."
              + ("" if ready else " It's in Research until you add a contact "
                 "with an email or phone — then its cadence starts."), "success")
        return redirect(url_for("account_detail", account_id=new_id))
    return render_template("account_form.html", account=None)


@app.route("/accounts/<int:account_id>")
def account_detail(account_id):
    return _render_account_detail(account_id)


def _render_account_detail(account_id, **extra):
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        # Untried people first, in the order they'd be tried; then the history.
        untried = _next_contacts(conn, account_id)
        contacts = untried + conn.execute(
            "SELECT * FROM contacts WHERE account_id = ? AND COALESCE(tried_status, '') != '' "
            "ORDER BY tried_at DESC, id DESC", (account_id,)).fetchall()
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
        finished = _finished_cadences(conn, account_id)
        finished = finished[0] if finished else None
        visits = conn.execute(
            "SELECT created_at, outcome FROM interactions WHERE account_id = ? "
            "AND interaction_type = ? ORDER BY created_at DESC, id DESC",
            (account_id, DOOR_KNOCK)).fetchall()
        steps = cadence.get_cadence_progress(conn, account_id)
        # What to do with this account now: the due step, else the next one.
        next_step = (next((st for st in steps if st["state"] == "due"), None)
                     or next((st for st in steps if st["state"] == "upcoming"), None))
        last_touch = conn.execute(
            "SELECT MAX(created_at) FROM interactions WHERE account_id=?",
            (account_id,)).fetchone()[0]
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
                           bids=bids, history_total=history_total, visits=visits,
                           finished=finished, TRIED_STATUSES=TRIED_STATUSES,
                           next_step=next_step, last_touch=last_touch,
                           RECYCLE_DAYS=RECYCLE_DAYS,
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
        if fields["email"].lower() != (existing["email"] or "").strip().lower():
            conn.execute("UPDATE accounts SET email_bounced='' WHERE id=?", (account_id,))
        started = cadence.start_cadence_if_ready(conn, account_id)
        conn.commit()
    finally:
        conn.close()
    flash("Account updated." + (" It has a contact now, so its cadence starts today."
                                if started else ""), "success")
    return redirect(url_for("account_detail", account_id=account_id))


def _restart_cadence(conn, ids: list[int]) -> list[dict]:
    """Start a fresh cadence from today for these accounts.

    Steps already logged would otherwise count as done and the new cycle would
    have nothing to do, so they're kept as history but retagged as General
    Notes. Returns the undo ops for everything this changed besides the
    account rows themselves (callers snapshot those)."""
    ph = ",".join("?" * len(ids))
    steps = [step for _, step in cadence.CADENCE_STEPS]
    sph = ",".join("?" * len(steps))
    retagged = [{"id": r["id"], "interaction_type": r["interaction_type"],
                 "notes": r["notes"]} for r in conn.execute(
        f"SELECT id, interaction_type, notes FROM interactions "
        f"WHERE account_id IN ({ph}) AND interaction_type IN ({sph})", (*ids, *steps))]
    dismissals = undo_module.capture(conn, "cadence_dismissals",
                                     f"account_id IN ({ph})", ids)
    conn.execute(
        f"UPDATE accounts SET cadence_start=?, prospecting_status=?, "
        f"pipeline_milestone=?, updated_at=? WHERE id IN ({ph})",
        (today_iso(), cadence.ACTIVE_STATUS, cadence.ACTIVE_MILESTONE, now_iso(), *ids))
    conn.execute(f"DELETE FROM cadence_dismissals WHERE account_id IN ({ph})", ids)
    conn.execute(
        f"UPDATE interactions SET interaction_type='General Note', "
        f"notes='[' || interaction_type || ' — previous cadence] ' || notes "
        f"WHERE account_id IN ({ph}) AND interaction_type IN ({sph})", (*ids, *steps))
    return [undo_module.update_op("interactions", retagged),
            undo_module.insert_op("cadence_dismissals", dismissals)]


@app.route("/accounts/<int:account_id>/next-contact", methods=["POST"])
def next_contact(account_id):
    """The current person's cadence ended without a reply: start the next."""
    contact_id = request.form.get("contact_id") or None
    status = request.form.get("status", "No reply")
    status = status if status in TRIED_STATUSES else "No reply"
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        who = _person_name(acct)
        rotated = _rotate_contact(conn, account_id, status, contact_id)
        if rotated:
            name, ops = rotated
            _offer_undo(conn, f"moving on from {who or 'the contact'} to {name}", ops)
            conn.commit()
            flash(f"{who or 'They'} marked {status.lower()}. {name}'s cadence starts today, "
                  f"and Email 1 mentions {who or 'the first contact'}.", "success")
        else:
            flash("Nobody untried is left at this company. Rest it for "
                  f"{RECYCLE_DAYS} days, or add someone new.", "warning")
    finally:
        conn.close()
    return redirect(request.form.get("next") or url_for("account_detail", account_id=account_id))


@app.route("/accounts/<int:account_id>/email-bounced", methods=["POST"])
def email_bounced(account_id):
    """The primary's email bounced: stop emailing them, keep calling and
    texting. The cadence skips its email steps from here on. Undoable;
    "undo" also un-marks it (e.g. after fixing a typo, edit the email)."""
    clear = bool(request.form.get("clear"))
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        _offer_undo(conn, ("un-marking" if clear else "marking") + f" {acct['email']} as bounced",
                    [undo_module.update_op("accounts", [{"id": account_id,
                                                         "email_bounced": acct["email_bounced"]}])])
        conn.execute("UPDATE accounts SET email_bounced=?, updated_at=? WHERE id=?",
                     ("" if clear else today_iso(), now_iso(), account_id))
        conn.commit()
    finally:
        conn.close()
    if clear:
        flash("Email marked as working again. Email steps are back in the cadence.", "success")
    else:
        flash(f"{acct['email']} marked as bounced. No more emails to "
              f"{acct['first_name'] or 'them'}: the cadence skips its email steps and "
              f"keeps the calls and texts.", "success")
    return redirect(request.form.get("next") or url_for("account_detail", account_id=account_id))


@app.route("/accounts/<int:account_id>/rest", methods=["POST"])
def rest_account(account_id):
    """Everyone has been tried: rest the company and bring it back later."""
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        before = undo_module.capture(conn, "accounts", "id=?", (account_id,))
        ops = []
        effect = _rest_account(conn, account_id, "No reply")
        _offer_undo(conn, f"resting {acct['company_name']}",
                    ops + [undo_module.update_op("accounts", before)])
        conn.commit()
    finally:
        conn.close()
    flash("Rested." + effect, "success")
    return redirect(request.form.get("next") or url_for("account_detail", account_id=account_id))


@app.route("/accounts/<int:account_id>/start-cadence", methods=["POST"])
def start_cadence(account_id):
    """Start the clock on an account waiting in Research, without a contact.

    Unlike a restart this keeps any progress: an account sent back to Research
    by a bad number carries on from the step it was on."""
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        if not (acct["cadence_start"] or "").strip():
            conn.execute("UPDATE accounts SET cadence_start=?, updated_at=? WHERE id=?",
                         (today_iso(), now_iso(), account_id))
            conn.commit()
    finally:
        conn.close()
    flash("Cadence started from today.", "success")
    return redirect(url_for("account_detail", account_id=account_id))


@app.route("/accounts/<int:account_id>/restart-cadence", methods=["POST"])
def restart_cadence(account_id):
    """Reset the cadence clock to today and clear prior step history.

    Also how an account waiting in Research is started by hand — say, to
    cold-call a switchboard before a named contact has been found."""
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        ops = [undo_module.update_op("accounts", [{
            "id": account_id, "cadence_start": acct["cadence_start"],
            "prospecting_status": acct["prospecting_status"],
            "pipeline_milestone": acct["pipeline_milestone"]}])]
        ops += _restart_cadence(conn, [account_id])
        _offer_undo(conn, "the cadence restart", ops)
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
            # Restarting also wipes the checked-off steps, so the undo has to
            # carry them too — otherwise it restores the dates but not them.
            # Same as restarting one account: logged steps are retagged as
            # history, or a worked account would restart with nothing to do.
            before = snapshot("cadence_start", "prospecting_status", "pipeline_milestone")
            _offer_undo(conn, f"restarting the cadence on {n} account(s)",
                        before + _restart_cadence(conn, ids))
            flash(f"Restarted the cadence on {n} account(s) from today.", "success")

        elif action == "delete":
            # A company has to be archived before it can be erased — enforced
            # here, not just by which menu shows the option, so a stray form
            # post can't wipe a live account. Still undoable, photos included.
            rows = [r for r in rows if r["archived_at"]]
            if not rows:
                flash("Only archived accounts can be deleted permanently. "
                      "Archive them first.", "warning")
                return redirect(back)
            refused = n - len(rows)
            ids = [r["id"] for r in rows]
            ph = ",".join("?" * len(ids))
            n = len(ids)
            if refused:
                flash(f"{refused} account(s) weren't archived and were left alone.",
                      "warning")
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
        outcome = request.form.get("outcome", "")
        effect = _log_interaction(conn, account_id, itype, notes, outcome)
        conn.commit()
    finally:
        conn.close()
    flash(f"Logged “{itype}”" + (f" — {outcome}" if outcome in ALL_OUTCOMES else "")
          + "." + effect, "success")
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
        outcome = request.form.get("outcome", row["outcome"] or "")
        if outcome not in ALL_OUTCOMES:
            outcome = ""
        # Editing only corrects the record; it doesn't re-run what the outcome
        # did at the time (moving the deal, etc.) — that would be surprising.
        conn.execute(
            "UPDATE interactions SET interaction_type=?, notes=?, outcome=?, "
            "created_at=? WHERE id=?",
            (itype, notes, outcome, created_at, interaction_id))
        _offer_undo(conn, "that edit", [undo_module.update_op("interactions", [
            {"id": row["id"], "interaction_type": row["interaction_type"],
             "notes": row["notes"], "outcome": row["outcome"] or "",
             "created_at": row["created_at"]}])])
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
    # A new person means a new address: a bounce belonged to the old one.
    conn.execute(
        f"UPDATE accounts SET {','.join(c + '=?' for c in CONTACT_COLS)}, "
        f"email_bounced='', updated_at=? WHERE id=?",
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
        _save_zoominfo_url(conn, account_id, paste)
        started = cadence.start_cadence_if_ready(conn, account_id)
        conn.commit()
    finally:
        conn.close()
    flash(f"Contact {person['first_name']} {person['last_name']} added"
          + (" as primary." if make_primary else ".") + parsed_note
          + (" The account had nobody to contact before, so its cadence starts "
             "today." if started else ""), "success")
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

    # Anything already typed into the form wins — pressing "Read the paste"
    # after typing the name must not wipe the name.
    typed = {c: request.form.get(c, "").strip() for c in CONTACT_COLS}
    for c in CONTACT_COLS:
        if typed[c]:
            parsed[c] = typed[c]

    guessed_name = False
    if not (parsed["first_name"] or parsed["last_name"]) and parsed["email"]:
        # Last resort: michael.delacruz@ gives up a name; mdelacruz@ doesn't,
        # and name_from_email leaves it alone rather than inventing one.
        first, last = importer.name_from_email(parsed["email"])
        if first or last:
            parsed["first_name"], parsed["last_name"] = first, last
            guessed_name = True
    if not parsed["seniority"] and parsed["title"]:
        parsed["seniority"] = importer.seniority_from_title(parsed["title"])

    found = [c for c in CONTACT_COLS if parsed.get(c)]
    label = lambda c: c.replace("_", " ").replace("linkedin url", "LinkedIn")
    if not found:
        flash("Nothing recognisable in that paste — no name, email or phone "
              "number. Try selecting the whole ZoomInfo page (Ctrl+A, Ctrl+C), "
              "or just type the details in below.", "warning")
    elif not (parsed["first_name"] or parsed["last_name"]):
        # ZoomInfo's "Contact Details" panel has no name in it — the name sits
        # higher up the page. Say what came through so it's clear what's left.
        flash("Got the " + ", ".join(label(c) for c in found)
              + " — but no name. ZoomInfo keeps the name at the top of the "
                "page, above the Contact Details panel: select the whole page "
                "(Ctrl+A, Ctrl+C) and it comes through with everything else. "
                "Or just type the name below.", "warning")
    else:
        msg = "Read: " + ", ".join(label(c) for c in found) + "."
        if guessed_name:
            msg += (" The name came from the email address, so check the "
                    "spelling.")
        missing = [label(c) for c in ("email", "work_phone", "title")
                   if not parsed.get(c)]
        if missing:
            msg += " Didn't find: " + ", ".join(missing) + " — add below if you have it."
        msg += " Check it over, then press Add Contact."
        flash(msg, "warning" if guessed_name else "info")
    return _render_account_detail(account_id, contact_prefill=parsed,
                                  contact_paste=paste)


@app.route("/accounts/<int:account_id>/contacts/roster", methods=["POST"])
def roster_contacts(account_id):
    """Find everyone on a pasted ZoomInfo employee list, to pick from.

    Nothing is saved here. A company's Employees tab lists the whole org, but
    only a handful fit the ICP — so this shows the list with the management
    level beside each name and you tick the three or four worth having.
    """
    paste = request.form.get("roster_paste", "").strip()
    if not paste:
        flash("Paste the ZoomInfo employee list into the box first.", "warning")
        return redirect(url_for("account_detail", account_id=account_id))
    people = importer.parse_contact_roster(paste)
    if not people:
        flash("No people found in that paste. On the company's Employees tab, "
              "press Ctrl+A then Ctrl+C and paste the whole page — each person "
              "is picked out by their ZoomInfo profile link.", "warning")
        return redirect(url_for("account_detail", account_id=account_id))

    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        existing = {(r["first_name"].lower(), r["last_name"].lower())
                    for r in conn.execute(
                        "SELECT first_name, last_name FROM contacts WHERE account_id=?",
                        (account_id,))}
        existing.add((acct["first_name"].lower(), acct["last_name"].lower()))
        # The one thing a preview keeps: where this company lives in ZoomInfo.
        _save_zoominfo_url(conn, account_id, paste)
        conn.commit()
    finally:
        conn.close()
    for person in people:
        person["already"] = (person["first_name"].lower(),
                             person["last_name"].lower()) in existing
    fresh = sum(1 for p in people if not p["already"])
    detailed = sum(1 for p in people if p["has_details"])
    msg = (f"Found {len(people)} people"
           + (f", {len(people) - fresh} already on this account" if fresh < len(people) else "")
           + ". Tick the ones worth contacting and press Add Selected.")
    if not detailed:
        msg += (" None of the rows were expanded, so no emails or phone numbers "
                "came through — expand the few you want on ZoomInfo before "
                "copying and they'll come with.")
    flash(msg, "info" if detailed else "warning")
    return _render_account_detail(account_id, roster=people,
                                  roster_detailed=detailed)


@app.route("/accounts/<int:account_id>/contacts/roster/add", methods=["POST"])
def roster_add(account_id):
    """Create a contact for each person ticked on the roster preview."""
    picked = request.form.getlist("pick")
    if not picked:
        flash("Tick at least one person first.", "warning")
        return redirect(url_for("account_detail", account_id=account_id))
    primary_pick = request.form.get("primary", "")

    conn = get_db()
    added, bare, skipped = [], [], 0
    try:
        acct = _account_or_404(conn, account_id)
        existing = {(r["first_name"].lower(), r["last_name"].lower())
                    for r in conn.execute(
                        "SELECT first_name, last_name FROM contacts WHERE account_id=?",
                        (account_id,))}
        existing.add((acct["first_name"].lower(), acct["last_name"].lower()))
        ts = now_iso()
        for idx in picked:
            person = {c: request.form.get(f"p{idx}_{c}", "").strip()
                      for c in CONTACT_COLS}
            if not (person["first_name"] or person["last_name"]):
                continue
            key = (person["first_name"].lower(), person["last_name"].lower())
            if key in existing:
                skipped += 1
                continue
            existing.add(key)
            make_primary = (idx == primary_pick) or not _has_primary_contact(acct)
            if make_primary:
                _demote_primary_to_contact(conn, acct)
                _set_primary_contact(conn, account_id, person)
                acct = _account_or_404(conn, account_id)   # it has one now
            else:
                conn.execute(
                    f"INSERT INTO contacts (account_id, {','.join(CONTACT_COLS)}, created_at) "
                    f"VALUES ({','.join('?' * (len(CONTACT_COLS) + 2))})",
                    (account_id, *(person[c] for c in CONTACT_COLS), ts))
            who = f"{person['first_name']} {person['last_name']}".strip()
            added.append(who)
            if not (person["email"] or person["work_phone"] or person["mobile_phone"]):
                bare.append(who)
        started = cadence.start_cadence_if_ready(conn, account_id)
        conn.commit()
    finally:
        conn.close()
    if added:
        msg = (f"Added {len(added)} contact(s): " + ", ".join(added) + "."
               + (f" {skipped} were already on the account." if skipped else ""))
        if bare:
            msg += (f" {len(bare)} came with no email or phone "
                    f"({', '.join(bare)}) — expand those rows on ZoomInfo and "
                    f"paste again, or add the details by hand.")
        if started:
            msg += " This account now has someone to contact, so its cadence starts today."
        flash(msg, "success")
    else:
        flash("Nothing added — those people are already on this account.", "warning")
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
    "my_phone": "your phone number", "my_title": "your title",
    "my_website": "your website", "my_email": "your email",
    "previous_contact": "the person contacted before",
}

# Sign-off lines a template may end with. With a signature set they're
# dropped, so "Best,\n{my_name}" becomes "Best,\n\n<signature>" rather than
# the name twice.
_SIGNOFF_LINE = re.compile(r"^\s*(\{my_(name|title|company|phone|website|email)\}\s*)+$")


def with_signature(body: str, settings: dict) -> str:
    """An email template body with the saved signature at the end.

    {signature} in a template places it explicitly; otherwise any trailing
    {my_name}/{my_company}/{my_phone} lines are replaced by it."""
    sig = (settings.get("email_signature") or "").strip("\n")
    if "{signature}" in body:
        return body.replace("{signature}", sig)
    if not sig.strip():
        return body
    lines = body.rstrip().split("\n")
    while lines and _SIGNOFF_LINE.match(lines[-1]):
        lines.pop()
    # A blank line between the sign-off ("Best,") and the signature.
    return "\n".join(lines).rstrip() + "\n\n" + sig


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
        "my_title": settings.get("my_title", ""),
        "my_website": settings.get("my_website", ""),
        "my_email": settings.get("my_email", ""),
        "previous_contact": (acct["previous_contact"]
                             if "previous_contact" in acct.keys() else "") or "",
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
                           CALL_APPS=CALL_APPS, EMAIL_APPS=EMAIL_APPS,
                           PHONE_EMAIL_APPS=PHONE_EMAIL_APPS,
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
                    "price_add_20", "daily_goal", "call_app", "email_app",
                    "email_app_phone", "send_from", "email_signature"):
            if key not in request.form:  # only touch submitted fields
                continue
            value = request.form.get(key, "").strip()
            if key == "call_app" and value not in CALL_APPS \
                    or key == "email_app" and value not in EMAIL_APPS \
                    or key == "email_app_phone" and value not in PHONE_EMAIL_APPS:
                continue
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
    """Render templates for an account, optionally filtered to a cadence step.

    A template that mentions {previous_contact} is for the second person
    onward, so it only shows once someone has been tried, and then first."""
    has_prev = bool(("previous_contact" in acct.keys() and acct["previous_contact"]))
    rows = [t for t in rows if has_prev or "{previous_contact}" not in (t["body"] or "")]
    if has_prev:
        rows = sorted(rows, key=lambda t: "{previous_contact}" not in (t["body"] or ""))
    scripts = []
    for t in rows:
        t_steps = [s.strip() for s in (t["steps"] or "").split(",") if s.strip()]
        if step and step not in t_steps:
            continue
        subject, missing_s = render_script(t["subject"], acct, settings)
        raw_body = with_signature(t["body"], settings) if t["kind"] == "email" else t["body"]
        body, missing_b = render_script(raw_body, acct, settings)
        if t["kind"] == "email":
            bounced = "email_bounced" in acct.keys() and bool(acct["email_bounced"])
            action_url, external = email_link("" if bounced else acct["email"], subject, body)
            missing_target = ("" if action_url else
                              "working email address (it bounced)" if bounced else "email address")
        elif t["kind"] == "text":
            action_url, external = phone_link(acct["mobile_phone"], "text", body)
            missing_target = "" if action_url else "mobile number"
        else:
            action_url, external = phone_link(
                acct["work_phone"] or acct["mobile_phone"], "call")
            missing_target = "" if action_url else "phone number"
        scripts.append({
            "template": t, "subject": subject, "body": body,
            "missing": list(dict.fromkeys(missing_s + missing_b)),
            "action_url": action_url, "log_steps": t_steps,
            "action_attrs": _attrs(action_url, external) if action_url else "",
            "missing_target": missing_target,
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


# An unpaced import bigger than this gets a warning pointing at Re-Pace.
UNPACED_WARNING = 30


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
                if not per_day and result["imported"] > UNPACED_WARNING:
                    # Unpaced, every account's Day 1 lands today and the rest
                    # of its steps land together after it — this is how a
                    # list turns into thousands of overdue tasks.
                    flash(f"All {result['imported']} accounts start their cadence "
                          f"today, so they'll all come due together. Use Re-Pace "
                          f"Cadence below to spread them out — about a fifth of "
                          f"the touches you can do in a day, e.g. 10.", "warning")
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


# ------------------------------------------------------- Merge duplicates

def _merge_accounts(conn, keep_id: int, loser_ids: list[int]) -> dict:
    """Fold duplicate accounts into one. Returns a summary; records an undo.

    Everything moves to the account being kept — people, history, roof
    reports, projects, checked-off steps — then the duplicates are deleted.
    Nothing is overwritten: the kept account's own details win, and a
    duplicate's details only fill gaps (its notes are appended, labelled).

    Undo restores all of it. Roof reports and projects are moved back by
    pointing them at their old account again, never by re-inserting them:
    re-inserting a row over itself would cascade-delete its photos and
    invoices.
    """
    keeper = conn.execute("SELECT * FROM accounts WHERE id=?", (keep_id,)).fetchone()
    keeper_before = {k: keeper[k] for k in keeper.keys()}
    ts = now_iso()

    def name_key(first, last):
        return ((first or "").strip().lower(), (last or "").strip().lower())

    people = {name_key(keeper["first_name"], keeper["last_name"])} - {("", "")}
    people |= {name_key(c["first_name"], c["last_name"]) for c in conn.execute(
        "SELECT first_name, last_name FROM contacts WHERE account_id=?", (keep_id,))}
    has_primary = bool((keeper["first_name"] or keeper["last_name"] or "").strip())

    updates: dict = {}
    notes = keeper["notes"] or ""
    losers, loser_contacts, loser_dismissals = [], [], []
    moved_children: dict[str, list] = {"interactions": [], "bids": [], "projects": []}
    new_contacts, new_dismissals = [], []
    summary = {"merged": [], "people": 0, "touches": 0, "reports": 0, "projects": 0}

    def current(field):
        return updates.get(field, keeper[field])

    for lid in loser_ids:
        loser = conn.execute("SELECT * FROM accounts WHERE id=?", (lid,)).fetchone()
        losers.append(dict(loser))
        loser_contacts += undo_module.capture(conn, "contacts", "account_id=?", (lid,))
        loser_dismissals += undo_module.capture(conn, "cadence_dismissals",
                                                "account_id=?", (lid,))
        summary["merged"].append(loser["company_name"])

        # The duplicate's primary person: becomes ours if we have nobody,
        # otherwise joins our contacts (unless they're already here).
        lkey = name_key(loser["first_name"], loser["last_name"])
        if lkey != ("", ""):
            if lkey not in people:
                if not has_primary:
                    for f in CONTACT_COLS:
                        updates[f] = loser[f] or ""
                    has_primary = True
                else:
                    cur = conn.execute(
                        f"INSERT INTO contacts (account_id, {','.join(CONTACT_COLS)}, created_at) "
                        f"VALUES ({','.join('?' * (len(CONTACT_COLS) + 2))})",
                        (keep_id, *(loser[f] or "" for f in CONTACT_COLS), ts))
                    new_contacts.append(cur.lastrowid)
                people.add(lkey)
                summary["people"] += 1
        else:
            # No named person — but a switchboard number or general email is
            # still worth keeping if we don't have one.
            for f in ("email", "work_phone", "mobile_phone"):
                if not (current(f) or "").strip() and (loser[f] or "").strip():
                    updates[f] = loser[f]

        for c in conn.execute("SELECT * FROM contacts WHERE account_id=?", (lid,)).fetchall():
            ck = name_key(c["first_name"], c["last_name"])
            if ck in people:
                continue                       # already here; goes with the duplicate
            conn.execute("UPDATE contacts SET account_id=? WHERE id=?", (keep_id, c["id"]))
            people.add(ck)
            summary["people"] += 1

        for table, label in (("interactions", "touches"), ("bids", "reports"),
                             ("projects", "projects")):
            for r in conn.execute(f"SELECT id FROM {table} WHERE account_id=?", (lid,)).fetchall():
                moved_children[table].append({"id": r["id"], "account_id": lid})
                summary[label] += 1
            conn.execute(f"UPDATE {table} SET account_id=? WHERE account_id=?", (keep_id, lid))

        for d in conn.execute("SELECT step_type, dismissed_at FROM cadence_dismissals "
                              "WHERE account_id=?", (lid,)).fetchall():
            cur = conn.execute("INSERT OR IGNORE INTO cadence_dismissals "
                               "(account_id, step_type, dismissed_at) VALUES (?,?,?)",
                               (keep_id, d["step_type"], d["dismissed_at"]))
            if cur.rowcount:
                new_dismissals.append(cur.lastrowid)

        # Company-level gaps.
        for f in ("num_properties", "matching_properties"):
            if loser[f] is not None and (current(f) is None or loser[f] > current(f)):
                updates[f] = loser[f]
        if not (current("next_follow_up") or "") and (loser["next_follow_up"] or ""):
            updates["next_follow_up"] = loser["next_follow_up"]
            updates["follow_up_note"] = loser["follow_up_note"] or ""
        if (current("preferred_contact") or "Unknown") == "Unknown" \
                and (loser["preferred_contact"] or "Unknown") != "Unknown":
            updates["preferred_contact"] = loser["preferred_contact"]
        if not (current("cadence_start") or "") and (loser["cadence_start"] or ""):
            updates["cadence_start"] = loser["cadence_start"]   # keep a running clock
        if (loser["notes"] or "").strip() and loser["notes"].strip() not in notes:
            notes = (notes + "\n\n" if notes else "") + \
                f"— merged from “{loser['company_name']}” —\n{loser['notes'].strip()}"

        conn.execute("DELETE FROM accounts WHERE id=?", (lid,))

    updates["notes"] = notes
    updates["updated_at"] = ts
    conn.execute(f"UPDATE accounts SET {','.join(k + '=?' for k in updates)} WHERE id=?",
                 (*updates.values(), keep_id))

    _offer_undo(conn, f"merging into “{keeper['company_name']}”", [
        undo_module.insert_op("accounts", losers),
        undo_module.insert_op("contacts", loser_contacts),
        undo_module.insert_op("cadence_dismissals", loser_dismissals),
        undo_module.update_op("interactions", moved_children["interactions"]),
        undo_module.update_op("bids", moved_children["bids"]),
        undo_module.update_op("projects", moved_children["projects"]),
        undo_module.delete_op("contacts", new_contacts),
        undo_module.delete_op("cadence_dismissals", new_dismissals),
        undo_module.update_op("accounts", [keeper_before]),
    ])
    return summary


@app.route("/duplicates/merge", methods=["POST"])
def merge_duplicates():
    back = url_for("import_page") + "#duplicates"
    try:
        keep_id = int(request.form.get("keep", ""))
        ids = sorted({int(i) for i in request.form.getlist("account_ids")})
    except ValueError:
        flash("Pick which account to keep first.", "warning")
        return redirect(back)
    losers = [i for i in ids if i != keep_id]
    if keep_id not in ids or not losers:
        flash("Pick which account to keep first.", "warning")
        return redirect(back)

    conn = get_db()
    try:
        rows = conn.execute(
            f"SELECT id, company_name, COALESCE(archived_at, '') AS archived_at "
            f"FROM accounts WHERE id IN ({','.join('?' * len(ids))})", ids).fetchall()
        # Only ever merge accounts that really are the same company: all
        # present, all active, all normalising to the same name. A tampered or
        # stale form can't fold two unrelated companies together.
        keys = {importer.normalize_company(r["company_name"]) for r in rows}
        if len(rows) != len(ids) or any(r["archived_at"] for r in rows) or len(keys) != 1:
            flash("Those accounts can't be merged — they're no longer duplicates "
                  "of each other. The list below is up to date.", "warning")
            return redirect(back)
        summary = _merge_accounts(conn, keep_id, losers)
        conn.commit()
        kept = conn.execute("SELECT company_name FROM accounts WHERE id=?",
                            (keep_id,)).fetchone()["company_name"]
    finally:
        conn.close()
    bits = [f"{summary[k]} {label}" for k, label in
            (("people", "people"), ("touches", "logged touches"),
             ("reports", "roof reports"), ("projects", "projects")) if summary[k]]
    flash(f"Merged {len(losers)} duplicate(s) into “{kept}”"
          + (": moved " + ", ".join(bits) if bits else "") + ".", "success")
    return redirect(url_for("account_detail", account_id=keep_id))


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
                 AND COALESCE(a.cadence_start, '') != ''
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
    per_day_raw = request.form.get("per_day", "").strip()
    per_day = int(per_day_raw) if per_day_raw.isdigit() and int(per_day_raw) > 0 else None
    contact_result = None
    duplicates = []
    conn = get_db()
    try:
        if not file or not file.filename:
            flash("Choose a .xlsx or .csv contact file first.", "danger")
        else:
            try:
                contact_result = importer.import_contacts(
                    conn, file, create_missing=create_missing, per_day=per_day,
                    record_undo=lambda ops: _offer_undo(
                        conn, f"Contact upload ({file.filename})", ops))
                msg = (f"Attached {contact_result['attached']} contact(s) to "
                       f"{contact_result['companies_matched']} account(s).")
                if contact_result["accounts_created"]:
                    msg += (f" Opened {contact_result['accounts_created']} new "
                            f"account(s) for companies you didn't have.")
                started = contact_result["cadence_started"]
                if started and per_day:
                    msg += (f" {started} cadence(s) start at {per_day}/business day, "
                            f"the last on {contact_result['last_start_date']}.")
                elif started:
                    msg += f" {started} cadence(s) start today."
                flash(msg, "success")
                if not per_day and started > UNPACED_WARNING:
                    flash(f"{started} accounts started their cadence today — that's "
                          f"{started} Day-1 tasks at once. Undo, then upload again "
                          f"with a per-day number, or use Re-Pace below.", "warning")
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


@app.route("/accounts/<int:account_id>/contacts/upload", methods=["POST"])
def upload_account_contacts(account_id):
    """A ZoomInfo export pulled for this one company, attached straight to
    it: no company-name matching, people already here skipped, undoable."""
    file = request.files.get("file")
    if not file or not file.filename:
        flash("Choose the .xlsx or .csv file ZoomInfo gave you first.", "warning")
        return redirect(url_for("account_detail", account_id=account_id))
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        try:
            res = importer.import_contacts(
                conn, file, to_account=account_id,
                record_undo=lambda ops: _offer_undo(
                    conn, f"Contact upload to {acct['company_name']}", ops))
        except ValueError as e:
            flash(str(e), "danger")
            return redirect(url_for("account_detail", account_id=account_id))
    finally:
        conn.close()
    msg = f"Added {res['attached']} contact(s) from {file.filename}."
    if res["skipped_duplicates"]:
        msg += f" {res['skipped_duplicates']} already on this account were skipped."
    if res["cadence_started"]:
        msg += " It has someone to contact now, so its cadence starts today."
    flash(msg, "success" if res["attached"] else "warning")
    return redirect(url_for("account_detail", account_id=account_id))


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
    notices = init_db()
    print(f"\n  Your data:  {db_module.DATA_DIR}")
    print(f"  If a page errors: {db_module.ERROR_LOG}")
    for line in notices:
        print(f"  ▸ {line}")
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
