"""Roof Coating CRM — lightweight local Flask app.

Run:  python3 app.py
Then open http://<your-local-ip>:8000 from any device on your Wi-Fi.
"""

import io
import socket
from datetime import datetime

from flask import (Flask, flash, redirect, render_template, request,
                   send_file, url_for)

import cadence
import importer
from db import (INTERACTION_TYPES, PIPELINE_MILESTONES,
                PREFERRED_CONTACT_METHODS, PROSPECTING_STATUSES,
                get_db, init_db, now_iso, today_iso)

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
    }


def _account_or_404(conn, account_id):
    acct = conn.execute("SELECT * FROM accounts WHERE id = ?", (account_id,)).fetchone()
    if acct is None:
        from flask import abort
        abort(404)
    return acct


def _account_fields_from_form(form):
    num_props = form.get("num_properties", "").strip()
    return {
        "company_name": form.get("company_name", "").strip(),
        "first_name": form.get("first_name", "").strip(),
        "last_name": form.get("last_name", "").strip(),
        "title": form.get("title", "").strip(),
        "num_properties": int(num_props) if num_props.isdigit() else None,
        "email": form.get("email", "").strip(),
        "work_phone": form.get("work_phone", "").strip(),
        "mobile_phone": form.get("mobile_phone", "").strip(),
        "preferred_contact": form.get("preferred_contact", "Unknown"),
        "notes": form.get("notes", "").strip(),
        "prospecting_status": form.get("prospecting_status", "Prospecting"),
        "pipeline_milestone": form.get("pipeline_milestone", "None / In Cadence"),
    }


# ---------------------------------------------------------------- Dashboard

@app.route("/")
def dashboard():
    conn = get_db()
    try:
        reminders = cadence.get_due_reminders(conn)
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
                               stats=stats, recent=recent)
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
                    email, work_phone, mobile_phone, preferred_contact, notes,
                    prospecting_status, pipeline_milestone, cadence_start,
                    created_at, updated_at)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
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
        interactions = conn.execute(
            "SELECT * FROM interactions WHERE account_id = ? "
            "ORDER BY created_at DESC, id DESC", (account_id,)).fetchall()
        steps = cadence.get_cadence_progress(conn, account_id)
        in_cadence = (acct["prospecting_status"] == cadence.ACTIVE_STATUS
                      and acct["pipeline_milestone"] == cadence.ACTIVE_MILESTONE)
    finally:
        conn.close()
    return render_template("account_detail.html", account=acct,
                           interactions=interactions, steps=steps,
                           in_cadence=in_cadence)


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
               title=?, num_properties=?, email=?, work_phone=?, mobile_phone=?,
               preferred_contact=?, notes=?, prospecting_status=?,
               pipeline_milestone=?, updated_at=? WHERE id=?""",
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
    return redirect(url_for("account_detail", account_id=account_id))


# ------------------------------------------------------------------- Import

@app.route("/import", methods=["GET", "POST"])
def import_page():
    result = None
    if request.method == "POST":
        file = request.files.get("file")
        if not file or not file.filename:
            flash("Choose a .xlsx or .csv file first.", "danger")
        else:
            conn = get_db()
            try:
                result = importer.import_accounts(conn, file)
                flash(f"Imported {result['imported']} account(s).", "success")
            except ValueError as e:
                flash(str(e), "danger")
            except Exception as e:
                flash(f"Import failed: {e}", "danger")
            finally:
                conn.close()
    return render_template("import.html", result=result)


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
    print_network_instructions()
    app.run(host="0.0.0.0", port=PORT, debug=False)
