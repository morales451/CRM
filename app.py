"""Roof Coating CRM — lightweight local Flask app.

Run:  python3 app.py
Then open http://<your-local-ip>:8000 from any device on your Wi-Fi.
"""

import io
import re
import socket
from datetime import datetime
from urllib.parse import quote

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
    "company": "company", "num_properties": "number of properties",
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


@app.route("/accounts/<int:account_id>/scripts")
def account_scripts(account_id):
    """Templates rendered for one account, optionally filtered to a cadence step."""
    step = request.args.get("step", "")
    conn = get_db()
    try:
        acct = _account_or_404(conn, account_id)
        settings = _get_settings(conn)
        rows = conn.execute(
            "SELECT * FROM templates ORDER BY sort_order, id").fetchall()
    finally:
        conn.close()

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
    print_network_instructions()
    app.run(host="0.0.0.0", port=PORT, debug=False)
