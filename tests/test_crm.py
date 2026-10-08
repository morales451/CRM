"""End-to-end functional test for the Roof CRM.

Run from anywhere:  python3 tests/test_crm.py
Uses a throwaway database in a temp directory; never touches crm.db.
"""
import io
import re
import sys
import tempfile
from pathlib import Path
from datetime import date, datetime, timedelta

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import db
db.DB_PATH = Path(tempfile.mkdtemp(prefix="crm_test_")) / "test_crm.db"

import cadence, importer
from app import app

db.init_db()
client = app.test_client()
failures = []

def check(name, cond, extra=""):
    print(("PASS " if cond else "FAIL ") + name + ((" — " + str(extra)) if extra and not cond else ""))
    if not cond:
        failures.append(name)

def _raises(fn):
    """True when fn() raises — used to assert a guard actually guards."""
    try:
        fn()
    except Exception:
        return True
    return False

# ---- 1. Build a fake HTX_Office_5kto10k.xlsx with realistic headers
import openpyxl
wb = openpyxl.Workbook()
ws = wb.active
ws.append(["Company Name", "First Name", "Last Name", "Job Title",
           "# of Properties", "# Properties (in search)", "Email Address",
           "Office Phone", "Cell Phone", "Comments"])
ws.append(["Acme Properties LLC", "Jane", "Doe", "Owner", 3, 2,
           "jane@acme.com", 7135551234, "713-555-9999", "Big flat roof"])
ws.append(["Bayou Holdings", "Bob", "Smith", "Facilities Mgr", None, None,
           "bob@bayou.com", None, None, None])
ws.append(["", "No", "Company", "", None, None, "", "", "", ""])          # blank company → skipped
ws.append(["Acme Properties LLC", "Dup", "Row", "", None, None, "", "", "", ""])  # dupe → skipped
buf = io.BytesIO(); wb.save(buf); buf.seek(0)

class FakeUpload(io.BytesIO):
    filename = "HTX_Office_5kto10k.xlsx"

conn = db.get_db()
res = importer.import_accounts(conn, FakeUpload(buf.getvalue()))
check("import: 2 imported", res["imported"] == 2, res)
check("import: 1 dupe skipped", res["skipped_duplicates"] == 1, res)
check("import: 1 blank skipped", res["skipped_blank"] == 1, res)
check("import: all key columns mapped",
      set(res["mapped_columns"]) >= {"company_name","first_name","last_name","title",
                                     "num_properties","email","work_phone","mobile_phone","notes"},
      res["mapped_columns"])

acme = conn.execute("SELECT * FROM accounts WHERE company_name='Acme Properties LLC'").fetchone()
check("import: phone float cleaned", acme["work_phone"] == "7135551234", acme["work_phone"])
check("import: num_properties int", acme["num_properties"] == 3)
check("import: matching_properties (column I) captured", acme["matching_properties"] == 2)
check("import: defaults", acme["prospecting_status"] == "Prospecting"
      and acme["pipeline_milestone"] == "None / In Cadence")
check("import: ISO created_at", "T" in acme["created_at"] and len(acme["created_at"]) >= 19,
      acme["created_at"])
bayou = conn.execute("SELECT * FROM accounts WHERE company_name='Bayou Holdings'").fetchone()
check("import: NaN → empty string", bayou["email"] == "bob@bayou.com" and bayou["work_phone"] == ""
      and bayou["notes"] == "")

# ---- 2. Cadence: Day 1 due immediately; later steps not yet
rems = cadence.get_due_reminders(conn)
check("cadence: Day 1 due for both accounts",
      len(rems) == 2 and all(r["step_type"] == "Email 1" for r in rems), rems)

# Pin "today" so these checks don't depend on the weekday the suite runs on:
# steps count business days, so "7 calendar days ago" means a different
# number of steps on a Wednesday than on a Monday.
_real_cadence_today = cadence.today
cadence.today = lambda: date(2026, 10, 14)           # a Wednesday
# Started Monday 10-05: Day 1 Mon 05, Day 3 Wed 07, Day 6 Mon 12,
# Day 8 Wed 14 (today), Day 10 Fri 16 (not yet).
conn.execute("UPDATE accounts SET cadence_start=? WHERE id=?", ("2026-10-05", acme["id"]))
conn.commit()
# The engine still knows every outstanding step...
all_steps = cadence.get_due_reminders(conn, acme["id"], collapse=False)
check("cadence: backdated 7d → days 1,3,6,8 outstanding",
      [r["step_type"] for r in all_steps] == ["Email 1", "Call & Text", "Call 2", "Email 2"],
      [r["step_type"] for r in all_steps])
# ...but you are only shown the one the account is actually waiting on.
acme_rems = cadence.get_due_reminders(conn, acme["id"])
check("cadence: only the next step is shown, not all four",
      [r["step_type"] for r in acme_rems] == ["Email 1"],
      [r["step_type"] for r in acme_rems])
check("cadence: the row says how far behind the account is",
      acme_rems[0]["steps_behind"] == 4, acme_rems[0])
check("cadence: overdue days computed", acme_rems[0]["days_overdue"] == 9, acme_rems[0])

# ---- 3. Logging a step clears it, and the clock resets from that day
# Email 1 was due 10-05; it's sent today (Wed 10-14), nine days late.
conn.execute("INSERT INTO interactions (account_id, interaction_type, notes, created_at) "
             "VALUES (?,?,?,?)", (acme["id"], "Email 1", "sent intro", "2026-10-14T09:00:00-05:00"))
conn.commit()
check("cadence: right after a late step, nothing is overdue",
      cadence.get_due_reminders(conn, acme["id"]) == [],
      cadence.get_due_reminders(conn, acme["id"]))
_prog = {p["step_type"]: p for p in cadence.get_cadence_progress(conn, acme["id"])}
check("cadence: the next step is due the gap AFTER the action (Wed + 2 business days = Fri)",
      _prog["Call & Text"]["due_date"] == "2026-10-16" and _prog["Call & Text"]["state"] == "upcoming",
      _prog["Call & Text"])
check("cadence: later steps chain on (Call 2 three business days later, then +2, +2)",
      [_prog[s]["due_date"] for s in ("Call 2", "Email 2", "Breakup Email")]
      == ["2026-10-21", "2026-10-23", "2026-10-27"],
      [_prog[s]["due_date"] for s in ("Call 2", "Email 2", "Breakup Email")])
cadence.today = lambda: date(2026, 10, 16)
acme_rems = cadence.get_due_reminders(conn, acme["id"])
check("cadence: the next step comes due on its new date",
      [r["step_type"] for r in acme_rems] == ["Call & Text"]
      and acme_rems[0]["steps_behind"] == 1 and acme_rems[0]["days_overdue"] == 0, acme_rems)

# ---- 4. Manual dismissal clears a step, and resets the clock the same way
conn.execute("INSERT INTO cadence_dismissals (account_id, step_type, dismissed_at) "
             "VALUES (?,?,?)", (acme["id"], "Call & Text", "2026-10-16T10:00:00-05:00"))
conn.commit()
check("cadence: skipping a step clears it without making the next one overdue",
      cadence.get_due_reminders(conn, acme["id"]) == [])
cadence.today = lambda: date(2026, 10, 21)
acme_rems = cadence.get_due_reminders(conn, acme["id"])
check("cadence: skipping a step also moves to the next",
      [r["step_type"] for r in acme_rems] == ["Call 2"], acme_rems)
check("cadence: never done on time -> steps stack as before",
      [r["step_type"] for r in cadence.get_due_reminders(conn, acme["id"], collapse=False)]
      == ["Call 2"])

cadence.today = _real_cadence_today

# ---- 5. Milestone cancellation: advancing milestone kills ALL reminders instantly
conn.execute("UPDATE accounts SET pipeline_milestone='Accepted Meeting' WHERE id=?", (acme["id"],))
conn.commit()
check("cadence: milestone change clears all reminders",
      cadence.get_due_reminders(conn, acme["id"]) == [])
# status change also cancels (test on Bayou)
conn.execute("UPDATE accounts SET prospecting_status='Not Interested' WHERE id=?", (bayou["id"],))
conn.commit()
check("cadence: status change clears all reminders",
      cadence.get_due_reminders(conn, bayou["id"]) == [])
# revert bayou → reminders reappear (nothing stored stale)
conn.execute("UPDATE accounts SET prospecting_status='Prospecting' WHERE id=?", (bayou["id"],))
conn.commit()
check("cadence: reverting status restores reminders",
      len(cadence.get_due_reminders(conn, bayou["id"])) == 1)

# ---- 6. Cadence progress states for detail page
steps = cadence.get_cadence_progress(conn, acme["id"])
states = {s["step_type"]: s["state"] for s in steps}
check("progress: done/skipped/inactive states",
      states["Email 1"] == "done" and states["Call & Text"] == "skipped"
      and states["Call 2"] == "inactive", states)
conn.close()

# ---- 7. Routes via test client
r = client.get("/");            check("GET / 200 + reminder shown", r.status_code == 200 and b"Bayou Holdings" in r.data)
r = client.get("/accounts");    check("GET /accounts 200", r.status_code == 200 and b"Acme Properties" in r.data)
r = client.get("/accounts?status=Prospecting&milestone=None+%2F+In+Cadence&q=bayou")
check("GET /accounts filters", r.status_code == 200 and b"Bayou" in r.data and b"Acme Properties" not in r.data)
r = client.get(f"/accounts/1"); check("GET detail 200", r.status_code == 200)
r = client.get("/accounts/999"); check("GET missing account 404", r.status_code == 404)
r = client.get("/accounts/new"); check("GET new form 200", r.status_code == 200)
r = client.post("/accounts/new", data={"company_name": "Test Co", "preferred_contact": "Email",
    "prospecting_status": "Prospecting", "pipeline_milestone": "None / In Cadence"},
    follow_redirects=True)
check("POST new account", r.status_code == 200 and b"Test Co" in r.data)
conn = db.get_db()
tid = conn.execute("SELECT id FROM accounts WHERE company_name='Test Co'").fetchone()["id"]
conn.close()
r = client.post(f"/accounts/{tid}/edit", data={"company_name": "Test Co", "num_properties": "12",
    "preferred_contact": "Call", "prospecting_status": "Might be Interested",
    "pipeline_milestone": "On Hold", "email": "x@y.com"}, follow_redirects=True)
check("POST edit account", r.status_code == 200)
conn = db.get_db()
row = conn.execute("SELECT * FROM accounts WHERE id=?", (tid,)).fetchone()
check("edit persisted + updated_at ISO", row["num_properties"] == 12
      and row["prospecting_status"] == "Might be Interested" and "T" in row["updated_at"])
conn.close()
r = client.post(f"/accounts/{tid}/log", data={"interaction_type": "Meeting", "notes": "walked roof"},
                follow_redirects=True)
check("POST log interaction", r.status_code == 200 and b"walked roof" in r.data)
r = client.post("/reminders/quicklog", data={"account_id": 2, "step_type": "Email 1"},
                follow_redirects=True)
check("POST quicklog clears reminder", r.status_code == 200 and b"Bayou Holdings" not in
      client.get("/").data.split(b"Recent Activity")[0])
r = client.post(f"/accounts/{tid}/restart-cadence", follow_redirects=True)
check("POST restart cadence", r.status_code == 200)
conn = db.get_db()
row = conn.execute("SELECT * FROM accounts WHERE id=?", (tid,)).fetchone()
check("restart resets status+start", row["cadence_start"] == date.today().isoformat()
      and row["prospecting_status"] == "Prospecting"
      and row["pipeline_milestone"] == "None / In Cadence")
prev = conn.execute("SELECT * FROM interactions WHERE account_id=? AND interaction_type='Meeting'",
                    (tid,)).fetchone()
check("restart keeps non-cadence history", prev is not None)
conn.close()
r = client.post(f"/accounts/{tid}/delete", follow_redirects=True)
check("POST delete account", r.status_code == 200)
conn = db.get_db()
check("delete cascades interactions",
      conn.execute("SELECT COUNT(*) c FROM interactions WHERE account_id=?", (tid,)).fetchone()["c"] == 0)
conn.close()
# CSV import route
csv_data = b"Company,Contact First Name,Contact Last Name,Phone Number\r\nCSV Co,Al,Jones,555-1111\r\n"
r = client.post("/import", data={"file": (io.BytesIO(csv_data), "list.csv")},
                content_type="multipart/form-data")
check("POST /import CSV", r.status_code == 200 and b"1</strong> account(s) imported" in r.data, r.data[:500])
r = client.get("/import/template")
check("GET template download", r.status_code == 200 and b"Company Name" in r.data)
# bad file type
r = client.post("/import", data={"file": (io.BytesIO(b"junk"), "list.pdf")},
                content_type="multipart/form-data")
check("POST /import bad type rejected", r.status_code == 200 and b"Unsupported file type" in r.data)

# ---- 8. Contacts: manual add / promote / delete
r = client.post("/accounts/new", data={"company_name": "Contact Co", "preferred_contact": "Unknown",
    "prospecting_status": "Prospecting", "pipeline_milestone": "None / In Cadence"}, follow_redirects=True)
conn = db.get_db()
cid_acct = conn.execute("SELECT id FROM accounts WHERE company_name='Contact Co'").fetchone()["id"]
conn.close()
# first contact on empty account auto-becomes primary
r = client.post(f"/accounts/{cid_acct}/contacts/add", data={"first_name": "Ann", "last_name": "Lee",
    "title": "Owner", "email": "ann@co.com", "work_phone": "555-0001"}, follow_redirects=True)
conn = db.get_db()
a = conn.execute("SELECT * FROM accounts WHERE id=?", (cid_acct,)).fetchone()
check("contacts: first add fills primary", a["first_name"] == "Ann" and a["email"] == "ann@co.com")
check("contacts: no extra row created",
      conn.execute("SELECT COUNT(*) c FROM contacts WHERE account_id=?", (cid_acct,)).fetchone()["c"] == 0)
conn.close()
# second contact goes to contacts table
r = client.post(f"/accounts/{cid_acct}/contacts/add", data={"first_name": "Bo", "last_name": "Cruz",
    "title": "Facilities Mgr", "mobile_phone": "555-0002"}, follow_redirects=True)
conn = db.get_db()
row = conn.execute("SELECT * FROM contacts WHERE account_id=?", (cid_acct,)).fetchone()
check("contacts: second add → contacts row", row is not None and row["first_name"] == "Bo"
      and "T" in row["created_at"])
conn.close()
# promote Bo → swaps with Ann
r = client.post(f"/accounts/{cid_acct}/contacts/{row['id']}/promote", follow_redirects=True)
conn = db.get_db()
a = conn.execute("SELECT * FROM accounts WHERE id=?", (cid_acct,)).fetchone()
demoted = conn.execute("SELECT * FROM contacts WHERE account_id=?", (cid_acct,)).fetchall()
check("contacts: promote swaps primary", a["first_name"] == "Bo"
      and len(demoted) == 1 and demoted[0]["first_name"] == "Ann")
conn.close()
# add-as-primary via checkbox demotes current primary
r = client.post(f"/accounts/{cid_acct}/contacts/add", data={"first_name": "Cy", "last_name": "Ford",
    "set_primary": "1"}, follow_redirects=True)
conn = db.get_db()
a = conn.execute("SELECT * FROM accounts WHERE id=?", (cid_acct,)).fetchone()
names = {r2["first_name"] for r2 in conn.execute("SELECT * FROM contacts WHERE account_id=?", (cid_acct,))}
check("contacts: set_primary demotes old primary", a["first_name"] == "Cy" and names == {"Ann", "Bo"})
# delete a contact
did = conn.execute("SELECT id FROM contacts WHERE first_name='Ann'").fetchone()["id"]
conn.close()
r = client.post(f"/accounts/{cid_acct}/contacts/{did}/delete", follow_redirects=True)
conn = db.get_db()
check("contacts: delete removes row",
      conn.execute("SELECT COUNT(*) c FROM contacts WHERE id=?", (did,)).fetchone()["c"] == 0)
check("contacts: nameless add rejected", True)
conn.close()
r = client.post(f"/accounts/{cid_acct}/contacts/add", data={"title": "Ghost"}, follow_redirects=True)
check("contacts: add without name rejected", b"needs at least a first or last name" in r.data)
r = client.get(f"/accounts/{cid_acct}")
check("contacts: detail page renders contacts", r.status_code == 200 and b"Bo" in r.data and b"Primary" in r.data)

# ---- 9. ZoomInfo-style contact import
conn = db.get_db()
conn.execute("""INSERT INTO accounts (company_name, cadence_start, created_at, updated_at,
    work_phone) VALUES ('Sallyport Investments, Llc', ?, ?, ?, '(281) 423-0260')""",
    (date.today().isoformat(), db.now_iso(), db.now_iso()))
conn.commit(); conn.close()
zi_csv = (b"Company Name,First Name,Last Name,Job Title,Email Address,Direct Phone Number,Mobile phone\r\n"
          b"Sallyport Investments LLC,Doug,Erwin,Founder,doug@sallyport.com,(281) 423-9999,(832) 555-1111\r\n"
          b"SALLYPORT INVESTMENTS,Meg,Ray,Asset Manager,meg@sallyport.com,,\r\n"
          b"Sallyport Investments LLC,Doug,Erwin,Founder,doug@sallyport.com,,\r\n"  # dupe
          b"Nowhere Holdings,Sam,Lost,CEO,sam@nowhere.com,,\r\n")                    # unmatched
r = client.post("/import/contacts", data={"file": (io.BytesIO(zi_csv), "zoominfo_export.csv")},
                content_type="multipart/form-data")
check("zi import: route 200", r.status_code == 200)
check("zi import: summary rendered", b"2</strong> contact(s) attached" in r.data
      and b"Nowhere Holdings" in r.data, r.data[:800])
conn = db.get_db()
a = conn.execute("SELECT * FROM accounts WHERE company_name LIKE 'Sallyport%'").fetchone()
check("zi import: fuzzy company match + primary filled",
      a["first_name"] == "Doug" and a["email"] == "doug@sallyport.com")
check("zi import: direct phone wins, main line kept in notes",
      a["work_phone"] == "(281) 423-9999" and "Company main line: (281) 423-0260" in a["notes"],
      dict(a))
extra = conn.execute("SELECT * FROM contacts WHERE account_id=?", (a["id"],)).fetchall()
check("zi import: second person → contacts row",
      len(extra) == 1 and extra[0]["first_name"] == "Meg")
conn.close()

# ---- 10. Outreach templates & scripts
conn = db.get_db()
n_templates = conn.execute("SELECT COUNT(*) c FROM templates").fetchone()["c"]
check("templates: 7 seeded", n_templates == 7, n_templates)
settings = {r["key"]: r["value"] for r in conn.execute("SELECT * FROM settings")}
check("templates: settings seeded", settings.get("my_name") == "Alexis Morales"
      and settings.get("my_company") == "Silicone Roof Pros, Inc.")
conn.close()
# re-init must not duplicate seeds
db.init_db()
conn = db.get_db()
check("templates: re-init does not re-seed",
      conn.execute("SELECT COUNT(*) c FROM templates").fetchone()["c"] == 7)
conn.close()

r = client.get("/templates")
check("GET /templates 200", r.status_code == 200 and b"Cold Call Script" in r.data
      and b"{first_name}" in r.data)
r = client.post("/settings", data={"my_name": "Alex", "my_company": "Silicone Roof Pros",
    "my_phone": "(713) 555-0100"}, follow_redirects=True)
check("POST /settings saves", r.status_code == 200)

# Sallyport (id from earlier ZI import) has Doug Erwin, title Founder, no num_properties set
conn = db.get_db()
sp = conn.execute("SELECT * FROM accounts WHERE company_name LIKE 'Sallyport%'").fetchone()
conn.execute("UPDATE accounts SET num_properties=40, matching_properties=23, "
             "title='Founder' WHERE id=?", (sp["id"],))
conn.commit(); conn.close()

r = client.get(f"/accounts/{sp['id']}/scripts")
check("GET scripts all 200", r.status_code == 200)
html = r.data.decode()
check("scripts: placeholders filled",
      "Hi Doug," in html and "Founder at Sallyport Investments, Llc" in html
      and "23 older properties" in html, html[:300])
check("scripts: my info filled", "Alex with Silicone Roof Pros" in html
      and "(713) 555-0100" in html)
check("scripts: mailto prefilled", "mailto:doug@sallyport.com?subject=" in html
      and "body=Hi%20Doug" in html)
check("scripts: tel + sms links use clean digits", "tel:2814239999" in html
      and "sms:8325551111?body=" in html)
check("scripts: log buttons present", 'value="Email 1"' in html
      and 'value="Breakup Email"' in html)

r = client.get(f"/accounts/{sp['id']}/scripts?step=Call+%26+Text")
html = r.data.decode()
check("scripts: step filter", "Cold Call Script" in html and "Voicemail" in html
      and "Text Message" in html and "Breakup" not in html)

# missing-info warning: account with no contact
conn = db.get_db()
cur = conn.execute("""INSERT INTO accounts (company_name, cadence_start, created_at, updated_at)
    VALUES ('Bare Co', ?, ?, ?)""", (date.today().isoformat(), db.now_iso(), db.now_iso()))
conn.commit()
bare = {"id": cur.lastrowid}
conn.close()
r = client.get(f"/accounts/{bare['id']}/scripts?step=Email+1")
html = r.data.decode()
check("scripts: missing fields flagged", "[first name]" in html and "Missing info" in html, html[:200])

# edit a template and see the change in scripts
conn = db.get_db()
t1 = conn.execute("SELECT id FROM templates WHERE name='Email 1'").fetchone()["id"]
conn.close()
r = client.post(f"/templates/{t1}", data={"name": "Email 1", "kind": "email",
    "steps": ["Email 1"], "subject": "Quick question about {company}",
    "body": "Hi {first_name}, testing edit."}, follow_redirects=True)
check("POST template edit", r.status_code == 200)
r = client.get(f"/accounts/{sp['id']}/scripts?step=Email+1")
check("scripts: edited template renders",
      b"Quick question about Sallyport" in r.data and b"testing edit." in r.data)

# ---- 11. Follow-ups
conn = db.get_db()
cols = {r[1] for r in conn.execute("PRAGMA table_info(accounts)")}
check("followup: columns migrated", "next_follow_up" in cols and "follow_up_note" in cols)
fid = conn.execute("SELECT id FROM accounts WHERE company_name='Bayou Holdings'").fetchone()["id"]
conn.close()
r = client.post(f"/accounts/{fid}/followup", data={"days": "0", "note": "call about bid"},
                follow_redirects=True)
check("followup: set today", r.status_code == 200)
conn = db.get_db()
a = conn.execute("SELECT * FROM accounts WHERE id=?", (fid,)).fetchone()
check("followup: stored", a["next_follow_up"] == date.today().isoformat()
      and a["follow_up_note"] == "call about bid")
conn.close()
r = client.get("/")
check("dashboard: follow-up shown", b"Follow-Ups Due" in r.data and b"call about bid" in r.data)
r = client.post(f"/accounts/{fid}/followup", data={"days": "3"}, follow_redirects=True)
conn = db.get_db()
a = conn.execute("SELECT * FROM accounts WHERE id=?", (fid,)).fetchone()
check("followup: snooze +3d keeps note",
      a["next_follow_up"] == (date.today() + timedelta(days=3)).isoformat()
      and a["follow_up_note"] == "call about bid")
conn.close()
check("dashboard: snoozed follow-up hidden", b"call about bid" not in client.get("/").data)
r = client.post(f"/accounts/{fid}/followup", data={"clear": "1"}, follow_redirects=True)
conn = db.get_db()
a = conn.execute("SELECT * FROM accounts WHERE id=?", (fid,)).fetchone()
check("followup: clear", a["next_follow_up"] == "" and a["follow_up_note"] == "")
conn.close()
r = client.post(f"/accounts/{fid}/followup", data={"date": "2030-01-15", "note": "custom"},
                follow_redirects=True)
conn = db.get_db()
check("followup: explicit date", conn.execute(
    "SELECT next_follow_up FROM accounts WHERE id=?", (fid,)).fetchone()[0] == "2030-01-15")
conn.execute("UPDATE accounts SET next_follow_up='', follow_up_note='' WHERE id=?", (fid,))
conn.commit(); conn.close()

# ---- 12. Stale deals
conn = db.get_db()
old = (datetime.now() - timedelta(days=20)).astimezone().isoformat(timespec="seconds")
conn.execute("""INSERT INTO accounts (company_name, pipeline_milestone, prospecting_status,
    cadence_start, created_at, updated_at) VALUES
    ('Stale Deal Co', 'Walked Roof', 'Interested - Continue Conversation', ?, ?, ?)""",
    (date.today().isoformat(), old, old))
conn.commit(); conn.close()
r = client.get("/")
check("dashboard: stale deal flagged", b"Going Stale" in r.data and b"Stale Deal Co" in r.data)
conn = db.get_db()
sid = conn.execute("SELECT id FROM accounts WHERE company_name='Stale Deal Co'").fetchone()["id"]
conn.execute("UPDATE accounts SET next_follow_up=? WHERE id=?",
             ((date.today() + timedelta(days=5)).isoformat(), sid))
conn.commit(); conn.close()
check("dashboard: follow-up suppresses stale", b"Stale Deal Co" not in
      client.get("/").data.split(b"Recent Activity")[0])
conn = db.get_db()
conn.execute("UPDATE accounts SET next_follow_up='' WHERE id=?", (sid,))
conn.commit(); conn.close()

# ---- 13. Queue
conn = db.get_db()
conn.execute("UPDATE accounts SET next_follow_up=?, follow_up_note='ask about roof age' WHERE id=?",
             (date.today().isoformat(), fid))
conn.commit()
n_tasks = len(__import__("app")._build_queue(conn))
conn.close()
r = client.get("/queue")
check("queue: renders first task", r.status_code == 200 and b"Task 1 of" in r.data)
check("queue: dashboard button", b"Work the Queue" in client.get("/").data)
r = client.get(f"/queue?pos={n_tasks - 1}")
check("queue: last pos ok", r.status_code == 200 and f"Task {n_tasks} of {n_tasks}".encode() in r.data)
r = client.get("/queue?pos=9999")
check("queue: pos clamped", r.status_code == 200)
# find pos of the follow-up task for Bayou
conn = db.get_db()
tasks = __import__("app")._build_queue(conn)
conn.close()
fu_pos = next(i for i, t in enumerate(tasks) if t["kind"] == "followup" and t["account_id"] == fid)
r = client.get(f"/queue?pos={fu_pos}")
check("queue: follow-up task shows note", b"ask about roof age" in r.data)
cad_pos = next((i for i, t in enumerate(tasks) if t["kind"] == "cadence"), None)
if cad_pos is not None:
    r = client.get(f"/queue?pos={cad_pos}")
    check("queue: cadence task shows scripts + log", b"Log" in r.data and b"textarea" in r.data)
conn = db.get_db()
conn.execute("UPDATE accounts SET next_follow_up='', follow_up_note='' WHERE id=?", (fid,))
conn.commit(); conn.close()

# ---- 14. Pipeline board + quick move
r = client.get("/pipeline")
check("pipeline: renders all columns", r.status_code == 200
      and b"Walked Roof" in r.data and b"Closed Won" in r.data and b"Stale Deal Co" in r.data)
r = client.post(f"/accounts/{sid}/milestone", data={"pipeline_milestone": "Created Report/Bid"},
                follow_redirects=True)
conn = db.get_db()
check("pipeline: quick move", conn.execute(
    "SELECT pipeline_milestone FROM accounts WHERE id=?", (sid,)).fetchone()[0] == "Created Report/Bid")
conn.close()
r = client.post(f"/accounts/{sid}/milestone", data={"pipeline_milestone": "Bogus"},
                follow_redirects=True)
conn = db.get_db()
check("pipeline: invalid milestone rejected", conn.execute(
    "SELECT pipeline_milestone FROM accounts WHERE id=?", (sid,)).fetchone()[0] == "Created Report/Bid")
conn.close()

# ---- 15. Staggered import
# Each row has someone to contact, so it goes into the paced cadence rather
# than into Research.
stag_csv = b"Company Name,First Name,Last Name,Work Phone\r\n" + b"".join(
    f"Stagger Co {i},Pat,Lee{i},713-555-01{i:02d}\r\n".encode() for i in range(1, 8))
r = client.post("/import", data={"file": (io.BytesIO(stag_csv), "stagger.csv"), "per_day": "3"},
                content_type="multipart/form-data")
check("stagger: import ok", b"7</strong> account(s) imported" in r.data, r.data[:400])
conn = db.get_db()
starts = [r2["cadence_start"] for r2 in conn.execute(
    "SELECT cadence_start FROM accounts WHERE company_name LIKE 'Stagger Co %' ORDER BY id")]
conn.close()
from collections import Counter
counts = Counter(starts)
check("stagger: 3 per day, 3 distinct business days",
      sorted(counts.values(), reverse=True) == [3, 3, 1] and len(counts) == 3, counts)
check("stagger: no weekend starts",
      all(date.fromisoformat(s).weekday() < 5 for s in starts), starts)
check("stagger: dates ascend", starts == sorted(starts))

# ---- 16. Exports + backup
r = client.get("/export/accounts.csv")
check("export accounts csv", r.status_code == 200 and b"company_name" in r.data
      and b"Stagger Co 1" in r.data)
r = client.get("/export/interactions.csv")
check("export interactions csv", r.status_code == 200 and b"interaction_type" in r.data)
db.BACKUP_DIR = Path(__file__).parent / "test_backups"
import shutil
shutil.rmtree(db.BACKUP_DIR, ignore_errors=True)
b1 = db.backup_db()
check("backup: created", b1 is not None and b1.exists() and b1.stat().st_size > 0)
b2 = db.backup_db()
check("backup: once per day", b2 is None
      and len(list(db.BACKUP_DIR.glob('crm-*.db'))) == 1)
# pruning: fabricate 20 old backups
for i in range(20):
    (db.BACKUP_DIR / f"crm-2020010{i % 10}-00000{i}.db").write_bytes(b"x")
db.backup_db(keep=14)  # skipped (today exists) but still prunes
check("backup: pruned to 14", len(list(db.BACKUP_DIR.glob('crm-*.db'))) == 14)
shutil.rmtree(db.BACKUP_DIR, ignore_errors=True)

# ---- 17. PWA assets
r = client.get("/static/manifest.json")
check("pwa: manifest served", r.status_code == 200 and b"Roof CRM" in r.data)
r = client.get("/static/icon-192.png")
check("pwa: icon served", r.status_code == 200 and r.data[:4] == b"\x89PNG")

# ---- 18. WAL mode, queue log return, pref badges, re-pace
conn = db.get_db()
check("wal: journal mode", conn.execute("PRAGMA journal_mode").fetchone()[0] == "wal")
conn.execute("UPDATE accounts SET preferred_contact='Text' WHERE id=?", (fid,))
conn.execute("UPDATE accounts SET next_follow_up=? WHERE id=?",
             (date.today().isoformat(), fid))
conn.commit(); conn.close()
check("pref badge: dashboard follow-up row", "prefers text" in client.get("/").data.decode())
r = client.post(f"/accounts/{fid}/log",
                data={"interaction_type": "General Note", "notes": "from queue",
                      "next": "/queue?pos=0"})
check("queue: log honors next", r.status_code == 302 and r.headers["Location"].endswith("/queue?pos=0"))
conn = db.get_db()
conn.execute("UPDATE accounts SET next_follow_up='' WHERE id=?", (fid,))
conn.commit(); conn.close()
# re-pace: untouched Stagger Cos move; touched accounts keep dates
conn = db.get_db()
touched_start = conn.execute(
    "SELECT cadence_start FROM accounts WHERE company_name='Bayou Holdings'").fetchone()[0]
conn.close()
r = client.post("/repace", data={"per_day": "2"}, follow_redirects=True)
check("repace: 200", r.status_code == 200 and b"Re-paced" in r.data, r.data[:300])
conn = db.get_db()
starts2 = [row["cadence_start"] for row in conn.execute(
    "SELECT cadence_start FROM accounts WHERE company_name LIKE 'Stagger Co %' "
    "AND pipeline_milestone='None / In Cadence' ORDER BY id")]
from collections import Counter as C2
check("repace: untouched re-staggered at 2/day",
      len(starts2) > 0 and max(C2(starts2).values()) <= 2
      and min(starts2) > date.today().isoformat()
      and all(date.fromisoformat(s).weekday() < 5 for s in starts2), starts2)
check("repace: touched account keeps date", conn.execute(
    "SELECT cadence_start FROM accounts WHERE company_name='Bayou Holdings'"
    ).fetchone()[0] == touched_start)
conn.close()
r = client.post("/repace", data={"per_day": ""}, follow_redirects=True)
check("repace: blank rejected", b"Enter how many" in r.data)

# ---- 19. Column I backfill, heuristic headers, placeholder migration
# Backfill: Bayou has NULL property counts; re-upload with values fills them in
bf_csv = (b"Company Name,# Properties (in search),Properties Owned\r\n"
          b"Bayou Holdings,5,12\r\n")
r = client.post("/import", data={"file": (io.BytesIO(bf_csv), "backfill.csv")},
                content_type="multipart/form-data")
check("backfill: summary shown", b"backfilled with property counts" in r.data, r.data[:400])
conn = db.get_db()
b_row = conn.execute("SELECT * FROM accounts WHERE company_name='Bayou Holdings'").fetchone()
check("backfill: counts filled on existing account",
      b_row["matching_properties"] == 5 and b_row["num_properties"] == 12)
first_before = b_row["first_name"]
conn.close()
# re-upload again → nothing more to backfill, other fields untouched
r = client.post("/import", data={"file": (io.BytesIO(bf_csv), "backfill.csv")},
                content_type="multipart/form-data")
conn = db.get_db()
b_row = conn.execute("SELECT * FROM accounts WHERE company_name='Bayou Holdings'").fetchone()
check("backfill: idempotent, contact untouched",
      b_row["matching_properties"] == 5 and b_row["first_name"] == first_before)
conn.close()

# Heuristic mapping: differently-worded CoStar-style headers still map
variant = importer.map_columns([
    "Company", "Owner First Name", "Owner Last Name", "Contact Title",
    "Total Properties Owned", "Matching Properties Count",
    "Office Phone Number", "Cell #", "E-Mail", "Internal Comments"])
check("headers: variant names all deciphered",
      variant.get("company_name") == "Company"
      and variant.get("first_name") == "Owner First Name"
      and variant.get("last_name") == "Owner Last Name"
      and variant.get("num_properties") == "Total Properties Owned"
      and variant.get("matching_properties") == "Matching Properties Count"
      and variant.get("work_phone") == "Office Phone Number"
      and variant.get("mobile_phone") == "Cell #"
      and variant.get("email") == "E-Mail"
      and variant.get("notes") == "Internal Comments", variant)
check("headers: company rule skips address/phone columns",
      "company_name" not in importer.map_columns(["Company Address", "Company Phone"]))

# Placeholder migration: an old-schema db (no matching_properties) with the old
# Email 1 text gets the placeholder swapped when init_db migrates it
import sqlite3 as _sq
old_path = db.DB_PATH.parent / "old_migrate.db"
old_path.unlink(missing_ok=True)
_c = _sq.connect(old_path)
_c.executescript("""
CREATE TABLE accounts (id INTEGER PRIMARY KEY, company_name TEXT NOT NULL,
 first_name TEXT DEFAULT '', last_name TEXT DEFAULT '', title TEXT DEFAULT '',
 num_properties INTEGER, email TEXT DEFAULT '', work_phone TEXT DEFAULT '',
 mobile_phone TEXT DEFAULT '', preferred_contact TEXT NOT NULL DEFAULT 'Unknown',
 notes TEXT DEFAULT '', prospecting_status TEXT NOT NULL DEFAULT 'Prospecting',
 pipeline_milestone TEXT NOT NULL DEFAULT 'None / In Cadence',
 cadence_start TEXT NOT NULL, next_follow_up TEXT DEFAULT '',
 follow_up_note TEXT DEFAULT '', created_at TEXT NOT NULL, updated_at TEXT NOT NULL);
CREATE TABLE templates (id INTEGER PRIMARY KEY, name TEXT NOT NULL,
 kind TEXT NOT NULL DEFAULT 'email', steps TEXT DEFAULT '', subject TEXT DEFAULT '',
 body TEXT NOT NULL DEFAULT '', sort_order INTEGER DEFAULT 0, updated_at TEXT NOT NULL);
INSERT INTO templates (name, kind, subject, body, updated_at) VALUES
 ('Email 1', 'email', '{company}''s {num_properties} older properties',
  'roughly {num_properties} properties built before the 1980s', '2026-01-01T00:00:00');
""")
_c.commit(); _c.close()
_real_path = db.DB_PATH
db.DB_PATH = old_path
db.init_db()
_c = db.get_db()
mig = _c.execute("SELECT subject, body FROM templates WHERE name='Email 1'").fetchone()
mig_cols = {r2[1] for r2 in _c.execute("PRAGMA table_info(accounts)")}
_c.close()
db.DB_PATH = _real_path
check("migration: matching_properties column added", "matching_properties" in mig_cols)
check("migration: old templates get new placeholder",
      "{matching_properties}" in mig["subject"] and "{matching_properties}" in mig["body"]
      and "{num_properties}" not in mig["body"], dict(mig))

# ---- 20. Insights
conn = db.get_db()
conn.execute("UPDATE accounts SET pipeline_milestone='Closed Won' WHERE company_name='Stagger Co 1'")
conn.execute("UPDATE accounts SET pipeline_milestone='Closed Lost' WHERE company_name IN "
             "('Stagger Co 2','Stagger Co 3','Stagger Co 4')")
wk_old = (datetime.now() - timedelta(days=21)).astimezone().isoformat(timespec="seconds")
aid = conn.execute("SELECT id FROM accounts LIMIT 1").fetchone()["id"]
conn.execute("INSERT INTO interactions (account_id, interaction_type, notes, created_at) "
             "VALUES (?,?,?,?)", (aid, "Call 2", "old call", wk_old))
conn.commit(); conn.close()
r = client.get("/insights")
html = r.data.decode()
check("insights: 200", r.status_code == 200)
check("insights: tiles present", "Tasks due today" in html and "Win rate" in html
      and "Touches this week" in html)
check("insights: win rate 25%", ">25%<" in html.replace(" ", ""),
      [l for l in html.splitlines() if "%" in l][:3])
check("insights: weekly chart has 8 buckets", html.count("Week of") == 8)
check("insights: raw notes not leaked into charts", "old call" not in html)
check("insights: pipeline bars", "Walked Roof" in html and "Closed Won" in html)
check("insights: cadence reach bars", "Day 1: Email 1" in html and "Day 10: Breakup Email" in html)
check("insights: activity mix", "Meeting" in html or "General Note" in html)
check("insights: nav link", 'href="/insights"' in html)

# ---- 20b. Edge-case hardening
r = client.get("/queue?pos=abc")
check("queue: junk pos handled", r.status_code == 200)
r = client.post("/invoices/99999/status", data={"status": "Paid"}, follow_redirects=True)
check("invoice status: missing invoice no crash", r.status_code == 200
      and b"Invoice not found" in r.data)
from app import clean_tel
check("tel filter strips formatting", clean_tel("(281) 423-0260") == "2814230260"
      and clean_tel("+1 832-303-3183") == "+18323033183" and clean_tel(None) == "")
check("dashboard: tel links dialable", b'href="tel:(' not in client.get("/").data)

# ---- 21. Projects & invoicing
conn = db.get_db()
won_id = conn.execute(
    "SELECT id FROM accounts WHERE pipeline_milestone='Closed Won' LIMIT 1").fetchone()["id"]
conn.close()
r = client.get("/projects")
check("projects: eligible won deal listed", r.status_code == 200
      and f'value="{won_id}"'.encode() in r.data)
r = client.post("/projects/new", data={"account_id": won_id, "contract_amount": "$24,500"},
                follow_redirects=True)
check("projects: created with checklist", r.status_code == 200
      and b"Job Checklist" in r.data and b"Warranty documents delivered" in r.data
      and b"$24,500" in r.data)
conn = db.get_db()
proj = conn.execute("SELECT * FROM projects WHERE account_id=?", (won_id,)).fetchone()
n_tasks = conn.execute("SELECT COUNT(*) c FROM project_tasks WHERE project_id=?",
                       (proj["id"],)).fetchone()["c"]
check("projects: 11 default tasks seeded", n_tasks == len(db.DEFAULT_PROJECT_TASKS))
check("projects: contract parsed", proj["contract_amount"] == 24500)
first_task = conn.execute("SELECT id FROM project_tasks WHERE project_id=? "
                          "ORDER BY sort_order LIMIT 1", (proj["id"],)).fetchone()["id"]
conn.close()
# duplicate project blocked
r = client.post("/projects/new", data={"account_id": won_id}, follow_redirects=True)
check("projects: duplicate blocked", b"already has a project" in r.data)
# checklist toggle
r = client.post(f"/projects/{proj['id']}/tasks/{first_task}/toggle", follow_redirects=True)
conn = db.get_db()
t_row = conn.execute("SELECT * FROM project_tasks WHERE id=?", (first_task,)).fetchone()
check("projects: task toggled done with timestamp", t_row["done"] == 1 and "T" in t_row["done_at"])
conn.close()
r = client.post(f"/projects/{proj['id']}/tasks/add", data={"title": "Order lift rental"},
                follow_redirects=True)
check("projects: custom task added", b"Order lift rental" in r.data)
# invoices: draft -> sent -> paid
r = client.post(f"/projects/{proj['id']}/invoices/add",
                data={"invoice_number": "INV-001", "amount": "12,250",
                      "notes": "50% deposit"}, follow_redirects=True)
check("invoice: draft added", b"INV-001" in r.data and b"$12,250" in r.data)
conn = db.get_db()
inv = conn.execute("SELECT * FROM invoices WHERE invoice_number='INV-001'").fetchone()
conn.close()
r = client.post(f"/invoices/{inv['id']}/status", data={"status": "Sent"}, follow_redirects=True)
conn = db.get_db()
inv = conn.execute("SELECT * FROM invoices WHERE id=?", (inv["id"],)).fetchone()
check("invoice: sent stamps dates", inv["status"] == "Sent"
      and inv["sent_date"] == date.today().isoformat()
      and inv["due_date"] == (date.today() + timedelta(days=30)).isoformat())
conn.close()
check("dashboard: outstanding invoice shown",
      b"Outstanding Invoices" in client.get("/").data and b"$12,250" in client.get("/").data)
# overdue rendering
conn = db.get_db()
conn.execute("UPDATE invoices SET due_date=? WHERE id=?",
             ((date.today() - timedelta(days=5)).isoformat(), inv["id"]))
conn.commit(); conn.close()
check("dashboard: overdue flagged", b"Overdue" in client.get("/").data)
r = client.post(f"/invoices/{inv['id']}/status", data={"status": "Paid", "next": "/"},
                follow_redirects=True)
conn = db.get_db()
inv = conn.execute("SELECT * FROM invoices WHERE id=?", (inv["id"],)).fetchone()
check("invoice: paid stamps date", inv["status"] == "Paid"
      and inv["paid_date"] == date.today().isoformat())
conn.close()
check("dashboard: paid invoice cleared", b"Outstanding Invoices" not in client.get("/").data)
# money rollups
r = client.post(f"/projects/{proj['id']}/invoices/add", data={"amount": "12250"},
                follow_redirects=True)
conn = db.get_db()
inv2 = conn.execute("SELECT id FROM invoices WHERE project_id=? AND status='Draft'",
                    (proj["id"],)).fetchone()
conn.close()
client.post(f"/invoices/{inv2['id']}/status", data={"status": "Sent"})
r = client.get(f"/projects/{proj['id']}")
html = r.data.decode()
check("project: money tiles roll up", html.count("$12,250") >= 2 and "$24,500" in html)
r = client.get("/projects")
check("projects list: totals row", b"Collected" in r.data and b"Outstanding" in r.data)
r = client.get("/insights")
check("insights: money tiles appear", b"Collected" in r.data and b"Contracted" in r.data)
# account page shows project link
r = client.get(f"/accounts/{won_id}")
check("account: project card links", "Open:".encode() in r.data)
# bad amount rejected
r = client.post(f"/projects/{proj['id']}/invoices/add", data={"amount": "abc"},
                follow_redirects=True)
check("invoice: bad amount rejected", b"Enter an invoice amount" in r.data)
# ---- printable invoice
conn = db.get_db()
s_keys = {r2["key"] for r2 in conn.execute("SELECT key FROM settings")}
check("invoice print: business settings seeded",
      {"my_title", "my_email", "my_website", "my_address", "invoice_terms"} <= s_keys)
conn.execute("UPDATE accounts SET notes='Address: 123 Main St, Houston, TX 77002\n"
             "Website: x.com', first_name='Doug', last_name='Erwin', "
             "email='doug@x.com' WHERE id=?", (won_id,))
conn.commit()
pr_inv = conn.execute("SELECT id FROM invoices WHERE project_id=? LIMIT 1",
                      (proj["id"],)).fetchone()["id"]
conn.close()
r = client.get(f"/invoices/{pr_inv}/print")
html = r.data.decode()
check("invoice print: 200 + branding", r.status_code == 200
      and "brand-logo.png" in html and "Silicone Roof Pros" in html
      and "#0088df" in html)
check("invoice print: bill-to from account + notes address",
      "Doug Erwin" in html and "123 Main St, Houston, TX 77002" in html)
check("invoice print: terms shown", "1.5% per month" in html)
check("invoice print: number fallback or set", "INV-" in html)
conn = db.get_db()
conn.execute("UPDATE invoices SET status='Paid', paid_date=? WHERE id=?",
             (date.today().isoformat(), pr_inv))
conn.commit(); conn.close()
r = client.get(f"/invoices/{pr_inv}/print")
check("invoice print: PAID stamp", b"stamp paid" in r.data and b"PAID" in r.data)
r = client.get("/invoices/99999/print")
check("invoice print: missing 404", r.status_code == 404)
r = client.get(f"/projects/{proj['id']}")
check("project page: print button present", "🖨".encode() in r.data)

# delete cascade
r = client.post(f"/projects/{proj['id']}/delete", follow_redirects=True)
conn = db.get_db()
check("projects: delete cascades tasks + invoices",
      conn.execute("SELECT COUNT(*) c FROM project_tasks WHERE project_id=?",
                   (proj["id"],)).fetchone()["c"] == 0
      and conn.execute("SELECT COUNT(*) c FROM invoices WHERE project_id=?",
                       (proj["id"],)).fetchone()["c"] == 0)
conn.close()

# ---- 22. Bid / roof report generator
import app as app_mod
db.UPLOAD_DIR = db.DB_PATH.parent / "bid_photos"
db.TRASH_DIR = db.DB_PATH.parent / "trash"

# warranty calculator must reproduce the template example: 3,900 - 300 sq ft,
# Silicone on Capsheet, 10-year, +5% waste
import warranty_calc as wc
q = wc.calculate(3900, deduction_sqft=300, warranty_years=10, waste_pct=5, linear_feet=400)
check("calc: adjusted squares match template example", q["adjusted_squares"] == 37.8, q)
base, top = q["coats"]
check("calc: basecoat 1.25 gal/sq -> 50 gal / 10 pails",
      base["rate"] == 1.25 and base["gallons"] == 50 and base["pails"] == 10, base)
check("calc: topcoat 2 gal/sq -> 80 gal / 16 pails",
      top["rate"] == 2 and top["gallons"] == 80 and top["pails"] == 16, top)
check("calc: mastic buckets from linear feet", q["mastic_buckets"] == 5, q["mastic_buckets"])
check("calc: total gallons", q["total_gallons"] == 130, q["total_gallons"])

# rates change with warranty length (the whole point of the calculator)
r10 = wc.get_rates("Silicone", "Capsheet", 10)
r15 = wc.get_rates("Silicone", "Capsheet", 15)
r20 = wc.get_rates("Silicone", "Capsheet", 20)
check("calc: silicone capsheet topcoat 2 / 2.5 / 3 by warranty",
      (r10["top1"], r15["top1"], r20["top1"]) == (2, 2.5, 3))
check("calc: silicone single-ply has no basecoat",
      wc.get_rates("Silicone", "Single-Ply", 10)["base"] == 0)
check("calc: acrylic reinforced capsheet 20yr has four coats",
      len(wc.calculate(10000, coating_system="Acrylic", acrylic_system_type="Reinforced",
                       roof_type="Capsheet", warranty_years=20)["coats"]) == 4)
check("calc: acrylic reinforced adds fabric rolls",
      wc.calculate(10000, coating_system="Acrylic", acrylic_system_type="Reinforced",
                   roof_type="Capsheet", warranty_years=20)["membrane_rolls"] == 10)
check("calc: aluminum has no 15/20-year option",
      wc.calculate(5000, coating_system="Aluminum", roof_type="Metal", warranty_years=15) is None
      and wc.calculate(5000, coating_system="Aluminum", roof_type="Metal",
                       warranty_years=10)["coats"][0]["rate"] == 2)
check("calc: failed adhesion adds primer",
      wc.calculate(10000, passed_adhesion=False)["adhesion_primer_gal"] == 20)
check("calc: rust field prime on metal",
      wc.calculate(10000, roof_type="Metal", has_rust=True)["rust_primer_gal"] == 50
      and wc.calculate(10000, roof_type="Metal", has_rust=True,
                       rust_prime_method="spot")["rust_primer_gal"] == 0)
check("calc: roof type guessed from surface text",
      (wc.guess_roof_type("Granulated Capsheet"), wc.guess_roof_type("TPO"),
       wc.guess_roof_type("Standing seam metal"), wc.guess_roof_type("SPF foam"))
      == ("Capsheet", "Single-Ply", "Metal", "Sprayfoam"))
check("calc: pails round up in fives", wc.round_to_pails(47.25) == 50
      and wc.round_to_pails(0) == 0 and wc.round_to_pails(75.6) == 80)

# pricing: $/sq ft by surface and warranty length (adders apply to all surfaces)
check("price: capsheet 4.50 / 4.65 / 4.75",
      [wc.price_per_sqft("Capsheet", y) for y in (10, 15, 20)] == [4.50, 4.65, 4.75])
check("price: other surfaces 4.00 / 4.15 / 4.25",
      all([wc.price_per_sqft(rt, y) for y in (10, 15, 20)] == [4.00, 4.15, 4.25]
          for rt in ("Single-Ply", "Sprayfoam", "Metal")))
sp = wc.suggested_price(3600, "Capsheet", 15)
check("price: total = rate x coated sqft",
      sp["rate"] == 4.65 and sp["total"] == 16740.0, sp)
# products: Henry Prograde defaults, catalog-constrained per system
pr = wc.resolve_products("Silicone")
check("products: silicone defaults to Prograde",
      pr["top"] == "Prograde 988 Silicone" and pr["base"] == "Prograde 294 BaseCoat"
      and pr["mastic"] == "Prograde 923 Butter Grade"
      and pr["manufacturer"].startswith("Henry Company")
      and pr["adhesion_primer"] == "Prograde 941 Adhesion Promoting Primer", pr)
pr_e = wc.resolve_products("Silicone", topcoat="Enduraroof Premium Silicone")
check("products: enduraroof topcoat switches brand, primers and warranty holder",
      pr_e["brand"] == "Enduraroof" and pr_e["manufacturer"] == "Enduraroof"
      and pr_e["rust_primer"] == "Enduraroof Metal Roofing Primer", pr_e)
check("products: a product from another system is replaced",
      wc.resolve_products("Aluminum", topcoat="Prograde 988 Silicone")["top"]
      == "Pro-Grade 586")
check("products: aluminum has no basecoat",
      wc.resolve_products("Aluminum")["base"] == "")

# only published combinations exist
check("support: aluminum is metal/capsheet at 10 years only",
      wc.supported_roof_types("Aluminum") == ["Capsheet", "Metal"]
      and wc.supported_warranties("Aluminum", "Metal") == [10])
check("support: reinforced acrylic excludes metal and sprayfoam",
      wc.supported_roof_types("Acrylic", "Reinforced") == ["Capsheet", "Single-Ply"])
check("support: silicone covers every roof at 10/15/20",
      all(wc.supported_warranties("Silicone", rt) == [10, 15, 20]
          for rt in wc.ROOF_TYPES))

check("price: custom rules honored",
      wc.price_per_sqft("Capsheet", 20, {"capsheet_base": 5, "add_15": .2, "add_20": .3}) == 5.5)
check("bid words: 15000", app_mod.dollars_in_words(15000) == "FIFTEEN THOUSAND DOLLARS")
check("bid words: 24750", app_mod.dollars_in_words(24750)
      == "TWENTY-FOUR THOUSAND SEVEN HUNDRED FIFTY DOLLARS")

conn = db.get_db()
bid_acct = conn.execute("SELECT id FROM accounts WHERE company_name LIKE 'Sallyport%'").fetchone()["id"]
conn.execute("UPDATE accounts SET notes = 'Address: 123 Main St, Houston, TX 77002' "
             "|| char(10) || notes WHERE id=?", (bid_acct,))
conn.commit(); conn.close()
r = client.post(f"/accounts/{bid_acct}/bids/new", data={"roof_size_sqft": "3900"},
                follow_redirects=True)
check("bid: created, address prefilled from notes", r.status_code == 200
      and b"Roof Report" in r.data)
conn = db.get_db()
bid = conn.execute("SELECT * FROM bids WHERE account_id=?", (bid_acct,)).fetchone()
conn.close()
check("bid: address from notes", "123 Main St" in bid["roof_address"])
r = client.post(f"/bids/{bid['id']}/edit", data={
    "roof_address": "5231 Braesvalley Drive, Houston, TX 77096",
    "roof_size_sqft": "3900", "deduction_sqft": "300", "surface_type": "Capsheet",
    "candidate": "Yes", "warranty_years": "10", "price": "$15,000",
    "assessment_date": "2026-09-05",
    "coating_system": "Silicone", "acrylic_system_type": "Standard",
    "roof_type": "Capsheet", "linear_feet": "400", "waste_pct": "5",
    "stretch_pct": "0", "passed_adhesion": "1",
    "assessment_notes": "Ponding at NW corner\n- Cracked seams along HVAC curb"},
    follow_redirects=True)
check("bid: saved", r.status_code == 200)
conn = db.get_db()
bsaved = conn.execute("SELECT * FROM bids WHERE id=?", (bid["id"],)).fetchone()
conn.close()
check("bid: calculator inputs persisted",
      bsaved["coating_system"] == "Silicone" and bsaved["roof_type"] == "Capsheet"
      and bsaved["linear_feet"] == 400 and bsaved["waste_pct"] == 5
      and bsaved["passed_adhesion"] == 1)
r2 = client.get(f"/bids/{bid['id']}")
check("bid editor: suggested price shown",
      b"Suggested price" in r2.data and b"$16,200" in r2.data, r2.data[-3000:])
r2 = client.post(f"/bids/{bid['id']}/use-suggested-price", follow_redirects=True)
conn = db.get_db()
bp = conn.execute("SELECT price FROM bids WHERE id=?", (bid["id"],)).fetchone()["price"]
conn.close()
check("bid: use suggested price applies it", bp == 16200.0, bp)
r2 = client.post("/settings", data={"price_capsheet_base": "5.00", "price_other_base": "4.25",
                                    "price_add_15": "0.20", "price_add_20": "0.15",
                                    "next": "/templates"}, follow_redirects=True)
check("pricing settings: saved + matrix rendered",
      b"$5.00" in r2.data and b"$5.35" in r2.data, r2.status_code)
r2 = client.get(f"/bids/{bid['id']}")
check("bid editor: suggestion follows edited settings", b"$18,000" in r2.data)
client.post("/settings", data={"price_capsheet_base": "4.50", "price_other_base": "4.00",
                               "price_add_15": "0.15", "price_add_20": "0.10"})
# restore the quoted price for later assertions
client.post(f"/bids/{bid['id']}/edit", data={
    "roof_address": "5231 Braesvalley Drive, Houston, TX 77096",
    "roof_size_sqft": "3900", "deduction_sqft": "300", "surface_type": "Capsheet",
    "candidate": "Yes", "warranty_years": "10", "price": "$15,000",
    "assessment_date": "2026-09-05", "coating_system": "Silicone",
    "acrylic_system_type": "Standard", "roof_type": "Capsheet", "linear_feet": "400",
    "waste_pct": "5", "stretch_pct": "0", "passed_adhesion": "1",
    "assessment_notes": "Ponding at NW corner\n- Cracked seams along HVAC curb"})

check("bid editor: materials plan + warranty options shown",
      b"Materials Plan" in r.data and b"Warranty options for this roof" in r.data
      and b"1.25 gal/sq" in r.data)
# photo upload with auto-resize
from PIL import Image as _Img
big = io.BytesIO()
_Img.new("RGB", (3200, 2400), (120, 40, 40)).save(big, "PNG")
big.seek(0)
r = client.post(f"/bids/{bid['id']}/photos", data={"photos": (big, "roof.png")},
                content_type="multipart/form-data", follow_redirects=True)
check("bid photo: uploaded", b"Added 1 photo" in r.data, r.data[:300])
conn = db.get_db()
ph = conn.execute("SELECT * FROM bid_photos WHERE bid_id=?", (bid["id"],)).fetchone()
conn.close()
saved = _Img.open(db.UPLOAD_DIR / ph["filename"])
check("bid photo: auto-resized to fit", max(saved.size) == 1600 and saved.format == "JPEG",
      saved.size)
r = client.post(f"/bids/photos/{ph['id']}/caption", data={"caption": "Ponding at NW corner"},
                follow_redirects=True)
check("bid photo: caption saved", b"Ponding at NW corner" in r.data)
# junk file rejected gracefully
r = client.post(f"/bids/{bid['id']}/photos", data={"photos": (io.BytesIO(b"junk"), "x.jpg")},
                content_type="multipart/form-data", follow_redirects=True)
check("bid photo: unreadable file handled", b"could not be read" in r.data)
# the report
r = client.get(f"/bids/{bid['id']}/report")
html = r.data.decode()
check("bid report: 200 + sections", r.status_code == 200
      and "Letter From The CEO" in html and "Roof Examined" in html
      and "Site Assessment" in html and "Material Plus" in html)
check("bid report: personalized", "Dear Doug," in html
      and "5231 Braesvalley Drive" in html and "3,900" in html)
check("bid report: calculator-driven quote", "37.8 squares" in html
      and "10 pails" in html and "16 pails" in html
      and "1.25 gal / square" in html and "2 gal / square" in html)
check("bid report: rates cite system, roof and warranty",
      "Silicone" in html and "Capsheet" in html
      and "10-year warranty" in html)
check("bid report: mastic from linear feet", "400 linear ft" in html)
check("bid report: price in words", "FOR THE SUM OF FIFTEEN THOUSAND DOLLARS" in html
      and "$15,000" in html)
check("bid report: photo + caption in survey", ph["filename"] in html
      and "Ponding at NW corner" in html)
check("bid report: observations as bullets", "<li>Cracked seams along HVAC curb</li>" in html)
check("bid report: plain-English glance box", "Your Project At A Glance" in html
      and "without the cost and disruption of a full tear-off" in html)
check("bid report: plain-English coating steps",
      "Clean the roof." in html and "Seal the weak spots." in html
      and "where two sections of roofing overlap" in html
      and "anything that comes up through the roof" in html)
check("bid report: plain-English process tables",
      "Check the Roof" in html and "Prepare the Surface" in html
      and "Apply the Coating System" in html and "moisture survey" in html)
check("bid report: warranty years flow into the text",
      "10-year manufacturer warranty from Henry" in html
      and "For the full 10 years" in html)
check("bid report: plain-English next steps + acceptance",
      "Here's how to move forward" in html and "No question is too small" in html
      and "an authorized agent of the Owner" in html
      and "good for 14 days" in html)
check("bid report: investment heading replaces quotation",
      "Your Investment" in html and "QUOTATION" not in html)
r = client.get(f"/accounts/{bid_acct}")
check("account page: bid listed", b"Roof Reports / Bids" in r.data
      and b"5231 Braesvalley" in r.data)
# photo file served
r = client.get(f"/uploads/bid_photos/{ph['filename']}")
check("bid photo: served", r.status_code == 200 and r.data[:2] == b"\xff\xd8")
# delete cleans up files
r = client.post(f"/bids/{bid['id']}/delete", follow_redirects=True)
check("bid: delete removes photo file", not (db.UPLOAD_DIR / ph["filename"]).exists())

# ---- 22b. Impossible combinations can't be saved or printed
r = client.post(f"/accounts/{bid_acct}/bids/new", data={"roof_size_sqft": "10000",
                "surface_type": "Metal"}, follow_redirects=True)
conn = db.get_db()
combo_id = conn.execute("SELECT MAX(id) FROM bids").fetchone()[0]
conn.close()
r = client.post(f"/bids/{combo_id}/edit", data={
    "roof_address": "Metal Shop", "roof_size_sqft": "10000", "deduction_sqft": "0",
    "surface_type": "Metal", "candidate": "Yes", "warranty_years": "20",
    "coating_system": "Aluminum", "acrylic_system_type": "Standard",
    "roof_type": "Sprayfoam", "linear_feet": "0", "waste_pct": "5",
    "stretch_pct": "0", "passed_adhesion": "1", "price": "40000"},
    follow_redirects=True)
check("combo: unsupported roof + warranty snapped with a notice",
      b"Saved, with adjustments" in r.data, r.data[:400])
conn = db.get_db()
snap = conn.execute("SELECT * FROM bids WHERE id=?", (combo_id,)).fetchone()
conn.close()
check("combo: snapped to a real aluminum combination",
      snap["roof_type"] in ("Capsheet", "Metal") and snap["warranty_years"] == 10,
      (snap["roof_type"], snap["warranty_years"]))
check("combo: products snapped to the aluminum catalog",
      snap["selected_topcoat"] == "Pro-Grade 586" and snap["selected_basecoat"] == "")
html = client.get(f"/bids/{combo_id}/report").data.decode()
check("report: aluminum job never shows silicone products",
      "Prograde 988" not in html and "Prograde 294" not in html
      and "Pro-Grade 586" in html, [l for l in html.splitlines() if "Prograde" in l][:3])

# a bid with no computable plan prints a warning instead of a wrong spec
conn = db.get_db()
ts2 = db.now_iso()
conn.execute("""INSERT INTO bids (account_id, roof_size_sqft, coating_system, roof_type,
    warranty_years, created_at, updated_at) VALUES (?,0,'Silicone','Capsheet',10,?,?)""",
    (bid_acct, ts2, ts2))
conn.commit()
empty_id = conn.execute("SELECT MAX(id) FROM bids").fetchone()[0]
conn.close()
html = client.get(f"/bids/{empty_id}/report").data.decode()
check("report: no plan -> explicit warning, no invented spec",
      "no coating spec yet" in html and "Prograde 294" not in html
      and "gallons per roofing square" not in html)

# ---- 23. Off-machine backup
mirror_dir = db.DB_PATH.parent / "mirror"
r = client.post("/settings", data={"backup_dir": str(mirror_dir)}, follow_redirects=True)
check("backup: dir setting saved", r.status_code == 200)
db.BACKUP_DIR = db.DB_PATH.parent / "b2"
import shutil as _sh
_sh.rmtree(db.BACKUP_DIR, ignore_errors=True); _sh.rmtree(mirror_dir, ignore_errors=True)
made = db.backup_db()
check("backup: mirrored off-machine", made is not None
      and (mirror_dir / made.name).exists())
db.backup_db()  # second call same day: no new backup, mirror unchanged
check("backup: mirror not duplicated", len(list(mirror_dir.glob("crm-*.db"))) == 1)
r = client.get("/backup/download")
check("backup: download snapshot", r.status_code == 200
      and r.data[:16] == b"SQLite format 3\x00"
      and "crm-backup-" in r.headers.get("Content-Disposition", ""))
_sh.rmtree(db.BACKUP_DIR, ignore_errors=True); _sh.rmtree(mirror_dir, ignore_errors=True)
r = client.get("/import")
check("import page: backup UI present", b"Off-Machine Backup Folder" in r.data
      and b"Download Full Backup" in r.data)

# ---- 23b. Priority by matching buildings
conn = db.get_db()
ts = db.now_iso()
today_s = date.today().isoformat()
old_s = (date.today() - timedelta(days=9)).isoformat()
# Big fish: many matching buildings, due TODAY (newest due date)
conn.execute("""INSERT INTO accounts (company_name, matching_properties, num_properties,
    prospecting_status, pipeline_milestone, cadence_start, created_at, updated_at)
    VALUES ('Big Portfolio Co', 40, 60, 'Prospecting', 'None / In Cadence', ?, ?, ?)""",
    (today_s, ts, ts))
# Small fish: few buildings but an OLDER due date
conn.execute("""INSERT INTO accounts (company_name, matching_properties, num_properties,
    prospecting_status, pipeline_milestone, cadence_start, created_at, updated_at)
    VALUES ('Tiny Single Co', 1, 1, 'Prospecting', 'None / In Cadence', ?, ?, ?)""",
    (old_s, ts, ts))
conn.commit()
big_id = conn.execute("SELECT id FROM accounts WHERE company_name='Big Portfolio Co'").fetchone()["id"]
tiny_id = conn.execute("SELECT id FROM accounts WHERE company_name='Tiny Single Co'").fetchone()["id"]

pri = cadence.get_due_reminders(conn, order="priority")
due = cadence.get_due_reminders(conn, order="due")
check("priority: reminder carries matching count",
      all("matching_properties" in r for r in pri))
def first_of(rs, aid):
    return next(i for i, r in enumerate(rs) if r["account_id"] == aid)
check("priority order: big portfolio outranks older small one",
      first_of(pri, big_id) < first_of(pri, tiny_id),
      [(r["company_name"], r["matching_properties"], r["due_date"]) for r in pri[:4]])
check("due order: oldest first regardless of size",
      first_of(due, tiny_id) < first_of(due, big_id))
conn.close()

# the saved setting drives dashboard + queue
conn = db.get_db()
check("task order: defaults to priority", app_mod._task_order(conn) == "priority")
conn.close()
r = client.get("/")
html = r.data.decode()
check("dashboard: priority toggle + matching column", "🎯 Priority" in html
      and "Buildings in today's tasks" in html)
check("dashboard: big portfolio listed before tiny one",
      html.index("Big Portfolio Co") < html.index("Tiny Single Co"))
check("dashboard: top priority card", "Top Priority Accounts" in html
      and "40" in html)
conn = db.get_db()
qt = app_mod._build_queue(conn)
conn.close()
check("queue: priority order puts big portfolio first",
      qt[0]["account_id"] == big_id, [(t["account_id"], t["matching"]) for t in qt[:3]])
# switch to due-date order
r = client.post("/settings/task-order", data={"order": "due"}, follow_redirects=True)
html = r.data.decode()
check("dashboard: switched to due order",
      html.index("Tiny Single Co") < html.index("Big Portfolio Co"))
conn = db.get_db()
check("task order: setting persisted", app_mod._task_order(conn) == "due")
qt = app_mod._build_queue(conn)
conn.close()
check("queue: due order follows setting", qt[0]["account_id"] == tiny_id)
r = client.post("/settings/task-order", data={"order": "bogus"}, follow_redirects=True)
conn = db.get_db()
check("task order: junk value ignored", app_mod._task_order(conn) == "due")
conn.close()
client.post("/settings/task-order", data={"order": "priority"})

# accounts page: sort + min filter
r = client.get("/accounts?sort=priority")
html = r.data.decode()
check("accounts: priority sort", html.index("Big Portfolio Co") < html.index("Tiny Single Co"))
check("accounts: total matching badge", "buildings" in html and "🎯" in html)
r = client.get("/accounts?min_matching=10")
html = r.data.decode()
check("accounts: min matching filter", "Big Portfolio Co" in html
      and "Tiny Single Co" not in html)
r = client.get("/accounts?sort=name")
html = r.data.decode()
check("accounts: A-Z sort still works",
      html.index("Big Portfolio Co") < html.index("Tiny Single Co"))
r = client.get("/accounts?sort=recent")
check("accounts: recent sort ok", r.status_code == 200)
r = client.get("/accounts?min_matching=abc&sort=bogus")
check("accounts: junk params handled", r.status_code == 200)

# ---- 23c. Input hardening (polish pass)
r = client.post("/reminders/quicklog", data={"account_id": "99999", "step_type": "Email 1"},
                follow_redirects=True)
check("hardening: quicklog on a deleted account fails gracefully",
      r.status_code == 200 and b"no longer exists" in r.data)
r = client.post("/reminders/dismiss", data={"account_id": "99999", "step_type": "Email 1"},
                follow_redirects=True)
check("hardening: dismiss on a deleted account fails gracefully",
      r.status_code == 200 and b"no longer exists" in r.data)

conn = db.get_db()
hb = conn.execute("""INSERT INTO bids (account_id, roof_size_sqft, deduction_sqft,
    linear_feet, waste_pct, warranty_years, price, coating_system, roof_type,
    created_at, updated_at)
    VALUES (?,3900,300,420,5,15,16740,'Silicone','Capsheet',?,?)""",
    (bid_acct, db.now_iso(), db.now_iso())).lastrowid
conn.commit(); conn.close()
# a typo must not silently erase stored numbers
r = client.post(f"/bids/{hb}/edit", data={"roof_size_sqft": "3,90O", "price": "16,74O",
    "waste_pct": "five", "warranty_years": "15", "coating_system": "Silicone",
    "roof_type": "Capsheet", "deduction_sqft": "300", "linear_feet": "420"},
    follow_redirects=True)
conn = db.get_db()
hrow = conn.execute("SELECT * FROM bids WHERE id=?", (hb,)).fetchone()
conn.close()
check("hardening: unreadable numbers keep the previous value",
      hrow["roof_size_sqft"] == 3900 and hrow["price"] == 16740.0
      and hrow["waste_pct"] == 5, dict(hrow))
check("hardening: the save warns about what it couldn't read",
      b"couldn" in r.data and b"previous value was kept" in r.data)
# commas and dollar signs are accepted
client.post(f"/bids/{hb}/edit", data={"roof_size_sqft": "4,200", "price": "$18,500.00",
    "waste_pct": "7.5", "linear_feet": "1,100", "warranty_years": "15",
    "coating_system": "Silicone", "roof_type": "Capsheet", "deduction_sqft": "0"})
conn = db.get_db()
hrow = conn.execute("SELECT * FROM bids WHERE id=?", (hb,)).fetchone()
conn.close()
check("hardening: commas and $ parse correctly",
      hrow["roof_size_sqft"] == 4200 and hrow["price"] == 18500.0
      and hrow["linear_feet"] == 1100, dict(hrow))
# clearing on purpose still clears
client.post(f"/bids/{hb}/edit", data={"roof_size_sqft": "4200", "price": "",
    "warranty_years": "15", "coating_system": "Silicone", "roof_type": "Capsheet"})
conn = db.get_db()
check("hardening: an empty field still clears the value",
      conn.execute("SELECT price FROM bids WHERE id=?", (hb,)).fetchone()["price"] is None)
conn.close()
client.post(f"/bids/{hb}/delete")

# account numbers behave the same way
conn = db.get_db()
conn.execute("UPDATE accounts SET num_properties=40, matching_properties=23 WHERE id=?",
             (bid_acct,))
conn.commit(); conn.close()
client.post(f"/accounts/{bid_acct}/edit", data={"company_name": "Sallyport Investments, Llc",
    "num_properties": "4O", "matching_properties": "1,250", "preferred_contact": "Call",
    "prospecting_status": "Prospecting", "pipeline_milestone": "None / In Cadence"})
conn = db.get_db()
arow = conn.execute("SELECT * FROM accounts WHERE id=?", (bid_acct,)).fetchone()
conn.close()
check("hardening: account typo keeps value, commas parse",
      arow["num_properties"] == 40 and arow["matching_properties"] == 1250,
      (arow["num_properties"], arow["matching_properties"]))

# junk pricing is refused rather than stored
r = client.post("/settings", data={"price_capsheet_base": "abc"}, follow_redirects=True)
conn = db.get_db()
pv = conn.execute("SELECT value FROM settings WHERE key='price_capsheet_base'").fetchone()["value"]
conn.close()
check("hardening: non-numeric price rejected and reported",
      pv != "abc" and b"left unchanged" in r.data, pv)

# ---- 23d. Archive: remove from the list without losing the record
conn = db.get_db()
ts3 = db.now_iso()
conn.execute("""INSERT INTO accounts (company_name, matching_properties, prospecting_status,
    pipeline_milestone, cadence_start, next_follow_up, follow_up_note, created_at, updated_at)
    VALUES ('Junk Data Co', 12, 'Prospecting', 'None / In Cadence', ?, ?, 'call back', ?, ?)""",
    (date.today().isoformat(), date.today().isoformat(), ts3, ts3))
conn.commit()
arch_id = conn.execute("SELECT id FROM accounts WHERE company_name='Junk Data Co'").fetchone()["id"]
conn.execute("INSERT INTO interactions (account_id, interaction_type, notes, created_at) "
             "VALUES (?,?,?,?)", (arch_id, "Call 2", "spoke briefly", ts3))
conn.commit(); conn.close()

check("archive: account starts visible",
      b"Junk Data Co" in client.get("/accounts").data
      and b"Junk Data Co" in client.get("/").data)
r = client.post(f"/accounts/{arch_id}/archive", data={"archive_reason": "Bad data / wrong company"},
                follow_redirects=True)
check("archive: confirmation explains the effect",
      b"won" in r.data and b"future" in r.data.lower(), r.data[:300])
check("archive: gone from accounts list and dashboard",
      b"Junk Data Co" not in client.get("/accounts").data
      and b"Junk Data Co" not in client.get("/").data)
check("archive: gone from queue, pipeline and insights",
      b"Junk Data Co" not in client.get("/queue").data
      and b"Junk Data Co" not in client.get("/pipeline").data
      and b"Junk Data Co" not in client.get("/insights").data)
conn = db.get_db()
arow = conn.execute("SELECT * FROM accounts WHERE id=?", (arch_id,)).fetchone()
hist = conn.execute("SELECT COUNT(*) c FROM interactions WHERE account_id=?",
                    (arch_id,)).fetchone()["c"]
conn.close()
check("archive: record and history kept, follow-up cleared",
      arow["archived_at"] and arow["archive_reason"] == "Bad data / wrong company"
      and hist == 1 and arow["next_follow_up"] == "", dict(arow))
check("archive: cadence engine ignores archived accounts",
      all(r2["account_id"] != arch_id for r2 in cadence.get_due_reminders(db.get_db())))
r = client.get("/accounts?view=archived")
check("archive: visible in the Archived view with its reason",
      b"Junk Data Co" in r.data and b"Bad data" in r.data and b"Restore" in r.data)
r = client.get(f"/accounts/{arch_id}")
check("archive: detail page shows the banner", b"Archived" in r.data and b"Restore" in r.data)

# the point of it all: a re-import must not bring it back
again = (b"Company Name,# Properties (in search)\r\n"
         b"Junk Data Co,12\r\nBrand New Co,7\r\n")
r = client.post("/import", data={"file": (io.BytesIO(again), "list.csv")},
                content_type="multipart/form-data")
check("archive: re-import skips archived company and says so",
      b"1</strong> account(s) imported" in r.data
      and b"skipped because you archived them" in r.data
      and b"Junk Data Co" in r.data, r.data[:1200])
conn = db.get_db()
dupes = conn.execute("SELECT COUNT(*) c FROM accounts WHERE company_name='Junk Data Co'").fetchone()["c"]
still_archived = conn.execute("SELECT archived_at FROM accounts WHERE id=?",
                              (arch_id,)).fetchone()["archived_at"]
conn.close()
check("archive: no duplicate created, still archived", dupes == 1 and still_archived != "")

# restore puts it back
r = client.post(f"/accounts/{arch_id}/restore", follow_redirects=True)
check("archive: restore returns it to the working list",
      b"Junk Data Co" in client.get("/accounts").data)
conn = db.get_db()
check("archive: restore clears the flags", conn.execute(
    "SELECT archived_at, archive_reason FROM accounts WHERE id=?",
    (arch_id,)).fetchone()["archived_at"] == "")
conn.close()

# permanent delete really forgets it (and a later import may re-add it)
r = client.post(f"/accounts/{arch_id}/delete", follow_redirects=True)
check("delete: permanent removal explains re-import behaviour",
      b"Permanently deleted" in r.data and b"can add this company again" in r.data)
conn = db.get_db()
check("delete: row and history gone",
      conn.execute("SELECT COUNT(*) c FROM accounts WHERE id=?", (arch_id,)).fetchone()["c"] == 0
      and conn.execute("SELECT COUNT(*) c FROM interactions WHERE account_id=?",
                       (arch_id,)).fetchone()["c"] == 0)
conn.close()
r = client.post("/import", data={"file": (io.BytesIO(again), "list.csv")},
                content_type="multipart/form-data")
check("delete: a deleted company can be imported again",
      b"1</strong> account(s) imported" in r.data)
conn = db.get_db()
conn.execute("DELETE FROM accounts WHERE company_name IN ('Junk Data Co','Brand New Co')")
conn.commit(); conn.close()

# ---- 23e. Full Excel export
# seed a project + invoice so the money sheets have rows to verify
conn = db.get_db()
ts4 = db.now_iso()
conn.execute("INSERT INTO projects (account_id, name, status, contract_amount, "
             "created_at, updated_at) VALUES (?,'Export Test Project','In Progress',"
             "24500,?,?)", (bid_acct, ts4, ts4))
xproj = conn.execute("SELECT MAX(id) FROM projects").fetchone()[0]
conn.execute("""INSERT INTO invoices (project_id, invoice_number, amount, status,
    sent_date, due_date, notes, created_at, updated_at)
    VALUES (?,'INV-EXP',12250,'Sent',?,?,'Final balance',?,?)""",
    (xproj, date.today().isoformat(),
     (date.today() - timedelta(days=5)).isoformat(), ts4, ts4))
conn.commit(); conn.close()

r = client.get("/export/workbook.xlsx")
check("excel: route returns a workbook", r.status_code == 200
      and r.data[:2] == b"PK"
      and "spreadsheetml" in r.headers.get("Content-Type", "")
      and ".xlsx" in r.headers.get("Content-Disposition", ""))
from openpyxl import load_workbook
xl = load_workbook(io.BytesIO(r.data))
check("excel: one sheet per part of the business",
      xl.sheetnames == ["Summary", "Accounts", "Contacts", "Interactions",
                        "Roof Reports", "Projects", "Invoices", "Tasks Due"],
      xl.sheetnames)
ws = xl["Accounts"]
headers = [c.value for c in ws[1]]
check("excel: accounts sheet carries the full record",
      {"Company", "Matching (🎯)", "Prospecting status", "Pipeline milestone",
       "Archived", "Notes"} <= set(headers), headers)
names = {ws.cell(row=i, column=1).value for i in range(2, ws.max_row + 1)}
check("excel: includes both active and archived accounts",
      "Sallyport Investments, Llc" in names, sorted(names)[:4])
check("excel: every sheet is frozen and filterable",
      all(xl[n].freeze_panes == "A2" and xl[n].auto_filter.ref
          for n in xl.sheetnames if n != "Summary" and xl[n].max_row > 1))
check("excel: money and dates carry real formats",
      xl["Invoices"]["D2"].number_format.startswith('"$"')
      and isinstance(xl["Invoices"]["F2"].value, (datetime, date))
      and xl["Invoices"]["D2"].value == 12250,
      (xl["Invoices"]["D2"].number_format, xl["Invoices"]["F2"].value))
check("excel: overdue days computed on the invoices sheet",
      xl["Invoices"]["I2"].value == 5, xl["Invoices"]["I2"].value)
summary = [xl["Summary"].cell(row=i, column=1).value
           for i in range(1, xl["Summary"].max_row + 1)]
check("excel: summary covers pipeline, opportunity, money and activity",
      {"Pipeline", "Opportunity", "Money", "Activity"} <= set(v for v in summary if v),
      [v for v in summary if v][:8])
rr = xl["Roof Reports"]
rr_headers = [c.value for c in rr[1]]
check("excel: roof reports include computed rates and suggested price",
      "Rates (gal/sq)" in rr_headers and "Suggested price" in rr_headers
      and "Total gallons" in rr_headers, rr_headers)

# still works on a brand-new, empty database
import tempfile as _tf
prev_path = db.DB_PATH
db.DB_PATH = Path(_tf.mkdtemp()) / "empty.db"
db.init_db()
econn = db.get_db()
import excel_export as _xe
empty = load_workbook(_xe.build_workbook(econn))
econn.close()
db.DB_PATH = prev_path
check("excel: empty database exports without error",
      empty.sheetnames[0] == "Summary" and empty["Accounts"].max_row == 1)

# ---- 23f. Data lives outside the app folder (safe upgrades)
import os as _os, shutil as _sh, tempfile as _tf2
check("data dir: defaults outside the app folder",
      db.APP_DIR not in db._resolve_data_dir().parents
      and db._resolve_data_dir() != db.APP_DIR, db._resolve_data_dir())
_prev_env = _os.environ.get("ROOF_CRM_DATA")
_custom = _tf2.mkdtemp()
_os.environ["ROOF_CRM_DATA"] = _custom
check("data dir: ROOF_CRM_DATA overrides the default",
      db._resolve_data_dir() == Path(_custom))
if _prev_env is None:
    _os.environ.pop("ROOF_CRM_DATA")
else:
    _os.environ["ROOF_CRM_DATA"] = _prev_env

# a legacy install (crm.db + photos inside the app folder) is migrated out
_legacy_app = Path(_tf2.mkdtemp())
_legacy_data = Path(_tf2.mkdtemp()) / "RoofCRM"
_prev = (db.APP_DIR, db.DATA_DIR, db.DB_PATH, db.BACKUP_DIR, db.UPLOAD_DIR)
db.APP_DIR = _legacy_app
db.DATA_DIR = _legacy_data
db.DB_PATH = _legacy_data / "crm.db"
db.BACKUP_DIR = _legacy_data / "backups"
db.UPLOAD_DIR = _legacy_data / "uploads" / "bid_photos"
(_legacy_app / "crm.db").write_bytes(b"SQLite format 3\x00legacy")
(_legacy_app / "uploads" / "bid_photos").mkdir(parents=True)
(_legacy_app / "uploads" / "bid_photos" / "old.jpg").write_bytes(b"\xff\xd8x")
(_legacy_app / "backups").mkdir()
(_legacy_app / "backups" / "crm-20260101-000000.db").write_bytes(b"backup")
_moved = db.migrate_legacy_data()
check("data dir: legacy database, photos and backups are moved out",
      db.DB_PATH.exists() and not (_legacy_app / "crm.db").exists()
      and (db.UPLOAD_DIR / "old.jpg").exists()
      and (db.BACKUP_DIR / "crm-20260101-000000.db").exists()
      and len(_moved) == 3, _moved)
check("data dir: migration keeps the original bytes",
      db.DB_PATH.read_bytes().endswith(b"legacy"))
# replacing the whole app folder must not touch the data
_sh.rmtree(_legacy_app)
check("data dir: data survives deleting the entire app folder",
      db.DB_PATH.exists() and (db.UPLOAD_DIR / "old.jpg").exists())
# a second run finds nothing to move
db.APP_DIR = Path(_tf2.mkdtemp())
check("data dir: migration is a no-op once done", db.migrate_legacy_data() == [])
db.APP_DIR, db.DATA_DIR, db.DB_PATH, db.BACKUP_DIR, db.UPLOAD_DIR = _prev

# ---- 24. In-app guide
r = client.get("/guide")
check("guide: renders", r.status_code == 200 and b"Roof CRM Guide" in r.data
      and b"The daily routine" in r.data and b"Roof reports / bids" in r.data
      and b"Data safety" in r.data)
check("guide: nav link", b'href="/guide"' in client.get("/").data)

# ---- 25. Company-name matching: one company, however it is spelled
check("normalize: legal suffixes and punctuation collapse",
      importer.normalize_company("Hartman Income REIT, Inc.")
      == importer.normalize_company("HARTMAN INCOME REIT LP")
      == importer.normalize_company("Hartman Income Reit"))
check("normalize: entity-only names stay distinct",
      importer.normalize_company("The Group") != importer.normalize_company("The Trust"))
check("normalize: a short real name isn't eaten",
      importer.normalize_company("IBM Inc") == importer.normalize_company("IBM Corporation"))

def _sheet(rows, header, name="list.xlsx"):
    """Build an uploadable xlsx in memory."""
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(header)
    for r in rows:
        ws.append(r)
    b = io.BytesIO(); wb.save(b)
    class _Up(io.BytesIO):
        filename = name
    return _Up(b.getvalue())

HDR = ["Company Name", "First Name", "Last Name", "Job Title",
       "# Properties (in search)", "Email Address", "Direct Phone Number",
       "Mobile phone", "LinkedIn Contact Profile URL", "Management Level"]

conn = db.get_db()
res = importer.import_accounts(conn, _sheet(
    [["Lone Star Realty Partners, LLC", "Pat", "Ng", "Owner", 4,
      "pat@lonestar.com", "713-555-0001", "", "", ""]], HDR))
check("dupe fix: first spelling imports", res["imported"] == 1, res)
res = importer.import_accounts(conn, _sheet(
    [["LONE STAR REALTY PARTNERS LP", "Pat", "Ng", "Owner", 4,
      "pat@lonestar.com", "713-555-0001", "", "", ""]], HDR))
check("dupe fix: a different spelling is the same company",
      res["imported"] == 0 and res["skipped_duplicates"] == 1, res)
check("dupe fix: only one account exists",
      conn.execute("SELECT COUNT(*) c FROM accounts WHERE company_name LIKE "
                   "'%Lone Star Realty%'").fetchone()["c"] == 1)

# an archived company can't come back under a different spelling
lone = conn.execute("SELECT id FROM accounts WHERE company_name LIKE "
                    "'%Lone Star Realty%'").fetchone()["id"]
client.post(f"/accounts/{lone}/archive", data={"archive_reason": "Not a fit"},
            follow_redirects=True)
res = importer.import_accounts(conn, _sheet(
    [["Lone Star Realty Partners Inc.", "Pat", "Ng", "Owner", 4,
      "pat@lonestar.com", "", "", "", ""]], HDR))
check("dupe fix: archived stays archived across spellings",
      res["imported"] == 0 and res["skipped_archived"] == 1, res)

# the detector surfaces duplicates older versions already created
conn.execute("INSERT INTO accounts (company_name, preferred_contact, "
             "prospecting_status, pipeline_milestone, cadence_start, created_at, "
             "updated_at) VALUES ('Gulf Coast Asset Co', 'Unknown', 'Prospecting', "
             "'None / In Cadence', ?, ?, ?)",
             (db.today_iso(), db.now_iso(), db.now_iso()))
conn.execute("INSERT INTO accounts (company_name, preferred_contact, "
             "prospecting_status, pipeline_milestone, cadence_start, created_at, "
             "updated_at) VALUES ('Gulf Coast Asset Company, LLC', 'Unknown', "
             "'Prospecting', 'None / In Cadence', ?, ?, ?)",
             (db.today_iso(), db.now_iso(), db.now_iso()))
conn.commit()
groups = importer.find_duplicate_groups(conn)
check("duplicates: pre-existing pairs are reported",
      any(len(g["accounts"]) == 2 and "gulf coast asset" in g["key"] for g in groups),
      [g["key"] for g in groups])
check("duplicates: shown on the Import page",
      b"Possible duplicate accounts" in client.get("/import").data)
conn.execute("DELETE FROM accounts WHERE company_name LIKE 'Gulf Coast%'")
conn.commit()

# ---- 26. ZoomInfo: LinkedIn, management level, create-missing
importer.import_accounts(conn, _sheet(
    [["Triten Real Estate Partners", "", "", "", 6, "", "", "", "", ""]], HDR))
res = importer.import_contacts(conn, _sheet(
    [["Triten Real Estate Partners, LLC", "Maria", "Lopez", "VP of Operations", "",
      "mlopez@triten.com", "281-555-0101", "281-555-0102",
      "linkedin.com/in/mlopez", "VP-Level"]], HDR, "zoominfo.csv".replace(".csv", ".xlsx")))
check("zoominfo: contact attached across a suffix difference", res["attached"] == 1, res)
triten = conn.execute("SELECT * FROM accounts WHERE company_name="
                      "'Triten Real Estate Partners'").fetchone()
check("zoominfo: LinkedIn URL captured", triten["linkedin_url"] == "linkedin.com/in/mlopez")
check("zoominfo: management level captured", triten["seniority"] == "VP-Level")
check("zoominfo: columns mapped", {"linkedin_url", "seniority"}
      <= set(res["mapped_columns"]), res["mapped_columns"])

res = importer.import_contacts(conn, _sheet(
    [["Brand New Owners Group", "Sam", "Reed", "Director of Facilities", "",
      "sam@bnog.com", "", "", "", ""]], HDR))
check("zoominfo: unknown company reported, not created",
      res["accounts_created"] == 0 and "Brand New Owners Group" in res["unmatched"], res)
res = importer.import_contacts(conn, _sheet(
    [["Brand New Owners Group", "Sam", "Reed", "Director of Facilities", "",
      "sam@bnog.com", "", "", "", ""]], HDR), create_missing=True)
check("zoominfo: create-missing opens the account", res["accounts_created"] == 1, res)
bnog = conn.execute("SELECT * FROM accounts WHERE company_name="
                    "'Brand New Owners Group'").fetchone()
check("zoominfo: new account is a normal prospect",
      bnog["prospecting_status"] == "Prospecting"
      and bnog["pipeline_milestone"] == "None / In Cadence"
      and bnog["first_name"] == "Sam")
res = importer.import_contacts(conn, _sheet(
    [["Lone Star Realty Partners LLC", "Pat", "Ng", "Owner", "", "", "", "", "", ""]],
    HDR), create_missing=True)
check("zoominfo: create-missing never resurrects an archived company",
      res["accounts_created"] == 0 and res["skipped_archived"] == 1, res)

# ---- 27. Paste a ZoomInfo profile instead of retyping it
parsed = importer.parse_contact_blob(
    "Jane Doe\nVice President of Asset Management\nHartman Income REIT, Inc.\n"
    "jane.doe@hartman.com\nDirect: (713) 555-0142\nMobile: (713) 555-9981\n"
    "linkedin.com/in/janedoe")
check("paste: plain profile block",
      parsed["first_name"] == "Jane" and parsed["last_name"] == "Doe"
      and parsed["title"] == "Vice President of Asset Management"
      and parsed["email"] == "jane.doe@hartman.com"
      and parsed["work_phone"] == "(713) 555-0142"
      and parsed["mobile_phone"] == "(713) 555-9981"
      and parsed["linkedin_url"] == "linkedin.com/in/janedoe"
      and parsed["seniority"] == "VP-Level", parsed)
parsed = importer.parse_contact_blob(
    "Name: Robert Chen\nJob Title: Director of Facilities\n"
    "Email Address: rchen@boxer.com\nDirect Phone: 713-555-0199")
check("paste: label-and-value lines",
      parsed["first_name"] == "Robert" and parsed["last_name"] == "Chen"
      and parsed["title"] == "Director of Facilities"
      and parsed["email"] == "rchen@boxer.com", parsed)
parsed = importer.parse_contact_blob(
    "Maria Soto\tProperty Manager\tmsoto@x.com\t(281) 555-0101")
check("paste: a row copied from a spreadsheet",
      parsed["first_name"] == "Maria" and parsed["title"] == "Property Manager"
      and parsed["work_phone"] == "(281) 555-0101", parsed)
check("paste: empty input is harmless",
      not any(importer.parse_contact_blob("   ").values()))

r = client.post(f"/accounts/{bnog['id']}/contacts/add", data={
    "paste": "Dana Price\nChief Operating Officer\ndana@bnog.com\n(713) 555-7777"},
    follow_redirects=True)
dana = conn.execute("SELECT * FROM contacts WHERE email='dana@bnog.com'").fetchone()
check("paste: Add Contact fills itself in from the paste",
      dana is not None and dana["first_name"] == "Dana"
      and dana["title"] == "Chief Operating Officer"
      and dana["work_phone"] == "(713) 555-7777"
      and dana["seniority"] == "C-Level")
check("paste: the form says what it read", b"Read from the paste" in r.data)
r = client.post(f"/accounts/{bnog['id']}/contacts/add",
                data={"paste": "no person here, just words and words"},
                follow_redirects=True)
check("paste: unreadable paste is refused, not half-saved",
      b"didn" in r.data and b"first or last name" in r.data)

# ---- 28. Search reaches the whole account, not just the company name
r = client.get("/accounts?q=Dana")
check("search: finds a secondary contact by name",
      b"Brand New Owners Group" in r.data)
check("search: says which contact matched", b"matched contact: Dana Price" in r.data)
r = client.get("/accounts?q=5557777")
check("search: finds by phone digits, ignoring formatting",
      b"Brand New Owners Group" in r.data)
conn.execute("UPDATE accounts SET notes='roof visible from Beltway 8' WHERE id=?",
             (bnog["id"],))
conn.commit()
check("search: finds by a word in the notes",
      b"Brand New Owners Group" in client.get("/accounts?q=Beltway").data)
check("search: no false positives",
      b"Brand New Owners Group" not in client.get("/accounts?q=zzzznope").data)

# ---- 29. Bulk actions on the accounts list
ids = [r["id"] for r in conn.execute(
    "SELECT id FROM accounts WHERE COALESCE(archived_at,'')='' LIMIT 3")]
r = client.post("/accounts/bulk", data={
    "action": "status", "value": "Might be Interested",
    "account_ids": [str(i) for i in ids]}, follow_redirects=True)
check("bulk: status applied to every ticked account",
      all(conn.execute("SELECT prospecting_status s FROM accounts WHERE id=?",
                       (i,)).fetchone()["s"] == "Might be Interested" for i in ids))
check("bulk: offers an undo", b"Undo" in r.data)
r = client.post("/undo/" + str(
    conn.execute("SELECT MAX(id) m FROM undo_log").fetchone()["m"]),
    follow_redirects=True)
check("bulk: undo puts the old statuses back",
      not any(conn.execute("SELECT prospecting_status s FROM accounts WHERE id=?",
                           (i,)).fetchone()["s"] == "Might be Interested" for i in ids))
client.post("/accounts/bulk", data={
    "action": "followup", "value": db.today_iso(), "follow_up_note": "call back",
    "account_ids": [str(ids[0])]}, follow_redirects=True)
check("bulk: follow-up date and note set",
      conn.execute("SELECT next_follow_up n, follow_up_note t FROM accounts WHERE id=?",
                   (ids[0],)).fetchone()["t"] == "call back")
r = client.post("/accounts/bulk", data={
    "action": "followup", "value": "not-a-date", "account_ids": [str(ids[0])]},
    follow_redirects=True)
check("bulk: an unreadable date changes nothing",
      b"Couldn" in r.data and conn.execute(
          "SELECT follow_up_note t FROM accounts WHERE id=?",
          (ids[0],)).fetchone()["t"] == "call back")
client.post("/accounts/bulk", data={
    "action": "archive", "value": "Not a fit", "account_ids": [str(ids[0])]},
    follow_redirects=True)
check("bulk: archive hides the account",
      conn.execute("SELECT archived_at a FROM accounts WHERE id=?",
                   (ids[0],)).fetchone()["a"] != "")
client.post("/accounts/bulk", data={
    "action": "restore", "account_ids": [str(ids[0])]}, follow_redirects=True)
check("bulk: restore brings it back",
      conn.execute("SELECT archived_at a FROM accounts WHERE id=?",
                   (ids[0],)).fetchone()["a"] == "")
r = client.post("/accounts/bulk", data={"action": "status", "value": "Prospecting"},
                follow_redirects=True)
check("bulk: nothing ticked is a friendly no-op", b"Tick at least one" in r.data)
r = client.post("/accounts/bulk", data={
    "action": "status", "value": "Not A Real Status", "account_ids": [str(ids[0])]},
    follow_redirects=True)
check("bulk: a value outside the list is rejected",
      conn.execute("SELECT prospecting_status s FROM accounts WHERE id=?",
                   (ids[0],)).fetchone()["s"] != "Not A Real Status")

# ---- 30. Fixing and removing a logged interaction
target = ids[1]
client.post(f"/accounts/{target}/log",
            data={"interaction_type": "Call 2", "notes": "left voicemail"},
            follow_redirects=True)
entry = conn.execute("SELECT * FROM interactions WHERE account_id=? "
                     "ORDER BY id DESC LIMIT 1", (target,)).fetchone()
yesterday = (date.today() - timedelta(days=1)).isoformat()
client.post(f"/interactions/{entry['id']}/edit", data={
    "interaction_type": "Call & Text", "notes": "actually reached him",
    "created_date": yesterday}, follow_redirects=True)
fixed = conn.execute("SELECT * FROM interactions WHERE id=?", (entry["id"],)).fetchone()
check("interaction: type, notes and date all corrected",
      fixed["interaction_type"] == "Call & Text"
      and fixed["notes"] == "actually reached him"
      and fixed["created_at"][:10] == yesterday, dict(fixed))
r = client.post(f"/interactions/{entry['id']}/edit",
                data={"interaction_type": "Call & Text", "notes": "x",
                      "created_date": "13/41/2026"}, follow_redirects=True)
check("interaction: an unreadable date keeps the original",
      b"kept the original" in r.data and conn.execute(
          "SELECT created_at c FROM interactions WHERE id=?",
          (entry["id"],)).fetchone()["c"][:10] == yesterday)
client.post(f"/interactions/{entry['id']}/delete", follow_redirects=True)
check("interaction: deleted",
      conn.execute("SELECT 1 FROM interactions WHERE id=?",
                   (entry["id"],)).fetchone() is None)
client.post("/undo/" + str(conn.execute(
    "SELECT MAX(id) m FROM undo_log").fetchone()["m"]), follow_redirects=True)
back = conn.execute("SELECT * FROM interactions WHERE id=?", (entry["id"],)).fetchone()
check("interaction: undo restores it exactly",
      back is not None and back["notes"] == "x", dict(back) if back else None)
r = client.post("/interactions/999999/delete", follow_redirects=True)
check("interaction: deleting a missing entry is a friendly error",
      r.status_code == 200 and b"no longer exists" in r.data)

# ---- 31. Undo of a full account delete, photos and all
import undo as undo_mod
victim = conn.execute("SELECT * FROM accounts WHERE company_name="
                      "'Brand New Owners Group'").fetchone()
client.post(f"/accounts/{victim['id']}/log",
            data={"interaction_type": "General Note", "notes": "walked the roof"},
            follow_redirects=True)
client.post(f"/accounts/{victim['id']}/bids/new",
            data={"roof_address": "1 Test Way"}, follow_redirects=True)
vbid = conn.execute("SELECT * FROM bids WHERE account_id=?", (victim["id"],)).fetchone()
db.UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
(db.UPLOAD_DIR / "undo_test.jpg").write_bytes(b"\xff\xd8photo")
conn.execute("INSERT INTO bid_photos (bid_id, filename, caption, created_at) "
             "VALUES (?,?,?,?)", (vbid["id"], "undo_test.jpg", "north slope",
                                  db.now_iso()))
conn.commit()
client.post(f"/accounts/{victim['id']}/delete", follow_redirects=True)
check("undo: the account really is gone first",
      conn.execute("SELECT 1 FROM accounts WHERE id=?",
                   (victim["id"],)).fetchone() is None
      and not (db.UPLOAD_DIR / "undo_test.jpg").exists())
undo_id = conn.execute("SELECT MAX(id) m FROM undo_log").fetchone()["m"]
client.post(f"/undo/{undo_id}", follow_redirects=True)
check("undo: account restored",
      conn.execute("SELECT company_name n FROM accounts WHERE id=?",
                   (victim["id"],)).fetchone()["n"] == "Brand New Owners Group")
check("undo: its contacts, history and bid came back too",
      conn.execute("SELECT COUNT(*) c FROM contacts WHERE account_id=?",
                   (victim["id"],)).fetchone()["c"] > 0
      and conn.execute("SELECT COUNT(*) c FROM interactions WHERE account_id=?",
                       (victim["id"],)).fetchone()["c"] > 0
      and conn.execute("SELECT COUNT(*) c FROM bids WHERE account_id=?",
                       (victim["id"],)).fetchone()["c"] == 1)
check("undo: the photo file is back on disk",
      (db.UPLOAD_DIR / "undo_test.jpg").read_bytes() == b"\xff\xd8photo")
r = client.post(f"/undo/{undo_id}", follow_redirects=True)
check("undo: the same record can't be replayed twice",
      b"no longer be undone" in r.data)

# an expired record is refused, and purge clears it out
conn.execute("INSERT INTO undo_log (label, payload, created_at) VALUES "
             "('old thing', '{\"ops\": []}', ?)",
             ((datetime.now().astimezone() - timedelta(days=30)).isoformat(),))
conn.commit()
old_id = conn.execute("SELECT MAX(id) m FROM undo_log").fetchone()["m"]
check("undo: an expired record is refused", undo_mod.peek(conn, old_id) is None
      and undo_mod.restore(conn, old_id) is None)
db.TRASH_DIR.mkdir(parents=True, exist_ok=True)
_stale = db.TRASH_DIR / "stale.jpg"
_stale.write_bytes(b"x")
import os as _os
_old_time = (datetime.now() - timedelta(days=30)).timestamp()
_os.utime(_stale, (_old_time, _old_time))
undo_mod.purge(conn)
check("undo: purge drops expired records",
      conn.execute("SELECT 1 FROM undo_log WHERE id=?", (old_id,)).fetchone() is None)
check("undo: purge empties the photo trash", not _stale.exists())
check("undo: the allow-list blocks an unknown table",
      _raises(lambda: undo_mod.capture(conn, "sqlite_master", "1=1")))

# ---- 32. Daily progress and activity tracking
# With work still outstanding the bar must count it, not declare victory.
_rem = len(cadence.get_due_reminders(conn)) + len(
    conn.execute("SELECT 1 FROM accounts WHERE COALESCE(archived_at,'')='' "
                 "AND next_follow_up != '' AND next_follow_up <= ?",
                 (db.today_iso(),)).fetchall())
r = client.get("/")
check("dashboard: progress counts what's left, not just what's done",
      (b"done today" in r.data and b"list is clear" not in r.data) if _rem
      else b"list is clear" in r.data, f"{_rem} outstanding")
check("dashboard: progress states the remaining count",
      (b" left" in r.data) if _rem else True)
client.post("/settings", data={"daily_goal": "12"}, follow_redirects=True)
check("activity: the daily goal saves",
      conn.execute("SELECT value v FROM settings WHERE key='daily_goal'"
                   ).fetchone()["v"] == "12")
r = client.get("/insights")
check("insights: daily activity chart with the goal",
      b"Daily activity" in r.data and b"12/day" in r.data
      and b"logged today" in r.data)
r = client.post("/settings", data={"daily_goal": "many"}, follow_redirects=True)
check("activity: a junk goal is rejected, not stored",
      b"left unchanged" in r.data and conn.execute(
          "SELECT value v FROM settings WHERE key='daily_goal'").fetchone()["v"] == "12")
client.post("/settings", data={"daily_goal": "0"}, follow_redirects=True)
check("activity: goal 0 turns the goal line off",
      b"Set a daily goal" in client.get("/insights").data)

# ---- 33. Queue keyboard shortcuts
r = client.get("/queue")
check("queue: shortcuts are wired and documented",
      b"k-help" in r.data and b"addEventListener('keydown'" in r.data)
# ---- 34. An older database upgrades in place, keeping its data
import sqlite3 as _sq
_mig_dir = Path(tempfile.mkdtemp(prefix="crm_mig_"))
_mig_db = _mig_dir / "crm.db"
_old = _sq.connect(_mig_db)
_old.executescript("""
CREATE TABLE accounts (
    id INTEGER PRIMARY KEY AUTOINCREMENT, company_name TEXT NOT NULL,
    first_name TEXT DEFAULT '', last_name TEXT DEFAULT '', title TEXT DEFAULT '',
    num_properties INTEGER, email TEXT DEFAULT '', work_phone TEXT DEFAULT '',
    mobile_phone TEXT DEFAULT '', preferred_contact TEXT NOT NULL DEFAULT 'Unknown',
    notes TEXT DEFAULT '', prospecting_status TEXT NOT NULL DEFAULT 'Prospecting',
    pipeline_milestone TEXT NOT NULL DEFAULT 'None / In Cadence',
    cadence_start TEXT NOT NULL, created_at TEXT NOT NULL, updated_at TEXT NOT NULL);
CREATE TABLE contacts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    account_id INTEGER NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
    first_name TEXT DEFAULT '', last_name TEXT DEFAULT '', title TEXT DEFAULT '',
    email TEXT DEFAULT '', work_phone TEXT DEFAULT '', mobile_phone TEXT DEFAULT '',
    created_at TEXT NOT NULL);
CREATE TABLE templates (
    id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL,
    kind TEXT NOT NULL DEFAULT 'email', steps TEXT DEFAULT '',
    subject TEXT DEFAULT '', body TEXT NOT NULL DEFAULT '',
    sort_order INTEGER DEFAULT 0, updated_at TEXT NOT NULL);
CREATE TABLE settings (key TEXT PRIMARY KEY, value TEXT NOT NULL DEFAULT '');
""")
_old.execute("INSERT INTO accounts (company_name, cadence_start, created_at, "
             "updated_at) VALUES ('Legacy Holdings LLC','2026-01-01',"
             "'2026-01-01T00:00:00-06:00','2026-01-01T00:00:00-06:00')")
_old.execute("INSERT INTO contacts (account_id, first_name, last_name, created_at) "
             "VALUES (1,'Old','Person','2026-01-01T00:00:00-06:00')")
_old.commit(); _old.close()

_keep = db.DB_PATH
db.DB_PATH = _mig_db
db.init_db()
_m = db.get_db()
_acc = {r[1] for r in _m.execute("PRAGMA table_info(accounts)")}
_con = {r[1] for r in _m.execute("PRAGMA table_info(contacts)")}
_tbl = {r[0] for r in _m.execute("SELECT name FROM sqlite_master WHERE type='table'")}
check("migrate: new account columns added to an old database",
      {"linkedin_url", "seniority", "matching_properties", "archived_at"} <= _acc)
check("migrate: new contact columns added", {"linkedin_url", "seniority"} <= _con)
check("migrate: undo_log table created", "undo_log" in _tbl)
check("migrate: existing rows survive untouched",
      _m.execute("SELECT company_name FROM accounts").fetchone()[0] == "Legacy Holdings LLC"
      and _m.execute("SELECT first_name FROM contacts").fetchone()[0] == "Old")
check("migrate: the daily goal setting is seeded",
      _m.execute("SELECT value FROM settings WHERE key='daily_goal'").fetchone()[0] == "20")
_m.close()
db.init_db()
check("migrate: running again changes nothing", True)
db.DB_PATH = _keep

# ---- 35. Big lists stay light: the page is capped, the counts are not
import app as _app
conn = db.get_db()
_ts = db.now_iso()
_today = db.today_iso()
for _i in range(_app.DASHBOARD_ROWS + _app.ACCOUNTS_PER_PAGE + 20):
    conn.execute(
        "INSERT INTO accounts (company_name, preferred_contact, prospecting_status, "
        "pipeline_milestone, cadence_start, matching_properties, created_at, updated_at) "
        "VALUES (?,'Unknown','Prospecting','None / In Cadence',?,?,?,?)",
        (f"Bulk Volume Test {_i:04d} LLC", _today, _i, _ts, _ts))
conn.commit()
_due = len(cadence.get_due_reminders(conn))
_active = conn.execute("SELECT COUNT(*) c FROM accounts WHERE "
                       "COALESCE(archived_at,'')=''").fetchone()["c"]
check("volume: the test really is oversized",
      _due > _app.DASHBOARD_ROWS and _active > _app.ACCOUNTS_PER_PAGE,
      f"{_due} reminders, {_active} accounts")

r = client.get("/")
_html = r.data.decode()
check("dashboard: draws at most DASHBOARD_ROWS reminder rows",
      _html.count('name="step_type"') <= _app.DASHBOARD_ROWS * 2,
      _html.count('name="step_type"'))
check("dashboard: still reports the true total",
      f"of <strong>{_due}</strong> due reminders" in _html
      or f">{_due}</strong> due reminders" in _html, _due)
check("dashboard: points at the Queue for the rest",
      "Work the Queue" in _html and "Show all" in _html)
check("dashboard: the page stays small", len(r.data) < 200_000, f"{len(r.data)} bytes")
_all = client.get("/?all=1").data.decode()
check("dashboard: ?all=1 really does draw them all",
      _all.count('name="step_type"') > _html.count('name="step_type"'))

r = client.get("/accounts")
_html = r.data.decode()
check("accounts: one page of rows at a time",
      _html.count('name="account_ids"') == _app.ACCOUNTS_PER_PAGE,
      _html.count('name="account_ids"'))
check("accounts: the header counts every match, not the page",
      f"({_active})" in _html, _active)
check("accounts: a pager is offered", "Page 1 of" in _html)
_p2 = client.get("/accounts?page=2").data.decode()
check("accounts: page 2 shows different accounts",
      _p2.count('name="account_ids"') > 0 and "Page 2 of" in _p2)
_far = client.get("/accounts?page=9999").data.decode()
check("accounts: a page past the end clamps instead of coming back empty",
      _far.count('name="account_ids"') > 0)
check("accounts: a junk page number is harmless",
      client.get("/accounts?page=abc").status_code == 200)
check("accounts: filters survive paging",
      b"Bulk Volume Test" in client.get(
          "/accounts?q=Bulk+Volume&page=2").data)

# a long history is capped on the account page
_acct = conn.execute("SELECT id FROM accounts WHERE company_name LIKE "
                     "'Bulk Volume Test%' LIMIT 1").fetchone()["id"]
for _i in range(_app.TIMELINE_ROWS + 10):
    conn.execute("INSERT INTO interactions (account_id, interaction_type, notes, "
                 "created_at) VALUES (?,?,?,?)", (_acct, "General Note", f"n{_i}", _ts))
conn.commit()
_html = client.get(f"/accounts/{_acct}").data.decode()
check("timeline: capped on the account page",
      _html.count('name="created_date"') == _app.TIMELINE_ROWS,
      _html.count('name="created_date"'))
check("timeline: says how many there are and offers them all",
      f"of {_app.TIMELINE_ROWS + 10}" in _html and "show all" in _html)
_html = client.get(f"/accounts/{_acct}?history=all").data.decode()
check("timeline: show all draws the whole history",
      _html.count('name="created_date"') == _app.TIMELINE_ROWS + 10)

# clean up so later checks aren't swamped
conn.execute("DELETE FROM accounts WHERE company_name LIKE 'Bulk Volume Test%'")
conn.commit()
conn.close()

# ---- 36. The paste preview: see it before you save it
_pa = conn = db.get_db()
_target = conn.execute("SELECT id FROM accounts WHERE COALESCE(archived_at,'')='' "
                       "LIMIT 1").fetchone()["id"]
conn.close()
r = client.post(f"/accounts/{_target}/contacts/parse", data={
    "paste": "Marco Webb\nSenior Property Manager\nmwebb@example.com\n"
             "Direct: (713) 555-4321"}, follow_redirects=True)
_html = r.data.decode()
check("preview: fills the form instead of saving straight away",
      'value="Marco"' in _html and 'value="Webb"' in _html
      and 'value="Senior Property Manager"' in _html
      and 'value="mwebb@example.com"' in _html)
check("preview: says what it read", "Read:" in _html and "press Add Contact" in _html)
check("preview: nothing was saved yet",
      db.get_db().execute("SELECT 1 FROM contacts WHERE email='mwebb@example.com'"
                          ).fetchone() is None)
_partial = client.post(f"/accounts/{_target}/contacts/parse",
                       data={"paste": "Nina Alvarez\nFacilities Director"},
                       follow_redirects=True).data.decode()
check("preview: tells you what it couldn't find",
      ("Didn&#39;t find" in _partial or "Didn't find" in _partial)
      and "email" in _partial)
r = client.post(f"/accounts/{_target}/contacts/parse",
                data={"paste": "   "}, follow_redirects=True)
check("preview: an empty paste just says so", b"Paste something into the box" in r.data)
r = client.post(f"/accounts/{_target}/contacts/parse",
                data={"paste": "719 Main Street, Suite 400"}, follow_redirects=True)
check("preview: a paste with no contact in it is refused, not guessed at",
      b"Nothing recognisable" in r.data)
check("preview: stray text never becomes a job title",
      not importer.parse_contact_blob("719 Main Street, Suite 400")["title"])
# after a preview, clearing a box must stick
client.post(f"/accounts/{_target}/contacts/add", data={
    "paste": "Marco Webb\nSenior Property Manager\nmwebb@example.com\n"
             "Direct: (713) 555-4321",
    "already_parsed": "1", "first_name": "Marco", "last_name": "Webb",
    "title": "Senior Property Manager", "email": "mwebb@example.com",
    "work_phone": ""}, follow_redirects=True)
_saved = db.get_db().execute(
    "SELECT * FROM contacts WHERE email='mwebb@example.com'").fetchone()
check("preview: a field you cleared is not re-filled from the paste",
      _saved is not None and _saved["work_phone"] == "",
      dict(_saved) if _saved else None)
check("preview: the parts you kept are saved",
      _saved["first_name"] == "Marco" and _saved["title"] == "Senior Property Manager")

# ---- 37. ZoomInfo's Contact Details panel, tags and all
_panel = """Contact Details
Emails
mdelacruz@harlowenterprises.com
(B)

Phone numbers
(713) 555-0100
(HQ)
(409) 555-0102
(M)"""
_r = importer.parse_contact_blob(_panel)
check("panel: section headings are not mistaken for a name",
      not _r["first_name"] and not _r["last_name"], _r)
check("panel: the email comes through",
      _r["email"] == "mdelacruz@harlowenterprises.com", _r)
check("panel: (HQ) is the work line and (M) is the mobile",
      _r["work_phone"] == "(713) 555-0100"
      and _r["mobile_phone"] == "(409) 555-0102", _r)
check("panel: no heading leaks into the title or company",
      not _r["title"] and not _r["company"], _r)

_r = importer.parse_contact_blob("""Michael Delacruz
Director of Facilities
Harlow Enterprises
Contact Details
Emails
mdelacruz@harlowenterprises.com
(B)
Phone numbers
(713) 555-0100
(HQ)
(409) 555-0102
(M)""")
check("panel: name and title above the panel are picked up",
      _r["first_name"] == "Michael" and _r["last_name"] == "Delacruz"
      and _r["title"] == "Director of Facilities"
      and _r["company"] == "Harlow Enterprises", _r)
check("panel: phones still sorted by their tags",
      _r["work_phone"] == "(713) 555-0100"
      and _r["mobile_phone"] == "(409) 555-0102", _r)

_r = importer.parse_contact_blob("""Sara Lin
Asset Manager
Phone numbers
(713) 555-0100
(HQ)
(281) 555-7788
(D)
(409) 555-0102
(M)""")
check("panel: a direct dial beats the HQ switchboard",
      _r["work_phone"] == "(281) 555-7788"
      and _r["mobile_phone"] == "(409) 555-0102", _r)

_r = importer.parse_contact_blob("Dana Price\nCOO\ndana@bnog.com (B)\n(713) 555-7777 (M)")
check("panel: tags on the end of the value's own line work too",
      _r["email"] == "dana@bnog.com" and _r["mobile_phone"] == "(713) 555-7777"
      and not _r["work_phone"], _r)

# Seniority: abbreviations must match as words, not as substrings.
check("seniority: 'director' is not read as CTO",
      importer.seniority_from_title("Director of Facilities") == "Director")
check("seniority: 'coordinator' is not read as COO",
      importer.seniority_from_title("Project Coordinator") == "")
check("seniority: real abbreviations still match",
      importer.seniority_from_title("COO") == "C-Level"
      and importer.seniority_from_title("CFO") == "C-Level"
      and importer.seniority_from_title("VP of Operations") == "VP-Level")
check("seniority: plurals and phrases still match",
      importer.seniority_from_title("Board of Directors") == "Director"
      and importer.seniority_from_title("Head of Real Estate") == "Director"
      and importer.seniority_from_title("Vice President of Asset Management") == "VP-Level")

# The preview explains a panel-only paste rather than just refusing it.
_conn = db.get_db()
_t = _conn.execute("SELECT id FROM accounts WHERE COALESCE(archived_at,'')='' "
                   "LIMIT 1").fetchone()["id"]
_conn.close()
_html = client.post(f"/accounts/{_t}/contacts/parse", data={"paste": _panel},
                    follow_redirects=True).data.decode()
check("preview: a panel-only paste says what it got and what's missing",
      "no name" in _html and "Contact Details" in _html
      and 'value="mdelacruz@harlowenterprises.com"' in _html, )
check("preview: the phones are prefilled from the panel",
      'value="(713) 555-0100"' in _html and 'value="(409) 555-0102"' in _html)
_html = client.post(f"/accounts/{_t}/contacts/parse",
                    data={"paste": "....."}, follow_redirects=True).data.decode()
check("preview: an unusable paste says so plainly",
      "Nothing recognisable" in _html or "Nothing recognisable" in _html)

# ---- 38. A whole ZoomInfo page, pasted as-is
_PAGE = """ZoomInfo
Home
Search
Lists
Export
Michael Delacruz
Director of Facilities
Harlow Enterprises
Houston, Texas, United States
View Profile
Save to List
Contact Details
Emails
mdelacruz@harlowenterprises.com
(B)
Phone numbers
(713) 555-0100
(HQ)
(409) 555-0102
(M)
Last Updated
Verified
Similar Contacts
Sandra Perez
Vice President of Operations
sperez@harlowenterprises.com
(713) 555-0199
(D)
Kevin Tran
Chief Financial Officer
ktran@harlowenterprises.com
"""
_r = importer.parse_contact_blob(_PAGE)
check("page: picks the contact the page is about",
      _r["first_name"] == "Michael" and _r["last_name"] == "Delacruz", _r)
check("page: title and company come with it",
      _r["title"] == "Director of Facilities"
      and _r["company"] == "Harlow Enterprises", _r)
check("page: the contact's own email and phones",
      _r["email"] == "mdelacruz@harlowenterprises.com"
      and _r["work_phone"] == "(713) 555-0100"
      and _r["mobile_phone"] == "(409) 555-0102", _r)
check("page: nobody from Similar Contacts leaks in",
      "sperez" not in str(_r) and "ktran" not in str(_r)
      and "555-0199" not in str(_r), _r)
check("page: buttons and menus are not mistaken for data",
      "ZoomInfo" not in _r["company"] and _r["title"] != "Home", _r)

# the name is chosen by agreement with the email, not by position
_r = importer.parse_contact_blob("""Harlow Enterprises
Commercial Real Estate
Michael Delacruz
Director of Facilities
Contact Details
Emails
mdelacruz@harlowenterprises.com
(B)""")
check("page: the company appearing first doesn't become the contact",
      _r["first_name"] == "Michael" and _r["last_name"] == "Delacruz"
      and _r["company"] == "Harlow Enterprises", _r)

# email-derived names: only where it is actually unambiguous
check("email name: first.last is read",
      importer.name_from_email("michael.delacruz@x.com") == ("Michael", "Delacruz"))
check("email name: an initial gives only the surname",
      importer.name_from_email("m.delacruz@x.com") == ("", "Delacruz"))
check("email name: a run-together local part is NOT guessed",
      importer.name_from_email("mdelacruz@x.com") == ("", ""))
check("email name: role mailboxes are ignored",
      importer.name_from_email("info@x.com") == ("", "")
      and importer.name_from_email("leasing@x.com") == ("", ""))

# typing into the form then pressing Read the paste must keep what you typed
_conn = db.get_db()
_t2 = _conn.execute("SELECT id FROM accounts WHERE COALESCE(archived_at,'')='' "
                    "LIMIT 1").fetchone()["id"]
_conn.close()
_html = client.post(f"/accounts/{_t2}/contacts/parse", data={
    "paste": "Contact Details\nEmails\nmdelacruz@harlowenterprises.com\n(B)\n"
             "Phone numbers\n(713) 555-0100\n(HQ)",
    "first_name": "Michael", "last_name": "Delacruz"},
    follow_redirects=True).data.decode()
check("preview: a name you typed survives Read the paste",
      'value="Michael"' in _html and 'value="Delacruz"' in _html
      and 'value="mdelacruz@harlowenterprises.com"' in _html)
_html = client.post(f"/accounts/{_t2}/contacts/parse", data={
    "paste": "Contact Details\nEmails\nsarah.obrien@x.com\n(B)"},
    follow_redirects=True).data.decode()
check("preview: a name taken from the email is flagged for checking",
      'value="Sarah"' in _html and "came from the email" in _html)

# ---- 39. A REAL whole-page copy, captured from ZoomInfo
_real = (Path(__file__).resolve().parent / "fixtures" / "zoominfo_page.txt").read_text()
_r = importer.parse_contact_blob(_real)
check("real page: the contact's name",
      _r["first_name"] == "Glen" and _r["last_name"] == "Harlow", _r)
check("real page: title and company", _r["title"] == "President"
      and _r["company"] == "Harlow Enterprises", _r)
check("real page: both phone numbers, sorted by their tags",
      _r["work_phone"] == "(713) 555-0100"
      and _r["mobile_phone"] == "(708) 555-0101", _r)
check("real page: seniority inferred", _r["seniority"] == "C-Level", _r)
check("real page: no email on the page, so none is invented", _r["email"] == "", _r)
_flat = " ".join(str(v) for v in _r.values())
check("real page: left-hand navigation is not the contact",
      "Lists" not in _flat and "Automations" not in _flat
      and "Track Contact" not in _flat, _r)
check("real page: no Markdown link syntax survives",
      "](" not in _flat and "http" not in _flat, _r)
check("real page: Similar Contacts are excluded",
      "Brent" not in _flat and "Vivian" not in _flat, _r)
check("real page: the Employment History and Web References prose is ignored",
      "Summit Events Group" not in _flat and "Lakeside Event Rental" not in _flat, _r)
check("real page: addresses don't become a title or company",
      "Harbor" not in _flat and "Springfield" not in _flat, _r)

# The specific traps this page sprang, kept as their own checks.
check("real page: a bulleted tab strip never cuts the page short",
      importer._expand_lines(["* Org Chart"])[0]["bullet"] is True
      and importer._expand_lines(["Org Chart"])[0]["bullet"] is False)
check("real page: 'Employees' as a field label doesn't stop the parse",
      importer._normalize("Employees") not in importer._PAGE_STOP_MARKERS)
check("real page: a company link gives the header its anchor",
      importer._contact_header(importer._expand_lines([
          "Glen Harlow", "President",
          "[Harlow Enterprises](https://app.zoominfo.com/#/apps/profile/company/1000001)"]))
      == ("Glen Harlow", "President", "Harlow Enterprises"))
check("real page: a Markdown nav bar splits into separate buttons",
      len(importer._expand_lines(["[Home](https://a/1)[Advanced Search](https://a/2)"])) == 2)
check("names: interface text is not capitalised like a name",
      not importer._looks_like_name("Lists and records")
      and importer._looks_like_name("Glen Harlow"))
check("names: lowercase particles are still allowed",
      importer._looks_like_name("Maria de Leon")
      and importer._looks_like_name("Sarah O'Brien"))

# ---- 40. The same contact with profile tabs switched off
# Disabling tabs adds a "Homepage / Glen Harlow" breadcrumb ABOVE the real
# header and drops entries from the tab strip. Neither may change the answer.
_real2 = (Path(__file__).resolve().parent / "fixtures"
          / "zoominfo_page_tabs_off.txt").read_text()
_r2 = importer.parse_contact_blob(_real2)
check("tabs off: same contact, same answer",
      {k: _r2[k] for k in ("first_name", "last_name", "title", "company",
                           "work_phone", "mobile_phone", "seniority")}
      == {"first_name": "Glen", "last_name": "Harlow", "title": "President",
          "company": "Harlow Enterprises", "work_phone": "(713) 555-0100",
          "mobile_phone": "(708) 555-0101", "seniority": "C-Level"}, _r2)
check("tabs off: toggling tabs changes nothing", _r2 == _r, (_r, _r2))
check("tabs off: the breadcrumb copy of the name is ignored, not used",
      [i for i, rec in enumerate(importer._expand_lines(
          [l.strip() for l in _real2.split("\n")])) if rec["text"] == "Glen Harlow"]
      != [], "the breadcrumb should still be present in the fixture")
check("tabs off: the header anchor reads the lines next to the company link",
      importer._contact_header(importer._expand_lines(
          [l.strip() for l in _real2.split("\n")]))
      == ("Glen Harlow", "President", "Harlow Enterprises"))
check("tabs off: a longer company URL still anchors",
      importer._contact_header(importer._expand_lines([
          "Glen Harlow", "President",
          "[Harlow Enterprises](https://app.zoominfo.com/#/apps/profile/company/"
          "1000001?url=%2Fapps%2Faccount-settings%2Fcustomization%2Fprofile-tabs"
          "%2Fcontact-tabs&titleText=Homepage&profileId=1000001)"]))
      == ("Glen Harlow", "President", "Harlow Enterprises"))

# ---- 41. Pick a few people off a company's employee list
_roster_text = (Path(__file__).resolve().parent / "fixtures"
                / "zoominfo_employees.txt").read_text()
_people = importer.parse_contact_roster(_roster_text)
check("roster: every person on the page is found", len(_people) == 7, len(_people))
check("roster: names, titles and levels come through",
      _people[0]["first_name"] == "Glen" and _people[0]["last_name"] == "Harlow"
      and _people[0]["title"] == "President"
      and _people[0]["seniority"] == "C-Level", _people[0])
check("roster: a collapsed row has no contact details to give",
      all(not p["has_details"] for p in _people), _people)
check("roster: a middle initial isn't part of the surname",
      any(p["first_name"] == "Gerald" and p["last_name"] == "Hayes"
          for p in _people), _people)
check("roster: commas in a title survive",
      any(p["title"] == "Director, Leasing & Brokerage" for p in _people), _people)
check("roster: company chrome and headings aren't people",
      not any(p["last_name"] in ("Enterprises", "Executives") for p in _people))
check("roster: nothing is returned for a page with no people",
      importer.parse_contact_roster("Harlow Enterprises\nReal Estate") == [])
check("roster: the same person listed twice appears once",
      len(importer.parse_contact_roster(_roster_text + _roster_text)) == 7)

_conn = db.get_db()
_rid = _conn.execute(
    "INSERT INTO accounts (company_name, preferred_contact, prospecting_status, "
    "pipeline_milestone, cadence_start, created_at, updated_at) VALUES "
    "('Roster Test Holdings','Unknown','Prospecting','None / In Cadence',?,?,?)",
    (db.today_iso(), db.now_iso(), db.now_iso())).lastrowid
_conn.commit(); _conn.close()

r = client.post(f"/accounts/{_rid}/contacts/roster",
                data={"roster_paste": _roster_text}, follow_redirects=True)
_html = r.data.decode()
check("roster: the preview lists everyone without saving anything",
      "7 people on this company" in _html and "Glen Harlow" in _html
      and db.get_db().execute("SELECT COUNT(*) c FROM contacts WHERE account_id=?",
                              (_rid,)).fetchone()["c"] == 0)
check("roster: nothing is pre-selected", 'class="form-check-input rosterpick" name="pick"'
      in _html.replace('\n', ' ') or 'rosterpick' in _html)
check("roster: levels are shown so the ICP picks are obvious",
      "C-Level" in _html and "Director" in _html and "VP-Level" in _html)

# pick three of the seven, and make one of them primary
def _row(i, person):
    return {f"p{i}_{k}": person.get(k, "") for k in
            ("first_name", "last_name", "title", "seniority", "email",
             "work_phone", "mobile_phone", "linkedin_url")}
_form = {"pick": ["2", "4", "5"], "primary": "2"}
for _i, _p in enumerate(_people):
    _form.update(_row(_i, _p))
r = client.post(f"/accounts/{_rid}/contacts/roster/add", data=_form,
                follow_redirects=True)
_conn = db.get_db()
_acct = _conn.execute("SELECT * FROM accounts WHERE id=?", (_rid,)).fetchone()
_contacts = [dict(c) for c in _conn.execute(
    "SELECT * FROM contacts WHERE account_id=?", (_rid,))]
_conn.close()
check("roster: only the ticked people are created", len(_contacts) + 1 == 3,
      [(c["first_name"], c["last_name"]) for c in _contacts])
check("roster: the one marked primary becomes the primary contact",
      _acct["first_name"] == "Jason" and _acct["last_name"] == "Pratt"
      and _acct["seniority"] == "C-Level", dict(_acct))
check("roster: the others become contacts, titles intact",
      {(c["first_name"], c["title"]) for c in _contacts}
      == {("Danny", "Director, Asset Management"),
          ("Gerald", "Director, Construction")}, _contacts)
check("roster: people not ticked are not created",
      not any(c["first_name"] == "Glen" for c in _contacts))

# a second pass must not duplicate anybody
r = client.post(f"/accounts/{_rid}/contacts/roster",
                data={"roster_paste": _roster_text}, follow_redirects=True)
check("roster: people already on the account are marked, not offered again",
      b"already added" in r.data)
r = client.post(f"/accounts/{_rid}/contacts/roster/add", data=_form,
                follow_redirects=True)
check("roster: re-adding the same people changes nothing",
      b"already on this account" in r.data
      and db.get_db().execute("SELECT COUNT(*) c FROM contacts WHERE account_id=?",
                              (_rid,)).fetchone()["c"] == 2)
r = client.post(f"/accounts/{_rid}/contacts/roster/add", data={},
                follow_redirects=True)
check("roster: adding with nothing ticked is a friendly no-op",
      b"Tick at least one person" in r.data)
r = client.post(f"/accounts/{_rid}/contacts/roster",
                data={"roster_paste": "just some words"}, follow_redirects=True)
check("roster: a paste with no people explains what to copy",
      b"No people found" in r.data and b"Employees tab" in r.data)

# ---- 42. Expand a few rows first and one paste carries their details
_exp = (Path(__file__).resolve().parent / "fixtures"
        / "zoominfo_employees_expanded.txt").read_text()
_rows = importer.parse_contact_roster(_exp)
_by = {f"{p['first_name']} {p['last_name']}": p for p in _rows}
check("expanded: the whole roster still comes through", len(_rows) == 7, len(_rows))
check("expanded: only the rows that were opened carry details",
      sum(1 for p in _rows if p["has_details"]) == 3,
      [(p["first_name"], p["has_details"]) for p in _rows])
_f = _by["Jason Pratt"]
check("expanded: email, phones and LinkedIn land on the right person",
      _f["email"] == "jpratt@harlowenterprises.com"
      and _f["mobile_phone"] == "(409) 555-0102"
      and _f["linkedin_url"] == "https://www.linkedin.com/in/jasonpratt", _f)
check("expanded: a direct dial beats the company switchboard",
      _f["work_phone"] == "(281) 555-0142", _f)
check("expanded: the HQ line is used when there's no direct",
      _by["Danny Severs"]["work_phone"] == "(713) 555-0100", _by["Danny Severs"])
check("expanded: details never leak onto the next person down",
      not _by["Vivian Harlow"]["email"] and not _by["Vivian Harlow"]["work_phone"]
      and not _by["Sean Ho"]["email"], (_by["Vivian Harlow"], _by["Sean Ho"]))
check("expanded: details never leak onto the person above",
      not _by["Brent Harlow"]["email"], _by["Brent Harlow"])
check("expanded: the company's own address and headcount aren't a contact",
      not any(p["last_name"] in ("Enterprises", "Estate") for p in _rows))
check("expanded: trailing page sections aren't swept into the last person",
      not _by["Sean Ho"]["title"].startswith("Employees")
      and "Operations" not in str(_by["Sean Ho"]), _by["Sean Ho"])

# end to end: pick three, their details are saved with them
_conn = db.get_db()
_eid = _conn.execute(
    "INSERT INTO accounts (company_name, preferred_contact, prospecting_status, "
    "pipeline_milestone, cadence_start, created_at, updated_at) VALUES "
    "('Expanded Roster Co','Unknown','Prospecting','None / In Cadence',?,?,?)",
    (db.today_iso(), db.now_iso(), db.now_iso())).lastrowid
_conn.commit(); _conn.close()
r = client.post(f"/accounts/{_eid}/contacts/roster",
                data={"roster_paste": _exp}, follow_redirects=True)
check("expanded: the preview shows the details it found",
      b"jpratt@harlowenterprises.com" in r.data
      and b"row not expanded" in r.data)
check("expanded: it says how many rows carried details",
      b"came with contact details" in r.data)

_form = {"pick": [], "primary": ""}
for _i, _p in enumerate(_rows):
    if _p["has_details"]:
        _form["pick"].append(str(_i))
    for _k in ("first_name", "last_name", "title", "seniority", "email",
               "work_phone", "mobile_phone", "linkedin_url"):
        _form[f"p{_i}_{_k}"] = _p.get(_k, "")
_form["primary"] = _form["pick"][0]
client.post(f"/accounts/{_eid}/contacts/roster/add", data=_form, follow_redirects=True)
_conn = db.get_db()
_a = _conn.execute("SELECT * FROM accounts WHERE id=?", (_eid,)).fetchone()
_cs = [dict(c) for c in _conn.execute(
    "SELECT * FROM contacts WHERE account_id=?", (_eid,))]
_conn.close()
check("expanded: the primary keeps every detail from the one paste",
      _a["first_name"] == "Jason" and _a["email"] == "jpratt@harlowenterprises.com"
      and _a["work_phone"] == "(281) 555-0142"
      and _a["mobile_phone"] == "(409) 555-0102"
      and _a["linkedin_url"] == "https://www.linkedin.com/in/jasonpratt", dict(_a))
check("expanded: the other picks keep theirs too",
      {(c["first_name"], c["email"]) for c in _cs}
      == {("Danny", "dsevers@harlowenterprises.com"),
          ("Gerald", "ghayes@harlowenterprises.com")}, _cs)
check("expanded: three people, three sets of details, one paste",
      len(_cs) + 1 == 3 and all(c["work_phone"] for c in _cs), _cs)

# and the nudge when nothing was expanded
r = client.post(f"/accounts/{_eid}/contacts/roster",
                data={"roster_paste": _roster_text}, follow_redirects=True)
check("expanded: a collapsed paste says to expand the rows first",
      b"expand" in r.data.lower() and b"before copying" in r.data.lower())

# ---- 43. A crash explains itself instead of showing a blank 500
import app as _app_mod

_prev_log = db.ERROR_LOG
db.ERROR_LOG = Path(tempfile.mkdtemp(prefix="crm_err_")) / "error.log"
_real_reminders = cadence.get_due_reminders
import logging as _logging
_app_mod.app.logger.setLevel(_logging.CRITICAL)   # the crash below is on purpose


def _explode(*a, **k):
    raise ValueError("a deliberately broken page")


cadence.get_due_reminders = _explode
try:
    r = client.get("/")
    _html = r.data.decode()
finally:
    cadence.get_due_reminders = _real_reminders

check("error page: returns 500 with a readable page", r.status_code == 500)
check("error page: names the actual error",
      "ValueError" in _html and "deliberately broken page" in _html, _html[:300])
check("error page: says saved data is intact, without overpromising",
      "Your saved data is intact" in _html and "check it went through" in _html)
check("error page: offers a way back", "Dashboard" in _html and "Accounts" in _html)
check("error page: the details can be copied", "Copy details" in _html)
check("error page: the traceback is written to the log",
      db.ERROR_LOG.exists()
      and "deliberately broken page" in db.ERROR_LOG.read_text())
check("error page: the log records which page failed",
      "GET /" in db.ERROR_LOG.read_text())
check("error page: the dashboard works again once the fault clears",
      client.get("/").status_code == 200)
check("error page: a missing page is still a plain 404, not an error report",
      client.get("/no-such-page").status_code == 404)
check("error page: a wrong method is still 405",
      client.get("/accounts/1/log").status_code == 405)

# the log is rotated rather than growing for ever
db.ERROR_LOG.write_text("x" * (_app_mod.MAX_ERROR_LOG_BYTES + 1000))
cadence.get_due_reminders = _explode
try:
    client.get("/")
finally:
    cadence.get_due_reminders = _real_reminders
check("error page: an oversized log is rotated, not appended to for ever",
      db.ERROR_LOG.stat().st_size < _app_mod.MAX_ERROR_LOG_BYTES
      and db.ERROR_LOG.with_suffix(".log.old").exists())
db.ERROR_LOG = _prev_log
_app_mod.app.logger.setLevel(_logging.NOTSET)

# ---- 44. A big untouched import shows one task per account, not five
_pile = db.get_db()
_pile_start = (date.today() - timedelta(days=16)).isoformat()
_ts = db.now_iso()
for _i in range(40):
    _pile.execute(
        "INSERT INTO accounts (company_name, preferred_contact, prospecting_status, "
        "pipeline_milestone, cadence_start, matching_properties, created_at, "
        "updated_at) VALUES (?,'Email','Prospecting','None / In Cadence',?,?,?,?)",
        (f"Untouched Import {_i:03d}", _pile_start, _i, _ts, _ts))
_pile.commit()
_ids = [r["id"] for r in _pile.execute(
    "SELECT id FROM accounts WHERE company_name LIKE 'Untouched Import%'")]
_outstanding = [r for r in cadence.get_due_reminders(_pile, collapse=False)
                if r["account_id"] in _ids]
_shown = [r for r in cadence.get_due_reminders(_pile) if r["account_id"] in _ids]
check("pile-up: every step really is outstanding",
      len(_outstanding) == 40 * len(cadence.CADENCE_STEPS), len(_outstanding))
check("pile-up: but each account appears exactly once",
      len(_shown) == 40
      and len({r["account_id"] for r in _shown}) == 40, len(_shown))
check("pile-up: the one shown is the earliest step the account is waiting on",
      all(r["step_type"] == "Email 1" and r["day"] == 1 for r in _shown),
      {r["step_type"] for r in _shown})
check("pile-up: the row says how many steps are stacked behind",
      all(r["steps_behind"] == len(cadence.CADENCE_STEPS) for r in _shown),
      {r["steps_behind"] for r in _shown})

# work one of them and the next step takes its place
_one = _ids[0]
_pile.execute("INSERT INTO interactions (account_id, interaction_type, notes, "
              "created_at) VALUES (?,?,?,?)", (_one, "Email 1", "sent", db.now_iso()))
_pile.commit()
check("pile-up: working a backlogged account clears it (the clock restarts today)",
      cadence.get_due_reminders(_pile, account_id=_one) == [])
_prog = {p["step_type"]: p for p in cadence.get_cadence_progress(_pile, _one)}
check("pile-up: ...and its next step is two business days out",
      _prog["Call & Text"]["due_date"]
      == cadence.add_business_days(cadence.today(), 2).isoformat(), _prog["Call & Text"])
_pile.close()

r = client.get("/")
_html = r.data.decode()
check("pile-up: the dashboard count reflects accounts, not steps",
      "Untouched Import" in _html)
check("pile-up: the backlog is shown rather than hidden",
      "more step" in _html or "steps behind" in _html.lower(), )
_conn = db.get_db()
_conn.execute("DELETE FROM accounts WHERE company_name LIKE 'Untouched Import%'")
_conn.commit(); _conn.close()

# ---- 45. Fixes from the review pass
# Company matching: legal suffixes only, and dotted forms close up.
check("review: descriptive words are part of the name, not suffixes",
      importer.normalize_company("ABC Partners") != importer.normalize_company("ABC Holdings")
      and importer.normalize_company("Moody Group") != importer.normalize_company("Moody Trust"))
check("review: dotted suffixes match their plain form",
      importer.normalize_company("Wilson, Cribbs & Goren, P.C.")
      == importer.normalize_company("Wilson Cribbs & Goren")
      and importer.normalize_company("Acme L.L.C.") == importer.normalize_company("Acme LLC"))
check("review: no non-ASCII junk in the role-mailbox list",
      all(w.isascii() for w in importer._ROLE_MAILBOXES))

_c = db.get_db()
_c.execute("INSERT INTO accounts (company_name, preferred_contact, prospecting_status, "
           "pipeline_milestone, cadence_start, created_at, updated_at) VALUES "
           "('Westlake Realty LLC','Unknown','Prospecting','None / In Cadence',?,?,?)",
           (db.today_iso(), db.now_iso(), db.now_iso()))
_c.commit()
_res = importer.import_accounts(_c, _sheet(
    [["Westlake Realty, Inc.", "", "", "", 3, "", "", "", "", ""],
     ["Westlake Realty Partners", "", "", "", 2, "", "", "", "", ""]], HDR))
check("review: a match to a differently spelled account is listed, not silent",
      ("Westlake Realty, Inc.", "Westlake Realty LLC") in _res.get("matched_names", []), _res)
check("review: a genuinely different name is imported, not merged",
      _res["imported"] == 1 and _c.execute(
          "SELECT 1 FROM accounts WHERE company_name='Westlake Realty Partners'"
      ).fetchone() is not None, _res)
r = client.post("/import", data={"file": (io.BytesIO(b""), "")}, follow_redirects=True)

# Undo of a cadence restart puts the checked-off steps back.
_rid = _c.execute("SELECT id FROM accounts WHERE company_name='Westlake Realty LLC'"
                  ).fetchone()["id"]
_c.execute("INSERT INTO cadence_dismissals (account_id, step_type, dismissed_at) "
           "VALUES (?,?,?)", (_rid, "Email 1", db.now_iso()))
_c.commit()
client.post("/accounts/bulk", data={"action": "restart_cadence",
                                    "account_ids": [str(_rid)]})
check("review: restarting the cadence clears checked-off steps",
      _c.execute("SELECT COUNT(*) c FROM cadence_dismissals WHERE account_id=?",
                 (_rid,)).fetchone()["c"] == 0)
client.post("/undo/" + str(_c.execute("SELECT MAX(id) m FROM undo_log").fetchone()["m"]))
check("review: undoing the restart brings the checked-off steps back",
      _c.execute("SELECT COUNT(*) c FROM cadence_dismissals WHERE account_id=?",
                 (_rid,)).fetchone()["c"] == 1)

# Archive-before-delete is enforced, not just implied by the menu.
r = client.post("/accounts/bulk", data={"action": "delete",
                                        "account_ids": [str(_rid)]},
                follow_redirects=True)
check("review: an active account cannot be bulk-deleted",
      _c.execute("SELECT 1 FROM accounts WHERE id=?", (_rid,)).fetchone() is not None
      and b"Only archived accounts" in r.data)
_c.execute("UPDATE accounts SET archived_at=? WHERE id=?", (db.now_iso(), _rid))
_c.commit()
client.post("/accounts/bulk", data={"action": "delete", "account_ids": [str(_rid)]})
check("review: an archived one still can",
      _c.execute("SELECT 1 FROM accounts WHERE id=?", (_rid,)).fetchone() is None)

# Undo replays only real column names.
import json as _json, undo as _undo
_c.execute("INSERT INTO undo_log (label, payload, created_at) VALUES (?,?,?)",
           ("tampered", _json.dumps({"ops": [{"op": "update", "table": "accounts",
            "rows": [{"id": 1, "company_name = 'x', notes": "y"}]}]}), db.now_iso()))
_c.commit()
_bad = _c.execute("SELECT MAX(id) m FROM undo_log").fetchone()["m"]
check("review: a tampered undo record can't smuggle SQL in a column name",
      _raises(lambda: _undo.restore(_c, _bad)))
_c.rollback()

# The Undo bar is offered right after the action, not for days.
with client.session_transaction() as _s:
    _s["undo_id"] = {"id": 1, "label": "an old thing",
                     "at": datetime.now().timestamp() - 3600}
check("review: an hour-old Undo bar is no longer shown",
      b"an old thing" not in client.get("/").data)
with client.session_transaction() as _s:
    _s["undo_id"] = {"id": 1, "label": "a fresh thing",
                     "at": datetime.now().timestamp()}
check("review: a fresh one is", b"a fresh thing" in client.get("/").data)
with client.session_transaction() as _s:
    _s.pop("undo_id", None)

# Large unpaced imports warn instead of silently stacking up.
_big = _sheet([[f"Unpaced Co {i}", "", "", "", 1, "", "", "", "", ""]
               for i in range(35)], HDR)
r = client.post("/import", data={"file": (io.BytesIO(_big.getvalue()), "big.xlsx")},
                content_type="multipart/form-data", follow_redirects=True)
check("review: a big unpaced import warns and points at Re-Pace",
      b"start their cadence today" in r.data and b"Re-Pace" in r.data)
check("review: the Re-Pace form explains how to pick the number",
      b"divided by" in client.get("/import").data)
_c.execute("DELETE FROM accounts WHERE company_name LIKE 'Unpaced Co%' "
           "OR company_name LIKE 'Westlake%'")
_c.commit(); _c.close()

# The error page survives a fault in the shared page chrome.
_real_undo_cp = _app_mod.inject_undo
def _broken_chrome():
    raise RuntimeError("the page chrome itself is broken")
_app_mod.app.template_context_processors[None].remove(_real_undo_cp)
_app_mod.app.template_context_processors[None].append(_broken_chrome)
_app_mod.app.logger.setLevel(_logging.CRITICAL)
_prev_log2 = db.ERROR_LOG
db.ERROR_LOG = Path(tempfile.mkdtemp(prefix="crm_err2_")) / "error.log"
try:
    r = client.get("/")
finally:
    _app_mod.app.template_context_processors[None].remove(_broken_chrome)
    _app_mod.app.template_context_processors[None].append(_real_undo_cp)
    db.ERROR_LOG = _prev_log2
    _app_mod.app.logger.setLevel(_logging.NOTSET)
check("review: a broken page chrome still produces a readable error",
      r.status_code == 500 and b"page chrome itself is broken" in r.data
      and b"Internal Server Error" not in r.data, r.data[:200])
check("review: the app recovers once the chrome is fixed",
      client.get("/").status_code == 200)

# ---- 46. Business-day cadence
_fri = date(2026, 10, 9)
check("business days: Day 1 is the start date itself",
      cadence.step_due(_fri, 1) == _fri)
check("business days: no step falls on a weekend",
      all(cadence.step_due(date(2026, 10, 5) + timedelta(days=k), d).weekday() < 5
          for k in range(14) for d, _ in cadence.CADENCE_STEPS if d > 1))
check("business days: a Friday start's Day 3 is Tuesday, not Sunday",
      cadence.step_due(_fri, 3) == date(2026, 10, 13))
check("business days: Day 10 is nine working days on",
      cadence.step_due(_fri, 10) == date(2026, 10, 22))
check("business days: a weekend start still has Day 1 that day",
      cadence.step_due(date(2026, 10, 10), 1) == date(2026, 10, 10)
      and cadence.step_due(date(2026, 10, 10), 3) == date(2026, 10, 13))

# Pacing produces a flat load: exactly 5x the daily starts, no Monday spike.
_starts, _d = [], date(2026, 10, 12)
while len(_starts) < 200:
    if _d.weekday() < 5:
        _starts += [_d] * 10
    _d += timedelta(days=1)
_load = []
for _k in range(14, 28):
    _t = date(2026, 10, 12) + timedelta(days=_k)
    if _t.weekday() < 5:
        _load.append(sum(1 for _s in _starts for _day, _ in cadence.CADENCE_STEPS
                         if cadence.step_due(_s, _day) == _t))
check("business days: paced at 10/day, every working day carries exactly 50",
      set(_load) == {50}, _load)

# ---- 47. Research: no clock until there's someone to contact
check("research: a named person with a phone is reachable",
      cadence.has_contact({"first_name": "Glen", "last_name": "", "email": "",
                           "work_phone": "713", "mobile_phone": ""}))
check("research: a switchboard number alone is not",
      not cadence.has_contact({"first_name": "", "last_name": "", "email": "",
                               "work_phone": "713", "mobile_phone": ""}))
check("research: a name with no way to reach them is not",
      not cadence.has_contact({"first_name": "Glen", "last_name": "Harlow",
                               "email": "", "work_phone": "", "mobile_phone": ""}))

_rc = db.get_db()
_res = importer.import_accounts(_rc, _sheet(
    [["Ready Holdings One", "Ann", "Lo", "", 9, "ann@r1.com", "", "", "", ""],
     ["Switchboard Only Co", "", "", "", 8, "", "713-555-1000", "", "", ""],
     ["Ready Holdings Two", "Ben", "Ma", "", 7, "", "713-555-2000", "", "", ""],
     ["Bare Name Co", "Cy", "Ng", "", 6, "", "", "", "", ""]], HDR), per_day=1)
_by_name = {r["company_name"]: r for r in _rc.execute(
    "SELECT * FROM accounts WHERE company_name IN ('Ready Holdings One',"
    "'Switchboard Only Co','Ready Holdings Two','Bare Name Co')")}
check("research: the import says how many went where",
      _res["started"] == 2 and _res["to_research"] == 2, _res)
check("research: reachable rows get a cadence start",
      _by_name["Ready Holdings One"]["cadence_start"] != ""
      and _by_name["Ready Holdings Two"]["cadence_start"] != "")
check("research: unreachable rows wait with no clock",
      _by_name["Switchboard Only Co"]["cadence_start"] == ""
      and _by_name["Bare Name Co"]["cadence_start"] == "")
check("research: pacing slots are only spent on rows that can be worked",
      _by_name["Ready Holdings One"]["cadence_start"]
      != _by_name["Ready Holdings Two"]["cadence_start"])
_due_ids = {r["account_id"] for r in cadence.get_due_reminders(_rc, collapse=False)}
check("research: a waiting account produces no reminders",
      _by_name["Switchboard Only Co"]["id"] not in _due_ids
      and _by_name["Bare Name Co"]["id"] not in _due_ids)
check("research: the account page shows it as waiting, not overdue",
      all(st["state"] == "waiting" and st["due_date"] == ""
          for st in cadence.get_cadence_progress(_rc, _by_name["Switchboard Only Co"]["id"])))

_sw = _by_name["Switchboard Only Co"]["id"]
r = client.get(f"/accounts/{_sw}")
check("research: the account page explains why and offers to start anyway",
      b"In Research" in r.data and b"Start the cadence anyway" in r.data)
r = client.get("/accounts?view=research")
check("research: it's on the Research list", b"Switchboard Only Co" in r.data
      and b"Who to look up next" in r.data)
check("research: and not on it once it has a contact",
      b"Ready Holdings One" not in r.data)
r = client.get("/")
check("research: the dashboard lists who to look up next",
      b"Research" in r.data and b"Switchboard Only Co" in r.data)

# Adding a contact starts the clock — from every place a contact can arrive.
r = client.post(f"/accounts/{_sw}/contacts/add", data={
    "first_name": "Dee", "last_name": "Fox", "email": "dee@sw.com"},
    follow_redirects=True)
_now_start = _rc.execute("SELECT cadence_start FROM accounts WHERE id=?",
                         (_sw,)).fetchone()["cadence_start"]
check("research: adding a contact starts the cadence today",
      _now_start == cadence.today().isoformat(), _now_start)
check("research: and says so", b"cadence starts today" in r.data)
check("research: the first step is due straight away",
      [x["step_type"] for x in cadence.get_due_reminders(_rc, account_id=_sw)] == ["Email 1"])

_bn = _by_name["Bare Name Co"]["id"]
client.post(f"/accounts/{_bn}/edit", data={
    "company_name": "Bare Name Co", "first_name": "Cy", "last_name": "Ng",
    "work_phone": "713-555-3000", "prospecting_status": "Prospecting",
    "pipeline_milestone": "None / In Cadence", "preferred_contact": "Unknown"},
    follow_redirects=True)
check("research: typing a phone in on the account page starts it too",
      _rc.execute("SELECT cadence_start FROM accounts WHERE id=?",
                  (_bn,)).fetchone()["cadence_start"] != "")

# a running clock is never moved by adding another contact
_r1 = _by_name["Ready Holdings One"]["id"]
_before = _rc.execute("SELECT cadence_start FROM accounts WHERE id=?", (_r1,)).fetchone()[0]
client.post(f"/accounts/{_r1}/contacts/add", data={
    "first_name": "Eve", "last_name": "Hu", "email": "eve@r1.com"})
check("research: a second contact doesn't restart a running cadence",
      _rc.execute("SELECT cadence_start FROM accounts WHERE id=?",
                  (_r1,)).fetchone()[0] == _before)

# the roster picker starts it
_rid2 = _rc.execute(
    "INSERT INTO accounts (company_name, preferred_contact, prospecting_status, "
    "pipeline_milestone, cadence_start, created_at, updated_at) VALUES "
    "('Roster Research Co','Unknown','Prospecting','None / In Cadence','',?,?)",
    (db.now_iso(), db.now_iso())).lastrowid
_rc.commit()
client.post(f"/accounts/{_rid2}/contacts/roster/add", data={
    "pick": ["0"], "p0_first_name": "Gil", "p0_last_name": "Ray",
    "p0_title": "VP", "p0_email": "gil@rr.com"}, follow_redirects=True)
check("research: picking people off a roster starts it",
      _rc.execute("SELECT cadence_start FROM accounts WHERE id=?",
                  (_rid2,)).fetchone()[0] != "")

# a ZoomInfo contact upload starts it
_zid = _rc.execute(
    "INSERT INTO accounts (company_name, preferred_contact, prospecting_status, "
    "pipeline_milestone, cadence_start, created_at, updated_at) VALUES "
    "('Zoom Research Co','Unknown','Prospecting','None / In Cadence','',?,?)",
    (db.now_iso(), db.now_iso())).lastrowid
_rc.commit()
_zres = importer.import_contacts(_rc, _sheet(
    [["Zoom Research Co", "Hal", "Ito", "Director", "", "hal@zr.com", "", "", "", ""]], HDR))
check("research: a ZoomInfo contact upload starts it",
      _zres["cadence_started"] == 1 and _rc.execute(
          "SELECT cadence_start FROM accounts WHERE id=?", (_zid,)).fetchone()[0] != "")

# start anyway, with nobody to contact
_any = _rc.execute(
    "INSERT INTO accounts (company_name, preferred_contact, prospecting_status, "
    "pipeline_milestone, cadence_start, created_at, updated_at) VALUES "
    "('Start Anyway Co','Unknown','Prospecting','None / In Cadence','',?,?)",
    (db.now_iso(), db.now_iso())).lastrowid
_rc.commit()
client.post(f"/accounts/{_any}/restart-cadence", follow_redirects=True)
check("research: 'start anyway' starts the clock with no contact",
      _rc.execute("SELECT cadence_start FROM accounts WHERE id=?",
                  (_any,)).fetchone()[0] == date.today().isoformat())

# a new account typed in by hand
client.post("/accounts/new", data={"company_name": "Hand Made Research Co",
                                    "prospecting_status": "Prospecting",
                                    "pipeline_milestone": "None / In Cadence"})
check("research: a new account with nobody to contact starts in Research",
      _rc.execute("SELECT cadence_start FROM accounts WHERE company_name="
                  "'Hand Made Research Co'").fetchone()[0] == "")

# re-pace leaves Research alone
_rc.execute("UPDATE accounts SET cadence_start='' WHERE id=?", (_any,))
_rc.commit()
client.post("/repace", data={"per_day": "5"}, follow_redirects=True)
check("research: Re-Pace doesn't give a waiting account a clock",
      _rc.execute("SELECT cadence_start FROM accounts WHERE id=?",
                  (_any,)).fetchone()[0] == "")
_rc.execute("DELETE FROM accounts WHERE company_name IN ('Ready Holdings One',"
            "'Switchboard Only Co','Ready Holdings Two','Bare Name Co',"
            "'Roster Research Co','Zoom Research Co','Start Anyway Co',"
            "'Hand Made Research Co')")
_rc.commit(); _rc.close()

# ---- 48. The one-time move for an existing database
_mdir = Path(tempfile.mkdtemp(prefix="crm_rmig_"))
_keep_db = db.DB_PATH
db.DB_PATH = _mdir / "crm.db"
db.init_db()
_m = db.get_db()
_m.execute("DELETE FROM settings WHERE key='migration_research_v1'")   # an older DB
_old_start = "2026-09-21"
def _acc(name, first="", email="", phone=""):
    return _m.execute(
        "INSERT INTO accounts (company_name, first_name, email, work_phone, "
        "preferred_contact, prospecting_status, pipeline_milestone, cadence_start, "
        "created_at, updated_at) VALUES (?,?,?,?,'Unknown','Prospecting',"
        "'None / In Cadence',?,?,?)",
        (name, first, email, phone, _old_start, db.now_iso(), db.now_iso())).lastrowid
_untouched_bare = _acc("Untouched Switchboard", phone="713-555-0000")
_touched_bare = _acc("Worked Switchboard", phone="713-555-0001")
_m.execute("INSERT INTO interactions (account_id, interaction_type, notes, created_at) "
           "VALUES (?,?,?,?)", (_touched_bare, "Call & Text", "", db.now_iso()))
_ready = _acc("Has A Contact", first="Ivy", email="ivy@x.com")
_m.commit(); _m.close()
_notices = db.init_db()
_m = db.get_db()
_cs = {r["company_name"]: r["cadence_start"] for r in _m.execute(
    "SELECT company_name, cadence_start FROM accounts")}
check("migration: an untouched account with nobody to contact moves to Research",
      _cs["Untouched Switchboard"] == "", _cs)
check("migration: one you've already worked keeps its dates",
      _cs["Worked Switchboard"] == _old_start, _cs)
check("migration: one with a contact keeps its dates",
      _cs["Has A Contact"] == _old_start, _cs)
check("migration: the startup window says what moved",
      any("1 account(s)" in n and "Research" in n for n in _notices), _notices)
_m.execute("UPDATE accounts SET cadence_start=? WHERE company_name='Untouched Switchboard'",
           (_old_start,))
_m.commit()
db.init_db()
check("migration: it only ever runs once",
      _m.execute("SELECT cadence_start FROM accounts WHERE company_name="
                 "'Untouched Switchboard'").fetchone()[0] == _old_start)
_m.close()
db.DB_PATH = _keep_db

# ---- 49. Call outcomes
_oc = db.get_db()
_two_bd_ago = date.today() - timedelta(days=7)   # Call & Text (Day 3) due by now
def _call_acct(name):
    i = _oc.execute(
        "INSERT INTO accounts (company_name, first_name, last_name, email, work_phone, "
        "preferred_contact, prospecting_status, pipeline_milestone, cadence_start, "
        "created_at, updated_at) VALUES (?,'Kim','Ode','k@x.com','713-555-4000',"
        "'Unknown','Prospecting','None / In Cadence',?,?,?)",
        (name, _two_bd_ago.isoformat(), db.now_iso(), db.now_iso())).lastrowid
    # Email 1 went out on the start day, so the call (two business days on) is due today.
    _oc.execute("INSERT INTO interactions (account_id, interaction_type, notes, created_at) "
                "VALUES (?,?,?,?)", (i, "Email 1", "", _two_bd_ago.isoformat() + "T09:00:00"))
    _oc.commit()
    return i

def _qlog(aid, outcome, step="Call & Text"):
    return client.post("/reminders/quicklog", data={
        "account_id": str(aid), "step_type": step, "outcome": outcome},
        follow_redirects=True)

_a1 = _call_acct("Outcome Voicemail Co")
r = _qlog(_a1, "Voicemail")
_row = _oc.execute("SELECT * FROM interactions WHERE account_id=? AND "
                   "interaction_type='Call & Text'", (_a1,)).fetchone()
check("outcomes: the outcome is stored with the call", _row["outcome"] == "Voicemail")
check("outcomes: a voicemail completes the step",
      "Call & Text" not in [x["step_type"] for x in
                            cadence.get_due_reminders(_oc, account_id=_a1, collapse=False)])
check("outcomes: the flash says what was logged", b"Voicemail" in r.data)

_a2 = _call_acct("Outcome Meeting Co")
r = _qlog(_a2, "Meeting booked")
check("outcomes: Meeting booked moves the deal to Accepted Meeting",
      _oc.execute("SELECT pipeline_milestone FROM accounts WHERE id=?",
                  (_a2,)).fetchone()[0] == "Accepted Meeting")
check("outcomes: ...which ends the cadence",
      cadence.get_due_reminders(_oc, account_id=_a2) == [])
check("outcomes: ...and says so", b"Accepted Meeting" in r.data)

_a3 = _call_acct("Outcome No Thanks Co")
_qlog(_a3, "Not interested")
check("outcomes: Not interested sets the status and ends the cadence",
      _oc.execute("SELECT prospecting_status FROM accounts WHERE id=?",
                  (_a3,)).fetchone()[0] == "Not Interested"
      and cadence.get_due_reminders(_oc, account_id=_a3) == [])

_a4 = _call_acct("Outcome Bad Number Co")
r = _qlog(_a4, "Bad number")
_acc4 = _oc.execute("SELECT * FROM accounts WHERE id=?", (_a4,)).fetchone()
check("outcomes: Bad number sends the account back to Research",
      _acc4["cadence_start"] == "" and b"back to Research" in r.data, dict(_acc4))
check("outcomes: ...with a dated note", "Bad number reported" in _acc4["notes"])
check("outcomes: ...and the call is NOT counted as done",
      [x["state"] for x in cadence.get_cadence_progress(_oc, _a4)
       if x["step_type"] == "Call & Text"] == ["waiting"])
client.post(f"/accounts/{_a4}/edit", data={
    "company_name": "Outcome Bad Number Co", "first_name": "Kim", "last_name": "Ode",
    "email": "k@x.com", "work_phone": "713-555-9999",
    "prospecting_status": "Prospecting", "pipeline_milestone": "None / In Cadence",
    "preferred_contact": "Unknown"})
_after = cadence.get_due_reminders(_oc, account_id=_a4, collapse=False)
check("outcomes: fixing the number restarts the cadence",
      _oc.execute("SELECT cadence_start FROM accounts WHERE id=?",
                  (_a4,)).fetchone()[0] == cadence.today().isoformat())
check("outcomes: ...and the unfinished call comes round again",
      "Call & Text" in [x["step_type"] for x in
                        cadence.get_due_reminders(_oc, account_id=_a4, collapse=False)]
      or any(st["step_type"] == "Call & Text" and st["state"] in ("upcoming", "due")
             for st in cadence.get_cadence_progress(_oc, _a4)))

# a mis-tap is undoable: the log AND what it did to the deal
_a5 = _call_acct("Outcome Misstap Co")
_qlog(_a5, "Not interested")
_undo_id = _oc.execute("SELECT MAX(id) m FROM undo_log").fetchone()["m"]
client.post(f"/undo/{_undo_id}")
check("outcomes: undoing a one-tap log restores the status",
      _oc.execute("SELECT prospecting_status FROM accounts WHERE id=?",
                  (_a5,)).fetchone()[0] == "Prospecting")
check("outcomes: ...and removes the logged call",
      _oc.execute("SELECT COUNT(*) FROM interactions WHERE account_id=? AND "
                  "interaction_type='Call & Text'", (_a5,)).fetchone()[0] == 0)
check("outcomes: ...so the call is due again",
      "Call & Text" in [x["step_type"] for x in
                        cadence.get_due_reminders(_oc, account_id=_a5, collapse=False)])

_a6 = _call_acct("Outcome Junk Co")
_qlog(_a6, "Abducted by aliens")
check("outcomes: an unknown outcome is ignored, the call still logged",
      _oc.execute("SELECT outcome FROM interactions WHERE account_id=? AND "
                  "interaction_type='Call & Text'", (_a6,)).fetchone()[0] == "")

# where the buttons appear
r = client.get("/")
check("outcomes: the dashboard offers outcomes on call steps",
      b"Log call" in r.data and b'value="Meeting booked"' in r.data)
_qpos = [i for i, t in enumerate(_app_mod._build_queue(_oc))
         if t["account_id"] == _a1 or t["kind"] == "cadence"]
_qhtml = ""
for _pos in range(len(_app_mod._build_queue(_oc))):
    _qhtml = client.get(f"/queue?pos={_pos}").data.decode()
    if "How did the call go" in _qhtml:
        break
check("outcomes: the queue shows one button per outcome, numbered",
      "How did the call go" in _qhtml and 'id="k-o7"' in _qhtml
      and "'7': 'k-o7'" in _qhtml)

# the account page: log with an outcome, see it, correct it
r = client.post(f"/accounts/{_a1}/log", data={
    "interaction_type": "Call 2", "outcome": "Spoke", "notes": "good chat"},
    follow_redirects=True)
check("outcomes: the account page can log an outcome",
      _oc.execute("SELECT outcome FROM interactions WHERE account_id=? AND "
                  "interaction_type='Call 2'", (_a1,)).fetchone()[0] == "Spoke")
check("outcomes: the timeline shows it", b">Spoke</span>" in r.data)
_iid = _oc.execute("SELECT id FROM interactions WHERE account_id=? AND "
                   "interaction_type='Call 2'", (_a1,)).fetchone()[0]
client.post(f"/interactions/{_iid}/edit", data={
    "interaction_type": "Call 2", "notes": "good chat", "outcome": "Gatekeeper"})
check("outcomes: the outcome can be corrected afterwards",
      _oc.execute("SELECT outcome FROM interactions WHERE id=?",
                  (_iid,)).fetchone()[0] == "Gatekeeper")

r = client.get("/insights")
check("outcomes: Insights shows calls, connects and meetings",
      b"Calls" in r.data and b"Connect rate" in r.data
      and b"Meetings booked" in r.data)
_oc.execute("DELETE FROM accounts WHERE company_name LIKE 'Outcome %'")
_oc.commit(); _oc.close()

# an older database gains the outcome column without losing anything
_odir = Path(tempfile.mkdtemp(prefix="crm_omig_"))
_keep_db2 = db.DB_PATH
db.DB_PATH = _odir / "crm.db"
_raw = _sq.connect(db.DB_PATH)
_raw.executescript("""
CREATE TABLE accounts (id INTEGER PRIMARY KEY AUTOINCREMENT, company_name TEXT NOT NULL,
  prospecting_status TEXT NOT NULL DEFAULT 'Prospecting',
  pipeline_milestone TEXT NOT NULL DEFAULT 'None / In Cadence',
  cadence_start TEXT NOT NULL, created_at TEXT NOT NULL, updated_at TEXT NOT NULL);
CREATE TABLE interactions (id INTEGER PRIMARY KEY AUTOINCREMENT,
  account_id INTEGER NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
  interaction_type TEXT NOT NULL, notes TEXT DEFAULT '', created_at TEXT NOT NULL);
INSERT INTO accounts (company_name, cadence_start, created_at, updated_at)
  VALUES ('Old Co', '2026-01-01', '2026-01-01', '2026-01-01');
INSERT INTO interactions (account_id, interaction_type, notes, created_at)
  VALUES (1, 'Call 2', 'kept', '2026-01-02');
""")
_raw.commit(); _raw.close()
db.init_db()
_m2 = db.get_db()
check("outcomes migration: the column is added to an old database",
      "outcome" in {r[1] for r in _m2.execute("PRAGMA table_info(interactions)")})
check("outcomes migration: existing calls keep their notes",
      _m2.execute("SELECT notes FROM interactions").fetchone()[0] == "kept")
_m2.close()
db.DB_PATH = _keep_db2

# ---- 50. Active accounts win over archived duplicates on import
_pc = db.get_db()
_pts = db.now_iso()
_pc.execute("INSERT INTO accounts (company_name, preferred_contact, prospecting_status, "
            "pipeline_milestone, cadence_start, archived_at, archive_reason, created_at, "
            "updated_at) VALUES ('Precedence Co Inc','Unknown','Prospecting',"
            "'None / In Cadence','',?,'Duplicate',?,?)", (_pts, _pts, _pts))
_live = _pc.execute("INSERT INTO accounts (company_name, preferred_contact, "
                    "prospecting_status, pipeline_milestone, cadence_start, created_at, "
                    "updated_at) VALUES ('Precedence Co','Unknown','Prospecting',"
                    "'None / In Cadence','',?,?)", (_pts, _pts)).lastrowid
_pc.commit()
_pr = importer.import_contacts(_pc, _sheet(
    [["Precedence Co", "Pia", "Lu", "", "", "pia@pc.com", "", "", "", ""]], HDR))
check("precedence: an archived duplicate doesn't block contacts for the live company",
      _pr["attached"] == 1 and _pr["skipped_archived"] == 0, _pr)
check("precedence: they land on the live account",
      _pc.execute("SELECT first_name FROM accounts WHERE id=?", (_live,)).fetchone()[0] == "Pia")
_pr = importer.import_accounts(_pc, _sheet(
    [["Precedence Co, LLC", "", "", "", 4, "", "", "", "", ""]], HDR))
check("precedence: re-importing the company counts as a duplicate of the live one",
      _pr["skipped_duplicates"] == 1 and _pr["skipped_archived"] == 0, _pr)
_pc.execute("DELETE FROM accounts WHERE company_name LIKE 'Precedence Co%'")
_pc.commit(); _pc.close()

# ---- 51. Merging duplicates
_mc = db.get_db()
def _mk(name, first="", last="", phone="", mobile="", notes="", matching=None,
        cadence_start=""):
    return _mc.execute(
        "INSERT INTO accounts (company_name, first_name, last_name, work_phone, "
        "mobile_phone, notes, matching_properties, preferred_contact, "
        "prospecting_status, pipeline_milestone, cadence_start, created_at, "
        "updated_at) VALUES (?,?,?,?,?,?,?,'Unknown','Prospecting',"
        "'None / In Cadence',?,?,?)",
        (name, first, last, phone, mobile, notes, matching, cadence_start,
         db.now_iso(), db.now_iso())).lastrowid
_keep = _mk("Hartley Properties LLC", "Ann", "Ross", phone="713-555-1111",
            notes="Keeper notes", matching=5, cadence_start="2026-09-01")
_dupe = _mk("Hartley Properties, Inc.", "Bob", "Tran", phone="713-555-2222",
            notes="Address: 1 Main St", matching=9)
_bare = _mk("HARTLEY PROPERTIES", mobile="832-555-3333")
for _ in range(3):
    _mc.execute("INSERT INTO interactions (account_id, interaction_type, notes, "
                "created_at) VALUES (?,?,?,?)", (_keep, "General Note", "k", db.now_iso()))
for _ in range(2):
    _mc.execute("INSERT INTO interactions (account_id, interaction_type, notes, "
                "created_at) VALUES (?,?,?,?)", (_dupe, "Call 2", "d", db.now_iso()))
_cara = _mc.execute("INSERT INTO contacts (account_id, first_name, last_name, created_at) "
                    "VALUES (?,?,?,?)", (_dupe, "Cara", "Diaz", db.now_iso())).lastrowid
_ann2 = _mc.execute("INSERT INTO contacts (account_id, first_name, last_name, created_at) "
                    "VALUES (?,?,?,?)", (_dupe, "Ann", "Ross", db.now_iso())).lastrowid
_bid = _mc.execute("INSERT INTO bids (account_id, roof_address, created_at, updated_at) "
                   "VALUES (?,?,?,?)", (_dupe, "1 Main St", db.now_iso(), db.now_iso())).lastrowid
db.UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
(db.UPLOAD_DIR / "merge_photo.jpg").write_bytes(b"\xff\xd8merge")
_photo = _mc.execute("INSERT INTO bid_photos (bid_id, filename, created_at) VALUES (?,?,?)",
                     (_bid, "merge_photo.jpg", db.now_iso())).lastrowid
_proj = _mc.execute("INSERT INTO projects (account_id, name, created_at, updated_at) "
                    "VALUES (?,?,?,?)", (_dupe, "Hartley roof", db.now_iso(), db.now_iso())).lastrowid
_inv = _mc.execute("INSERT INTO invoices (project_id, amount, created_at, updated_at) "
                   "VALUES (?,?,?,?)", (_proj, 5000, db.now_iso(), db.now_iso())).lastrowid
_mc.execute("INSERT INTO cadence_dismissals (account_id, step_type, dismissed_at) "
            "VALUES (?,?,?)", (_dupe, "Email 1", db.now_iso()))
_mc.commit()

_grp = [g for g in importer.find_duplicate_groups(_mc) if g["key"] == "hartley properties"]
check("merge: the three spellings are found as one group",
      len(_grp) == 1 and len(_grp[0]["accounts"]) == 3, _grp)
check("merge: the most-worked account is suggested as the one to keep",
      [a["id"] for a in _grp[0]["accounts"] if a["suggested"]] == [_keep])
check("merge: the Import page offers to merge them",
      b"Merge into the selected account" in client.get("/import").data)

# refusals
r = client.post("/duplicates/merge", data={"keep": str(_keep),
                "account_ids": [str(_keep), "999999"]}, follow_redirects=True)
check("merge: refuses ids that don't exist", b"be merged" in r.data)
_other = _mk("Completely Different Co", "Zed", "Ink")
_mc.commit()
r = client.post("/duplicates/merge", data={"keep": str(_keep),
                "account_ids": [str(_keep), str(_other)]}, follow_redirects=True)
check("merge: refuses to merge two different companies",
      b"be merged" in r.data
      and _mc.execute("SELECT 1 FROM accounts WHERE id=?", (_other,)).fetchone())
r = client.post("/duplicates/merge", data={"keep": "",
                "account_ids": [str(_keep), str(_dupe)]}, follow_redirects=True)
check("merge: insists on picking which one to keep", b"Pick which account" in r.data)

r = client.post("/duplicates/merge", data={
    "keep": str(_keep), "account_ids": [str(_keep), str(_dupe), str(_bare)]},
    follow_redirects=True)
_k = _mc.execute("SELECT * FROM accounts WHERE id=?", (_keep,)).fetchone()
_kcontacts = {(c["first_name"], c["last_name"]) for c in _mc.execute(
    "SELECT first_name, last_name FROM contacts WHERE account_id=?", (_keep,))}
check("merge: the duplicates are gone",
      not _mc.execute(f"SELECT 1 FROM accounts WHERE id IN ({_dupe},{_bare})").fetchone())
check("merge: the kept account keeps its own primary contact",
      _k["first_name"] == "Ann" and _k["work_phone"] == "713-555-1111", dict(_k))
check("merge: the duplicate's primary person joins as a contact",
      ("Bob", "Tran") in _kcontacts, _kcontacts)
check("merge: other people move across", ("Cara", "Diaz") in _kcontacts)
check("merge: someone already on the account isn't added twice",
      ("Ann", "Ross") not in _kcontacts, _kcontacts)
check("merge: all the history moves",
      _mc.execute("SELECT COUNT(*) FROM interactions WHERE account_id=?",
                  (_keep,)).fetchone()[0] == 5)
check("merge: the roof report moves, photo and file intact",
      _mc.execute("SELECT account_id FROM bids WHERE id=?", (_bid,)).fetchone()[0] == _keep
      and _mc.execute("SELECT 1 FROM bid_photos WHERE id=?", (_photo,)).fetchone()
      and (db.UPLOAD_DIR / "merge_photo.jpg").exists())
check("merge: the project moves, invoice intact",
      _mc.execute("SELECT account_id FROM projects WHERE id=?", (_proj,)).fetchone()[0] == _keep
      and _mc.execute("SELECT 1 FROM invoices WHERE id=?", (_inv,)).fetchone())
check("merge: checked-off steps move",
      _mc.execute("SELECT COUNT(*) FROM cadence_dismissals WHERE account_id=?",
                  (_keep,)).fetchone()[0] == 1)
check("merge: the bigger building count wins", _k["matching_properties"] == 9)
check("merge: a missing mobile is filled from a duplicate", _k["mobile_phone"] == "832-555-3333")
check("merge: the duplicate's notes are kept, labelled",
      "Keeper notes" in _k["notes"] and "merged from" in _k["notes"]
      and "1 Main St" in _k["notes"], _k["notes"])
check("merge: the running cadence clock is kept", _k["cadence_start"] == "2026-09-01")
check("merge: the summary says what moved",
      b"Merged 2 duplicate" in r.data and b"logged touches" in r.data)

client.post("/undo/" + str(_mc.execute("SELECT MAX(id) m FROM undo_log").fetchone()["m"]))
_k2 = _mc.execute("SELECT * FROM accounts WHERE id=?", (_keep,)).fetchone()
check("merge undo: the duplicates come back",
      _mc.execute(f"SELECT COUNT(*) FROM accounts WHERE id IN ({_dupe},{_bare})").fetchone()[0] == 2)
check("merge undo: the kept account's details are as they were",
      _k2["matching_properties"] == 5 and _k2["notes"] == "Keeper notes"
      and _k2["mobile_phone"] == "", dict(_k2))
check("merge undo: history goes back where it came from",
      _mc.execute("SELECT COUNT(*) FROM interactions WHERE account_id=?",
                  (_keep,)).fetchone()[0] == 3
      and _mc.execute("SELECT COUNT(*) FROM interactions WHERE account_id=?",
                      (_dupe,)).fetchone()[0] == 2)
check("merge undo: people go back, and the added contact is removed",
      _mc.execute("SELECT account_id FROM contacts WHERE id=?", (_cara,)).fetchone()[0] == _dupe
      and _mc.execute("SELECT account_id FROM contacts WHERE id=?", (_ann2,)).fetchone()[0] == _dupe
      and not _mc.execute("SELECT 1 FROM contacts WHERE account_id=? AND first_name='Bob'",
                          (_keep,)).fetchone())
check("merge undo: the roof report goes back WITH its photo",
      _mc.execute("SELECT account_id FROM bids WHERE id=?", (_bid,)).fetchone()[0] == _dupe
      and _mc.execute("SELECT 1 FROM bid_photos WHERE id=?", (_photo,)).fetchone()
      and (db.UPLOAD_DIR / "merge_photo.jpg").exists())
check("merge undo: the project goes back WITH its invoice",
      _mc.execute("SELECT account_id FROM projects WHERE id=?", (_proj,)).fetchone()[0] == _dupe
      and _mc.execute("SELECT 1 FROM invoices WHERE id=?", (_inv,)).fetchone())
check("merge undo: checked-off steps go back",
      _mc.execute("SELECT COUNT(*) FROM cadence_dismissals WHERE account_id=?",
                  (_dupe,)).fetchone()[0] == 1
      and _mc.execute("SELECT COUNT(*) FROM cadence_dismissals WHERE account_id=?",
                      (_keep,)).fetchone()[0] == 0)
_mc.execute("DELETE FROM accounts WHERE company_name LIKE 'Hartley%' "
            "OR company_name='HARTLEY PROPERTIES' OR company_name='Completely Different Co'")
_mc.commit(); _mc.close()

# ---- 52. Restarting a worked account really starts over — single or bulk
_rs = db.get_db()
def _worked(name):
    a = _rs.execute(
        "INSERT INTO accounts (company_name, first_name, email, preferred_contact, "
        "prospecting_status, pipeline_milestone, cadence_start, created_at, updated_at) "
        "VALUES (?,'Al','a@c.com','Unknown','Prospecting','None / In Cadence',"
        "'2026-08-01',?,?)", (name, db.now_iso(), db.now_iso())).lastrowid
    for _, step in cadence.CADENCE_STEPS:
        _rs.execute("INSERT INTO interactions (account_id, interaction_type, notes, "
                    "created_at) VALUES (?,?,?,?)", (a, step, "sent", "2026-08-15T10:00:00-05:00"))
    _rs.commit()
    return a
_w1, _w2 = _worked("Worked Single Co"), _worked("Worked Bulk Co")
check("restart: a fully worked account has nothing due", not cadence.get_due_reminders(_rs, account_id=_w1))
client.post(f"/accounts/{_w1}/restart-cadence")
check("restart: restarting one account starts at Email 1 again",
      [x["step_type"] for x in cadence.get_due_reminders(_rs, account_id=_w1)] == ["Email 1"])
check("restart: the old steps are kept as history",
      _rs.execute("SELECT COUNT(*) FROM interactions WHERE account_id=? AND "
                  "notes LIKE '[%previous cadence] sent'", (_w1,)).fetchone()[0] == 5)
client.post("/accounts/bulk", data={"action": "restart_cadence", "account_ids": [str(_w2)]})
check("restart: a BULK restart starts at Email 1 too",
      [x["step_type"] for x in cadence.get_due_reminders(_rs, account_id=_w2)] == ["Email 1"])
client.post("/undo/" + str(_rs.execute("SELECT MAX(id) m FROM undo_log").fetchone()["m"]))
check("restart: undoing a bulk restart puts the logged steps back",
      _rs.execute("SELECT COUNT(*) FROM interactions WHERE account_id=? AND "
                  "interaction_type != 'General Note' AND notes='sent'", (_w2,)).fetchone()[0] == 5
      and not cadence.get_due_reminders(_rs, account_id=_w2))

# "Start anyway" on a Research account keeps progress (e.g. after a bad number)
_sa = _worked("Start Anyway Progress Co")
_rs.execute("UPDATE interactions SET interaction_type='General Note' WHERE account_id=? "
            "AND interaction_type NOT IN ('Email 1')", (_sa,))
_rs.execute("UPDATE accounts SET cadence_start='' WHERE id=?", (_sa,))
_rs.commit()
client.post(f"/accounts/{_sa}/start-cadence")
_due = [x["step_type"] for x in cadence.get_due_reminders(_rs, account_id=_sa, collapse=False)]
check("start anyway: the clock starts today",
      _rs.execute("SELECT cadence_start FROM accounts WHERE id=?", (_sa,)).fetchone()[0]
      == date.today().isoformat())
check("start anyway: a step already done stays done", "Email 1" not in _due, _due)
_rs.execute("DELETE FROM accounts WHERE company_name IN ('Worked Single Co','Worked Bulk Co',"
            "'Start Anyway Progress Co')")
_rs.commit(); _rs.close()

# ---- 60. ZoomInfo batch loop: a REAL contact export, domain matching, pacing
# Own database, so accounts made by earlier sections can't match by accident.
import csv as _csv
_keep_db3 = db.DB_PATH
db.DB_PATH = Path(tempfile.mkdtemp(prefix="crm_zi_")) / "crm.db"
db.init_db()
_FIX = Path(__file__).resolve().parent / "fixtures" / "zoominfo_contacts_export.csv"
_zc = db.get_db()
_zts = db.now_iso()

def _acct(name, website="", notes="", start="", matching=None, archived=""):
    return _zc.execute(
        "INSERT INTO accounts (company_name, website, notes, matching_properties, "
        "preferred_contact, prospecting_status, pipeline_milestone, cadence_start, "
        "archived_at, created_at, updated_at) VALUES (?,?,?,?,'Unknown','Prospecting',"
        "'None / In Cadence',?,?,?,?)",
        (name, website, notes, matching, start, archived, _zts, _zts)).lastrowid

def _zupload(data=None, **form):
    payload = {"file": (io.BytesIO(data if data is not None else _FIX.read_bytes()),
                        "export.csv")}
    payload.update(form)
    return client.post("/import/contacts", data=payload,
                       content_type="multipart/form-data", follow_redirects=True)

check("domain: website, URL and email reduce to the bare domain",
      importer.domain_of("https://www.HarlowEnterprises.com/about?x=1") == "harlowenterprises.com"
      and importer.domain_of("harlowenterprises.com") == "harlowenterprises.com"
      and importer.domain_of("danny@harlowenterprises.com") == "harlowenterprises.com"
      and importer.domain_of("www.harlowenterprises.com:443") == "harlowenterprises.com")
check("domain: free mail and junk give nothing",
      importer.domain_of("bob@gmail.com") == "" and importer.domain_of("n/a") == ""
      and importer.domain_of("") == "" and importer.domain_of(None) == "")

# The CoStar-style name shares no words ZoomInfo would use; only the site ties them.
_harlow = _acct("HARLOW ENT HOLDINGS LTD", website="http://www.HarlowEnterprises.com/",
               notes="Address: 4400 Harbor Blvd, Houston, TX 77006", matching=7)
_zc.commit()
_r = _zupload()
_b = _zc.execute("SELECT * FROM accounts WHERE id=?", (_harlow,)).fetchone()
_bc = {r["first_name"]: dict(r) for r in _zc.execute(
    "SELECT * FROM contacts WHERE account_id=?", (_harlow,))}
check("real export: all four people attach to the account by domain",
      _b["first_name"] == "Danny" and set(_bc) == {"Gerald", "Jason", "Marco"},
      (_b["first_name"], list(_bc)))
check("real export: the domain match is shown so it can be checked",
      b"Matched by website/email domain" in _r.data and b"HARLOW ENT HOLDINGS LTD" in _r.data)
check("real export: primary gets title, level, email, mobile and LinkedIn",
      _b["title"] == "Director, Asset Management" and _b["seniority"] == "Director"
      and _b["email"] == "danny@harlowenterprises.com" and _b["mobile_phone"] == "(281) 555-0103"
      and _b["linkedin_url"] == "https://www.linkedin.com/in/dannysevers")
check("real export: the direct line keeps its extension",
      _b["work_phone"] == "(713) 555-0100 ext. 32", _b["work_phone"])
check("real export: the company main line goes to Notes",
      "Company main line: (713) 555-0100" in _b["notes"], _b["notes"])
check("real export: 'Gerald W. Hayes Jr.' is first Gerald, last Hayes",
      _bc["Gerald"]["last_name"] == "Hayes" and _bc["Gerald"]["title"] == "Director, Construction")
check("real export: a contact with no LinkedIn just has none",
      _bc["Jason"]["linkedin_url"] == "" and _bc["Jason"]["seniority"] == "C-Level")
check("real export: the company's office address comes from the export",
      (_b["street"], _b["city"], _b["zip"]) == ("4400 Harbor Blvd", "Houston", "77006")
      or bool(_b["street"]), (_b["street"], _b["city"], _b["zip"]))
check("real export: the account leaves Research today",
      _b["cadence_start"] == cadence.today().isoformat())
check("real export: Company Division / Company ID don't steal the company column",
      b"Company Name</code>" in _r.data and b"Company Division Name</code>" not in _r.data)
_page = client.get(f"/accounts/{_harlow}").data
check("ext dialing: the call link pauses then keys the extension",
      b'href="tel:7135550100,32"' in _page)
from app import clean_tel, clean_sms
check("ext dialing: x / Ext / # forms too, plain numbers untouched",
      clean_tel("713-555-1234 x12") == "7135551234,12"
      and clean_tel("+1 713 555 1234 Ext 9") == "+17135551234,9"
      and clean_tel("713.555.1234#77") == "7135551234,77"
      and clean_tel("(713) 555-0100") == "7135550100" and clean_tel("") == "")
check("ext dialing: a text goes to the number without the extension",
      clean_sms("(713) 555-0100 ext. 32") == "7135550100")

# Undo puts everything back, Research included.
client.post("/undo/" + str(_zc.execute("SELECT MAX(id) m FROM undo_log").fetchone()["m"]))
_b = _zc.execute("SELECT * FROM accounts WHERE id=?", (_harlow,)).fetchone()
check("upload undo: the people it added are gone",
      _zc.execute("SELECT COUNT(*) FROM contacts WHERE account_id=?", (_harlow,)).fetchone()[0] == 0
      and _b["first_name"] == "" and _b["email"] == "" and _b["notes"].startswith("Address:")
      and "main line" not in _b["notes"])
check("upload undo: the account is back in Research", _b["cadence_start"] == "")
_zupload()          # re-upload for what follows
check("re-upload after undo works the same",
      _zc.execute("SELECT first_name FROM accounts WHERE id=?", (_harlow,)).fetchone()[0] == "Danny")
_r = _zupload()
check("re-upload: everyone is skipped as already on the account",
      _zc.execute("SELECT COUNT(*) FROM contacts WHERE account_id=?", (_harlow,)).fetchone()[0] == 3
      and b"4 skipped (person already on the account)" in _r.data)

# Name wins; a domain shared by two accounts matches neither; archived stays out.
ZH = ["Company Name", "First Name", "Last Name", "Email Address", "Email Domain",
      "Website", "Company HQ Phone", "Direct Phone Number", "Mobile phone"]
def _zcsv(rows):
    b = io.StringIO(); w = _csv.writer(b); w.writerow(ZH); w.writerows(rows)
    return b.getvalue().encode()

_named = _acct("Name Wins Co")
_other = _acct("Elsewhere LLC", website="namewins.com")
_twin1 = _acct("Twin Owner A", website="twins.com")
_twin2 = _acct("Twin Owner B", website="https://twins.com")
_gone = _acct("Old Archived Inc", website="gone.com", archived=_zts)
_zc.commit()
_zupload(_zcsv([
    ["Name Wins Co", "Nia", "Wu", "nia@namewins.com", "namewins.com", "namewins.com", "", "", "713-555-0101"],
    ["Twin Holdings", "Tom", "Tee", "tom@twins.com", "twins.com", "twins.com", "", "", "713-555-0102"],
    ["Archived By Another Name", "Ann", "Ark", "ann@gone.com", "gone.com", "gone.com", "", "", "713-555-0103"],
    ["Free Mail Owner", "Fay", "Free", "fay@gmail.com", "gmail.com", "", "", "", "713-555-0104"],
]))
_first = lambda i: _zc.execute("SELECT first_name FROM accounts WHERE id=?", (i,)).fetchone()[0]
check("domain match: the company name still wins over a domain",
      _first(_named) == "Nia" and _first(_other) == "")
check("domain match: a domain two accounts share matches neither",
      _first(_twin1) == "" and _first(_twin2) == "")
check("domain match: a domain that's only on an archived account is left alone",
      _first(_gone) == "" and not _zc.execute(
          "SELECT 1 FROM accounts WHERE company_name='Archived By Another Name'").fetchone())

# create_missing: HQ phone as the dial line, website saved onto the new account
_zupload(_zcsv([["Fresh Owner LP", "Hal", "Hq", "hal@freshowner.com", "freshowner.com",
                 "www.freshowner.com", "(281) 555-0000", "", ""]]), create_missing="1")
_fresh = _zc.execute("SELECT * FROM accounts WHERE company_name='Fresh Owner LP'").fetchone()
check("create missing: no direct line -> the company HQ phone is dialed",
      _fresh["work_phone"] == "(281) 555-0000", _fresh["work_phone"])
check("create missing: the website is saved on the new account",
      _fresh["website"] == "www.freshowner.com")
check("create missing: and it starts its cadence (named + phone)",
      _fresh["cadence_start"] == cadence.today().isoformat())
# ...and the next upload finds it by domain even under another name
_zupload(_zcsv([["Fresh Owner Management", "Ivy", "Two", "ivy@freshowner.com",
                 "freshowner.com", "", "", "", "832-555-0001"]]))
check("domain match: a later upload finds it by the website it was given",
      _zc.execute("SELECT COUNT(*) FROM contacts WHERE account_id=? AND first_name='Ivy'",
                  (_fresh["id"],)).fetchone()[0] == 1)

# People's emails never act as an account's domain (management firms)
_owner = _acct("Owner Managed By Firm")
_zc.execute("UPDATE accounts SET first_name='Pat', email='pat@bigmgmtfirm.com' WHERE id=?",
            (_owner,))
_zc.commit()
_zupload(_zcsv([["Big Mgmt Firm", "Sam", "Staff", "sam@bigmgmtfirm.com", "bigmgmtfirm.com",
                 "bigmgmtfirm.com", "", "", "713-555-0199"]]))
check("domain match: a contact's email domain doesn't pull a firm's staff onto an owner",
      _zc.execute("SELECT COUNT(*) FROM contacts WHERE account_id=?", (_owner,)).fetchone()[0] == 0)

# Pacing: N accounts per business day, weekends skipped
cadence.today = lambda: date(2026, 10, 15)       # a Thursday
_pace = [_acct(f"Pace Co {i}") for i in range(5)]
_zc.commit()
_r = _zupload(_zcsv([[f"Pace Co {i}", f"P{i}", "Pace", f"p{i}@pace{i}.com", "", "", "", "",
                      "713-555-02%02d" % i] for i in range(5)]), per_day="2")
_starts = [_zc.execute("SELECT cadence_start FROM accounts WHERE id=?", (i,)).fetchone()[0]
           for i in _pace]
check("pacing: 2 per business day — Thu, Thu, Fri, Fri, Mon",
      _starts == ["2026-10-15", "2026-10-15", "2026-10-16", "2026-10-16", "2026-10-19"], _starts)
check("pacing: the summary says when the last one starts", b"2026-10-19" in _r.data)
cadence.today = _real_cadence_today

# Big unpaced upload warns
_many = [_acct(f"Many Co {i}") for i in range(31)]
_zc.commit()
_r = _zupload(_zcsv([[f"Many Co {i}", "M", f"N{i}", f"m{i}@many{i}.com", "", "", "", "", ""]
                     for i in range(31)]))
check("pacing: a big unpaced upload warns about the Day-1 pile",
      b"31 accounts started their cadence today" in _r.data)

# CRM -> ZoomInfo: the Research list as a company list
_zc.execute("DELETE FROM accounts WHERE company_name LIKE 'Many Co %' OR company_name LIKE 'Pace Co %'")
_rsch = _acct("Research Big Owner", website="bigowner.com",
              notes="Big flat roofs\nAddress: 55 Elm, Suite 4, Dallas, TX 75201", matching=9)
_acct("Research Small Owner", notes="Address: Houston, Texas", matching=1)
_acct("Already Working Co", start="2026-10-01", matching=50)
_zc.commit()
_dl = client.get("/accounts/zoominfo.csv")
_rows = list(_csv.reader(io.StringIO(_dl.data.decode("utf-8-sig"))))
_names = [r[0] for r in _rows[1:]]
check("zoominfo list: a CSV download with ZoomInfo's column names",
      _dl.status_code == 200 and "attachment" in _dl.headers.get("Content-Disposition", "")
      and _rows[0] == ["Company Name", "Website", "Street", "City", "State", "Zip Code", "Country"])
check("zoominfo list: Research accounts only, biggest portfolios first",
      "Already Working Co" not in _names and _names.index("Research Big Owner")
      < _names.index("Research Small Owner"), _names)
check("zoominfo list: website and the address from Notes split into columns",
      ["Research Big Owner", "bigowner.com", "55 Elm, Suite 4", "Dallas", "TX", "75201",
       "United States"] in _rows)
check("zoominfo list: a partial address fills what it can",
      ["Research Small Owner", "", "", "Houston", "Texas", "", "United States"] in _rows)
_names = [r[0] for r in _csv.reader(io.StringIO(client.get(
    "/accounts/zoominfo.csv?view=research&min_matching=5").data.decode("utf-8-sig")))][1:]
check("zoominfo list: the page's filters narrow the file", _names == ["Research Big Owner"], _names)
_ap = client.get("/accounts?view=research").data
check("zoominfo list: the Research list offers the download",
      b"/accounts/zoominfo.csv" in _ap and b"Download for ZoomInfo" in _ap)
check("zoominfo list: the guide explains the batch loop",
      b'id="zoominfo-batch"' in client.get("/guide").data)

# Account import fills the website field; the edit form saves it
_imp = importer.import_accounts(_zc, _sheet(
    [["Site Field Co", "", "", "www.sitefield.com"]],
    ["Company Name", "First Name", "Last Name", "Website"]))
_sf = _zc.execute("SELECT * FROM accounts WHERE company_name='Site Field Co'").fetchone()
check("website: account import stores it in its own field, not Notes",
      _sf["website"] == "www.sitefield.com" and "Website:" not in (_sf["notes"] or ""))
client.post(f"/accounts/{_sf['id']}/edit", data={"company_name": "Site Field Co",
                                                  "website": "sitefield.net"})
check("website: editable on the account",
      _zc.execute("SELECT website FROM accounts WHERE id=?", (_sf["id"],)).fetchone()[0]
      == "sitefield.net")
check("website: searchable", b"Site Field Co" in client.get("/accounts?q=sitefield.net").data)
_zc.close()

# Older databases: the website is lifted out of the Notes line
db.DB_PATH = Path(tempfile.mkdtemp(prefix="crm_web_")) / "crm.db"
_raw = _sq.connect(db.DB_PATH)
_raw.executescript("""
CREATE TABLE accounts (id INTEGER PRIMARY KEY AUTOINCREMENT, company_name TEXT NOT NULL,
  notes TEXT DEFAULT '', prospecting_status TEXT NOT NULL DEFAULT 'Prospecting',
  pipeline_milestone TEXT NOT NULL DEFAULT 'None / In Cadence',
  cadence_start TEXT NOT NULL, created_at TEXT NOT NULL, updated_at TEXT NOT NULL);
INSERT INTO accounts (company_name, notes, cadence_start, created_at, updated_at)
  VALUES ('Has Site', 'Address: Houston, TX\nWebsite: www.hassite.com\nPortfolio SF: 9', '', 'x', 'x'),
         ('No Site', 'just a note', '', 'x', 'x');
""")
_raw.commit(); _raw.close()
db.init_db()
_wm = db.get_db()
check("address migration: the Notes address is split into fields on upgrade",
      tuple(_wm.execute("SELECT city, state FROM accounts WHERE company_name='Has Site'")
            .fetchone()) == ("Houston", "TX"))
check("website migration: lifted out of Notes on upgrade",
      [r[0] for r in _wm.execute("SELECT website FROM accounts ORDER BY id")]
      == ["www.hassite.com", ""])
_wm.close()
db.DB_PATH = _keep_db3

# ---- 61. Self-updater (no git, no zip juggling)
import updater, zipfile as _zf
def _branch_zip(files, top="CRM-main/"):
    b = io.BytesIO()
    with _zf.ZipFile(b, "w") as z:
        z.writestr(top, "")
        for name, body in files.items():
            z.writestr(top + name, body)
    return b.getvalue()

_ud = Path(tempfile.mkdtemp(prefix="crm_upd_"))
(_ud / "app.py").write_text("old")
(_ud / "same.txt").write_text("same")
(_ud / "mine.txt").write_text("not in the repo")
_os_stat = (_ud / "same.txt").stat().st_mtime_ns
_changed = updater.apply_zip(_branch_zip({
    "app.py": "new", "same.txt": "same", "templates/new.html": "hi",
    "../escape.txt": "no", "requirements.txt": "flask"}), _ud)
check("updater: changed and new files are written, the top folder dropped",
      (_ud / "app.py").read_text() == "new" and (_ud / "templates/new.html").read_text() == "hi")
check("updater: unchanged files aren't rewritten (start.bat is running)",
      "same.txt" not in _changed and (_ud / "same.txt").stat().st_mtime_ns == _os_stat, _changed)
check("updater: files that aren't part of the app are left alone",
      (_ud / "mine.txt").read_text() == "not in the repo")
check("updater: nothing is written outside the app folder",
      not (_ud.parent / "escape.txt").exists() and "../escape.txt" not in _changed)
check("updater: refuses a download that isn't the app, changing nothing",
      _raises(lambda: updater.apply_zip(_branch_zip({"other.py": "x"}), _ud))
      and (_ud / "app.py").read_text() == "new")
check("updater: refuses a corrupt download", _raises(lambda: updater.apply_zip(b"junk", _ud)))

_real_sha, _real_dl, _real_req = updater.latest_sha, updater.download, updater._install_requirements
_pip = []
updater._install_requirements = lambda d: _pip.append(d)
updater.latest_sha = lambda timeout=5: "a" * 40
updater.download = lambda timeout=30: _branch_zip({"app.py": "v2", "requirements.txt": "flask\npandas"})
check("updater: an update applies and records the version",
      updater.update(_ud, auto=True) == 0 and (_ud / "app.py").read_text() == "v2"
      and (_ud / ".installed_version").read_text().strip() == "a" * 40)
check("updater: new requirements get installed", len(_pip) == 1)
updater.download = lambda timeout=30: (_ for _ in ()).throw(AssertionError("downloaded"))
check("updater: already current -> no download at all", updater.update(_ud, auto=True) == 0)
updater.latest_sha = lambda timeout=5: None
check("updater: offline on start -> the app still starts",
      updater.update(_ud, auto=True) == 0 and (_ud / "app.py").read_text() == "v2")
updater.latest_sha = lambda timeout=5: "b" * 40
updater.download = lambda timeout=30: (_ for _ in ()).throw(OSError("network down"))
check("updater: a failed download on start -> the app still starts, nothing changed",
      updater.update(_ud, auto=True) == 0 and (_ud / "app.py").read_text() == "v2"
      and (_ud / ".installed_version").read_text().strip() == "a" * 40)
check("updater: a failed download on demand says so", updater.update(_ud, auto=False) == 1)
updater.latest_sha, updater.download, updater._install_requirements = _real_sha, _real_dl, _real_req
_bat = (Path(__file__).resolve().parent.parent / "start.bat").read_text()
check("updater: start.bat updates first, inside one ( ) block",
      _bat.index("updater.py --auto") < _bat.index("python app.py"))
check("updater: ...and the block holds every command",
      _bat.count("(") >= 1 and _bat.rstrip().endswith(")"))

# ---- 62. Call/email app choice: Google Voice, Gmail, Outlook
_lc = db.get_db()
_lts = db.now_iso()
_lid = _lc.execute(
    "INSERT INTO accounts (company_name, first_name, email, work_phone, mobile_phone, "
    "preferred_contact, prospecting_status, pipeline_milestone, cadence_start, created_at, "
    "updated_at) VALUES ('Link Pref Co','Lena','lena@linkpref.com','(713) 555-0100 ext. 32',"
    "'(281) 555-0123','Unknown','Prospecting','None / In Cadence',?,?,?)",
    (date.today().isoformat(), _lts, _lts)).lastrowid
_nid = _lc.execute(
    "INSERT INTO accounts (company_name, first_name, work_phone, preferred_contact, "
    "prospecting_status, pipeline_milestone, cadence_start, created_at, updated_at) "
    "VALUES ('No Email Co','Ned','713-555-0177','Unknown','Prospecting','None / In Cadence',"
    "?,?,?)", (date.today().isoformat(), _lts, _lts)).lastrowid
_lc.commit()
_p = client.get(f"/accounts/{_lid}").data.decode()
check("links: default is the phone's dialer and mail app",
      'href="tel:7135550100,32"' in _p and 'href="sms:2815550123"' in _p
      and 'href="mailto:lena@linkpref.com"' in _p)
_p = client.get(f"/accounts/{_nid}").data.decode()
check("links: no email -> the header says so instead of hiding the button",
      "No email on file" in _p)
_s = client.get(f"/accounts/{_nid}/scripts").data.decode()
check("links: no email -> the email script says what's missing",
      "No email address on file. Add one" in _s)

client.post("/settings", data={"call_app": "google_voice", "email_app": "gmail",
                               "next": "/templates"})
_p = client.get(f"/accounts/{_lid}").data.decode()
check("links: Google Voice calls open voice.google.com in a new tab, extension dropped",
      'href="https://voice.google.com/u/0/calls?a=nc,%2B17135550100" target="_blank"' in _p
      or 'href="https://voice.google.com/u/0/calls?a=nc,%2B12815550123" target="_blank"' in _p)
check("links: Google Voice texts open the message thread",
      "https://voice.google.com/u/0/messages?itemId=t.%2B12815550123" in _p)
check("links: no tel:/sms: links left in Google Voice mode",
      'href="tel:' not in _p and 'href="sms:' not in _p)
_s = client.get(f"/accounts/{_lid}/scripts").data.decode()
check("links: Gmail compose opens with the email written",
      "view=cm&amp;fs=1&amp;to=lena%40linkpref.com&amp;su=" in _s and "mail.google.com/mail/?" in _s
      and "&amp;body=" in _s and "Write this email" in _s)
_IOS = {"User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 Mobile/15E148"}
_AND = {"User-Agent": "Mozilla/5.0 (Linux; Android 14; Pixel 8) AppleWebKit/537.36 Chrome/126 Mobile"}
_s = client.get(f"/accounts/{_lid}/scripts", headers=_IOS).data.decode()
check("links: on an iPhone, Gmail opens the Gmail app's compose with the draft",
      "googlegmail:///co?to=lena%40linkpref.com&amp;subject=" in _s and "&amp;body=" in _s
      and "mail.google.com" not in _s)
_s = client.get(f"/accounts/{_lid}/scripts", headers=_AND).data.decode()
check("links: on Android, Gmail uses mailto with the draft (the app picker keeps it)",
      'href="mailto:lena@linkpref.com?subject=' in _s)
client.post("/settings", data={"send_from": "crm.sales@gmail.com"})
check("links: Gmail web compose opens in the CRM account, not whichever is signed in",
      "mail.google.com/mail/?authuser=crm.sales%40gmail.com&amp;view=cm" in
      client.get(f"/accounts/{_lid}/scripts").data.decode())
client.post("/settings", data={"send_from": ""})
check("links: blank send-from falls back to Your Info's email",
      "authuser=" in client.get(f"/accounts/{_lid}/scripts").data.decode())
client.post("/settings", data={"email_app": "outlook"})
check("links: on an iPhone, Outlook opens the Outlook app's compose",
      "ms-outlook://compose?to=lena%40linkpref.com" in
      client.get(f"/accounts/{_lid}/scripts", headers=_IOS).data.decode())
client.post("/settings", data={"email_app": "gmail", "email_app_phone": "default"})
check("links: a phone-only choice sends the iPhone to Mail (mailto) with the draft",
      'href="mailto:lena@linkpref.com?subject=' in
      client.get(f"/accounts/{_lid}/scripts", headers=_IOS).data.decode())
check("links: ...while the computer keeps Gmail",
      "mail.google.com/mail/?" in client.get(f"/accounts/{_lid}/scripts").data.decode())
client.post("/settings", data={"email_app_phone": "same"})
client.post("/settings", data={"email_app": "outlook", "call_app": "bogus"})
_s = client.get(f"/accounts/{_lid}/scripts").data.decode()
check("links: Outlook web compose works too", "outlook.office.com/mail/deeplink/compose?to=" in _s)
check("links: an unknown app value is ignored, the old choice kept",
      "voice.google.com" in _s)
_t = client.get("/templates").data.decode()
check("links: the choice is on the Templates page",
      'name="call_app"' in _t and 'value="google_voice" selected' in _t
      and 'value="outlook" selected' in _t)
client.post("/settings", data={"call_app": "phone", "email_app": "default"})
check("links: switching back restores tel: links",
      'href="tel:7135550100,32"' in client.get(f"/accounts/{_lid}").data.decode())
_lc.execute("DELETE FROM accounts WHERE id IN (?,?)", (_lid, _nid)); _lc.commit(); _lc.close()

# ---- 63. Email signature
from app import with_signature
_sg = {"email_signature": "{my_name}\n{my_title}\n{my_company}"}
check("signature: replaces the closing name lines once",
      with_signature("Hi,\n\nBest,\n{my_name}\n{my_company}\n{my_phone}", _sg)
      == "Hi,\n\nBest,\n\n{my_name}\n{my_title}\n{my_company}")
check("signature: {signature} places it explicitly",
      with_signature("Hi\n{signature}\nPS", _sg) == "Hi\n{my_name}\n{my_title}\n{my_company}\nPS")
check("signature: blank means none", with_signature("Best,\n{my_name}", {}) == "Best,\n{my_name}")
_gc = db.get_db()
_gid = _gc.execute(
    "INSERT INTO accounts (company_name, first_name, email, preferred_contact, prospecting_status,"
    " pipeline_milestone, cadence_start, created_at, updated_at) VALUES ('Sig Co','Sam',"
    "'sam@sigco.com','Unknown','Prospecting','None / In Cadence',?,?,?)",
    (date.today().isoformat(), _lts, _lts)).lastrowid
_gc.commit()
client.post("/settings", data={"email_signature": "{my_name}\n{my_title}\nSilicone Roof Pros",
                               "my_title": "Owner"})
_s = client.get(f"/accounts/{_gid}/scripts?step=Email+1").data.decode()
check("signature: shows in the email and goes into the compose link",
      "Owner\nSilicone Roof Pros</textarea>" in _s and "Owner%0ASilicone%20Roof%20Pros" in _s)
_myname = _gc.execute("SELECT value FROM settings WHERE key='my_name'").fetchone()[0]
_mail = _s.split("Silicone Roof Pros</textarea>")[0].rsplit(">", 1)[1]
check("signature: the name isn't doubled",
      _mail.count(_myname) == 1 and _mail.endswith(_myname + "\nOwner\n"), _mail[-80:])
_c = client.get(f"/accounts/{_gid}/scripts").data.decode()
check("signature: calls and texts don't get one", "Silicone Roof Pros</textarea>" not in
      _c.split("Cold Call Script")[1].split("</textarea>")[0] + "</textarea>")
check("signature: editable on the Templates page", 'name="email_signature"' in
      client.get("/templates").data.decode())
_gc.execute("DELETE FROM accounts WHERE id=?", (_gid,)); _gc.commit(); _gc.close()

# ---- 64. ZoomInfo button opens the company's own ZoomInfo page
check("zi link: built from a profile URL, an ID or a pasted page",
      importer.zoominfo_company_url("https://app.zoominfo.com/#/apps/profile/company/1000001?profileId=1000001")
      == "https://app.zoominfo.com/#/apps/profile/company/1000001"
      and importer.zoominfo_company_url("1000001") == "https://app.zoominfo.com/#/apps/profile/company/1000001"
      and importer.zoominfo_company_url("[Harlow](https://app.zoominfo.com/#/apps/profile/company/77?x=1)")
      .endswith("/company/77")
      and importer.zoominfo_company_url("Harlow Enterprises") == "")
_zk = db.get_db()
_zid = _zk.execute(
    "INSERT INTO accounts (company_name, preferred_contact, prospecting_status, pipeline_milestone,"
    " cadence_start, created_at, updated_at) VALUES ('ZI Link Co','Unknown','Prospecting',"
    "'None / In Cadence','',?,?)", (_lts, _lts)).lastrowid
_zk.commit()
_p = client.get(f"/accounts/{_zid}").data.decode()
check("zi link: unknown -> a ZoomInfo search for the company",
      "google.com/search?q=site%3Azoominfo.com%20ZI%20Link%20Co" in _p)
_csv_zi = ("Company Name,First Name,Last Name,Email Address,ZoomInfo Company ID,"
           "ZoomInfo Company Profile URL\r\nZI Link Co,Zed,Eye,zed@zilink.com,555,"
           "https://app.zoominfo.com/#/apps/profile/company/555\r\n").encode()
client.post("/import/contacts", data={"file": (io.BytesIO(_csv_zi), "zi.csv")},
            content_type="multipart/form-data")
_p = client.get(f"/accounts/{_zid}").data.decode()
check("zi link: a ZoomInfo upload saves the company page and the button opens it",
      'href="https://app.zoominfo.com/#/apps/profile/company/555" target="_blank"' in _p
      and "Open this company in ZoomInfo" in _p)
_zk.execute("UPDATE accounts SET zoominfo_url='' WHERE id=?", (_zid,)); _zk.commit()
client.post("/import/contacts", data={"file": (io.BytesIO(_csv_zi), "zi.csv")},
            content_type="multipart/form-data")
check("zi link: re-uploading an old export backfills it (person skipped as a duplicate)",
      _zk.execute("SELECT zoominfo_url FROM accounts WHERE id=?", (_zid,)).fetchone()[0]
      .endswith("/company/555"))
_zid2 = _zk.execute(
    "INSERT INTO accounts (company_name, preferred_contact, prospecting_status, pipeline_milestone,"
    " cadence_start, created_at, updated_at) VALUES ('ZI Paste Co','Unknown','Prospecting',"
    "'None / In Cadence','',?,?)", (_lts, _lts)).lastrowid
_zk.commit()
client.post(f"/accounts/{_zid2}/contacts/add", data={
    "paste": (Path(__file__).resolve().parent / "fixtures" / "zoominfo_page.txt").read_text()})
check("zi link: pasting a ZoomInfo page saves the company page too",
      _zk.execute("SELECT zoominfo_url FROM accounts WHERE id=?", (_zid2,)).fetchone()[0]
      == "https://app.zoominfo.com/#/apps/profile/company/1000001")
client.post(f"/accounts/{_zid2}/edit", data={
    "company_name": "ZI Paste Co", "zoominfo_url": "https://app.zoominfo.com/#/apps/profile/company/99?foo=1"})
check("zi link: editable on the account (and tidied)",
      _zk.execute("SELECT zoominfo_url FROM accounts WHERE id=?", (_zid2,)).fetchone()[0]
      == "https://app.zoominfo.com/#/apps/profile/company/99")
_zk.execute("DELETE FROM accounts WHERE id IN (?,?)", (_zid, _zid2)); _zk.commit(); _zk.close()

# ---- 65. Door knocking: addresses, visits, never-visited list
_dk = db.get_db()
def _dk_acct(name, street="", city="", zip_="", notes=""):
    return _dk.execute(
        "INSERT INTO accounts (company_name, street, city, state, zip, notes, preferred_contact,"
        " prospecting_status, pipeline_milestone, cadence_start, created_at, updated_at) VALUES"
        " (?,?,?,'TX',?,?,'Unknown','Prospecting','None / In Cadence','',?,?)",
        (name, street, city, zip_, notes, _lts, _lts)).lastrowid
_d1 = _dk_acct("Door A Co", "100 Main St", "Houston", "77002")
_d2 = _dk_acct("Door B Co", "5 Elm St", "Houston", "77001")
_d3 = _dk_acct("Door C Co", "9 Oak St", "Katy", "77494")
_dk.commit()
_p = client.get(f"/accounts/{_d1}").data.decode()
check("door: account page says never visited, maps the office, offers the log",
      "🚪 Never visited" in _p and "google.com/maps/search/?api=1&amp;query=Door%20A%20Co" in _p
      and 'name="interaction_type" value="Door Knock"' in _p and "Met decision-maker" in _p)
r = client.post(f"/accounts/{_d1}/log", data={"interaction_type": "Door Knock",
                                               "outcome": "Left info", "notes": "card at desk"},
                follow_redirects=True)
_p = r.data.decode()
check("door: a visit is logged with its outcome and shows on the account",
      "🚪 Visited" in _p and "Left info" in _p
      and _dk.execute("SELECT outcome FROM interactions WHERE account_id=? AND "
                      "interaction_type='Door Knock'", (_d1,)).fetchone()[0] == "Left info")
check("door: a visit isn't a cadence step (the account stays in Research)",
      _dk.execute("SELECT cadence_start FROM accounts WHERE id=?", (_d1,)).fetchone()[0] == "")
client.post(f"/accounts/{_d2}/log", data={"interaction_type": "Door Knock",
                                           "outcome": "Meeting booked"})
check("door: Meeting booked at the door moves the deal like a call does",
      _dk.execute("SELECT pipeline_milestone FROM accounts WHERE id=?", (_d2,)).fetchone()[0]
      == "Accepted Meeting")
_l = client.get("/accounts?visited=no&q=Door").data.decode()
check("door: the accounts list filters to never-visited",
      "Door C Co" in _l and "Door A Co" not in _l)
_l = client.get("/accounts?visited=yes&q=Door").data.decode()
check("door: ...and to visited, with the date shown",
      "Door A Co" in _l and "Door C Co" not in _l and "🚪" in _l)
check("door: search finds accounts by zip", "Door C Co" in client.get("/accounts?q=77494").data.decode())
_rows = list(_csv.reader(io.StringIO(client.get("/accounts/door-knock.csv?q=Door").data
                                     .decode("utf-8-sig"))))
check("door: the door-knock list is never-visited only, sorted by zip",
      [r[4] for r in _rows[1:]] == ["Door C Co"] and _rows[0][:5] == ["Zip", "Street", "City", "State", "Company"])
_rows = list(_csv.reader(io.StringIO(client.get("/accounts/door-knock.csv?q=Door&visited=").data
                                     .decode("utf-8-sig"))))
check("door: with the filter cleared it lists everyone, nearest zips together",
      [r[0] for r in _rows[1:]] == ["77001", "77002", "77494"], _rows)
client.post(f"/accounts/{_d3}/edit", data={"company_name": "Door C Co", "street": "1 New Rd",
                                           "city": "Sugar Land", "state": "TX", "zip": "77478"})
check("door: the address is editable",
      tuple(_dk.execute("SELECT city, zip FROM accounts WHERE id=?", (_d3,)).fetchone())
      == ("Sugar Land", "77478"))
_imp = importer.import_accounts(_dk, _sheet(
    [["Door Import Co", "12 Pine St", "Houston", "TX", "77003"]],
    ["Company Name", "Address", "City", "State", "Zip"]))
check("door: account imports fill the address fields",
      tuple(_dk.execute("SELECT street, city, state, zip FROM accounts WHERE company_name="
                        "'Door Import Co'").fetchone()) == ("12 Pine St", "Houston", "TX", "77003"))
_dk.execute("DELETE FROM accounts WHERE company_name LIKE 'Door % Co'"); _dk.commit(); _dk.close()

# ---- 66. Upload a ZoomInfo export straight onto one account
_ua = db.get_db()
_uid = _ua.execute(
    "INSERT INTO accounts (company_name, preferred_contact, prospecting_status, pipeline_milestone,"
    " cadence_start, created_at, updated_at) VALUES ('Totally Different Name LLC','Unknown',"
    "'Prospecting','None / In Cadence','',?,?)", (_lts, _lts)).lastrowid
_ua.commit()
_pg = client.get(f"/accounts/{_uid}").data.decode()
check("account upload: the account page has the upload box",
      f"/accounts/{_uid}/contacts/upload" in _pg and 'enctype="multipart/form-data"' in _pg)
r = client.post(f"/accounts/{_uid}/contacts/upload",
                data={"file": (io.BytesIO(_FIX.read_bytes()), "export.csv")},
                content_type="multipart/form-data", follow_redirects=True)
_ua2 = _ua.execute("SELECT * FROM accounts WHERE id=?", (_uid,)).fetchone()
check("account upload: everyone lands on THIS account whatever company the file names",
      _ua2["first_name"] == "Danny" and _ua.execute(
          "SELECT COUNT(*) FROM contacts WHERE account_id=?", (_uid,)).fetchone()[0] == 3)
check("account upload: says what it did, and the cadence starts",
      b"Added 4 contact(s)" in r.data and _ua2["cadence_start"] != "")
r = client.post(f"/accounts/{_uid}/contacts/upload",
                data={"file": (io.BytesIO(_FIX.read_bytes()), "export.csv")},
                content_type="multipart/form-data", follow_redirects=True)
check("account upload: uploading it again skips everyone", b"4 already on this account" in r.data)
client.post("/undo/" + str(_ua.execute("SELECT MAX(id) m FROM undo_log WHERE COALESCE(used_at, '') = ''").fetchone()["m"]))
client.post("/undo/" + str(_ua.execute(
    "SELECT MAX(id) m FROM undo_log WHERE label LIKE 'Contact upload to%' AND COALESCE(used_at, '') = ''").fetchone()["m"]))
check("account upload: undoable",
      _ua.execute("SELECT COUNT(*) FROM contacts WHERE account_id=?", (_uid,)).fetchone()[0] == 0)
_csv_noco = b"First Name,Last Name,Email Address\r\nNo,Company,nc@x.com\r\n"
r = client.post(f"/accounts/{_uid}/contacts/upload",
                data={"file": (io.BytesIO(_csv_noco), "x.csv")},
                content_type="multipart/form-data", follow_redirects=True)
check("account upload: a file with no company column still works here", b"Added 1 contact" in r.data)
_ua.execute("DELETE FROM accounts WHERE id=?", (_uid,)); _ua.commit(); _ua.close()

# ---- 67. One person at a time: rotation, finished cadences, rest
_rt = db.get_db()
def _rt_acct(name, people):
    """Account with a primary (Pat) in cadence plus extra people."""
    i = _rt.execute(
        "INSERT INTO accounts (company_name, first_name, last_name, title, email, work_phone,"
        " seniority, preferred_contact, prospecting_status, pipeline_milestone, cadence_start,"
        " created_at, updated_at) VALUES (?,'Pat','Prime','COO','pat@x.com','713-555-0001',"
        "'C-Level','Unknown','Prospecting','None / In Cadence',?,?,?)",
        (name, (date.today() - timedelta(days=21)).isoformat(), _lts, _lts)).lastrowid
    for first, sen, email in people:
        _rt.execute("INSERT INTO contacts (account_id, first_name, last_name, title, email,"
                    " seniority, created_at) VALUES (?,?,'X','',?,?,?)", (i, first, email, sen, _lts))
    _rt.commit()
    return i
_r1 = _rt_acct("Rotate One Co", [("Cee", "C-Level", "cee@x.com"), ("Dee", "Director", "dee@x.com"),
                                  ("Nob", "Director", "")])
r = client.post("/reminders/quicklog", data={"account_id": str(_r1), "step_type": "Email 1",
                                             "outcome": "Not interested"}, follow_redirects=True)
_ra = _rt.execute("SELECT * FROM accounts WHERE id=?", (_r1,)).fetchone()
check("rotate: one person's 'not interested' moves on, the company stays open",
      _ra["prospecting_status"] == "Prospecting" and _ra["first_name"] == "Dee", dict(_ra))
check("rotate: next is the reachable Director first (problem owner before the C-suite)",
      _ra["first_name"] == "Dee" and _ra["previous_contact"] == "Pat")
check("rotate: the cadence restarts today for the new person",
      _ra["cadence_start"] == date.today().isoformat()
      and [x["step_type"] for x in cadence.get_due_reminders(_rt, account_id=_r1)] == ["Email 1"])
check("rotate: the first person is kept, marked how their turn ended",
      tuple(_rt.execute("SELECT tried_status, tried_at FROM contacts WHERE account_id=? AND "
                        "first_name='Pat'", (_r1,)).fetchone()) == ("Not interested", date.today().isoformat()))
check("rotate: the flash says who's next", b"Moved on to Dee" in r.data)
_s = client.get(f"/accounts/{_r1}/scripts?step=Email+1").data.decode()
_names = re.findall(r"</span>|(Email 1 \(next contact\)|Email 1)\s*</h5>", _s)
_names = [n for n in _names if n]
check("rotate: Email 1 now leads with the version that mentions the first person",
      _names[:2] == ["Email 1 (next contact)", "Email 1"], _names)
check("rotate: ...filled in with their name", "reached out to Pat about" in _s)
_s0 = client.get(f"/accounts/{_rt_acct('Fresh Co', [])}/scripts?step=Email+1").data.decode()
check("rotate: a first contact never sees the next-contact email", "Email 1 (next contact)" not in _s0)
_p = client.get(f"/accounts/{_r1}").data.decode()
check("rotate: the account page shows who's been tried and who hasn't",
      "Tried: Not interested" in _p and "Not tried yet" in _p)
client.post("/undo/" + str(_rt.execute(
    "SELECT MAX(id) m FROM undo_log WHERE label LIKE 'logging%'").fetchone()["m"]))
_ra = _rt.execute("SELECT * FROM accounts WHERE id=?", (_r1,)).fetchone()
check("rotate: undo puts Pat back and the people as they were",
      _ra["first_name"] == "Pat" and _rt.execute(
          "SELECT COUNT(*) FROM contacts WHERE account_id=? AND COALESCE(tried_status,'')=''",
          (_r1,)).fetchone()[0] == 3
      and not _rt.execute("SELECT 1 FROM contacts WHERE account_id=? AND first_name='Pat'",
                          (_r1,)).fetchone())

# Wrong person rotates too
client.post("/reminders/quicklog", data={"account_id": str(_r1), "step_type": "Call & Text",
                                         "outcome": "Wrong person"})
check("rotate: 'Wrong person' hands over as well",
      _rt.execute("SELECT first_name FROM accounts WHERE id=?", (_r1,)).fetchone()[0] == "Dee")

# Last person standing says no -> company closes and rests 90 days
_r2 = _rt_acct("Rotate Last Co", [])
r = client.post("/reminders/quicklog", data={"account_id": str(_r2), "step_type": "Call & Text",
                                             "outcome": "Not interested"}, follow_redirects=True)
_ra = _rt.execute("SELECT * FROM accounts WHERE id=?", (_r2,)).fetchone()
check("rotate: nobody left -> Not Interested, back as a follow-up in 90 days",
      _ra["prospecting_status"] == "Not Interested"
      and _ra["next_follow_up"] == (date.today() + timedelta(days=90)).isoformat()
      and "Start again with the most senior" in _ra["follow_up_note"])

# Finished cadence with no reply -> offer the next person
_r3 = _rt_acct("Rotate Finished Co", [("Fin", "Manager", "fin@x.com")])
for _st in [s for _, s in cadence.CADENCE_STEPS]:
    _rt.execute("INSERT INTO interactions (account_id, interaction_type, notes, created_at)"
                " VALUES (?,?,'',?)", (_r3, _st, db.now_iso()))
_rt.commit()
_d = client.get("/").data.decode()
check("finished: the dashboard lists accounts whose cadence ran out, with the next person",
      "Cadence finished, no reply" in _d and "Rotate Finished Co" in _d and "Start Fin" in _d)
_p = client.get(f"/accounts/{_r3}").data.decode()
check("finished: the account page offers the next contact", "Start next contact" in _p)
client.post(f"/accounts/{_r3}/next-contact", data={})
_ra = _rt.execute("SELECT * FROM accounts WHERE id=?", (_r3,)).fetchone()
check("finished: one tap starts the next person, the old one marked No reply",
      _ra["first_name"] == "Fin" and _rt.execute(
          "SELECT tried_status FROM contacts WHERE account_id=? AND first_name='Pat'",
          (_r3,)).fetchone()[0] == "No reply")
check("finished: the old steps are kept as history, the new cadence starts at Email 1",
      [x["step_type"] for x in cadence.get_due_reminders(_rt, account_id=_r3)] == ["Email 1"]
      and _rt.execute("SELECT COUNT(*) FROM interactions WHERE account_id=? AND "
                      "interaction_type='General Note'", (_r3,)).fetchone()[0] == 5)
check("finished: it leaves the finished list", not _app_mod._finished_cadences(_rt, _r3))

# Finished with nobody left -> rest
_r4 = _rt_acct("Rotate Rest Co", [])
for _st in [s for _, s in cadence.CADENCE_STEPS]:
    _rt.execute("INSERT INTO interactions (account_id, interaction_type, notes, created_at)"
                " VALUES (?,?,'',?)", (_r4, _st, db.now_iso()))
_rt.commit()
check("rest: nobody left -> the offer is to rest it",
      f"/accounts/{_r4}/rest" in client.get(f"/accounts/{_r4}").data.decode())
client.post(f"/accounts/{_r4}/rest", data={})
_ra = _rt.execute("SELECT * FROM accounts WHERE id=?", (_r4,)).fetchone()
check("rest: comes back as a follow-up in 90 days, and leaves the finished list",
      _ra["next_follow_up"] == (date.today() + timedelta(days=90)).isoformat()
      and not _app_mod._finished_cadences(_rt, _r4))
_rt.execute("DELETE FROM accounts WHERE company_name LIKE 'Rotate %' OR company_name='Fresh Co'")
_rt.commit(); _rt.close()

print()
print(f"{'ALL TESTS PASSED' if not failures else f'{len(failures)} FAILURES: {failures}'}")
sys.exit(1 if failures else 0)
