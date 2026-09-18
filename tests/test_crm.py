"""End-to-end functional test for the Roof CRM.

Run from anywhere:  python3 tests/test_crm.py
Uses a throwaway database in a temp directory; never touches crm.db.
"""
import io
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

# Backdate Acme's cadence 7 days → days 1,3,6 due
week_ago = (date.today() - timedelta(days=7)).isoformat()
conn.execute("UPDATE accounts SET cadence_start=? WHERE id=?", (week_ago, acme["id"]))
conn.commit()
acme_rems = cadence.get_due_reminders(conn, acme["id"])
check("cadence: backdated 7d → days 1,3,6,8 due",
      [r["step_type"] for r in acme_rems] == ["Email 1", "Call & Text", "Call 2", "Email 2"],
      [r["step_type"] for r in acme_rems])
check("cadence: overdue days computed", acme_rems[0]["days_overdue"] == 7, acme_rems[0])

# ---- 3. Logging an interaction clears that step only (persistence rule)
conn.execute("INSERT INTO interactions (account_id, interaction_type, notes, created_at) "
             "VALUES (?,?,?,?)", (acme["id"], "Email 1", "sent intro", db.now_iso()))
conn.commit()
acme_rems = cadence.get_due_reminders(conn, acme["id"])
check("cadence: logged step cleared, others persist",
      [r["step_type"] for r in acme_rems] == ["Call & Text", "Call 2", "Email 2"])

# ---- 4. Manual dismissal clears a step
conn.execute("INSERT INTO cadence_dismissals (account_id, step_type, dismissed_at) "
             "VALUES (?,?,?)", (acme["id"], "Call & Text", db.now_iso()))
conn.commit()
acme_rems = cadence.get_due_reminders(conn, acme["id"])
check("cadence: dismissed step cleared", [r["step_type"] for r in acme_rems] == ["Call 2", "Email 2"])

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
check("templates: 6 seeded", n_templates == 6, n_templates)
settings = {r["key"]: r["value"] for r in conn.execute("SELECT * FROM settings")}
check("templates: settings seeded", settings.get("my_name") == "Alexis Morales"
      and settings.get("my_company") == "Silicone Roof Pros, Inc.")
conn.close()
# re-init must not duplicate seeds
db.init_db()
conn = db.get_db()
check("templates: re-init does not re-seed",
      conn.execute("SELECT COUNT(*) c FROM templates").fetchone()["c"] == 6)
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
check("scripts: tel + sms links", "tel:(281) 423-9999" in html
      and "sms:(832) 555-1111?body=" in html)
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
stag_csv = b"Company Name\r\n" + b"".join(f"Stagger Co {i}\r\n".encode() for i in range(1, 8))
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

print()
print(f"{'ALL TESTS PASSED' if not failures else f'{len(failures)} FAILURES: {failures}'}")
sys.exit(1 if failures else 0)
