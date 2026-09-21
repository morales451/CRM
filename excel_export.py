"""Full-CRM Excel export — one workbook, one sheet per part of the business.

Sheets: Summary, Accounts, Contacts, Interactions, Roof Reports, Projects,
Invoices, Tasks Due. Every sheet gets a frozen, filterable header row and
sensible column widths so the whole file is scannable without formatting.
"""

import io
from datetime import date, datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import cadence
import warranty_calc

HEADER_FILL = PatternFill("solid", fgColor="1D4E89")
HEADER_FONT = Font(color="FFFFFF", bold=True)
TITLE_FONT = Font(bold=True, size=14, color="1D4E89")
SECTION_FONT = Font(bold=True, color="1D4E89")
MONEY = '"$"#,##0.00'
MONEY0 = '"$"#,##0'
DATEFMT = "yyyy-mm-dd"


def _clean(value):
    """Excel-safe value; ISO timestamps become real dates."""
    if value is None:
        return ""
    if isinstance(value, str) and len(value) >= 10:
        head = value[:10]
        if len(head) == 10 and head[4] == "-" and head[7] == "-":
            try:
                return date.fromisoformat(head)
            except ValueError:
                return value
    return value


def _add_sheet(wb, title, headers, rows, widths=None, money_cols=(),
               int_cols=(), first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title[:31]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for row in rows:
        ws.append([_clean(v) for v in row])
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows) + 1}"
    for idx, header in enumerate(headers, start=1):
        letter = get_column_letter(idx)
        if widths and idx <= len(widths) and widths[idx - 1]:
            ws.column_dimensions[letter].width = widths[idx - 1]
        else:
            longest = max([len(str(header))]
                          + [len(str(r[idx - 1])) for r in rows[:200]
                             if idx <= len(r) and r[idx - 1] is not None] or [10])
            ws.column_dimensions[letter].width = min(max(longest + 2, 10), 46)
        fmt = (MONEY if idx in money_cols else
               "#,##0" if idx in int_cols else None)
        if fmt:
            for cell in ws[letter][1:]:
                cell.number_format = fmt
        for cell in ws[letter][1:]:
            if isinstance(cell.value, date):
                cell.number_format = DATEFMT
    ws.row_dimensions[1].height = 28
    return ws


def _summary_sheet(wb, conn):
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = "Roof CRM — Full Export"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Generated {datetime.now().astimezone().strftime('%b %d, %Y %I:%M %p')}"
    ws["A2"].font = Font(color="666666")

    def one(sql, *params):
        row = conn.execute(sql, params).fetchone()
        return (row[0] if row and row[0] is not None else 0)

    active = "COALESCE(archived_at, '') = ''"
    money = conn.execute(
        """SELECT COALESCE(SUM(CASE WHEN status != 'Draft' THEN amount END), 0),
                  COALESCE(SUM(CASE WHEN status = 'Paid' THEN amount END), 0),
                  COALESCE(SUM(CASE WHEN status = 'Sent' THEN amount END), 0)
           FROM invoices""").fetchone()
    due = cadence.get_due_reminders(conn)

    blocks = [
        ("Pipeline", [
            ("Active accounts", one(f"SELECT COUNT(*) FROM accounts WHERE {active}"), None),
            ("Archived accounts", one("SELECT COUNT(*) FROM accounts "
                                      "WHERE COALESCE(archived_at, '') != ''"), None),
            ("In cold cadence", one(f"SELECT COUNT(*) FROM accounts WHERE {active} "
                                    "AND prospecting_status='Prospecting' "
                                    "AND pipeline_milestone='None / In Cadence'"), None),
            ("Interested", one(f"SELECT COUNT(*) FROM accounts WHERE {active} "
                               "AND prospecting_status LIKE 'Interested%'"), None),
            ("Closed won", one(f"SELECT COUNT(*) FROM accounts WHERE {active} "
                               "AND pipeline_milestone='Closed Won'"), None),
            ("Closed lost", one(f"SELECT COUNT(*) FROM accounts WHERE {active} "
                                "AND pipeline_milestone='Closed Lost'"), None),
        ]),
        ("Opportunity", [
            ("Matching buildings (all active accounts)",
             one(f"SELECT COALESCE(SUM(matching_properties),0) FROM accounts WHERE {active}"), None),
            ("Matching buildings in today's tasks",
             sum(r["matching_properties"] or 0 for r in due), None),
            ("Tasks due today", len(due), None),
            ("Follow-ups due", one(f"SELECT COUNT(*) FROM accounts WHERE {active} "
                                   "AND next_follow_up != '' AND next_follow_up <= ?",
                                   date.today().isoformat()), None),
        ]),
        ("Money", [
            ("Projects", one("SELECT COUNT(*) FROM projects"), None),
            ("Contracted", one("SELECT COALESCE(SUM(contract_amount),0) FROM projects"), MONEY0),
            ("Invoiced", money[0], MONEY0),
            ("Collected", money[1], MONEY0),
            ("Outstanding", money[2], MONEY0),
        ]),
        ("Activity", [
            ("Interactions logged (all time)", one("SELECT COUNT(*) FROM interactions"), None),
            ("Roof reports / bids", one("SELECT COUNT(*) FROM bids"), None),
            ("Quoted value of open bids",
             one("""SELECT COALESCE(SUM(b.price),0) FROM bids b
                    JOIN accounts a ON a.id = b.account_id
                    WHERE COALESCE(a.archived_at,'') = ''
                      AND a.pipeline_milestone NOT IN ('Closed Won','Closed Lost')"""), MONEY0),
        ]),
    ]
    row = 4
    for name, entries in blocks:
        ws.cell(row=row, column=1, value=name).font = SECTION_FONT
        row += 1
        for label, value, fmt in entries:
            ws.cell(row=row, column=1, value=label)
            cell = ws.cell(row=row, column=2, value=value)
            if fmt:
                cell.number_format = fmt
            row += 1
        row += 1
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 16
    ws.freeze_panes = "A4"


def build_workbook(conn) -> io.BytesIO:
    wb = Workbook()
    _summary_sheet(wb, conn)

    # --- Accounts -----------------------------------------------------
    rows = conn.execute(
        """SELECT a.*, (SELECT MAX(created_at) FROM interactions i
                        WHERE i.account_id = a.id) AS last_activity,
                  (SELECT COUNT(*) FROM interactions i
                   WHERE i.account_id = a.id) AS touches
           FROM accounts a
           ORDER BY COALESCE(a.archived_at,'') != '',
                    COALESCE(a.matching_properties,0) DESC,
                    a.company_name COLLATE NOCASE""").fetchall()
    _add_sheet(wb, "Accounts",
               ["Company", "Matching (🎯)", "Total properties", "First name",
                "Last name", "Title", "Email", "Work phone", "Mobile phone",
                "Preferred contact", "Prospecting status", "Pipeline milestone",
                "Cadence start", "Next follow-up", "Follow-up note", "Touches",
                "Last activity", "Archived", "Archive reason", "Notes", "Created"],
               [[r["company_name"], r["matching_properties"], r["num_properties"],
                 r["first_name"], r["last_name"], r["title"], r["email"],
                 r["work_phone"], r["mobile_phone"], r["preferred_contact"],
                 r["prospecting_status"], r["pipeline_milestone"], r["cadence_start"],
                 r["next_follow_up"], r["follow_up_note"], r["touches"],
                 r["last_activity"], "Yes" if r["archived_at"] else "",
                 r["archive_reason"], (r["notes"] or "").replace("\n", " · "),
                 r["created_at"]] for r in rows],
               widths=[34, 12, 12, 14, 14, 20, 26, 16, 16, 12, 22, 20, 12, 12,
                       24, 9, 14, 10, 20, 46, 12],
               int_cols=(2, 3, 16))

    # --- Contacts -----------------------------------------------------
    rows = conn.execute(
        """SELECT a.company_name, c.* FROM contacts c
           JOIN accounts a ON a.id = c.account_id
           ORDER BY a.company_name COLLATE NOCASE, c.last_name""").fetchall()
    _add_sheet(wb, "Contacts",
               ["Company", "First name", "Last name", "Title", "Email",
                "Work phone", "Mobile phone", "Added"],
               [[r["company_name"], r["first_name"], r["last_name"], r["title"],
                 r["email"], r["work_phone"], r["mobile_phone"], r["created_at"]]
                for r in rows],
               widths=[34, 14, 14, 24, 26, 16, 16, 12])

    # --- Interactions -------------------------------------------------
    rows = conn.execute(
        """SELECT a.company_name, i.* FROM interactions i
           JOIN accounts a ON a.id = i.account_id
           ORDER BY i.created_at DESC, i.id DESC""").fetchall()
    _add_sheet(wb, "Interactions",
               ["Date", "Company", "Type", "Notes"],
               [[r["created_at"], r["company_name"], r["interaction_type"],
                 r["notes"]] for r in rows],
               widths=[12, 34, 16, 70])

    # --- Roof reports / bids -----------------------------------------
    rows = conn.execute(
        """SELECT b.*, a.company_name FROM bids b
           JOIN accounts a ON a.id = b.account_id
           ORDER BY b.updated_at DESC""").fetchall()
    bid_rows = []
    for b in rows:
        plan = warranty_calc.calculate(
            b["roof_size_sqft"] or 0, coating_system=b["coating_system"] or "Silicone",
            roof_type=b["roof_type"] or "Capsheet",
            warranty_years=b["warranty_years"] or 10,
            acrylic_system_type=b["acrylic_system_type"] or "Standard",
            deduction_sqft=b["deduction_sqft"] or 0, linear_feet=b["linear_feet"] or 0,
            waste_pct=b["waste_pct"] or 0, stretch_pct=b["stretch_pct"] or 0,
            passed_adhesion=bool(b["passed_adhesion"]), has_rust=bool(b["has_rust"]),
            rust_prime_method=b["rust_prime_method"] or "field",
            topcoat=b["selected_topcoat"] or "", basecoat=b["selected_basecoat"] or "",
            butter_grade=b["selected_butter_grade"] or "")
        net = plan["net_sqft"] if plan else (b["roof_size_sqft"] or 0) - (b["deduction_sqft"] or 0)
        suggested = (warranty_calc.suggested_price(
            net, b["roof_type"] or "Capsheet", b["warranty_years"] or 10)["total"]
            if net > 0 else None)
        rates = " · ".join(f"{c['short']} {c['rate']}" for c in plan["coats"]) if plan else ""
        bid_rows.append([
            b["company_name"], b["roof_address"], b["roof_size_sqft"],
            b["deduction_sqft"], net, b["surface_type"], b["roof_type"],
            b["coating_system"], b["warranty_years"], b["candidate"],
            rates, plan["total_gallons"] if plan else "",
            plan["mastic_buckets"] if plan else "",
            b["price"], suggested, b["assessment_date"], b["updated_at"]])
    _add_sheet(wb, "Roof Reports",
               ["Company", "Roof address", "Roof sq ft", "Minus sq ft", "Coated sq ft",
                "Surface", "Roof category", "System", "Warranty (yrs)", "Candidate",
                "Rates (gal/sq)", "Total gallons", "Mastic pails", "Quoted price",
                "Suggested price", "Assessed", "Updated"],
               bid_rows,
               widths=[30, 34, 11, 11, 12, 18, 13, 11, 12, 11, 22, 12, 12, 14, 14, 12, 12],
               money_cols=(14, 15), int_cols=(3, 4, 5, 12, 13))

    # --- Projects -----------------------------------------------------
    rows = conn.execute(
        """SELECT p.*, a.company_name,
                  (SELECT COUNT(*) FROM project_tasks t
                   WHERE t.project_id = p.id AND t.done = 1) AS done,
                  (SELECT COUNT(*) FROM project_tasks t
                   WHERE t.project_id = p.id) AS total,
                  COALESCE((SELECT SUM(amount) FROM invoices v
                            WHERE v.project_id = p.id AND v.status != 'Draft'), 0) AS invoiced,
                  COALESCE((SELECT SUM(amount) FROM invoices v
                            WHERE v.project_id = p.id AND v.status = 'Paid'), 0) AS paid,
                  COALESCE((SELECT SUM(amount) FROM invoices v
                            WHERE v.project_id = p.id AND v.status = 'Sent'), 0) AS outstanding
           FROM projects p JOIN accounts a ON a.id = p.account_id
           ORDER BY p.updated_at DESC""").fetchall()
    _add_sheet(wb, "Projects",
               ["Project", "Company", "Status", "Contract", "Invoiced", "Collected",
                "Outstanding", "Checklist", "Start", "Completion", "Notes"],
               [[r["name"], r["company_name"], r["status"], r["contract_amount"],
                 r["invoiced"], r["paid"], r["outstanding"],
                 f"{r['done']}/{r['total']}", r["start_date"], r["completion_date"],
                 (r["notes"] or "").replace("\n", " · ")] for r in rows],
               widths=[36, 30, 14, 13, 13, 13, 13, 10, 12, 12, 46],
               money_cols=(4, 5, 6, 7))

    # --- Invoices -----------------------------------------------------
    today = date.today().isoformat()
    rows = conn.execute(
        """SELECT i.*, p.name AS project_name, a.company_name
           FROM invoices i JOIN projects p ON p.id = i.project_id
           JOIN accounts a ON a.id = p.account_id
           ORDER BY i.status != 'Sent', i.due_date, i.created_at""").fetchall()
    _add_sheet(wb, "Invoices",
               ["Invoice #", "Company", "Project", "Amount", "Status", "Sent",
                "Due", "Paid", "Days overdue", "Notes"],
               [[r["invoice_number"], r["company_name"], r["project_name"],
                 r["amount"], r["status"], r["sent_date"], r["due_date"],
                 r["paid_date"],
                 ((date.fromisoformat(today) - date.fromisoformat(r["due_date"])).days
                  if r["status"] == "Sent" and r["due_date"] and r["due_date"] < today
                  else ""),
                 r["notes"]] for r in rows],
               widths=[14, 30, 34, 13, 10, 12, 12, 12, 13, 40],
               money_cols=(4,))

    # --- Tasks due today ----------------------------------------------
    due = cadence.get_due_reminders(conn, order="priority")
    followups = conn.execute(
        """SELECT * FROM accounts WHERE COALESCE(archived_at,'') = ''
           AND next_follow_up != '' AND next_follow_up <= ?
           ORDER BY COALESCE(matching_properties,0) DESC, next_follow_up""",
        (today,)).fetchall()
    task_rows = [["Cadence", r["company_name"], r["matching_properties"],
                  f"Day {r['day']}: {r['step_type']}", r["due_date"],
                  r["days_overdue"] or "",
                  f"{r['first_name']} {r['last_name']}".strip(),
                  r["mobile_phone"] or r["work_phone"], r["email"]]
                 for r in due]
    task_rows += [["Follow-up", f["company_name"], f["matching_properties"],
                   f["follow_up_note"] or "Follow up", f["next_follow_up"],
                   (date.fromisoformat(today) - date.fromisoformat(f["next_follow_up"])).days
                   or "",
                   f"{f['first_name']} {f['last_name']}".strip(),
                   f["mobile_phone"] or f["work_phone"], f["email"]]
                  for f in followups]
    _add_sheet(wb, "Tasks Due",
               ["Kind", "Company", "Matching (🎯)", "What", "Due", "Days overdue",
                "Contact", "Phone", "Email"],
               task_rows,
               widths=[11, 34, 12, 26, 12, 13, 20, 16, 26],
               int_cols=(3,))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
