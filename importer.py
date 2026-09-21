"""Excel/CSV account importer with fuzzy column mapping.

Headers are normalized (lowercased, non-alphanumerics stripped) and matched
against synonym lists, so files like HTX_Office_5kto10k.xlsx import without
needing exact column names.
"""

import re
from datetime import date, timedelta

import pandas as pd

from db import now_iso

# field -> normalized header synonyms
COLUMN_SYNONYMS = {
    "company_name": [
        "companyname", "company", "account", "accountname", "business",
        "businessname", "organization", "organisation", "propertyname",
        "buildingname",
    ],
    "first_name": ["firstname", "first", "contactfirstname", "givenname"],
    "last_name": ["lastname", "last", "surname", "contactlastname", "familyname"],
    "title": ["title", "jobtitle", "position", "role", "contacttitle"],
    "num_properties": [
        "numberofproperties", "numproperties", "properties", "propertycount",
        "ofproperties", "numberproperties", "totalproperties", "propertiesowned",
    ],
    # How many of their properties fit the sales criteria (e.g. pre-1980s
    # buildings in the searched size band) — CoStar's "# Properties (in search)"
    "matching_properties": [
        "propertiesinsearch", "matchingproperties", "propertiesmatching",
        "propertiesmatchingcriteria", "qualifyingproperties", "matchingcount",
    ],
    "email": ["email", "emailaddress", "workemail", "contactemail", "mail"],
    "work_phone": [
        "workphone", "phone", "phonenumber", "officephone", "businessphone",
        "directphone", "directphonenumber", "telephone", "companyphone",
        "mainphone", "hqphone", "work",
    ],
    "mobile_phone": [
        "mobilephone", "mobile", "cell", "cellphone", "cellular",
        "mobilenumber", "cellnumber",
    ],
    "notes": ["notes", "note", "comments", "comment", "description", "remarks"],
}

# Context columns (e.g. from CoStar-style property exports) that don't map to
# account fields directly but are worth keeping — they get appended to Notes.
CONTEXT_SYNONYMS = {
    "address": ["companyaddress", "address", "streetaddress", "mailingaddress"],
    "city": ["city", "town"],
    "state": ["statecountry", "state", "stateprovince"],
    "zip": ["zipcode", "zip", "postalcode"],
    "website": ["website", "url", "web", "companywebsite"],
    "portfolio_sf": ["portfoliosf", "totalsf", "buildingsf"],
}


def _normalize(header: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(header).lower())


# Fallback keyword rules for headers that aren't an exact synonym — lets
# differently-worded CoStar/ZoomInfo exports map without editing the file.
# Ordered: more specific rules claim their column before broader ones.
_KEYWORD_TESTS = [
    ("matching_properties", lambda n: "propert" in n and any(
        k in n for k in ("search", "matching", "criteria", "qualif"))),
    ("num_properties", lambda n: "propert" in n and any(
        k in n for k in ("owned", "total", "count", "number", "num"))),
    ("mobile_phone", lambda n: "mobile" in n or "cell" in n),
    ("work_phone", lambda n: "phone" in n and not any(
        k in n for k in ("mobile", "cell", "fax"))),
    ("email", lambda n: "email" in n),
    ("first_name", lambda n: "first" in n and "name" in n),
    ("last_name", lambda n: ("last" in n and "name" in n) or "surname" in n),
    ("title", lambda n: "title" in n),
    ("company_name", lambda n: ("company" in n or "account" in n) and not any(
        k in n for k in ("address", "phone", "website", "type",
                         "city", "state", "zip", "email"))),
    ("notes", lambda n: "note" in n or "comment" in n),
]


def map_columns(columns, synonym_table=COLUMN_SYNONYMS) -> dict:
    """Map dataframe columns to fields. Returns {field: original_column}.

    Exact (normalized) synonym matches win; for account fields, a keyword
    heuristic then fills anything still unmapped, so column names only need
    to MEAN the same thing, not match the original file.
    """
    mapping = {}
    normalized = {_normalize(c): c for c in columns}
    for field, synonyms in synonym_table.items():
        for syn in synonyms:
            if syn in normalized:
                mapping[field] = normalized[syn]
                break
    if synonym_table is COLUMN_SYNONYMS:
        used = set(mapping.values())
        for field, test in _KEYWORD_TESTS:
            if field in mapping:
                continue
            for norm, orig in normalized.items():
                if orig not in used and test(norm):
                    mapping[field] = orig
                    used.add(orig)
                    break
    return mapping


def read_file(file_storage) -> pd.DataFrame:
    name = (file_storage.filename or "").lower()
    if name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(file_storage, engine="openpyxl")
    elif name.endswith(".csv"):
        df = pd.read_csv(file_storage)
    else:
        raise ValueError("Unsupported file type — upload a .xlsx or .csv file.")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _clean(value) -> str:
    if value is None or pd.isna(value):
        return ""
    s = str(value).strip()
    # Excel often stores phone numbers / counts as floats: "7135551234.0"
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return "" if s.lower() in ("nan", "none", "null") else s


def _clean_phone(value) -> str:
    # CoStar-style exports suffix numbers with "(p)" / "(f)" / "(m)" markers
    return re.sub(r"\s*\([pfmo]\)\s*$", "", _clean(value), flags=re.I)


def _clean_int(value):
    s = _clean(value)
    if not s:
        return None
    try:
        return int(float(s))
    except ValueError:
        return None


def _next_business_day(d: date) -> date:
    while d.weekday() >= 5:  # Sat=5, Sun=6
        d += timedelta(days=1)
    return d


def import_accounts(conn, file_storage, per_day: int | None = None) -> dict:
    """Import rows as accounts. Returns summary dict.

    Defaults for every imported account:
      prospecting_status = 'Prospecting'
      pipeline_milestone = 'None / In Cadence'
      cadence_start      = today  (so they enter the cadence immediately)

    per_day: if set, stagger cadence starts so only that many accounts enter
    the cadence per business day (weekends skipped) — keeps the daily task
    list workable on big imports.

    Duplicate rule: skip a row when an account with the same company name
    (case-insensitive) already exists.
    """
    df = read_file(file_storage)
    mapping = map_columns(df.columns)
    context_mapping = map_columns(
        [c for c in df.columns if c not in mapping.values()], CONTEXT_SYNONYMS)

    if "company_name" not in mapping:
        raise ValueError(
            "Could not find a company-name column. Expected a header like "
            "'Company Name', 'Company', 'Account Name', or 'Business'. "
            f"Found columns: {', '.join(str(c) for c in df.columns)}"
        )

    existing = {
        row["company_name"].strip().lower(): row["id"]
        for row in conn.execute("SELECT id, company_name FROM accounts")
    }

    imported = skipped_dupe = skipped_blank = backfilled = 0
    ts = now_iso()
    start_date = _next_business_day(date.today()) if per_day else date.today()
    start = start_date.isoformat()

    for _, row in df.iterrows():
        company = _clean(row.get(mapping["company_name"]))
        if not company:
            skipped_blank += 1
            continue
        def row_int(name):
            return _clean_int(row.get(mapping[name])) if name in mapping else None

        if company.lower() in existing:
            # Duplicate: don't re-import, but fill in property counts the
            # account doesn't have yet (lets a re-upload backfill new columns).
            acct_id = existing[company.lower()]
            updated = conn.execute(
                """UPDATE accounts SET
                     num_properties = COALESCE(num_properties, ?),
                     matching_properties = COALESCE(matching_properties, ?),
                     updated_at = ?
                   WHERE id = ? AND (
                     (num_properties IS NULL AND ? IS NOT NULL) OR
                     (matching_properties IS NULL AND ? IS NOT NULL))""",
                (row_int("num_properties"), row_int("matching_properties"), ts,
                 acct_id, row_int("num_properties"),
                 row_int("matching_properties"))).rowcount
            if updated:
                backfilled += 1
            skipped_dupe += 1
            continue

        def field(name):
            return _clean(row.get(mapping[name])) if name in mapping else ""

        def ctx(name):
            return _clean(row.get(context_mapping[name])) if name in context_mapping else ""

        # Keep useful location/context columns by appending them to Notes.
        notes = field("notes")
        addr_parts = [p for p in (ctx("address"), ctx("city"), ctx("state")) if p]
        addr = ", ".join(addr_parts) + (f" {ctx('zip')}" if ctx("zip") and addr_parts else "")
        extras = []
        if addr:
            extras.append(f"Address: {addr}")
        if ctx("website"):
            extras.append(f"Website: {ctx('website')}")
        sf = _clean_int(row.get(context_mapping["portfolio_sf"])) if "portfolio_sf" in context_mapping else None
        if sf:
            extras.append(f"Portfolio SF: {sf:,}")
        if extras:
            notes = (notes + "\n" if notes else "") + "\n".join(extras)

        cur = conn.execute(
            """INSERT INTO accounts
               (company_name, first_name, last_name, title, num_properties,
                matching_properties, email, work_phone, mobile_phone,
                preferred_contact, notes, prospecting_status,
                pipeline_milestone, cadence_start, created_at, updated_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (company, field("first_name"), field("last_name"), field("title"),
             row_int("num_properties"), row_int("matching_properties"),
             field("email"),
             _clean_phone(row.get(mapping["work_phone"])) if "work_phone" in mapping else "",
             _clean_phone(row.get(mapping["mobile_phone"])) if "mobile_phone" in mapping else "",
             "Unknown", notes,
             "Prospecting", "None / In Cadence", start, ts, ts),
        )
        existing[company.lower()] = cur.lastrowid
        imported += 1
        if per_day and imported % per_day == 0:
            start_date = _next_business_day(start_date + timedelta(days=1))
            start = start_date.isoformat()

    conn.commit()
    return {
        "imported": imported,
        "backfilled": backfilled,
        "skipped_duplicates": skipped_dupe,
        "skipped_blank": skipped_blank,
        "total_rows": len(df),
        "per_day": per_day,
        "last_start_date": start,
        "mapped_columns": {f: str(c) for f, c in mapping.items()},
        "context_columns": {f: str(c) for f, c in context_mapping.items()},
        "unmapped_columns": [str(c) for c in df.columns
                             if c not in mapping.values()
                             and c not in context_mapping.values()],
    }


def _company_key(name: str) -> str:
    """Normalize a company name for matching: lowercase, strip punctuation
    and common legal suffixes (LLC, Inc, LP, ...)."""
    s = re.sub(r"[^a-z0-9 ]", "", str(name).lower())
    s = re.sub(r"\b(llc|llp|lp|inc|incorporated|corp|corporation|co|company|ltd|limited)\b",
               "", s)
    return re.sub(r"\s+", " ", s).strip()


def import_contacts(conn, file_storage) -> dict:
    """Attach a contact export (e.g. from ZoomInfo) to existing accounts.

    Rows are matched to accounts by normalized company name. The first
    contact for an account with no primary contact becomes the primary
    (filling the account's own contact fields); the rest become contacts
    rows. Duplicate people (same email, or same first+last name) on an
    account are skipped. Unmatched company names are reported back.
    """
    df = read_file(file_storage)
    mapping = map_columns(df.columns)

    if "company_name" not in mapping:
        raise ValueError(
            "Could not find a company column to match contacts against. "
            f"Found columns: {', '.join(str(c) for c in df.columns)}")
    if "first_name" not in mapping and "last_name" not in mapping:
        raise ValueError(
            "Could not find a contact name column (e.g. 'First Name' / "
            f"'Last Name'). Found columns: {', '.join(str(c) for c in df.columns)}")

    accounts_by_key: dict[str, int] = {}
    for row in conn.execute("SELECT id, company_name FROM accounts"):
        accounts_by_key.setdefault(_company_key(row["company_name"]), row["id"])

    attached = skipped_dupe = skipped_blank = 0
    unmatched: dict[str, int] = {}
    matched_accounts: set[int] = set()
    ts = now_iso()

    for _, row in df.iterrows():
        company = _clean(row.get(mapping["company_name"]))
        if not company:
            skipped_blank += 1
            continue
        account_id = accounts_by_key.get(_company_key(company))
        if account_id is None:
            unmatched[company] = unmatched.get(company, 0) + 1
            continue

        def field(name):
            return _clean(row.get(mapping[name])) if name in mapping else ""

        person = {
            "first_name": field("first_name"),
            "last_name": field("last_name"),
            "title": field("title"),
            "email": field("email"),
            "work_phone": _clean_phone(row.get(mapping["work_phone"])) if "work_phone" in mapping else "",
            "mobile_phone": _clean_phone(row.get(mapping["mobile_phone"])) if "mobile_phone" in mapping else "",
        }
        if not (person["first_name"] or person["last_name"]):
            skipped_blank += 1
            continue

        acct = conn.execute("SELECT * FROM accounts WHERE id=?", (account_id,)).fetchone()
        existing_people = [dict(acct)] + [
            dict(r) for r in conn.execute(
                "SELECT * FROM contacts WHERE account_id=?", (account_id,))]
        name_key = (person["first_name"].lower(), person["last_name"].lower())
        is_dupe = any(
            (person["email"] and p["email"].lower() == person["email"].lower())
            or (p["first_name"].lower(), p["last_name"].lower()) == name_key
            for p in existing_people)
        if is_dupe:
            skipped_dupe += 1
            continue

        has_primary = any(acct[c] for c in ("first_name", "last_name", "email"))
        if not has_primary:
            # The person's direct phone beats the company main line; keep the
            # main line in Notes so it isn't lost.
            work_phone = person["work_phone"] or acct["work_phone"]
            notes = acct["notes"]
            if person["work_phone"] and acct["work_phone"] \
                    and person["work_phone"] != acct["work_phone"]:
                notes = (notes + "\n" if notes else "") \
                    + f"Company main line: {acct['work_phone']}"
            conn.execute(
                "UPDATE accounts SET first_name=?, last_name=?, title=?, email=?, "
                "work_phone=?, mobile_phone=?, notes=?, updated_at=? WHERE id=?",
                (person["first_name"], person["last_name"], person["title"],
                 person["email"], work_phone, person["mobile_phone"], notes,
                 ts, account_id))
        else:
            conn.execute(
                "INSERT INTO contacts (account_id, first_name, last_name, title, "
                "email, work_phone, mobile_phone, created_at) VALUES (?,?,?,?,?,?,?,?)",
                (account_id, *person.values(), ts))
        matched_accounts.add(account_id)
        attached += 1

    conn.commit()
    return {
        "attached": attached,
        "companies_matched": len(matched_accounts),
        "skipped_duplicates": skipped_dupe,
        "skipped_blank": skipped_blank,
        "unmatched": unmatched,          # {company name: row count}
        "total_rows": len(df),
        "mapped_columns": {f: str(c) for f, c in mapping.items()},
    }
