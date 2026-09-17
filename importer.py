"""Excel/CSV account importer with fuzzy column mapping.

Headers are normalized (lowercased, non-alphanumerics stripped) and matched
against synonym lists, so files like HTX_Office_5kto10k.xlsx import without
needing exact column names.
"""

import re
import pandas as pd

from db import now_iso, today_iso

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
    "email": ["email", "emailaddress", "workemail", "contactemail", "mail"],
    "work_phone": [
        "workphone", "phone", "phonenumber", "officephone", "businessphone",
        "directphone", "telephone", "companyphone", "mainphone", "work",
    ],
    "mobile_phone": [
        "mobilephone", "mobile", "cell", "cellphone", "cellular",
        "mobilenumber", "cellnumber",
    ],
    "notes": ["notes", "note", "comments", "comment", "description", "remarks"],
}


def _normalize(header: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(header).lower())


def map_columns(columns) -> dict:
    """Map dataframe columns to account fields. Returns {field: original_column}."""
    mapping = {}
    normalized = {_normalize(c): c for c in columns}
    for field, synonyms in COLUMN_SYNONYMS.items():
        for syn in synonyms:
            if syn in normalized:
                mapping[field] = normalized[syn]
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


def _clean_int(value):
    s = _clean(value)
    if not s:
        return None
    try:
        return int(float(s))
    except ValueError:
        return None


def import_accounts(conn, file_storage) -> dict:
    """Import rows as accounts. Returns summary dict.

    Defaults for every imported account:
      prospecting_status = 'Prospecting'
      pipeline_milestone = 'None / In Cadence'
      cadence_start      = today  (so they enter the cadence immediately)

    Duplicate rule: skip a row when an account with the same company name
    (case-insensitive) already exists.
    """
    df = read_file(file_storage)
    mapping = map_columns(df.columns)

    if "company_name" not in mapping:
        raise ValueError(
            "Could not find a company-name column. Expected a header like "
            "'Company Name', 'Company', 'Account Name', or 'Business'. "
            f"Found columns: {', '.join(str(c) for c in df.columns)}"
        )

    existing = {
        row["company_name"].strip().lower()
        for row in conn.execute("SELECT company_name FROM accounts")
    }

    imported = skipped_dupe = skipped_blank = 0
    ts = now_iso()
    start = today_iso()

    for _, row in df.iterrows():
        company = _clean(row.get(mapping["company_name"]))
        if not company:
            skipped_blank += 1
            continue
        if company.lower() in existing:
            skipped_dupe += 1
            continue

        def field(name):
            return _clean(row.get(mapping[name])) if name in mapping else ""

        conn.execute(
            """INSERT INTO accounts
               (company_name, first_name, last_name, title, num_properties,
                email, work_phone, mobile_phone, preferred_contact, notes,
                prospecting_status, pipeline_milestone, cadence_start,
                created_at, updated_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (company, field("first_name"), field("last_name"), field("title"),
             _clean_int(row.get(mapping["num_properties"])) if "num_properties" in mapping else None,
             field("email"), field("work_phone"), field("mobile_phone"),
             "Unknown", field("notes"),
             "Prospecting", "None / In Cadence", start, ts, ts),
        )
        existing.add(company.lower())
        imported += 1

    conn.commit()
    return {
        "imported": imported,
        "skipped_duplicates": skipped_dupe,
        "skipped_blank": skipped_blank,
        "total_rows": len(df),
        "mapped_columns": {f: str(c) for f, c in mapping.items()},
        "unmapped_columns": [str(c) for c in df.columns if c not in mapping.values()],
    }
