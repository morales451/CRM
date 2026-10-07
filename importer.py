"""Excel/CSV account importer with fuzzy column mapping.

Headers are normalized (lowercased, non-alphanumerics stripped) and matched
against synonym lists, so files like HTX_Office_5kto10k.xlsx import without
needing exact column names.
"""

import re
from datetime import date, timedelta

import pandas as pd

import cadence
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
    # ZoomInfo exports these; they make a contact far more useful to work.
    "linkedin_url": [
        "linkedincontactprofileurl", "linkedinurl", "linkedin",
        "linkedinprofile", "linkedinprofileurl", "contactlinkedin",
    ],
    "seniority": [
        "managementlevel", "seniority", "joblevel", "level", "contactlevel",
        "senioritylevel",
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
    ("linkedin_url", lambda n: "linkedin" in n),
    ("seniority", lambda n: "managementlevel" in n or "seniority" in n
     or "joblevel" in n),
    ("company_name", lambda n: ("company" in n or "account" in n) and not any(
        k in n for k in ("address", "phone", "website", "type", "linkedin",
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

    # Matched on the NORMALIZED name so "Boxer Property Corp" and
    # "Boxer Property, Corp." are the same company — otherwise a re-pull from
    # CoStar silently creates duplicates and un-archives companies you removed.
    # An ACTIVE account for a company always wins over an archived one with the
    # same name: archiving a duplicate must not block the company it duplicated.
    # Active accounts are read first so they claim each key.
    existing, archived, existing_names = {}, {}, {}
    active_keys: set[str] = set()
    for row in conn.execute(
            "SELECT id, company_name, COALESCE(archived_at, '') AS archived_at "
            "FROM accounts ORDER BY (COALESCE(archived_at, '') != ''), id"):
        key = normalize_company(row["company_name"])
        existing.setdefault(key, row["id"])
        existing_names.setdefault(key, row["company_name"])
        if row["archived_at"]:
            archived.setdefault(key, row["company_name"])
        else:
            active_keys.add(key)

    imported = skipped_dupe = skipped_blank = backfilled = 0
    skipped_archived = 0
    started = to_research = 0
    archived_names = []
    # Rows matched to an account spelled differently — listed back so a match
    # is never silent and a wrong one can be spotted.
    matched_names: list[tuple[str, str]] = []
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

        def ctx(name):
            return _clean(row.get(context_mapping[name])) if name in context_mapping else ""

        key = normalize_company(company)
        if key in archived and key not in active_keys:
            # Deliberately removed from the working list — don't resurrect it.
            skipped_archived += 1
            if len(archived_names) < 25:
                archived_names.append(archived[key])
            continue

        if key in existing:
            # Duplicate: don't re-import, but fill in property counts the
            # account doesn't have yet (lets a re-upload backfill new columns).
            acct_id = existing[key]
            known = existing_names.get(key, "")
            if (known and known.strip().lower() != company.strip().lower()
                    and len(matched_names) < 50):
                matched_names.append((company, known))
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
            if ctx("website"):
                updated += conn.execute(
                    "UPDATE accounts SET website=?, updated_at=? "
                    "WHERE id=? AND COALESCE(website, '') = ''",
                    (ctx("website"), ts, acct_id)).rowcount
            if updated:
                backfilled += 1
            skipped_dupe += 1
            continue

        def field(name):
            return _clean(row.get(mapping[name])) if name in mapping else ""

        # Keep useful location/context columns by appending them to Notes.
        notes = field("notes")
        addr_parts = [p for p in (ctx("address"), ctx("city"), ctx("state")) if p]
        addr = ", ".join(addr_parts) + (f" {ctx('zip')}" if ctx("zip") and addr_parts else "")
        extras = []
        if addr:
            extras.append(f"Address: {addr}")
        sf = _clean_int(row.get(context_mapping["portfolio_sf"])) if "portfolio_sf" in context_mapping else None
        if sf:
            extras.append(f"Portfolio SF: {sf:,}")
        if extras:
            notes = (notes + "\n" if notes else "") + "\n".join(extras)

        person = {
            "first_name": field("first_name"), "last_name": field("last_name"),
            "email": field("email"),
            "work_phone": _clean_phone(row.get(mapping["work_phone"])) if "work_phone" in mapping else "",
            "mobile_phone": _clean_phone(row.get(mapping["mobile_phone"])) if "mobile_phone" in mapping else "",
        }
        # Someone to contact -> into the cadence (paced). Nobody yet -> into
        # Research with no clock; it starts the day a contact is added.
        ready = cadence.has_contact(person)
        cur = conn.execute(
            """INSERT INTO accounts
               (company_name, first_name, last_name, title, num_properties,
                matching_properties, email, work_phone, mobile_phone,
                linkedin_url, seniority, website,
                preferred_contact, notes, prospecting_status,
                pipeline_milestone, cadence_start, created_at, updated_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (company, person["first_name"], person["last_name"], field("title"),
             row_int("num_properties"), row_int("matching_properties"),
             person["email"], person["work_phone"], person["mobile_phone"],
             field("linkedin_url"), field("seniority"), ctx("website"),
             "Unknown", notes,
             "Prospecting", "None / In Cadence", start if ready else "", ts, ts),
        )
        existing[key] = cur.lastrowid
        existing_names.setdefault(key, company)
        imported += 1
        if not ready:
            to_research += 1
            continue                      # research rows don't use a pacing slot
        started += 1
        if per_day and started % per_day == 0:
            start_date = _next_business_day(start_date + timedelta(days=1))
            start = start_date.isoformat()

    conn.commit()
    return {
        "imported": imported,
        "backfilled": backfilled,
        "skipped_archived": skipped_archived,
        "archived_names": archived_names,
        "matched_names": matched_names,
        "started": started,
        "to_research": to_research,
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


_FILLER_WORDS = {"the", "a", "an", "of", "and"}

# Legal-entity suffixes ONLY. Descriptive words like Partners, Holdings, Group,
# Trust and REIT are part of a company's name — stripping them made "ABC
# Partners" and "ABC Holdings" the same company, and on import the second one
# was silently skipped as a duplicate. A missed duplicate shows up on the
# Import page's duplicates list; a false merge just quietly loses an account.
_LEGAL_SUFFIXES = (r"\b(llc|llp|lllp|lp|inc|incorporated|corp|corporation|co"
                   r"|company|ltd|limited|plc|pllc|pc)\b")


def normalize_company(name: str) -> str:
    """The key two spellings of the same company have in common.

    Lowercases, drops punctuation and strips legal-entity suffixes, so
    "Hartman Income REIT, Inc.", "Hartman Income Reit Inc" and "HARTMAN
    INCOME REIT LP" all collapse to "hartman income reit".

    Used for EVERY company match in the app — account import duplicates, the
    archived-account block list, and attaching ZoomInfo contacts — so a
    CoStar list that spells a name slightly differently next month can't
    create a second account or resurrect one you archived.

    Falls back to the plain lowercased name when stripping would leave
    nothing distinctive — "The Group" and "The Trust" must not collapse into
    each other just because "group" and "trust" are entity words.
    """
    # Periods go first so dotted abbreviations close up — "P.C." and
    # "L.L.C." must become "pc" and "llc", not "p c" and "l l c", or the
    # suffix list never sees them. Other punctuation becomes a space, so
    # "Smith,Jones" stays two words.
    base = str(name).lower().replace(".", "")
    base = re.sub(r"[^a-z0-9 ]", " ", base)
    base = re.sub(r"\s+", " ", base).strip()
    stripped = re.sub(r"\s+", " ", re.sub(_LEGAL_SUFFIXES, "", base)).strip()
    # Nothing but filler left ("The Group" -> "the") means the entity word WAS
    # the name; keep the full name so two such companies stay distinct.
    if not stripped or all(w in _FILLER_WORDS for w in stripped.split()):
        return base
    return stripped


# Kept as the old internal name so existing callers keep working.
_company_key = normalize_company


PERSON_FIELDS = ("first_name", "last_name", "title", "email",
                 "work_phone", "mobile_phone", "linkedin_url", "seniority")

# Company-level columns a contact export carries alongside each person.
# Matched only against columns the person fields didn't already claim.
CONTACT_COMPANY_SYNONYMS = {
    "website": ["website", "companywebsite", "companyurl", "web", "url",
                "domain", "companydomain"],
    "email_domain": ["emaildomain"],
    "hq_phone": ["companyhqphone", "hqphone", "companyphone", "mainphone",
                 "companymainphone"],
}

# Free mailbox providers. Two people on gmail.com are not colleagues, so
# these never tie a contact to a company.
_PUBLIC_MAIL_DOMAINS = {
    "gmail.com", "googlemail.com", "yahoo.com", "ymail.com", "hotmail.com",
    "outlook.com", "live.com", "msn.com", "aol.com", "icloud.com", "me.com",
    "mac.com", "comcast.net", "att.net", "sbcglobal.net", "verizon.net",
    "bellsouth.net", "cox.net", "charter.net", "protonmail.com", "proton.me",
}


def domain_of(value) -> str:
    """The bare domain in a website, URL or email address — "" if none.

    "https://www.HarlowEnterprises.com/about", "harlowenterprises.com" and
    "danny@harlowenterprises.com" all give "harlowenterprises.com". Free
    mailbox domains (gmail.com, ...) give "" because they say nothing about
    which company someone works for.
    """
    v = _clean(value).lower()
    if not v:
        return ""
    if "@" in v:
        v = v.rsplit("@", 1)[1]
    v = re.sub(r"^[a-z][a-z0-9+.-]*://", "", v)
    v = re.split(r"[/?#:\s]", v, 1)[0].strip(".")
    if v.startswith("www."):
        v = v[4:]
    if not re.fullmatch(r"[a-z0-9-]+(\.[a-z0-9-]+)+", v):
        return ""
    return "" if v in _PUBLIC_MAIL_DOMAINS else v


# Account columns a contact import can change, snapshotted for its undo.
_UNDO_ACCOUNT_COLS = ("first_name", "last_name", "title", "email", "work_phone",
                      "mobile_phone", "linkedin_url", "seniority", "notes",
                      "website", "cadence_start", "updated_at")


def import_contacts(conn, file_storage, create_missing: bool = False,
                    per_day: int | None = None, record_undo=None) -> dict:
    """Attach a contact export (e.g. from ZoomInfo) to accounts.

    Each row is matched to an account by normalized company name (see
    normalize_company); failing that, by web domain — the export's Email
    Domain / Website columns or the person's email, against each account's
    website. A domain shared by two of your accounts matches neither, and
    every domain match is listed back.

    The first contact for an account with no primary contact becomes the
    primary (filling the account's own contact fields); the rest become
    contacts rows. Duplicate people (same email, or same first+last name) on
    an account are skipped.

    create_missing=True opens a new account for any company in the file you
    don't already have, so a ZoomInfo pull can seed accounts instead of
    reporting them as unmatched. Archived companies are NEVER re-created or
    attached to — archiving means "stop showing me this company".

    per_day staggers the cadence starts of accounts leaving Research, so a
    300-company upload doesn't land 300 Day-1 tasks on one morning.
    record_undo(ops) is called before the commit with what undoes it all.
    """
    df = read_file(file_storage)
    mapping = map_columns(df.columns)
    company_cols = map_columns(
        [c for c in df.columns if c not in mapping.values()],
        CONTACT_COMPANY_SYNONYMS)

    if "company_name" not in mapping:
        raise ValueError(
            "Could not find a company column to match contacts against. "
            f"Found columns: {', '.join(str(c) for c in df.columns)}")
    if "first_name" not in mapping and "last_name" not in mapping:
        raise ValueError(
            "Could not find a contact name column (e.g. 'First Name' / "
            f"'Last Name'). Found columns: {', '.join(str(c) for c in df.columns)}")

    accounts_by_key: dict[str, int] = {}
    archived: dict[str, str] = {}
    names: dict[int, str] = {}
    active_domains: dict[str, set[int]] = {}
    archived_domains: dict[str, str] = {}
    # Domains come from each account's WEBSITE only. The emails of people on
    # an account aren't trusted for this: an owner's property manager often
    # works for a management firm, and that firm's domain would pull the
    # firm's own staff onto the owner's account.
    for row in conn.execute(
            "SELECT id, company_name, COALESCE(website, '') AS website, "
            "COALESCE(archived_at, '') AS archived_at FROM accounts"):
        key = normalize_company(row["company_name"])
        d = domain_of(row["website"])
        if row["archived_at"]:
            archived.setdefault(key, row["company_name"])
            if d:
                archived_domains.setdefault(d, row["company_name"])
        else:
            accounts_by_key.setdefault(key, row["id"])
            names[row["id"]] = row["company_name"]
            if d:
                active_domains.setdefault(d, set()).add(row["id"])

    attached = skipped_dupe = skipped_blank = skipped_archived = 0
    accounts_created = 0
    unmatched: dict[str, int] = {}
    archived_names: list[str] = []
    # (name in the file, account it was attached to) for every match made on
    # the web domain rather than the name — never silent, easy to spot.
    domain_matches: list[tuple[str, str]] = []
    matched_accounts: list[int] = []      # first-match order, for pacing
    before: dict[int, dict] = {}          # undo snapshots of changed accounts
    new_contacts: list[int] = []
    new_accounts: list[int] = []
    ts = now_iso()

    for _, row in df.iterrows():
        company = _clean(row.get(mapping["company_name"]))
        if not company:
            skipped_blank += 1
            continue

        def field(name):
            return _clean(row.get(mapping[name])) if name in mapping else ""

        def company_field(name):
            return _clean(row.get(company_cols[name])) if name in company_cols else ""

        person = {f: field(f) for f in PERSON_FIELDS}
        person["work_phone"] = (_clean_phone(row.get(mapping["work_phone"]))
                                if "work_phone" in mapping else "")
        person["mobile_phone"] = (_clean_phone(row.get(mapping["mobile_phone"]))
                                  if "mobile_phone" in mapping else "")
        if not (person["first_name"] or person["last_name"]):
            skipped_blank += 1
            continue
        website = company_field("website")
        row_domains = [d for d in (domain_of(company_field("email_domain")),
                                   domain_of(website), domain_of(person["email"]))
                       if d]

        key = normalize_company(company)
        account_id = accounts_by_key.get(key)
        if account_id is None:
            for d in row_domains:
                ids = active_domains.get(d, set())
                if len(ids) == 1:
                    account_id = next(iter(ids))
                    pair = (company, names.get(account_id, ""))
                    if pair not in domain_matches and len(domain_matches) < 50:
                        domain_matches.append(pair)
                    break
                if ids:
                    break                 # shared by several accounts: ambiguous
        if account_id is None:
            # Only an archived record has this name or domain: leave it alone.
            gone = archived.get(key) or next(
                (archived_domains[d] for d in row_domains if d in archived_domains),
                None)
            if gone:
                skipped_archived += 1
                if gone not in archived_names and len(archived_names) < 25:
                    archived_names.append(gone)
                continue
        if account_id is None:
            if not create_missing:
                unmatched[company] = unmatched.get(company, 0) + 1
                continue
            # Opened in Research; the contact attached below starts its clock.
            cur = conn.execute(
                """INSERT INTO accounts
                   (company_name, website, preferred_contact, prospecting_status,
                    pipeline_milestone, cadence_start, created_at, updated_at)
                   VALUES (?, ?, 'Unknown', 'Prospecting', 'None / In Cadence', '', ?, ?)""",
                (company, website, ts, ts))
            account_id = cur.lastrowid
            accounts_by_key[key] = account_id
            names[account_id] = company
            d = domain_of(website)
            if d:
                active_domains.setdefault(d, set()).add(account_id)
            new_accounts.append(account_id)
            accounts_created += 1

        acct = conn.execute("SELECT * FROM accounts WHERE id=?", (account_id,)).fetchone()
        existing_people = [dict(acct)] + [
            dict(r) for r in conn.execute(
                "SELECT * FROM contacts WHERE account_id=?", (account_id,))]
        name_key = (person["first_name"].lower(), person["last_name"].lower())
        is_dupe = any(
            (person["email"] and (p["email"] or "").lower() == person["email"].lower())
            or ((p["first_name"] or "").lower(), (p["last_name"] or "").lower()) == name_key
            for p in existing_people)
        if is_dupe:
            skipped_dupe += 1
            continue

        if account_id not in before and account_id not in new_accounts:
            before[account_id] = {"id": account_id,
                                  **{c: acct[c] for c in _UNDO_ACCOUNT_COLS}}
        if website and not (acct["website"] or "").strip():
            conn.execute("UPDATE accounts SET website=? WHERE id=?",
                         (website, account_id))
            if domain_of(website):
                active_domains.setdefault(domain_of(website), set()).add(account_id)
        has_primary = any(acct[c] for c in ("first_name", "last_name", "email"))
        if not has_primary:
            # The person's direct phone beats the company main line; keep the
            # main line in Notes so it isn't lost. No direct line -> dial the
            # main line (the account's, else the export's HQ phone).
            main_line = acct["work_phone"] or _clean_phone(company_field("hq_phone"))
            work_phone = person["work_phone"] or main_line
            notes = acct["notes"]
            if person["work_phone"] and main_line \
                    and person["work_phone"] != main_line \
                    and main_line not in (notes or ""):
                notes = (notes + "\n" if notes else "") \
                    + f"Company main line: {main_line}"
            conn.execute(
                "UPDATE accounts SET first_name=?, last_name=?, title=?, email=?, "
                "work_phone=?, mobile_phone=?, linkedin_url=?, seniority=?, "
                "notes=?, updated_at=? WHERE id=?",
                (person["first_name"], person["last_name"], person["title"],
                 person["email"], work_phone, person["mobile_phone"],
                 person["linkedin_url"], person["seniority"], notes,
                 ts, account_id))
        else:
            cur = conn.execute(
                "INSERT INTO contacts (account_id, first_name, last_name, title, "
                "email, work_phone, mobile_phone, linkedin_url, seniority, "
                "created_at) VALUES (?,?,?,?,?,?,?,?,?,?)",
                (account_id, *(person[f] for f in PERSON_FIELDS), ts))
            new_contacts.append(cur.lastrowid)
        if account_id not in matched_accounts:
            matched_accounts.append(account_id)
        attached += 1

    # Anyone who just got a reachable contact leaves Research — today, or
    # per_day at a time across the coming business days.
    day = _next_business_day(cadence.today())
    on_day = cadence_started = 0
    last_start = ""
    for a in matched_accounts:
        if per_day and on_day == per_day:
            day = _next_business_day(day + timedelta(days=1))
            on_day = 0
        start = day.isoformat() if per_day else cadence.today().isoformat()
        if cadence.start_cadence_if_ready(conn, a, start):
            cadence_started += 1
            on_day += 1
            last_start = start
    if record_undo is not None:
        record_undo([
            # Created accounts go first; their contacts go with them (cascade).
            {"op": "delete", "table": "accounts",
             "rows": [{"id": i} for i in new_accounts]},
            {"op": "delete", "table": "contacts",
             "rows": [{"id": i} for i in new_contacts]},
            {"op": "update", "table": "accounts", "rows": list(before.values())},
        ])
    conn.commit()
    return {
        "attached": attached,
        "companies_matched": len(matched_accounts),
        "accounts_created": accounts_created,
        "cadence_started": cadence_started,
        "per_day": per_day,
        "last_start_date": last_start,
        "domain_matches": domain_matches,
        "skipped_duplicates": skipped_dupe,
        "skipped_blank": skipped_blank,
        "skipped_archived": skipped_archived,
        "archived_names": archived_names,
        "unmatched": unmatched,          # {company name: row count}
        "total_rows": len(df),
        "mapped_columns": {f: str(c) for f, c in
                           {**mapping, **company_cols}.items()},
    }


_US_STATES = {
    "alabama", "alaska", "arizona", "arkansas", "california", "colorado",
    "connecticut", "delaware", "florida", "georgia", "hawaii", "idaho",
    "illinois", "indiana", "iowa", "kansas", "kentucky", "louisiana", "maine",
    "maryland", "massachusetts", "michigan", "minnesota", "mississippi",
    "missouri", "montana", "nebraska", "nevada", "new hampshire",
    "new jersey", "new mexico", "new york", "north carolina", "north dakota",
    "ohio", "oklahoma", "oregon", "pennsylvania", "rhode island",
    "south carolina", "south dakota", "tennessee", "texas", "utah", "vermont",
    "virginia", "washington", "west virginia", "wisconsin", "wyoming",
    "district of columbia",
}


def _is_state(part: str) -> bool:
    p = part.strip().lower()
    return bool(re.fullmatch(r"[a-z]{2}", p)) or p in _US_STATES


def address_from_notes(notes: str) -> dict:
    """Street / city / state / zip from the "Address: ..." line an account
    import writes into Notes ("123 Main St, Houston, TX 77002" — any part
    may be missing). Best effort; blanks where it can't tell."""
    out = {"street": "", "city": "", "state": "", "zip": ""}
    m = re.search(r"^Address:\s*(.+?)\s*$", notes or "", re.M)
    if not m:
        return out
    line = m.group(1)
    z = re.search(r"\s(\d{5}(?:-\d{4})?)$", line)
    if z:
        out["zip"] = z.group(1)
        line = line[:z.start()]
    parts = [p.strip() for p in line.split(",") if p.strip()]
    if parts and _is_state(parts[-1]):
        out["state"] = parts.pop()
        if parts:
            out["city"] = parts.pop()
    elif len(parts) >= 2:
        out["city"] = parts.pop()
    out["street"] = ", ".join(parts)
    return out


# --------------------------------------------------------- Paste a profile

_EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
_LINKEDIN_RE = re.compile(
    r"(?:https?://)?(?:[\w-]+\.)?linkedin\.com/\S+", re.I)
_LABEL_RE = re.compile(r"^\s*([A-Za-z][A-Za-z /#.'-]{1,30})\s*[:\-–]\s*(.+?)\s*$")

# Words that mean a line is a job title, not a person's name.
_TITLE_HINTS = (
    "president", "vp", "vice", "director", "manager", "officer", "chief",
    "ceo", "cfo", "coo", "cto", "owner", "principal", "partner", "head",
    "lead", "supervisor", "coordinator", "administrator", "executive",
    "asset", "property", "facilities", "operations", "real estate",
    "engineer", "analyst", "specialist", "associate", "controller",
)

# ZoomInfo's "Management Level" values, matched loosely against a title.
# Ordered: "Vice President of Asset Management" must read VP-Level, not
# C-Level, so the VP rule is tested before the one matching "president".
#
# Each rule carries two kinds of hint, because a plain substring test is
# wrong for abbreviations: "cto" hides inside "dire-cto-r" and "coo" inside
# "coo-rdinator". Abbreviations must match as whole words; the spelled-out
# terms match as word prefixes so "Directors" still reads as Director.
_SENIORITY_RULES = (
    ("VP-Level", ("vp", "svp", "evp", "avp"),
     ("vice president", "vice-president")),
    ("C-Level", ("ceo", "cfo", "coo", "cto", "cio", "cmo", "cro"),
     ("chief", "president", "owner", "principal", "founder", "partner")),
    ("Director", (), ("director", "head of")),
    ("Manager", (), ("manager", "supervisor", "superintendent")),
)

_LABELS = {
    "full_name": ("name", "full name", "contact", "contact name", "person"),
    "first_name": ("first name", "first", "given name"),
    "last_name": ("last name", "last", "surname", "family name"),
    "title": ("title", "job title", "position", "role"),
    "email": ("email", "email address", "work email", "e-mail"),
    "work_phone": ("direct", "direct phone", "direct phone number", "phone",
                   "work phone", "office", "office phone", "work", "tel",
                   "telephone", "company phone", "hq phone"),
    "mobile_phone": ("mobile", "mobile phone", "cell", "cell phone",
                     "mobile number"),
    "linkedin_url": ("linkedin", "linkedin url", "linkedin profile"),
    "seniority": ("management level", "seniority", "job level", "level"),
    "company": ("company", "company name", "account", "employer"),
}
_LABEL_LOOKUP = {alias: field for field, aliases in _LABELS.items()
                 for alias in aliases}

# ZoomInfo's own panel headings. They sit on their own lines and are not data —
# without this, "Contact Details" became the contact's name.
_SECTION_HEADINGS = {
    _normalize(h) for h in (
        "Contact Details", "Contact Detail", "Contact Info",
        "Contact Information", "Details", "Overview", "About",
        "Emails", "Email", "Email Address", "Email Addresses",
        "Phone Numbers", "Phone Number", "Phone", "Phones",
        "Direct Dial", "Direct Phone", "HQ Phone", "Mobile Phone",
        "Company Details", "Company Detail", "Company Information",
        "Location", "Locations", "Education", "Employment History",
        "Work History", "Web References", "Technologies", "Skills",
        "Social", "Social Media", "Websites", "Bio", "Summary",
    )
}

# The short type tag ZoomInfo prints beside (or under) each value.
_TYPE_TAGS = {
    "m": "mobile", "c": "mobile", "cell": "mobile", "mobile": "mobile",
    "d": "direct", "dd": "direct", "direct": "direct",
    "hq": "hq", "headquarters": "hq",
    "o": "work", "b": "work", "w": "work", "work": "work",
    "office": "work", "business": "work",
    "p": "personal", "personal": "personal", "h": "personal",
}

# A line that is nothing but a tag: "(M)", "[HQ]", "M".
_TAG_ONLY_RE = re.compile(r"^\s*[\[(]?\s*[A-Za-z]{1,12}\s*[\])]?\s*$")
# ...or the same tag parked at the end of the value's own line.
_INLINE_TAG_RE = re.compile(r"\s*[\[(]\s*([A-Za-z]{1,12})\s*[\])]\s*$")


def _remainder(line: str) -> str:
    """What's left of a line after lifting an email/LinkedIn URL out of it.

    Returns "" for leftovers that carry no information — a bare separator
    ("|", "-") or the label the value belonged to ("Email Address:") —
    so they can't be mistaken for a name or a company later on."""
    line = line.strip(" \t|,;:-\u2013")
    if not any(ch.isalnum() for ch in line):
        return ""
    if line.strip().rstrip(":").strip().lower() in _LABEL_LOOKUP:
        return ""
    return line


def _looks_like_name(line: str) -> bool:
    """A person's name: 2-4 words, letters only, no title words."""
    if not line or len(line) > 60 or any(ch.isdigit() for ch in line):
        return False
    words = line.replace(",", " ").split()
    if not 1 < len(words) <= 4:
        return False
    low = line.lower()
    if _normalize(line) in _SECTION_HEADINGS:
        return False
    if any(hint in low for hint in _TITLE_HINTS):
        return False
    if not all(re.fullmatch(r"[A-Za-z][A-Za-z.'\-]*", w) for w in words):
        return False
    # Real names are capitalised. Interface text like "Lists and records"
    # isn't — which is how a nav item once became a contact's name.
    return all(w[0].isupper() or w.lower() in _NAME_PARTICLES for w in words)


def _split_name(full: str) -> tuple[str, str]:
    full = full.strip()
    if "," in full:                       # "Doe, Jane"
        last, _, first = full.partition(",")
        return first.strip(), last.strip()
    parts = full.split()
    if len(parts) == 1:
        return parts[0], ""
    # Drop a middle initial: "Gerald W. Hayes" is Gerald Jones, not
    # Gerald "W. Hayes". Particles like "de" and "van" are kept.
    middles = [p for p in parts[1:-1]
               if not re.fullmatch(r"[A-Za-z]\.?", p)]
    return parts[0], " ".join(middles + [parts[-1]])


def seniority_from_title(title: str) -> str:
    """Best-guess ZoomInfo management level from a job title."""
    low = (title or "").lower()
    for level, abbreviations, terms in _SENIORITY_RULES:
        if any(re.search(rf"\b{re.escape(a)}\b", low) for a in abbreviations):
            return level
        if any(re.search(rf"\b{re.escape(t)}", low) for t in terms):
            return level
    return ""


# A page copied out of the browser arrives as Markdown: every link shows up
# as [text](url). The URL is worth more than the text — ZoomInfo links the
# company to /profile/company/ and other people to /profile/person/.
_MD_LINK_RE = re.compile(r"\[([^\]]*)\]\(([^)]*)\)")
_COMPANY_LINK = "/profile/company/"
_PERSON_LINK = "/profile/person/"

# Bullets Markdown puts in front of list items.
_BULLET_RE = re.compile(r"^\s*[*+\u2022\u2013-]\s+")

# Name particles that are lowercase in a real name.
_NAME_PARTICLES = {"de", "del", "della", "van", "von", "der", "den", "da",
                   "di", "la", "le", "bin", "al", "ter", "ten", "dos", "do"}


# Everything below one of these headings belongs to OTHER people or to
# unrelated sections, so a whole-page copy is cut off here. Without this, a
# colleague's phone number from "Similar Contacts" could land on your contact.
_PAGE_STOP_MARKERS = {
    _normalize(h) for h in (
        "Similar Contacts", "Similar Profiles", "People Also Viewed",
        "Related Contacts", "Other Contacts", "Org Chart",
        "Organizational Chart", "Colleagues", "Coworkers",
        # NOT "Employees": on a contact page that is a company field label
        # sitting above the headcount, and stopping there threw away the
        # Contact Details panel below it. It stays in the chrome list instead.
        "Company Contacts", "More Contacts", "Recommended Contacts",
        "Frequently Viewed", "Similar Companies", "Competitors",
        "Related Companies", "News", "Recent News", "Scoops", "Funding",
        "Technologies Used", "Intent", "Employment History",
        "Web References", "Activity Feed", "About", "Recent Activity",
    )
}

# Stops only once the contact's own details have been seen. ZoomInfo puts
# Location directly under the phone numbers, but a different layout might put
# it higher, and cutting the page before the phones would be worse.
_SOFT_STOP_MARKERS = {_normalize(h) for h in ("Location", "Locations", "CRM")}

# ZoomInfo's own navigation, buttons, tabs and field labels. Dropped so they
# can't be mistaken for a name, a title or a company. Taken from a real
# whole-page copy — "Lists and records" was becoming the contact's name.
_UI_CHROME = {
    _normalize(c) for c in (
        # left-hand navigation and toolbar
        "Navigation", "Home", "Search", "Advanced Search", "Signals",
        "Lists and records", "Lists", "My Lists", "Automations", "Alerts",
        "Suggest Update", "Tag", "Tags", "Add tag", "Add Tag", "Track Contact",
        "Export", "Exports", "Tabs", "Contact Profile", "Overview",
        "Technologies", "Save", "Saved", "Save to List", "Add to List",
        "Share", "Copy", "Copied", "Print", "Edit", "Delete", "View Profile",
        "View Full Profile", "See More", "Show More", "Show more",
        "Show Less", "More", "Less", "Back", "Next", "Previous", "Close",
        "Cancel", "Done", "Upgrade", "Upgrade Now", "Request",
        "Request Contact", "Feedback", "Report an Issue",
        "Report Inaccuracy", "Suggest an Edit", "Settings", "Help",
        "Support", "Log Out", "Sign Out", "Profile", "Dashboard",
        "Notifications", "Filters", "Filter", "Sort", "Clear", "Select All",
        "Actions", "Enrich", "Engage", "Connect", "Connect Now", "Follow",
        "Following", "Verified", "Premium", "Add Note", "Notes",
        "ZoomInfo", "Copy to Clipboard", "Reveal", "Click to Reveal",
        "Effort", "Engagement",
        # field labels that sit on their own line above their value
        "Website", "Industry", "Revenue", "Employees", "Local", "HQ",
        "Account Owner", "Match Date", "Last Updated", "Updated",
        "Accuracy", "Confidence", "Contact Details", "Phone numbers",
        "Emails", "Email", "Current Role Start Date", "Current Company Start Date",
        "Notice Provided Date", "Salesforce", "HubSpot", "Dynamics",
        "Enter search terms or select from the recent searches in the open dialog",
    )
}


# Addresses that belong to a company, not a person.
_ROLE_MAILBOXES = {
    "info", "sales", "admin", "administration", "contact", "contactus",
    "office", "leasing", "support", "hello", "team", "accounting",
    "accountspayable", "accountsreceivable", "ap", "ar", "billing", "hr",
    "careers", "jobs", "noreply", "donotreply", "help", "service",
    "services", "main", "mail", "enquiries", "inquiries", "general",
    "maintenance", "property", "management", "operations", "reception",
}


def name_from_email(email: str) -> tuple[str, str]:
    """Best-effort first and last name from an email address.

    Only the cases that are actually unambiguous: a local part separated by
    ".", "_" or "-", like michael.delacruz@ (first and last) or
    m.delacruz@ (an initial, so only the last name).

    A run-together local part is deliberately NOT guessed at. "mdelacruz"
    looks like M. Delacruz, but "delacruz" looks identical to a machine, and
    filling in "Elacruz" is worse than leaving the box empty.

    Returns ("", "") whenever it isn't sure.
    """
    local = str(email or "").split("@")[0].strip().lower()
    local = re.sub(r"\d+$", "", local)                   # trailing jsmith2
    if not local or local.replace(".", "").replace("_", "").replace("-", "") \
            in _ROLE_MAILBOXES:
        return "", ""
    parts = [p for p in re.split(r"[._-]+", local) if p]
    if len(parts) < 2:
        return "", ""
    first, last = parts[0], parts[-1]
    if not last.isalpha() or len(last) < 2:
        return "", ""
    # A single leading letter is an initial, not a first name.
    first_name = first.title() if first.isalpha() and len(first) > 1 else ""
    return first_name, last.title()


def _best_values(out: dict, emails: list, phones: list) -> None:
    """Choose which of the addresses and numbers found actually go on the
    contact. Tags decide it: a business address beats a personal one, and a
    direct dial beats the company switchboard."""
    if not out["email"] and emails:
        out["email"] = next((e[0] for e in emails if e[1] == "work"), emails[0][0])
    if not phones:
        return
    mobiles = [p[0] for p in phones if p[1] == "mobile"]
    direct = [p[0] for p in phones if p[1] == "direct"]
    work = [p[0] for p in phones if p[1] in ("work", "hq")]
    untagged = [p[0] for p in phones if not p[1]]
    work_order = direct + work + untagged
    if not out["work_phone"] and work_order:
        out["work_phone"] = work_order[0]
    mobile_order = mobiles + [p for p in untagged if p != out["work_phone"]]
    if not out["mobile_phone"] and mobile_order:
        out["mobile_phone"] = mobile_order[0]


def _tag_of(line: str) -> str:
    """The type tag on a line of its own: "(M)", "[HQ]", "D"."""
    return (_TYPE_TAGS.get(re.sub(r"[^a-z]", "", line.lower()))
            if _TAG_ONLY_RE.match(line) else "") or ""


def _expand_lines(raw_lines):
    """Turn pasted page text into records: {"text", "url", "bullet"}.

    A browser copy is Markdown, so "[Harlow Enterprises](…/profile/company/1)"
    arrives as one line. The URL is the useful part — it says whether the line
    names the company, another person, or nothing in particular. A line made
    only of links becomes one record per link, so a navigation bar collapses
    into individual buttons that the chrome list can drop.
    """
    out = []
    for line in raw_lines:
        line = line.strip()
        bullet = bool(_BULLET_RE.match(line))   # test BEFORE stripping it off
        line = _BULLET_RE.sub("", line)
        links = _MD_LINK_RE.findall(line)
        if not links:
            out.append({"text": line, "url": "", "bullet": bullet})
            continue
        stripped = _MD_LINK_RE.sub("", line).strip(" |\t")
        if not stripped:
            # Nothing but links: each one stands on its own.
            for text, url in links:
                out.append({"text": text.strip(" ,|"), "url": url,
                            "bullet": bullet})
        else:
            # Links embedded in a sentence: keep the sentence, drop the URLs.
            out.append({"text": _MD_LINK_RE.sub(r"\1", line).strip(),
                        "url": links[0][1], "bullet": bullet})
    return out


def _contact_header(records):
    """(name, title, company) read off a ZoomInfo contact page's header.

    The company is a link to /profile/company/, and ZoomInfo stacks the header
    as name, job title, company. Anchoring on that link beats guessing from
    word lists: it survives new nav items, renamed buttons and a page with no
    email on it at all.
    """
    for i, rec in enumerate(records):
        if _COMPANY_LINK not in (rec["url"] or "") or not rec["text"]:
            continue
        company = rec["text"]
        above = [r["text"] for r in records[:i] if r["text"]]
        if len(above) >= 2 and _looks_like_name(above[-2]):
            return above[-2], above[-1], company
        if above and _looks_like_name(above[-1]):
            return above[-1], "", company
        return "", "", company
    return "", "", ""


def parse_contact_blob(text: str) -> dict:
    """Pull a contact out of text copied from a ZoomInfo profile.

    Handles the shapes people actually paste:

    * ZoomInfo's own **Contact Details** panel, where the headings ("Emails",
      "Phone numbers") sit on their own lines and each value is followed by a
      type tag on the NEXT line — (B) business, (HQ) head office, (D) direct,
      (M) mobile. Those tags decide which number is the mobile and which is
      the work line, instead of it coming down to the order they appear in.
    * A plain stack of lines: name, title, company, email, phones.
    * "Label: value" pairs.
    * A row copied out of a spreadsheet (tab separated).

    Anything it can't place is left blank rather than guessed at — the form is
    still shown, so nothing ends up silently wrong.

    Returns the contact fields plus "company" (only to report which company
    the paste named; it never re-points the contact).
    """
    out = {f: "" for f in PERSON_FIELDS}
    out["company"] = ""
    if not text or not text.strip():
        return out

    # A single tab-separated line is a spreadsheet row; treat cells as lines.
    raw_lines = [l.strip() for l in text.replace("\r", "").split("\n")]
    if len([l for l in raw_lines if l]) == 1 and "\t" in text:
        raw_lines = [c.strip() for c in text.split("\t")]
    records = _expand_lines(raw_lines)
    name_line, title_line, company_line = _contact_header(records)
    if name_line:
        out["first_name"], out["last_name"] = _split_name(name_line)
    if title_line:
        out["title"] = title_line
    if company_line:
        out["company"] = company_line

    lines = []
    seen_detail = False
    for rec in records:
        line, url = rec["text"], rec["url"]
        if not line:
            continue
        key = _normalize(line)
        # A bullet is a list item ("* Org Chart" is a tab, not a heading), so
        # it can never cut the page short — the tab strip sits ABOVE the
        # contact's phone numbers, and stopping there lost them entirely.
        if not rec["bullet"]:
            if key in _PAGE_STOP_MARKERS:
                break                  # the rest belongs to other people
            if key in _SOFT_STOP_MARKERS and seen_detail:
                break
        if key in _UI_CHROME:
            continue                   # a button or a field label, not data
        if _PERSON_LINK in url:
            continue                   # somebody else's profile
        if len(line) <= 2 and line.isalpha():
            continue                   # avatar initials ("A", "AM")
        if _EMAIL_RE.search(line) or len(re.sub(r"\D", "", line)) >= 10:
            seen_detail = True
        lines.append(line)

    emails: list[list] = []   # [[address, tag], ...]
    phones: list[list] = []   # [[number, tag], ...]
    leftovers: list[str] = []
    last = None               # the list a trailing type tag belongs to

    def take_tag(line):
        """A line that is only a type tag, e.g. "(M)" or "HQ"."""
        return _TYPE_TAGS.get(re.sub(r"[^a-z]", "", line.lower())) \
            if _TAG_ONLY_RE.match(line) else None

    for line in lines:
        if _normalize(line) in _SECTION_HEADINGS:
            last = None
            continue

        tag = take_tag(line)
        if tag:
            # Belongs to the value on the line above.
            if last:
                last[1] = last[1] or tag
            continue

        # A tag can also sit at the end of the value's own line.
        inline = _INLINE_TAG_RE.search(line)
        inline_tag = ""
        if inline:
            found = _TYPE_TAGS.get(re.sub(r"[^a-z]", "", inline.group(1).lower()))
            if found:
                inline_tag = found
                line = line[:inline.start()].strip()

        li = _LINKEDIN_RE.search(line)
        if li and not out["linkedin_url"]:
            out["linkedin_url"] = li.group(0).rstrip(".,;")
            line = _remainder(line.replace(li.group(0), ""))
            if not line:
                continue
        em = _EMAIL_RE.search(line)
        if em:
            emails.append([em.group(0), inline_tag])
            last = emails[-1]
            line = _remainder(line.replace(em.group(0), ""))
            if not line:
                continue

        m = _LABEL_RE.match(line)
        field = _LABEL_LOOKUP.get(m.group(1).strip().lower()) if m else None
        if field:
            value = m.group(2).strip()
            if field == "full_name":
                if not (out["first_name"] or out["last_name"]):
                    out["first_name"], out["last_name"] = _split_name(value)
            elif field in ("work_phone", "mobile_phone"):
                phones.append([_clean_phone(value),
                               "mobile" if field == "mobile_phone" else "work"])
                last = phones[-1]
            elif field == "company":
                out["company"] = out["company"] or value
            elif not out.get(field):
                out[field] = value
            continue

        digits = re.sub(r"\D", "", line)
        if len(digits) >= 10 and len(re.sub(r"[\d\s().+\-x/]", "", line)) <= 12:
            low = line.lower()
            phone_tag = inline_tag or (
                "mobile" if ("mobile" in low or "cell" in low)
                else "direct" if "direct" in low
                else "work" if ("office" in low or "work" in low) else "")
            phones.append([_clean_phone(line), phone_tag])
            last = phones[-1]
            continue

        leftovers.append(line)
        last = None


    _best_values(out, emails, phones)

    # Unlabeled lines, in ZoomInfo's own order: name, then title, then company.
    #
    # A whole-page copy contains several name-like lines — the contact, the
    # company, section headings, sometimes colleagues. The email settles it:
    # the person whose surname shows up in the address is the one the page is
    # about. mdelacruz@harlowenterprises.com picks "Michael Delacruz" out of
    # the page and ignores everything else that merely looks like a name.
    if not (out["first_name"] or out["last_name"]):
        candidates = [(i, l) for i, l in enumerate(leftovers) if _looks_like_name(l)]
        chosen = None
        local = out["email"].split("@")[0].lower() if out["email"] else ""
        local_letters = re.sub(r"[^a-z]", "", local)
        if local_letters:
            best = 0
            for i, line in candidates:
                first, last = _split_name(line)
                score = 0
                surname = re.sub(r"[^a-z]", "", last.lower())
                if surname and len(surname) > 1 and surname in local_letters:
                    score += 2
                    # ...and the initial in front of it seals it.
                    if first and local_letters.startswith(first[0].lower()):
                        score += 1
                given = re.sub(r"[^a-z]", "", first.lower())
                if given and len(given) > 1 and given in local_letters:
                    score += 1
                if score > best:
                    best, chosen = score, (i, line)
        if chosen is None and candidates:
            chosen = candidates[0]
        if chosen:
            i, line = chosen
            out["first_name"], out["last_name"] = _split_name(line)
            # The job title is the line directly under the name.
            nxt = leftovers[i + 1] if i + 1 < len(leftovers) else ""
            if nxt and not out["title"] and not _looks_like_name(nxt):
                out["title"] = nxt
                leftovers.pop(i + 1)
            leftovers.pop(i)
    # Only read a title or company out of the leftovers once the paste has
    # proved it really is a contact. Otherwise a stray line — an address, a
    # mis-selected paragraph — would quietly become somebody's job title.
    is_a_contact = bool(out["first_name"] or out["last_name"]
                        or out["email"] or out["work_phone"] or out["mobile_phone"])
    if is_a_contact:
        if not out["title"] and leftovers:
            out["title"] = leftovers.pop(0)
        if not out["company"] and leftovers:
            out["company"] = leftovers.pop(0)
    if not out["seniority"]:
        out["seniority"] = seniority_from_title(out["title"])
    return out


# --------------------------------------------- A whole roster in one paste

# Lines that follow a name but are never a job title.
_NOT_A_TITLE = {_normalize(t) for t in (
    "View Profile", "Save", "Export", "Add to List", "Email", "Phone",
    "Mobile", "Direct", "Show More", "Show Less", "Contact", "Connect",
    "Verified", "LinkedIn", "Twitter", "Facebook",
)}


def _person_block(records) -> dict:
    """Read one person's rows off a company employee list.

    A row shows the name and title, and — once that row is expanded — their
    email and phone numbers too. Everything between one person's profile link
    and the next belongs to that person, so expanding the few rows you care
    about and copying the page once gets their details as well as their names.
    """
    out = {f: "" for f in PERSON_FIELDS}
    emails, phones = [], []
    last = None

    for rec in records:
        line, url = rec["text"].strip(), rec["url"] or ""
        if not line:
            continue
        # The URL is worth looking at before the text is judged: the link
        # labelled "LinkedIn" is on the not-a-title list, but its address is
        # exactly what we want.
        if "linkedin.com" in url.lower():
            out["linkedin_url"] = out["linkedin_url"] or url
            continue

        key = _normalize(line)
        if key in _UI_CHROME or key in _NOT_A_TITLE:
            last = None
            continue
        if len(line) <= 2 and line.isalpha():
            continue                               # avatar initial

        tag = _tag_of(line)
        if tag:
            if last:
                last[1] = last[1] or tag           # belongs to the line above
            continue

        inline = _INLINE_TAG_RE.search(line)
        inline_tag = ""
        if inline:
            found = _TYPE_TAGS.get(re.sub(r"[^a-z]", "", inline.group(1).lower()))
            if found:
                inline_tag = found
                line = line[:inline.start()].strip()

        li = _LINKEDIN_RE.search(line)
        if li:
            out["linkedin_url"] = out["linkedin_url"] or li.group(0).rstrip(".,;")
            line = _remainder(line.replace(li.group(0), ""))
            if not line:
                continue

        em = _EMAIL_RE.search(line)
        if em:
            emails.append([em.group(0), inline_tag])
            last = emails[-1]
            continue

        digits = re.sub(r"\D", "", line)
        if len(digits) >= 10 and len(re.sub(r"[\d\s().+\-x/]", "", line)) <= 12:
            low = line.lower()
            phones.append([_clean_phone(line), inline_tag or (
                "mobile" if ("mobile" in low or "cell" in low)
                else "direct" if "direct" in low else "")])
            last = phones[-1]
            continue

        if not out["title"]:
            out["title"] = line                    # the line under the name
        last = None

    _best_values(out, emails, phones)
    return out


def parse_contact_roster(text: str) -> list[dict]:
    """Everyone on a pasted ZoomInfo company page, in one go.

    The company's Employees tab lists the whole org — name, job title and
    management level — without opening anybody's profile. Each person there is
    a link to /profile/person/, which makes them unambiguous to pick out, and
    everything up to the next person's link belongs to them.

    Contact details are collapsed on that page by default. **Expand the rows
    you want before copying** and their email and phone numbers come through
    in the same single paste — which is the whole point: pick four people,
    expand four rows, copy once.

    Returns a list in page order, without repeats, each with the contact
    fields plus "has_details" saying whether anything beyond the title came
    through for that person.
    """
    if not text or not text.strip():
        return []
    records = _expand_lines([l.strip() for l in text.replace("\r", "").split("\n")])

    # Where each person starts, and where the roster itself ends.
    starts = [i for i, rec in enumerate(records)
              if _PERSON_LINK in (rec["url"] or "")
              and rec["text"].strip() and _looks_like_name(rec["text"].strip())]
    if not starts:
        return []
    end = len(records)
    for i in range(starts[-1] + 1, len(records)):
        rec = records[i]
        if not rec["bullet"] and _normalize(rec["text"]) in _PAGE_STOP_MARKERS:
            end = i
            break

    people, seen = [], set()
    for n, i in enumerate(starts):
        stop = starts[n + 1] if n + 1 < len(starts) else end
        first, last_name = _split_name(records[i]["text"].strip())
        key = (first.lower(), last_name.lower())
        if key in seen:
            continue
        seen.add(key)
        person = _person_block(records[i + 1:stop])
        person["first_name"], person["last_name"] = first, last_name
        person["seniority"] = person["seniority"] or seniority_from_title(person["title"])
        person["has_details"] = bool(person["email"] or person["work_phone"]
                                     or person["mobile_phone"])
        people.append(person)
    return people


# ------------------------------------------------------ Duplicate accounts

def find_duplicate_groups(conn) -> list[dict]:
    """Active accounts that normalize to the same company.

    Older versions matched on the exact name, so a CoStar re-pull spelled
    "… Corp." instead of "… Corp" could create a second account. This finds
    what slipped through before the fix so it can be cleaned up by hand.
    """
    groups: dict[str, list] = {}
    for row in conn.execute(
            """SELECT a.id, a.company_name, a.matching_properties,
                      a.prospecting_status, a.pipeline_milestone, a.created_at,
                      a.first_name, a.last_name,
                      (SELECT COUNT(*) FROM interactions i WHERE i.account_id = a.id) AS touches,
                      (SELECT COUNT(*) FROM contacts c WHERE c.account_id = a.id)
                        + (CASE WHEN TRIM(COALESCE(a.first_name,'') || COALESCE(a.last_name,'')) != ''
                                THEN 1 ELSE 0 END) AS people
               FROM accounts a
               WHERE COALESCE(a.archived_at, '') = '' ORDER BY a.id"""):
        groups.setdefault(normalize_company(row["company_name"]), []).append(dict(row))
    out = []
    for key, accts in groups.items():
        if len(accts) < 2:
            continue
        # Suggest keeping the one that's been worked most (then the one with
        # most people on it, then the oldest) — that's where the history is.
        best = max(accts, key=lambda a: (a["touches"], a["people"], -a["id"]))
        for a in accts:
            a["suggested"] = a["id"] == best["id"]
        out.append({"key": key, "accounts": accts})
    return out
