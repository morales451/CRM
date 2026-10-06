"""Take-backs for destructive actions.

Every delete (and every bulk change) records what it did into the `undo_log`
table before committing. Clicking Undo replays the record in reverse:
deleted rows are re-inserted with their original ids, changed columns are
written back, and bid photos are moved out of the trash folder.

Records expire after db.UNDO_RETENTION_DAYS so the log — and the photo
trash — can't grow forever.
"""

import json
import shutil
from datetime import datetime, timedelta

import db as db_module
from db import now_iso

# Only these tables can appear in an undo payload. The payloads are written
# by this app's own code, but an allow-list keeps a corrupted or hand-edited
# record from turning into arbitrary SQL.
ALLOWED_TABLES = {
    "accounts", "contacts", "interactions", "bids", "bid_photos",
    "projects", "project_tasks", "invoices", "cadence_dismissals",
}


def capture(conn, table: str, where: str, params=()) -> list[dict]:
    """Snapshot rows so they can be re-inserted later. Order matters:
    capture parents before children so the restore re-inserts them first."""
    if table not in ALLOWED_TABLES:
        raise ValueError(f"table not allowed in undo: {table}")
    return [dict(r) for r in conn.execute(
        f"SELECT * FROM {table} WHERE {where}", params)]


def insert_op(table: str, rows: list[dict]) -> dict:
    return {"op": "insert", "table": table, "rows": rows}


def update_op(table: str, rows: list[dict]) -> dict:
    """rows must each carry 'id' plus the columns to restore."""
    return {"op": "update", "table": table, "rows": rows}


def files_op(names: list[str]) -> dict:
    """Bid photo filenames sitting in the trash folder, to move back."""
    return {"op": "files", "names": names}


def record(conn, label: str, ops: list[dict]) -> int | None:
    """Store an undo record. Returns its id, or None if there's nothing to undo."""
    ops = [o for o in ops
           if (o["op"] == "files" and o["names"]) or o.get("rows")]
    if not ops:
        return None
    cur = conn.execute(
        "INSERT INTO undo_log (label, payload, created_at) VALUES (?,?,?)",
        (label, json.dumps({"ops": ops}), now_iso()))
    return cur.lastrowid


def trash_photo(filename: str) -> bool:
    """Move a bid photo into the trash folder so an undo can bring it back."""
    src = db_module.UPLOAD_DIR / filename
    if not src.exists():
        return False
    db_module.TRASH_DIR.mkdir(parents=True, exist_ok=True)
    shutil.move(str(src), str(db_module.TRASH_DIR / filename))
    return True


def _restore_photo(filename: str) -> None:
    src = db_module.TRASH_DIR / filename
    if src.exists():
        db_module.UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        shutil.move(str(src), str(db_module.UPLOAD_DIR / filename))


def peek(conn, undo_id) -> str | None:
    """The label of an unused, unexpired undo record — or None."""
    row = conn.execute(
        "SELECT label, created_at, used_at FROM undo_log WHERE id=?",
        (undo_id,)).fetchone()
    if row is None or row["used_at"]:
        return None
    if _expired(row["created_at"]):
        return None
    return row["label"]


def _expired(created_at: str) -> bool:
    try:
        age = datetime.now().astimezone() - datetime.fromisoformat(created_at)
    except (TypeError, ValueError):
        return True
    return age > timedelta(days=db_module.UNDO_RETENTION_DAYS)


def restore(conn, undo_id) -> str | None:
    """Replay an undo record. Returns its label on success, None if the
    record is missing, already used, or expired."""
    row = conn.execute(
        "SELECT * FROM undo_log WHERE id=?", (undo_id,)).fetchone()
    if row is None or row["used_at"] or _expired(row["created_at"]):
        return None
    ops = json.loads(row["payload"])["ops"]
    for op in ops:
        if op["op"] == "files":
            for name in op["names"]:
                _restore_photo(name)
            continue
        table = op["table"]
        if table not in ALLOWED_TABLES:
            raise ValueError(f"table not allowed in undo: {table}")
        if op["op"] == "insert":
            for data in op["rows"]:
                cols = list(data)
                conn.execute(
                    f"INSERT OR REPLACE INTO {table} ({','.join(cols)}) "
                    f"VALUES ({','.join('?' * len(cols))})",
                    [data[c] for c in cols])
        elif op["op"] == "update":
            for data in op["rows"]:
                cols = [c for c in data if c != "id"]
                if not cols:
                    continue
                conn.execute(
                    f"UPDATE {table} SET {','.join(c + '=?' for c in cols)} "
                    f"WHERE id=?", [data[c] for c in cols] + [data["id"]])
    conn.execute("UPDATE undo_log SET used_at=? WHERE id=?", (now_iso(), undo_id))
    conn.commit()
    return row["label"]


def purge(conn, retention_days: int | None = None) -> int:
    """Drop expired undo records and the trashed photos they held.
    Returns how many records were dropped."""
    days = db_module.UNDO_RETENTION_DAYS if retention_days is None else retention_days
    cutoff = (datetime.now().astimezone() - timedelta(days=days)).isoformat()
    doomed = conn.execute(
        "SELECT id, payload FROM undo_log WHERE created_at < ?", (cutoff,)).fetchall()
    for row in doomed:
        try:
            ops = json.loads(row["payload"])["ops"]
        except (ValueError, KeyError):
            ops = []
        for op in ops:
            if op.get("op") == "files":
                for name in op.get("names", []):
                    (db_module.TRASH_DIR / name).unlink(missing_ok=True)
    conn.execute("DELETE FROM undo_log WHERE created_at < ?", (cutoff,))
    # Sweep anything left behind in the trash by an older version or a crash.
    if db_module.TRASH_DIR.is_dir():
        stale = (datetime.now() - timedelta(days=days)).timestamp()
        for photo in db_module.TRASH_DIR.iterdir():
            try:
                if photo.is_file() and photo.stat().st_mtime < stale:
                    photo.unlink()
            except OSError:
                pass
    conn.commit()
    return len(doomed)
