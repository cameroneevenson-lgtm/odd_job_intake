"""Persistent registry of one-off job intakes.

Pure data-access module - no Qt, no Flask - so the desktop Job Intake page and
the Outlook-facing listener can share one store.

Storage is SQLite. The previous JSON file was read-modify-write with no lock,
and this store has three concurrent writers: the Flask request thread, its
background extraction worker, and the Qt UI. Two of them saving at once meant
the second silently discarded the first's changes. Each entry is still a JSON
blob in the `data` column, so rows remain hand-inspectable, but the primary key
claims a job atomically and updates run inside a transaction.
"""

from __future__ import annotations

from contextlib import closing
from datetime import datetime
import json
from pathlib import Path
import sqlite3
from typing import Any

from paths import JOB_INTAKE_REGISTRY_PATH


STATUS_NEW = "new"
STATUS_RPD_CREATED = "rpd_created"
STATUS_PARTS_IMPORTED = "parts_imported"
STATUS_BLOCKS_SENT = "blocks_sent"
STATUS_ERROR = "error"

VALID_STATUSES = (
    STATUS_NEW,
    STATUS_RPD_CREATED,
    STATUS_PARTS_IMPORTED,
    STATUS_BLOCKS_SENT,
    STATUS_ERROR,
)

VALID_SOURCES = ("manual", "outlook")

# Async intake state, separate from `status`. `status` tracks how far the job
# has progressed toward the machine (RPD made, parts imported, blocks sent);
# `state` tracks whether the *extraction* that runs on a background thread is
# still going. The macro polls on this one, so both endings must be reachable -
# a failure that never leaves "running" is a five-minute wait for nothing.
STATE_QUEUED = "queued"
STATE_RUNNING = "running"
STATE_SUCCEEDED = "succeeded"
STATE_FAILED = "failed"

VALID_STATES = (STATE_QUEUED, STATE_RUNNING, STATE_SUCCEEDED, STATE_FAILED)
TERMINAL_STATES = (STATE_SUCCEEDED, STATE_FAILED)


def _now_text() -> str:
    return datetime.now().isoformat(timespec="seconds")


def entry_key(job_number: str, label: str | None = None) -> str:
    """A job number alone identifies a fresh one-off; a labeled one-off under
    an existing job folder is identified by number + label together."""
    number = str(job_number or "").strip().upper()
    label_text = str(label or "").strip()
    return f"{number}::{label_text.casefold()}" if label_text else number


def new_entry(
    *,
    job_number: str,
    label: str | None = None,
    source: str = "manual",
    email_subject: str = "",
    email_sender: str = "",
) -> dict[str, Any]:
    if source not in VALID_SOURCES:
        raise ValueError(f"Invalid intake source: {source}")
    now = _now_text()
    return {
        "key": entry_key(job_number, label),
        "job_number": str(job_number or "").strip().upper(),
        "label": str(label or "").strip() or None,
        "po_number": None,
        "received_at": now,
        "updated_at": now,
        "source": source,
        "status": STATUS_NEW,
        # A manually created intake is already done extracting by the time it
        # is written; only the listener's async path moves through queued.
        "state": STATE_SUCCEEDED if source == "manual" else STATE_QUEUED,
        "email_subject": email_subject,
        "email_sender": email_sender,
        "job_folder": None,
        "rpd_path": None,
        "csv_log_path": None,
        "error": None,
        "due_date": None,
        "laser_hours": None,
        "bend_hours": None,
        "attachments": [],
        "material_qty": [],
    }


def _resolve_path(path: Path | None) -> Path:
    # Resolve the module global at call time (not as a default arg) so tests
    # that monkeypatch JOB_INTAKE_REGISTRY_PATH are honored.
    return path if path is not None else JOB_INTAKE_REGISTRY_PATH


def _database_path(path: Path | None) -> Path:
    """The SQLite file for a given registry path.

    Callers still pass the historical .json path around, so the database sits
    beside it under the same stem. That keeps every existing caller and test
    working unchanged while the storage underneath is transactional.
    """
    resolved = _resolve_path(path)
    return resolved.with_suffix(".db")


# Databases whose journal mode has already been settled this process. Only an
# optimisation - the mode is persisted in the file itself.
_WAL_ENABLED: set[str] = set()


def _connect(path: Path | None):
    """A connection with the settings this store depends on.

    The registry is written by a Flask worker, its request threads and the Qt
    UI at once. The old JSON store was an unlocked read-modify-write, so two
    writers could each read, each modify, and the second silently discard the
    first's work. SQLite makes that a transaction instead.
    """
    database = _database_path(path)
    database.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(
        str(database),
        timeout=30.0,  # wait for another writer rather than failing instantly
        isolation_level=None,  # explicit transactions, no implicit commits
        check_same_thread=False,
    )
    # Set the wait first: everything below can hit a lock, and until this is in
    # place a busy database raises instead of waiting.
    connection.execute("PRAGMA busy_timeout=30000")
    if str(database) not in _WAL_ENABLED:
        try:
            connection.execute("PRAGMA journal_mode=WAL")
        except sqlite3.OperationalError:
            # Switching journal mode needs exclusive access and does *not*
            # honour busy_timeout, so it fails outright when another connection
            # is already open. That is fine to skip: the mode is stored in the
            # file, so whoever opened it first has already set it, and the
            # default rollback journal is correct either way - only slower.
            pass
        _WAL_ENABLED.add(str(database))
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS entries (
            key         TEXT PRIMARY KEY,
            received_at TEXT NOT NULL DEFAULT '',
            data        TEXT NOT NULL
        )
        """
    )
    _migrate_json_store(connection, _resolve_path(path))
    return connection


def _migrate_json_store(connection, json_path: Path) -> None:
    """One-off import of the old JSON registry, then rename it aside.

    Renamed rather than deleted so nothing is destroyed by an upgrade, and so a
    second start does not re-import entries that have since been changed.
    """
    if not json_path.exists():
        return
    try:
        payload = json.loads(json_path.read_text(encoding="utf-8"))
        records = payload.get("entries", []) if isinstance(payload, dict) else []
    except (OSError, json.JSONDecodeError):
        records = []

    for record in records:
        if not isinstance(record, dict):
            continue
        key = str(record.get("key", "") or "")
        if not key:
            continue
        connection.execute(
            "INSERT OR IGNORE INTO entries (key, received_at, data) VALUES (?, ?, ?)",
            (key, str(record.get("received_at", "") or ""), json.dumps(record)),
        )
    try:
        json_path.replace(json_path.with_suffix(".json.migrated"))
    except OSError:
        pass


def load_entries(path: Path | None = None) -> list[dict[str, Any]]:
    """Newest first, matching how the queue is read in the UI."""
    with closing(_connect(path)) as connection:
        rows = connection.execute(
            "SELECT data FROM entries ORDER BY received_at DESC"
        ).fetchall()
    entries: list[dict[str, Any]] = []
    for (data,) in rows:
        try:
            record = json.loads(data)
        except json.JSONDecodeError:
            continue
        if isinstance(record, dict):
            entries.append(record)
    return entries


def get_entry(key: str, path: Path | None = None) -> dict[str, Any] | None:
    with closing(_connect(path)) as connection:
        row = connection.execute(
            "SELECT data FROM entries WHERE key = ?", (key,)
        ).fetchone()
    if row is None:
        return None
    try:
        record = json.loads(row[0])
    except json.JSONDecodeError:
        return None
    return record if isinstance(record, dict) else None


def append_entry(entry: dict[str, Any], path: Path | None = None) -> dict[str, Any]:
    """Claim a job number (+ label) for this intake.

    The primary key is what makes the claim atomic: two intakes racing for the
    same job cannot both read "not there yet" and both insert. The loser gets
    the same ValueError the old read-then-check produced.
    """
    key = str(entry.get("key", "") or "")
    if not key:
        raise ValueError("Registry entries need a key (use new_entry()).")
    with closing(_connect(path)) as connection:
        try:
            connection.execute(
                "INSERT INTO entries (key, received_at, data) VALUES (?, ?, ?)",
                (key, str(entry.get("received_at", "") or ""), json.dumps(entry)),
            )
        except sqlite3.IntegrityError as exc:
            raise ValueError(f"A job intake already exists for {key}.") from exc
    return entry


def set_state(
    key: str,
    state: str,
    path: Path | None = None,
    **fields: Any,
) -> dict[str, Any]:
    """Move an intake through the async state machine."""
    if state not in VALID_STATES:
        raise ValueError(f"Invalid intake state: {state}")
    return update_entry(key, path=path, state=state, **fields)


def entry_state(entry: dict[str, Any] | None) -> str:
    """The async state of an entry, defaulting old rows to a terminal one.

    Entries written before `state` existed are finished by definition - they
    were saved synchronously - so reporting them as terminal is correct rather
    than merely convenient.
    """
    if not entry:
        return STATE_FAILED
    state = str(entry.get("state", "") or "")
    return state if state in VALID_STATES else STATE_SUCCEEDED


def fail_interrupted_entries(path: Path | None = None) -> list[str]:
    """Close out intakes whose worker died with the process.

    Nothing will ever finish them - the thread that was running them is gone -
    so leaving them `running` means a poller waits out its whole timeout on a
    job that stopped when the app was last closed. Called at startup.
    """
    stranded: list[str] = []
    for entry in load_entries(path):
        if entry_state(entry) in (STATE_QUEUED, STATE_RUNNING):
            key = str(entry.get("key", "") or "")
            if not key:
                continue
            set_state(
                key,
                STATE_FAILED,
                path=path,
                status=STATUS_ERROR,
                error="Intake was interrupted before it finished; file the job again.",
            )
            stranded.append(key)
    return stranded


def update_entry(key: str, path: Path | None = None, **fields: Any) -> dict[str, Any]:
    """Merge fields into an entry inside a single transaction.

    BEGIN IMMEDIATE takes the write lock before reading, so two writers cannot
    both read the same record and have the second discard the first's changes -
    which is exactly what the Flask worker and the Qt UI could do to each other
    with the old JSON store.
    """
    status = fields.get("status")
    if status is not None and status not in VALID_STATUSES:
        raise ValueError(f"Invalid intake status: {status}")

    with closing(_connect(path)) as connection:
        connection.execute("BEGIN IMMEDIATE")
        try:
            row = connection.execute(
                "SELECT data FROM entries WHERE key = ?", (key,)
            ).fetchone()
            if row is None:
                raise ValueError(f"Job intake was not found: {key}")
            record = json.loads(row[0])
            record.update(fields)
            record["updated_at"] = _now_text()
            connection.execute(
                "UPDATE entries SET received_at = ?, data = ? WHERE key = ?",
                (str(record.get("received_at", "") or ""), json.dumps(record), key),
            )
            connection.execute("COMMIT")
        except Exception:
            connection.execute("ROLLBACK")
            raise
    return record


def delete_entry(key: str, path: Path | None = None) -> None:
    with closing(_connect(path)) as connection:
        cursor = connection.execute("DELETE FROM entries WHERE key = ?", (key,))
        if cursor.rowcount == 0:
            raise ValueError(f"Job intake was not found: {key}")
