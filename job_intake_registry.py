"""Persistent registry of one-off job intakes.

Pure data-access module - no Qt, no Flask - so the desktop Job Intake page
and the (future) Outlook-facing listener can share one store. Both writers
must go through this module's atomic save; the file is also hand-inspectable
JSON like the rest of _runtime.
"""

from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path
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


def _empty_store() -> dict[str, Any]:
    return {"version": 1, "entries": []}


def _load_store(path: Path) -> dict[str, Any]:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (FileNotFoundError, OSError, json.JSONDecodeError):
        return _empty_store()
    return payload if isinstance(payload, dict) else _empty_store()


def _save_store(store: dict[str, Any], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(".json.tmp")
    temp_path.write_text(json.dumps(store, indent=2, sort_keys=True), encoding="utf-8")
    temp_path.replace(path)


def _entry_records(store: dict[str, Any]) -> list[dict[str, Any]]:
    records = store.get("entries", [])
    if not isinstance(records, list):
        return []
    return [record for record in records if isinstance(record, dict)]


def _resolve_path(path: Path | None) -> Path:
    # Resolve the module global at call time (not as a default arg) so tests
    # that monkeypatch JOB_INTAKE_REGISTRY_PATH are honored.
    return path if path is not None else JOB_INTAKE_REGISTRY_PATH


def load_entries(path: Path | None = None) -> list[dict[str, Any]]:
    """Newest first, matching how the queue is read in the UI."""
    records = _entry_records(_load_store(_resolve_path(path)))
    return sorted(records, key=lambda record: str(record.get("received_at", "")), reverse=True)


def get_entry(key: str, path: Path | None = None) -> dict[str, Any] | None:
    for record in _entry_records(_load_store(_resolve_path(path))):
        if str(record.get("key", "")) == key:
            return record
    return None


def append_entry(entry: dict[str, Any], path: Path | None = None) -> dict[str, Any]:
    store_path = _resolve_path(path)
    key = str(entry.get("key", "") or "")
    if not key:
        raise ValueError("Registry entries need a key (use new_entry()).")
    store = _load_store(store_path)
    records = _entry_records(store)
    if any(str(record.get("key", "")) == key for record in records):
        raise ValueError(f"A job intake already exists for {key}.")
    records.append(entry)
    store["entries"] = records
    _save_store(store, store_path)
    return entry


def update_entry(key: str, path: Path | None = None, **fields: Any) -> dict[str, Any]:
    store_path = _resolve_path(path)
    status = fields.get("status")
    if status is not None and status not in VALID_STATUSES:
        raise ValueError(f"Invalid intake status: {status}")
    store = _load_store(store_path)
    records = _entry_records(store)
    for record in records:
        if str(record.get("key", "")) == key:
            record.update(fields)
            record["updated_at"] = _now_text()
            store["entries"] = records
            _save_store(store, store_path)
            return record
    raise ValueError(f"Job intake was not found: {key}")


def delete_entry(key: str, path: Path | None = None) -> None:
    store_path = _resolve_path(path)
    store = _load_store(store_path)
    records = _entry_records(store)
    remaining = [record for record in records if str(record.get("key", "")) != key]
    if len(remaining) == len(records):
        raise ValueError(f"Job intake was not found: {key}")
    store["entries"] = remaining
    _save_store(store, store_path)
