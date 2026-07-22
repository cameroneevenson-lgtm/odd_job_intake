"""One-off job intake orchestration: folder/RPD creation, PO hint extraction,
RADAN import CSV building, and block transfer - no Qt in this module.

RADAN work is reached through the embedded truck_nest_explorer services module
(passed in as a parameter), matching how the rest of master_app borrows sibling
app behavior instead of importing it directly.
"""

from __future__ import annotations

import csv
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re
import shutil
from typing import Any

import job_intake_registry
from paths import (
    BATTLESHIELD_ROOT,
    EXPLORER_TEMPLATE_PATH,
    INVENTOR_TO_RADAN_DIR,
    JOB_PREFIX_TO_ROOT,
    MACHINE_EIA_BATTLESHIELD_ROOT,
)


FALLBACK_MATERIALS = (
    "Aluminum 3003 CHK",
    "Aluminum 5052",
    "Mild Steel-A36",
    "Stainless Steel",
)

# description_rules.csv shows a strict material-family -> strategy correlation;
# the user picks a material and the strategy follows without being asked.
MATERIAL_DEFAULT_STRATEGY = {
    "Aluminum 3003 CHK": "Air",
    "Aluminum 5052": "Air",
    "Mild Steel-A36": "O2",
    "Stainless Steel": "N2",
}

UNIT_CHOICES = ("in", "mm", "swg")

# Every sampled PFF PO number is 4 digits, dash, 3 digits, optionally prefixed
# (PO-, Q-) and/or revision-suffixed (-R1).
PO_NUMBER_PATTERN = re.compile(r"(?:PO-|Q-)?\d{4}-\d{3}(?:-R\d+)?")

# PyMuPDF linearizes the PO table one cell per text line: a filled row comes
# through as three consecutive lines ("<line#>", "<qty>", "<description>"),
# while empty rows collapse to bare line numbers. Matching is therefore
# description-anchored: find the stem-matching description line, then read
# the qty from the integer lines immediately above it.

PO_DATE_PATTERN = re.compile(
    r"(January|February|March|April|May|June|July|August|September|October|November|December)"
    r"\s+\d{1,2},\s+\d{4}"
)


class JobIntakeError(RuntimeError):
    pass


@dataclass(frozen=True)
class JobPaths:
    """Resolved on-disk shape for a one-off job, mirroring the shop's real
    convention (e.g. M59919): DXFs flat in the intake dir, next to an inner
    project dir holding the RPD and its nests/remnants folders."""

    job_number: str
    label: str | None
    root_name: str
    release_root: Path
    job_dir: Path
    intake_dir: Path
    project_name: str
    project_dir: Path
    rpd_path: Path


def resolve_job_root(job_number: str) -> str:
    number = str(job_number or "").strip().upper()
    if len(number) < 2 or not number[1:].isdigit():
        raise JobIntakeError(
            f"Job numbers look like M59919 (prefix letter + digits); got: {job_number!r}"
        )
    root_name = JOB_PREFIX_TO_ROOT.get(number[0])
    if root_name is None:
        prefixes = ", ".join(sorted(JOB_PREFIX_TO_ROOT))
        raise JobIntakeError(f"Job prefix {number[0]!r} is not one of: {prefixes}")
    return root_name


def job_folder_exists(job_number: str) -> bool:
    number = str(job_number or "").strip().upper()
    root_name = resolve_job_root(number)
    return (BATTLESHIELD_ROOT / root_name / number).exists()


def resolve_job_paths(job_number: str, label: str | None = None) -> JobPaths:
    number = str(job_number or "").strip().upper()
    root_name = resolve_job_root(number)
    release_root = BATTLESHIELD_ROOT / root_name
    job_dir = release_root / number
    label_text = str(label or "").strip()

    if label_text:
        # Existing job number (real truck or prior one-off): the one-off nests
        # under its own label folder, mirroring the truck/kit naming shape.
        intake_dir = job_dir / label_text
        project_name = f"{number} {label_text}"
    else:
        intake_dir = job_dir
        project_name = number

    project_dir = intake_dir / project_name
    return JobPaths(
        job_number=number,
        label=label_text or None,
        root_name=root_name,
        release_root=release_root,
        job_dir=job_dir,
        intake_dir=intake_dir,
        project_name=project_name,
        project_dir=project_dir,
        rpd_path=project_dir / f"{project_name}.rpd",
    )


def create_job_folders(paths: JobPaths) -> None:
    if paths.label is None and paths.job_dir.exists():
        raise JobIntakeError(
            f"{paths.job_dir} already exists - give this one-off a Label so it "
            "gets its own subfolder instead of mixing into the existing job."
        )
    if paths.label is not None and paths.intake_dir.exists():
        raise JobIntakeError(f"{paths.intake_dir} already exists - pick a different Label.")
    for folder in (paths.intake_dir, paths.project_dir, paths.project_dir / "nests", paths.project_dir / "remnants"):
        folder.mkdir(parents=True, exist_ok=True)


def copy_attachments(paths: JobPaths, source_files: list[Path]) -> list[dict[str, Any]]:
    """DXFs and reference docs (PO PDFs etc.) land flat in the intake dir,
    matching the M59919 convention."""
    attachments: list[dict[str, Any]] = []
    for source in source_files:
        source = Path(source)
        if not source.exists():
            raise JobIntakeError(f"Attachment was not found: {source}")
        target = paths.intake_dir / source.name
        shutil.copy2(source, target)
        attachments.append(
            {"filename": source.name, "saved_path": str(target), "size": target.stat().st_size}
        )
    return attachments


def create_intake(
    job_number: str,
    label: str | None,
    files: list[Path],
    *,
    source: str = "manual",
    email_subject: str = "",
    email_sender: str = "",
) -> dict[str, Any]:
    """Create the job folder, copy attachments in, scrape whatever the PO PDFs
    give up, seed one part row per DXF, and register the intake.

    This is the whole intake sequence and the only copy of it: the desktop page
    and the 127.0.0.1 listener both call this, differing only in ``source``.
    It deliberately stops short of any RADAN work - cloning the RPD, importing
    parts, and sending blocks stay explicit user actions on the desktop page,
    so an Outlook-driven intake can never block on RADAN COM automation.
    """
    paths = resolve_job_paths(job_number, label or None)
    create_job_folders(paths)
    attachments = copy_attachments(paths, list(files))

    dxf_names = [
        attachment["filename"]
        for attachment in attachments
        if str(attachment["filename"]).casefold().endswith(".dxf")
    ]
    dxf_stems = [Path(name).stem for name in dxf_names]

    po_number: str | None = None
    due_date: date | None = None
    due_note: str | None = None
    line_items: dict[str, dict[str, Any]] = {}
    unmatched: list[str] = []
    for attachment in attachments:
        filename = str(attachment["filename"])
        if filename.casefold().endswith(".dxf") or not filename.casefold().endswith(".pdf"):
            continue
        hints = extract_po_hints(Path(attachment["saved_path"]), dxf_stems)
        po_number = po_number or hints.po_number
        due_date = due_date or hints.due_date
        due_note = due_note or hints.due_note
        for stem, hint in hints.line_items.items():
            line_items.setdefault(stem, hint)
        unmatched.extend(line for line in hints.unmatched_lines if line not in unmatched)

    material_qty = []
    for name in dxf_names:
        hint = line_items.get(Path(name).stem, {})
        material_qty.append(
            {
                "filename": name,
                "material": "",
                "thickness": 0.0,
                "qty": int(hint.get("qty") or 1),
                "unit": "in",
                "strategy": "",
                "po_ref": str(hint.get("raw_description", "") or ""),
            }
        )

    entry = job_intake_registry.new_entry(
        job_number=job_number,
        label=label or None,
        source=source,
        email_subject=email_subject,
        email_sender=email_sender,
    )
    entry["job_folder"] = str(paths.intake_dir)
    entry["po_number"] = po_number
    entry["due_date"] = due_date.isoformat() if due_date else None
    entry["due_note"] = due_note
    entry["attachments"] = attachments
    entry["material_qty"] = material_qty
    entry["po_unmatched"] = unmatched
    job_intake_registry.append_entry(entry)
    return entry


# --- RPD template clone ------------------------------------------------------
# Modeled on truck_nest_explorer's proven template-clone approach (literal XML
# text substitution on the raw template bytes, no RADAN needed), re-keyed from
# truck/kit to job/label. Kept independent of that app's private helpers.

_TEMPLATE_ENCODINGS = ("utf-8-sig", "utf-16", "utf-16-le", "utf-16-be", "cp1252")


def _decode_template_bytes(data: bytes) -> str:
    for encoding in _TEMPLATE_ENCODINGS:
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise JobIntakeError("The RPD template has an unsupported text encoding.")


def _replace_exact_xml_text(text: str, old_value: str, new_value: str) -> str:
    pattern = rf">(\s*){re.escape(old_value)}(\s*)<"
    return re.sub(pattern, lambda match: f">{match.group(1)}{new_value}{match.group(2)}<", text)


def _replace_xml_element_value(text: str, tag_name: str, new_value: str) -> str:
    pattern = rf"(<{tag_name}>)(.*?)(</{tag_name}>)"
    return re.sub(pattern, lambda match: f"{match.group(1)}{new_value}{match.group(3)}", text, flags=re.DOTALL)


def clone_rpd_template(paths: JobPaths, template_path: Path = EXPLORER_TEMPLATE_PATH) -> Path:
    if paths.rpd_path.exists():
        raise JobIntakeError(f"An RPD already exists: {paths.rpd_path}")
    if not template_path.exists():
        raise JobIntakeError(f"The RPD template is missing: {template_path}")

    text = _decode_template_bytes(template_path.read_bytes())
    text = _replace_exact_xml_text(text, "Template.rpd", paths.rpd_path.name)
    text = _replace_exact_xml_text(text, "Template", paths.project_name)
    text = _replace_xml_element_value(text, "JobName", paths.project_name)
    text = _replace_xml_element_value(text, "NestFolder", str(paths.project_dir / "nests"))
    text = _replace_xml_element_value(text, "RemnantSaveFolder", str(paths.project_dir / "remnants"))
    text = re.sub(r'encoding="[^"]+"', 'encoding="utf-8"', text, count=1)

    paths.project_dir.mkdir(parents=True, exist_ok=True)
    (paths.project_dir / "nests").mkdir(exist_ok=True)
    (paths.project_dir / "remnants").mkdir(exist_ok=True)
    paths.rpd_path.write_text(text, encoding="utf-8")
    return paths.rpd_path


# --- Material list -----------------------------------------------------------


EXPECTED_DESCRIPTIONS_FILENAME = "expected_laser_descriptions.csv"
DESCRIPTION_RULES_FILENAME = "description_rules.csv"


# Resolved at call time, not as module-level constants. Binding these at
# import would freeze the directory (defeating monkeypatching in tests) and,
# more importantly, is the same class of bug the registry invariant warns
# about. Everything below re-reads from disk on every call so that materials
# added to the shop's CSV appear without restarting anything.
def _expected_descriptions_path() -> Path:
    return INVENTOR_TO_RADAN_DIR / EXPECTED_DESCRIPTIONS_FILENAME


def _description_rules_path() -> Path:
    return INVENTOR_TO_RADAN_DIR / DESCRIPTION_RULES_FILENAME


def _read_expected_descriptions() -> list[str]:
    """The shop's authoritative list of laser descriptions.

    If a description isn't in this file it doesn't exist as far as intake is
    concerned. The file is maintained outside this repo and changes, so it is
    read on demand rather than cached at import.
    """
    descriptions: list[str] = []
    try:
        with _expected_descriptions_path().open(newline="", encoding="utf-8-sig") as handle:
            for row in csv.DictReader(handle):
                text = str(row.get("Description", "") or "").strip()
                if text:
                    descriptions.append(text)
    except OSError:
        return []
    return descriptions


def _read_description_rules() -> dict[str, dict[str, Any]]:
    """description -> {material, thickness, strategy}, keyed for lookup.

    This file translates a description into the values RADAN needs; the
    expected-descriptions file decides which of them are offered.
    """
    rules: dict[str, dict[str, Any]] = {}
    try:
        with _description_rules_path().open(newline="", encoding="utf-8-sig") as handle:
            for row in csv.DictReader(handle):
                description = str(row.get("Description", "") or "").strip()
                material = str(row.get("Material", "") or "").strip()
                if not description or not material:
                    continue
                try:
                    thickness = float(str(row.get("Thickness", "") or "").strip())
                except ValueError:
                    continue
                if thickness <= 0:
                    continue
                rules[_normalize_match_key(description)] = {
                    "material": material,
                    "thickness": thickness,
                    "strategy": str(row.get("Strategy", "") or "").strip(),
                }
    except OSError:
        return {}
    return rules


def material_thickness_catalog() -> dict[str, tuple[float, ...]]:
    """material -> the thicknesses actually available for it, ascending.

    Built by taking every description the shop says is valid and translating
    it through the rules table. A material/thickness pair the shop doesn't
    stock therefore can't be picked at all, which is the point.

    Read fresh on every call: both CSVs are maintained outside this repo and
    are expected to change, so the grid must reflect the file as it is now
    rather than as it was at import time.
    """
    rules = _read_description_rules()
    catalog: dict[str, set[float]] = {}
    for description in _read_expected_descriptions():
        rule = rules.get(_normalize_match_key(description))
        if rule is None:
            continue
        material = str(rule["material"])
        # FTQ is a forced per-part override elsewhere (ftq_parts.csv), never a
        # choice the user makes here.
        if "FTQ" in material.upper():
            continue
        catalog.setdefault(material, set()).add(float(rule["thickness"]))

    if not catalog:
        # The files are missing or reshaped; fall back so intake still works.
        return {material: () for material in FALLBACK_MATERIALS}
    return {
        material: tuple(sorted(thicknesses))
        for material, thicknesses in sorted(catalog.items(), key=lambda pair: pair[0].casefold())
    }


def material_choices() -> tuple[str, ...]:
    """Materials the shop's expected-descriptions file actually allows."""
    return tuple(material_thickness_catalog().keys())


def thickness_choices(material: str) -> tuple[float, ...]:
    """Thicknesses available for one material, per the expected descriptions."""
    return material_thickness_catalog().get(str(material or "").strip(), ())


def default_strategy_for_material(material: str) -> str:
    exact = MATERIAL_DEFAULT_STRATEGY.get(material)
    if exact:
        return exact
    lowered = material.casefold()
    if "stainless" in lowered:
        return "N2"
    if "steel" in lowered:
        return "O2"
    return "Air"


# --- PO hint extraction ------------------------------------------------------


@dataclass(frozen=True)
class POHints:
    po_number: str | None
    due_date: date | None
    # Date Required sometimes carries urgency text (RUSH, ASAP) instead of a
    # date - surfaced so the UI can show it, since it can't become a due_date.
    due_note: str | None
    # dxf stem -> {"qty": int, "raw_description": str}
    line_items: dict[str, dict[str, Any]]
    # Table rows that matched no attached DXF. Every real PO line should have
    # at least a partial DXF match, so anything here is worth showing the
    # user - a missing attachment, or an order-wide note like
    # "ALL MATERIAL 1/8 MILD STEEL".
    unmatched_lines: tuple[str, ...] = ()


# Footer/header labels that follow the table's bare line numbers in the text
# and would otherwise be mistaken for line items.
_PO_NON_ITEM_PREFIXES = (
    "sub-total",
    "subtotal",
    "sales tax",
    "total",
    "prepared",
    "special instructions",
    "laser order",
    "laser quote",
    "ship to",
    "drawing number",
    "description",
    "pricing",
)


def _normalize_match_key(text: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(text or "").casefold())


def _parse_po_date(text: str) -> date | None:
    try:
        return datetime.strptime(text, "%B %d, %Y").date()
    except ValueError:
        return None


def extract_po_hints(pdf_path: Path, dxf_stems: list[str]) -> POHints:
    """Best-effort scrape of a customer PO PDF. Real POs vary wildly (blank
    Drawing Number columns, bare part codes, RUSH instead of a date, typos),
    so every step degrades to None/empty instead of raising."""
    try:
        import fitz
    except ImportError as exc:
        raise JobIntakeError("PyMuPDF (fitz) is required for PO extraction.") from exc

    try:
        with fitz.open(str(pdf_path)) as doc:
            page_texts = [str(page.get_text("text") or "") for page in doc]
    except Exception:
        return POHints(po_number=None, due_date=None, due_note=None, line_items={}, unmatched_lines=())
    full_text = "\n".join(page_texts)

    po_number = None
    po_match = PO_NUMBER_PATTERN.search(full_text)
    if po_match is not None:
        po_number = po_match.group(0)

    # The order date and due date both render as "Month D, YYYY"; extraction
    # order is unreliable, so only claim a due date when at least two dates
    # are present (due = latest). A lone date is usually just the order date.
    dates = sorted(
        {parsed for match in PO_DATE_PATTERN.finditer(full_text) if (parsed := _parse_po_date(match.group(0)))}
    )
    due_date = dates[-1] if len(dates) >= 2 else None
    due_note = None
    if due_date is None:
        note_match = re.search(r"\b(RUSH|ASAP)\b", full_text, re.IGNORECASE)
        if note_match is not None:
            due_note = note_match.group(1).upper()

    # Assign each description line to the longest-prefix-matching DXF stem so
    # "End Cap D" doesn't swallow "End Cap D_2"'s row. A stem can also match
    # its Drawing Number cell later in the text ("Clip-End.DXF"), which has no
    # qty above it - prefer the first match that carries a qty.
    normalized_stems = sorted(
        ((stem, _normalize_match_key(stem)) for stem in dxf_stems if _normalize_match_key(stem)),
        key=lambda pair: len(pair[1]),
        reverse=True,
    )
    lines = [line.strip() for line in full_text.splitlines()]
    line_items: dict[str, dict[str, Any]] = {}
    unmatched: list[str] = []
    for index, line in enumerate(lines):
        normalized_line = _normalize_match_key(line)
        if not normalized_line:
            continue
        matched = next(
            (stem for stem, normalized_stem in normalized_stems if normalized_line.startswith(normalized_stem)),
            None,
        )
        # A filled table row linearizes as "<line#>", "<qty>", "<description>";
        # only claim a qty when both integer cells are present above, otherwise
        # a bare table line number would be misread as the quantity.
        qty = None
        if index >= 2 and lines[index - 1].isdigit() and lines[index - 2].isdigit():
            qty = int(lines[index - 1])
        if matched is not None:
            existing = line_items.get(matched)
            if existing is None or (existing.get("qty") is None and qty is not None):
                line_items[matched] = {"qty": qty, "raw_description": line}
            continue
        # Unmatched-candidate detection: table content follows a digit cell
        # (its line number or qty), has letters, and isn't a footer label.
        lowered = line.casefold()
        if (
            index >= 1
            and lines[index - 1].isdigit()
            and len(line) >= 3
            and any(char.isalpha() for char in line)
            and not any(lowered.startswith(prefix) for prefix in _PO_NON_ITEM_PREFIXES)
            and _parse_po_date(line) is None
        ):
            unmatched.append(line)

    return POHints(
        po_number=po_number,
        due_date=due_date,
        due_note=due_note,
        line_items=line_items,
        unmatched_lines=tuple(dict.fromkeys(unmatched)),
    )


# --- RADAN import CSV --------------------------------------------------------


def build_import_csv_rows(entry: dict[str, Any]) -> list[list[str]]:
    """The exact 6-column, headerless shape radan_automation's
    read_import_csv expects: dxf_path, qty, material, thickness, unit, strategy."""
    attachments_by_name = {
        str(attachment.get("filename", "")): str(attachment.get("saved_path", ""))
        for attachment in entry.get("attachments", [])
    }
    rows: list[list[str]] = []
    problems: list[str] = []
    for part in entry.get("material_qty", []):
        filename = str(part.get("filename", "") or "")
        saved_path = attachments_by_name.get(filename, "")
        material = str(part.get("material", "") or "").strip()
        thickness = part.get("thickness")
        qty = part.get("qty")
        if not saved_path or not Path(saved_path).exists():
            problems.append(f"{filename}: the copied DXF is missing")
            continue
        if not material:
            problems.append(f"{filename}: pick a material")
            continue
        if not isinstance(thickness, (int, float)) or thickness <= 0:
            problems.append(f"{filename}: thickness must be greater than zero")
            continue
        if not isinstance(qty, int) or qty <= 0:
            problems.append(f"{filename}: quantity must be at least 1")
            continue
        unit = str(part.get("unit", "") or "in").strip().casefold()
        strategy = str(part.get("strategy", "") or "").strip() or default_strategy_for_material(material)
        rows.append([saved_path, str(qty), material, str(thickness), unit, strategy])
    if problems:
        raise JobIntakeError("Fix these part rows first:\n" + "\n".join(problems))
    if not rows:
        raise JobIntakeError("There are no DXF part rows to import.")
    return rows


def write_import_csv(rows: list[list[str]], csv_path: Path) -> Path:
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    with csv_path.open("w", newline="", encoding="utf-8") as handle:
        csv.writer(handle).writerows(rows)
    return csv_path


def launch_radan_import(
    explorer_services: Any,
    *,
    paths: JobPaths,
    csv_path: Path,
    log_path: Path,
) -> Any:
    """Real RADAN COM conversion via the proven headless import pipeline;
    symbols land flat in the intake dir next to their DXFs (M59919 shape).
    Never uses the experimental lab symbol writer."""
    return explorer_services.launch_radan_csv_import(
        csv_path,
        paths.intake_dir,
        project_path=paths.rpd_path,
        log_path=log_path,
    )


def send_job_blocks_to_machine(
    explorer_services: Any,
    *,
    paths: JobPaths,
    progress_cb: Any = None,
    should_cancel_cb: Any = None,
) -> Any:
    """Same verified copy-then-delete block transfer the truck workflow uses,
    pointed at this job's prefix root on both the L: and machine sides."""
    return explorer_services.send_project_block_files_to_machine(
        paths.project_dir,
        paths.release_root,
        machine_root=MACHINE_EIA_BATTLESHIELD_ROOT / paths.root_name,
        progress_cb=progress_cb,
        should_cancel_cb=should_cancel_cb,
    )
