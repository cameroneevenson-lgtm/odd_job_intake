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
import json
from pathlib import Path
import re
import shutil
from typing import Any

import job_intake_registry
from paths import (
    APP_DIR,
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

# What gets pulled in when an email points at a folder instead of attaching
# files. Mirrors the attachment allow-list: geometry, the prints and BOM the
# scrapes read, and a spreadsheet BOM that inventor_to_radan can consume.
INGESTED_SUFFIXES = (".dxf", ".pdf", ".csv", ".xlsx")

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


# Drive-letter or UNC paths in free text. Kept here rather than in the Outlook
# macro so the rules can change without re-importing and re-signing VBA on
# every machine - the macro sends the raw body and Python decides.
# Runs to the end of the line rather than to the first space, because real
# shop paths contain spaces - "W:\LASER\For Battleshield Fabrication" is one
# of this shop's own roots. The trailing words are trimmed back below.
_PATH_IN_TEXT = re.compile(
    r"(?:[A-Za-z]:[\\/]|\\\\[^\s\\/]+[\\/])[^<>\"'|?*\r\n]*",
)


def paths_in_text(text: str) -> list[str]:
    """Folder/file paths mentioned in an email body, longest first.

    A path can't be delimited reliably in prose - it may contain spaces and may
    be followed by more words - so each match yields the full run plus every
    shorter version of it with trailing words removed. The caller keeps
    whichever actually exists, which is the only real test.
    """
    found: list[str] = []
    for match in _PATH_IN_TEXT.finditer(str(text or "")):
        candidate = match.group(0).strip()
        while candidate and len(candidate) > 3:
            trimmed = candidate.rstrip(" .,;:)]}>")
            if trimmed not in found:
                found.append(trimmed)
            # Drop the last whitespace-separated word and try again.
            if " " not in trimmed.strip():
                break
            candidate = trimmed.rsplit(" ", 1)[0]
    # Longest first so a full path wins over the parent folder it contains.
    return sorted(found, key=len, reverse=True)


def create_intake(
    job_number: str,
    label: str | None,
    files: list[Path],
    *,
    source: str = "manual",
    email_subject: str = "",
    email_sender: str = "",
    email_body: str = "",
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

    # Some jobs arrive as a path on W: instead of attachments - "please
    # manufacture the parts at the following path". Pull the work files in so
    # everything downstream sees one job folder, exactly as if they had been
    # attached.
    collected = list(files)
    ingested_from: list[str] = []
    for candidate in paths_in_text(email_body):
        folder = Path(candidate)
        if not folder.is_dir():
            continue
        pulled = [
            item
            for item in sorted(folder.iterdir())
            if item.is_file() and item.suffix.casefold() in INGESTED_SUFFIXES
        ]
        if pulled:
            collected.extend(pulled)
            ingested_from.append(str(folder))
        # Only the first folder that actually holds work files; the candidate
        # list contains parent paths of it too.
        break

    attachments = copy_attachments(paths, collected)

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

    attachment_paths = {
        str(attachment["filename"]): Path(str(attachment["saved_path"]))
        for attachment in attachments
    }

    # A parts list, if one came with the job. Its DESCRIPTION column is the
    # shop's own description verbatim, so it resolves material/thickness
    # exactly rather than by flattening - and it carries the quantity the
    # prints defer to ("QTY: AS PER BOM").
    bom_rows: dict[str, BomRow] = {}
    available_materials = set(material_choices())
    for filename, path in attachment_paths.items():
        if not filename.casefold().endswith(".pdf"):
            continue
        found = extract_bom_rows(path, [Path(name).stem for name in dxf_names])
        if found:
            bom_rows = found
            break

    # The PO scrape ran over the BOM and the prints too, reporting their rows
    # as "unmatched PO lines". Those lines were understood, just by a different
    # reader, so drop them before any real warnings are added below - filtering
    # afterwards would also strip warnings that legitimately mention the BOM.
    if bom_rows:
        accounted = {_normalize_match_key(row.description) for row in bom_rows.values()}
        accounted.update(_normalize_match_key(row.part) for row in bom_rows.values())
        unmatched = [
            line
            for line in unmatched
            if _normalize_match_key(line) not in accounted
            and not _QTY_DEFERRED.search(line)
        ]

    material_qty = []
    for name in dxf_names:
        hint = line_items.get(Path(name).stem, {})
        po_qty = hint.get("qty")

        # Second pass over the DXF's own text. The PO always wins where it
        # said something; this only fills gaps, which is the common case for
        # jobs emailed without a PO at all.
        dxf_hints = extract_dxf_hints(attachment_paths[name])

        # Third pass: the drawing print sharing this DXF's stem. In practice
        # this is the richest source - laser DXFs are usually pure geometry
        # while the print's title block carries MATERIAL, GAUGE and QTY.
        print_hints = PrintHints(material=None, thickness=None, qty=None)
        stem = Path(name).stem

        # Try the identically-named PDF first, then any other PDF, looking for
        # the page whose title block names this part. A per-part print often
        # doesn't exist, and a drawing set is routinely one multi-page PDF with
        # a different part on each page.
        ordered_pdfs = sorted(
            (
                (other, path)
                for other, path in attachment_paths.items()
                if other.casefold().endswith(".pdf")
            ),
            key=lambda pair: Path(pair[0]).stem.casefold() != stem.casefold(),
        )
        for other, path in ordered_pdfs:
            found = extract_print_hints(path, part_stem=stem)
            if found.material is not None or found.qty is not None or found.qty_unknown:
                print_hints = found
                break

        # The BOM wins: its description is the shop's own string, not a
        # customer's wording that had to be interpreted.
        bom_row = bom_rows.get(Path(name).stem)
        bom_material = bom_thickness = None
        if bom_row is not None and bom_row.material:
            if bom_row.material in available_materials:
                bom_material = bom_row.material
                bom_thickness = bom_row.thickness
            else:
                note = (
                    f"{name}: the BOM asks for {bom_row.material} "
                    f'("{bom_row.description}"), which isn\'t in the shop\'s laser list'
                )
                if note not in unmatched:
                    unmatched.append(note)

        material = bom_material or print_hints.material or dxf_hints.material
        thickness = bom_thickness or print_hints.thickness or dxf_hints.thickness
        material_source = (
            (bom_row.description if bom_material else None)
            or print_hints.material_source_text
            or dxf_hints.material_source_text
        )

        # A print that names a material/thickness the catalog doesn't list is
        # worth saying out loud - it usually means the shop's own CSV hasn't
        # caught up, and silently dropping it looks like the parser failing.
        if (
            print_hints.material is not None
            and print_hints.thickness is None
            and print_hints.thickness_source_text
        ):
            note = (
                f"{name}: the print asks for {print_hints.thickness_source_text} "
                f"{print_hints.material}, which isn't in the shop's laser list"
            )
            if note not in unmatched:
                unmatched.append(note)

        # Quantity precedence: the PO is a commitment, the print is a drawing
        # note. Neither may be invented - a blank QTY or a "REFER TO BOM" is
        # recorded as unknown rather than defaulting to 1, because silently
        # cutting one of a forty-off part is the expensive failure here.
        resolved_qty = (
            po_qty
            or (bom_row.qty if bom_row is not None else None)
            or print_hints.qty
            or dxf_hints.qty
        )
        qty_unknown = resolved_qty is None

        material_qty.append(
            {
                "filename": name,
                "material": material or "",
                "thickness": thickness or 0.0,
                "qty": int(resolved_qty or 1),
                "unit": "in",
                "strategy": default_strategy_for_material(material) if material else "",
                "po_ref": str(hint.get("raw_description", "") or ""),
                # Shown as reference so the user can see what the drawing said,
                # the same way po_ref shows what the PO said.
                "dxf_ref": " | ".join(dxf_hints.raw_lines),
                # The customer's own wording for the material, kept beside the
                # prediction so the translation is visible and checkable.
                "material_source_text": material_source or "",
                # A prediction is never trusted on its own: the user has to
                # confirm it before parts can be sent to RADAN. Rows with no
                # prediction start unconfirmed too - picking a material by hand
                # is itself the confirmation.
                "material_confirmed": False,
                # True when nothing stated a quantity, so the 1 above is a
                # placeholder rather than an answer.
                "qty_unknown": qty_unknown,
                "qty_source_text": str(print_hints.qty_source_text or ""),
            }
        )

    entry = job_intake_registry.new_entry(
        job_number=job_number,
        label=label or None,
        source=source,
        email_subject=email_subject,
        email_sender=email_sender,
    )
    # Kept whole so later passes (material/qty scraped from the message, or a
    # W: folder named in it) can re-read it without the email being needed
    # again. The macro sends it raw precisely so this stays changeable here.
    entry["email_body"] = email_body
    # Keep only paths that exist, and drop any that is merely a parent of
    # another kept one - the candidate list deliberately includes prefixes.
    existing = [path for path in paths_in_text(email_body) if Path(path).exists()]
    entry["ingested_from"] = ingested_from
    entry["source_paths"] = [
        path
        for path in existing
        if not any(other != path and other.startswith(path) for other in existing)
    ]
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
                    # Kept so callers can tokenise the shop's own wording -
                    # the key is normalised beyond recognition.
                    "description": description,
                    "material": material,
                    "thickness": thickness,
                    "strategy": str(row.get("Strategy", "") or "").strip(),
                }
    except OSError:
        return {}
    return rules


def material_thickness_catalog() -> dict[str, tuple[float, ...]]:
    """material -> the thicknesses available for it, ascending.

    Built from description_rules.csv, which is the mapping the CAM side
    actually uses. expected_laser_descriptions.csv is deliberately *not*
    consulted: it lists a much narrower set (14 rows against 36) and gating on
    it rejected materials that turn up in real BOMs - 3003 checker plate came
    through a customer parts list and was refused for not being on that list.

    Read fresh on every call: the file is maintained outside this repo and is
    expected to change, so the grid must reflect it as it is now rather than
    as it was at import time.
    """
    # Built from the de-duped Material/Thickness/Strategy columns. The
    # Description column is only a lookup key - its wording varies per
    # customer and it says nothing about what the shop stocks.
    rules = _read_description_rules()
    # Grouped case-insensitively: the file spells the same material more than
    # one way (it does the same with strategies - AIR vs Air), and two casings
    # of one material would otherwise become two entries in the drop-down.
    grouped: dict[str, set[float]] = {}
    spellings: dict[str, dict[str, int]] = {}
    for rule in rules.values():
        material = str(rule["material"]).strip()
        # FTQ is a forced per-part override elsewhere (ftq_parts.csv), never a
        # choice the user makes here.
        if "FTQ" in material.upper():
            continue
        key = material.casefold()
        grouped.setdefault(key, set()).add(float(rule["thickness"]))
        spellings.setdefault(key, {})
        spellings[key][material] = spellings[key].get(material, 0) + 1

    catalog: dict[str, set[float]] = {}
    for key, thicknesses in grouped.items():
        # The most common spelling wins, so the grid shows what the file mostly
        # says rather than whichever row happened to be read first.
        best = max(spellings[key].items(), key=lambda pair: pair[1])[0]
        catalog[best] = thicknesses

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


def _strategies_from_rules() -> dict[str, str]:
    """material -> strategy, as recorded in inventor_to_radan's rules table.

    That file is the truth for what the CAM side expects, so it decides which
    strategy a material gets rather than anything hardcoded here.
    """
    counts: dict[str, dict[str, int]] = {}
    for rule in _read_description_rules().values():
        material = str(rule.get("material", "") or "").strip()
        strategy = str(rule.get("strategy", "") or "").strip()
        if material and strategy:
            counts.setdefault(material, {})
            counts[material][strategy] = counts[material].get(strategy, 0) + 1

    resolved: dict[str, str] = {}
    for material, spellings in counts.items():
        winner = max(spellings.items(), key=lambda pair: pair[1])[0]
        # The file spells the same strategy inconsistently (AIR vs Air). Where
        # our known-good constant matches case-insensitively, keep its casing:
        # the truth decides *which* strategy, not how to capitalise it, and
        # this casing is what the RADAN import was verified against.
        known = MATERIAL_DEFAULT_STRATEGY.get(material)
        if known and known.casefold() == winner.casefold():
            winner = known
        resolved[material] = winner
    return resolved


def default_strategy_for_material(material: str) -> str:
    material = str(material or "").strip()
    from_truth = _strategies_from_rules().get(material)
    if from_truth:
        return from_truth
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


# --- DXF text extraction (fallback when the PO gives nothing) ----------------
#
# Second-best source for qty/material after the PO, and the only one available
# when a DXF arrives with no accompanying document. Text in the drawing is the
# designer's own value rather than a salesperson's retyping, so where it is
# unambiguous it is more trustworthy than the PO - but it is still best-effort
# and never overrides anything the PO or the user supplied.


@dataclass(frozen=True)
class DXFHints:
    material: str | None       # canonical prediction, only when unambiguous
    thickness: float | None    # only when it exists for that material
    qty: int | None
    # Every line that contributed, for the grid's reference column - so the
    # user can see what the drawing actually said.
    raw_lines: tuple[str, ...] = ()
    # The customer's own wording that produced `material`. Shown next to the
    # prediction so the user is verifying a visible translation rather than
    # trusting an opaque guess.
    material_source_text: str | None = None


# 12MB covers the shop's DXFs comfortably; the cap just stops a pathological
# file from being slurped into memory during intake.
_DXF_READ_LIMIT = 12 * 1024 * 1024

_DXF_QTY_PATTERN = re.compile(
    r"\b(?:QTY|QUANTITY|REQD|REQUIRED)\b\s*[:.\-]?\s*(\d{1,4})\b", re.IGNORECASE
)
# "2 OFF", "4 PLCS", "3 REQ'D"
_DXF_QTY_SUFFIX_PATTERN = re.compile(
    r"\b(\d{1,4})\s*(?:OFF|PLCS?|PCS?|REQ'?D)\b", re.IGNORECASE
)
_DXF_THICKNESS_PATTERN = re.compile(
    # Leading-dot decimals first: drawings write ".125 THK" far more often than
    # "0.125 THK", and without this alternative the dot was skipped and the
    # value read as 125 inches.
    r"(\d+\s*/\s*\d+|\d*\.\d+|\d+)\s*(?:\"|IN\b|INCH)?\s*(?:THK|THICK)",
    re.IGNORECASE,
)


def _dxf_text_values(dxf_path: Path) -> list[str]:
    """Text carried by TEXT/MTEXT entities.

    DXF is line-pairs of group code then value; codes 1 and 3 hold entity text.
    Reading the codes directly avoids a CAD dependency for what is a flat scan.
    """
    try:
        raw = dxf_path.read_bytes()[:_DXF_READ_LIMIT]
    except OSError:
        return []
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError:
        # DXFs are commonly cp1252; latin-1 never fails and preserves bytes.
        text = raw.decode("latin-1", errors="replace")

    lines = text.splitlines()
    values: list[str] = []
    for index in range(0, len(lines) - 1):
        code = lines[index].strip()
        if code in ("1", "3"):
            value = lines[index + 1].strip()
            # MTEXT carries formatting codes like \pxqc; \fArial|b0; - strip
            # them so the words survive but the markup doesn't.
            value = re.sub(r"\\[A-Za-z][^;\\]*;?", " ", value)
            value = re.sub(r"[{}]", " ", value)
            value = re.sub(r"\s+", " ", value).strip()
            if value:
                values.append(value)
    return values


# Structural words shared by every description; they identify no material.
_DESCRIPTION_STOPWORDS = frozenset(
    {"plate", "sheet", "thk", "thick", "general", "inch", "steel", "aly"}
)

MATERIAL_ALIASES_FILENAME = "material_aliases.csv"


def _material_aliases_path() -> Path:
    return APP_DIR / MATERIAL_ALIASES_FILENAME


def _read_material_aliases() -> dict[str, str]:
    """Drawing vocabulary -> the canonical material RADAN expects.

    Every drawing speaks a different dialect - ALUM, ALUMINIUM, AL ALY, 5052,
    CRS, A-36, 44W - and all of it has to be flattened to the exact string the
    CAM side accepts. The shop's own description file only covers *its* words,
    so this table carries the customer wording that file will never contain.

    Edit the CSV to teach it a new dialect; it is read on every call, so new
    aliases take effect without a restart. An alias pointing at a material the
    authoritative catalog doesn't list is ignored - that file still decides
    what exists.
    """
    aliases: dict[str, str] = {}
    try:
        with _material_aliases_path().open(newline="", encoding="utf-8-sig") as handle:
            for row in csv.DictReader(handle):
                alias = _normalize_match_key(row.get("Alias", ""))
                material = str(row.get("Material", "") or "").strip()
                if alias and material:
                    aliases[alias] = material
    except OSError:
        return {}
    return aliases


def _material_tokens() -> dict[str, str]:
    """token -> material, for tokens that identify exactly one material.

    Column A (Description) is deliberately not read. It is only a lookup key
    whose wording varies per customer, and tokenising it harvested "new",
    "tool" and "mm" as if they identified materials - a drawing note reading
    NEW then claimed an aluminium.

    The material names in column B already carry the grade ("Aluminum 5052" ->
    5052, "Mild Steel-A36" -> a36). Anything else a drawing might say belongs
    in material_aliases.csv, where it is a deliberate entry rather than an
    accident of how somebody worded a description. Tokens shared by more than
    one material are dropped, so an ambiguous word can never decide.
    """
    candidates: dict[str, set[str]] = {}
    available = set(material_choices())

    def _add(token: str, material: str) -> None:
        token = token.casefold()
        if len(token) < 2 or token in _DESCRIPTION_STOPWORDS:
            return
        candidates.setdefault(token, set()).add(material)

    for material in available:
        for token in re.findall(r"[A-Za-z0-9]+", material):
            _add(token, material)

    for alias, material in _read_material_aliases().items():
        if material in available:
            _add(alias, material)

    return {
        token: next(iter(materials))
        for token, materials in candidates.items()
        if len(materials) == 1
    }


# --- material fingerprinting -------------------------------------------------
#
# A deliberately lossy hash of a material phrase. Unlike a normal hash,
# collisions are the entire point: every way a drawing can write the same
# material must land in the same bucket, so that one verified answer covers all
# of them. "MATL: ALUMINIUM 5052", "AL ALY 5052-H32" and "5052 alum plate" are
# meant to collide.
#
# It works by throwing away everything that isn't material-defining - noise
# words, dimensions, tempers, punctuation, word order - and mapping what
# survives onto equivalence classes. The fingerprint is then the sorted set of
# classes, so ordering can't create two buckets for one material.

# Words that appear on drawings regardless of material.
_FINGERPRINT_NOISE = frozenset({
    "material", "matl", "mat", "mtl", "made", "from", "of", "grade", "spec",
    "plate", "sheet", "bar", "tube", "flat", "stock", "thk", "thick", "thickness",
    "ga", "gauge", "gage", "typ", "all", "parts", "part", "note", "notes", "see",
    "req", "reqd", "required", "qty", "quantity", "off", "pcs", "pc", "plcs",
    "finish", "mill", "raw", "general", "min", "max", "nom", "ref", "approx",
})

# Tempers and heat treatments: they qualify a material without changing which
# one it is, so they must not split the bucket.
_FINGERPRINT_TEMPERS = re.compile(r"^(?:h\d{2,3}|t\d{1,2}|o|f|hr|cr|ann|annealed)$")

# Equivalence classes. Everything on the right collapses to the key, which is
# what makes the collisions happen on purpose.
_FINGERPRINT_CLASSES: dict[str, frozenset[str]] = {
    "ALU": frozenset({
        "al", "alu", "alum", "alumin", "aluminum", "aluminium", "aly", "alloy",
    }),
    "STEEL": frozenset({"steel", "stl", "st"}),
    "MILD": frozenset({"mild", "ms", "crs", "hrs", "coldrolled", "hotrolled", "carbon"}),
    "STAINLESS": frozenset({"ss", "sst", "stainless", "inox"}),
    "GALV": frozenset({"galv", "galvanised", "galvanized", "gi"}),
    "CHK": frozenset({"chk", "checker", "checkered", "chequer", "chequered", "tread"}),
}

_FINGERPRINT_TOKEN_TO_CLASS = {
    token: klass for klass, tokens in _FINGERPRINT_CLASSES.items() for token in tokens
}


# Multi-word forms, collapsed before tokenising so both halves survive as one.
_FINGERPRINT_PHRASES = (
    (re.compile(r"\bcold\s*roll(?:ed)?\b"), "crs"),
    (re.compile(r"\bhot\s*roll(?:ed)?\b"), "hrs"),
    (re.compile(r"\bstainless\s*steel\b"), "stainless"),
    (re.compile(r"\bmild\s*steel\b"), "mild"),
    (re.compile(r"\bcarbon\s*steel\b"), "mild"),
    (re.compile(r"\bcheck(?:er|ered|quer|quered)\s*plate\b"), "chk"),
    (re.compile(r"\bal\s*aly\b"), "aluminum"),
)

# A grade designation: 3-4 digit alloy (5052, 304), or a letter/number pair
# (A36, 44W). Anything not matching this and not in a class is dropped.
_FINGERPRINT_GRADE = re.compile(r"^(?:\d{3,4}|[a-z]\d{2,3}|\d{2,3}[a-z])$")


def material_fingerprint(text: str) -> str:
    """A collision-seeking fingerprint of a material phrase.

    Unlike a normal hash, collisions are the whole point: every way a drawing
    can write the same material must land in the same bucket. Returns "" when
    nothing material-defining survives, which is the signal that the text said
    nothing useful rather than that it hashed to empty.

    Everything not recognisably a material class or a grade is discarded. That
    strictness is deliberate - keeping unknown words by default let noise like
    "astm", "csa" and "is" split buckets that should have collided.
    """
    raw = str(text or "")
    if not raw.strip():
        return ""

    working = raw.casefold()
    # Dimensions and thicknesses first: ".125", "0.250", "1/4" would otherwise
    # survive as grade-shaped digits and split identical materials apart.
    working = re.sub(r"\d*\.\d+", " ", working)
    working = re.sub(r"\d+\s*/\s*\d+", " ", working)
    # "80 X 120" is a size, not a grade - strip it before the letter/number
    # recombination below turns "X 120" into a grade-shaped "x120".
    working = re.sub(r"\d+\s*x\s*\d+", " ", working)
    # Punctuation to space, so "M.S." and "A-36" survive as tokens at all.
    working = re.sub(r"[^a-z0-9]+", " ", working)
    # "M S" -> "ms": single letters run together are an abbreviation, not words.
    working = re.sub(r"\b([a-z])\s+([a-z])\b(?!\s*[a-z])", r"\1\2", working)
    # "A 36" / "44 W" -> "a36" / "44w" before the grade test sees them.
    working = re.sub(r"\b([a-z])\s+(\d{2,3})\b", r"\1\2", working)
    working = re.sub(r"\b(\d{2,3})\s+([a-z])\b", r"\1\2", working)

    for pattern, replacement in _FINGERPRINT_PHRASES:
        working = pattern.sub(f" {replacement} ", working)

    parts: set[str] = set()
    for token in working.split():
        if token in _FINGERPRINT_NOISE or _FINGERPRINT_TEMPERS.match(token):
            continue
        klass = _FINGERPRINT_TOKEN_TO_CLASS.get(token)
        if klass is not None:
            parts.add(klass)
        elif _FINGERPRINT_GRADE.match(token):
            parts.add(token)
        # Everything else is discarded on purpose.

    # STEEL is a generic qualifier. It only stands alone when nothing more
    # specific was found, so that "MILD STEEL"/"MS" collide, and so that
    # "ASTM A36 STEEL" collides with a bare "A36".
    grades = {part for part in parts if _FINGERPRINT_GRADE.match(part)}
    if grades or parts & {"MILD", "STAINLESS", "GALV"}:
        parts.discard("STEEL")

    if not parts:
        return ""
    return "+".join(sorted(parts))


MATERIAL_MEMORY_FILENAME = "material_fingerprints.json"
MATERIAL_MEMORY_PATH = APP_DIR / "_runtime" / MATERIAL_MEMORY_FILENAME


def _material_memory_path() -> Path:
    # Read the module global at call time so tests can monkeypatch it - the
    # same reason the registry resolves its path this way. Binding it as a
    # default argument or capturing it at import let a test write learned
    # material wordings into the real _runtime store.
    return MATERIAL_MEMORY_PATH


def _load_material_memory() -> dict[str, dict[str, Any]]:
    """fingerprint -> {material: times_verified, ...}."""
    try:
        payload = json.loads(_material_memory_path().read_text(encoding="utf-8"))
    except (FileNotFoundError, OSError, json.JSONDecodeError):
        return {}
    entries = payload.get("entries") if isinstance(payload, dict) else None
    return entries if isinstance(entries, dict) else {}


def learn_material_fingerprint(source_text: str, material: str) -> str:
    """Record that a human confirmed `source_text` means `material`.

    This is the payoff of making the user verify: their correction is the
    training signal. Next time any drawing writes the material the same way -
    or any of the countless other ways that collide onto the same fingerprint -
    it is predicted from experience rather than from the hand-seeded aliases.

    Returns the fingerprint learned, or "" if the text carried nothing.
    """
    fingerprint = material_fingerprint(source_text)
    material = str(material or "").strip()
    if not fingerprint or not material:
        return ""

    entries = _load_material_memory()
    bucket = entries.setdefault(fingerprint, {})
    if not isinstance(bucket, dict):
        bucket = {}
    bucket[material] = int(bucket.get(material, 0)) + 1
    entries[fingerprint] = bucket

    path = _material_memory_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(".json.tmp")
    temp_path.write_text(
        json.dumps({"version": 1, "entries": entries}, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    temp_path.replace(path)
    return fingerprint


def recall_material(source_text: str) -> str | None:
    """The material previously verified for this text's fingerprint.

    Returns None when the bucket has never been seen, or when past
    verifications disagree - a fingerprint that has been confirmed as two
    different materials is a collision that shouldn't have happened, and
    guessing between them is worse than asking.
    """
    fingerprint = material_fingerprint(source_text)
    if not fingerprint:
        return None
    bucket = _load_material_memory().get(fingerprint)
    if not isinstance(bucket, dict) or not bucket:
        return None
    materials = [name for name, count in bucket.items() if int(count or 0) > 0]
    if len(materials) != 1:
        return None
    material = materials[0]
    return material if material in set(material_choices()) else None


def material_memory_conflicts() -> dict[str, dict[str, int]]:
    """Fingerprints verified as more than one material.

    A real collision between materials that aren't the same thing - either the
    fingerprint is too lossy, or somebody verified a row wrongly. Surfaced so
    it can be looked at rather than silently degrading predictions.
    """
    return {
        fingerprint: bucket
        for fingerprint, bucket in _load_material_memory().items()
        if isinstance(bucket, dict)
        and len([name for name, count in bucket.items() if int(count or 0) > 0]) > 1
    }


def _material_grades() -> dict[str, set[str]]:
    """material -> the grade numbers that legitimately belong to it.

    Used to stop a generic word claiming a specific alloy: "ALUM" alone maps
    to the shop's aluminium, but "6061 ALUM" names an alloy that isn't it.
    Derived from the same two sources as everything else, so it widens as the
    shop's files do.
    """
    grades: dict[str, set[str]] = {}
    for token, material in _material_tokens().items():
        if token[0].isdigit() and _FINGERPRINT_GRADE.match(token):
            grades.setdefault(material, set()).add(token)
    for alias, material in _read_material_aliases().items():
        for part in material_fingerprint(alias).split("+"):
            if part and part[0].isdigit() and _FINGERPRINT_GRADE.match(part):
                grades.setdefault(material, set()).add(part)
    return grades


def _match_material_in_text(text: str) -> str | None:
    """Canonical material for this text, only when exactly one matches.

    Two layers, both flattening drawing dialect onto the one string the CAM
    expects: the editable alias table (customer wording), then tokens derived
    from the shop's own descriptions. Whatever the source, the result must be
    a material the authoritative catalog lists, or it doesn't exist.

    Deliberately conservative: the standing rule is that material stays the
    user's choice, so anything ambiguous returns None and stays manual.
    """
    raw = str(text or "")
    if not raw.strip():
        return None

    # What a human already confirmed for this fingerprint beats both the
    # hand-seeded aliases and the derived tokens: it is evidence from this
    # shop's own drawings rather than a guess about how they might be worded.
    remembered = recall_material(raw)
    if remembered is not None:
        return remembered

    available = set(material_choices())
    matches: set[str] = set()

    # Aliases are matched as normalised phrases so multi-word entries
    # ("cold rolled", "checker plate") work, not just single tokens.
    haystack = _normalize_match_key(raw)
    for alias, material in _read_material_aliases().items():
        if alias and alias in haystack and material in available:
            matches.add(material)

    words = {word.casefold() for word in re.findall(r"[A-Za-z0-9]+", raw)}
    for token, material in _material_tokens().items():
        if token in words and material in available:
            matches.add(material)

    if len(matches) != 1:
        return None
    material = matches.pop()

    # A generic word ("ALUM") must not claim a specific alloy the drawing did
    # not ask for. If the text names a grade that isn't one of this material's
    # known grades, it is a different alloy - 6061-T6 ALUM is not Aluminum
    # 5052 - so predict nothing and let the user decide.
    known = _material_grades().get(material, set())
    stated = {
        part
        for part in material_fingerprint(raw).split("+")
        if part and part[0].isdigit() and _FINGERPRINT_GRADE.match(part)
    }
    if stated and known and not (stated & known):
        return None
    return material


# Sheet gauges. A gauge number is not a thickness - it means different things
# per material family, so a gauge can only be converted once the material is
# known. Manufacturers' Standard Gauge for steel, Brown & Sharpe for aluminium.
#
# Worked example of why nothing here is trusted directly: 11ga steel is .118
# nominal, the CAM side wants .12, and the shop floor calls it "1/8" or ".125".
# Four numbers for one sheet. That is exactly why every value below is passed
# through snap_thickness() onto whatever the catalog actually lists, instead of
# being used as a thickness in its own right - "11 GA", "1/8" and ".125" all
# have to end up at the same catalog entry.
_STEEL_GAUGES = {
    7: 0.1793, 8: 0.1644, 9: 0.1495, 10: 0.1345, 11: 0.118, 12: 0.1046,
    13: 0.0897, 14: 0.0747, 15: 0.0673, 16: 0.0598, 17: 0.0538, 18: 0.0478,
    19: 0.0418, 20: 0.0359, 21: 0.0329, 22: 0.0299, 24: 0.0239, 26: 0.0179,
}
_ALUMINUM_GAUGES = {
    7: 0.1443, 8: 0.1285, 9: 0.1144, 10: 0.1019, 11: 0.09, 12: 0.0808,
    13: 0.0720, 14: 0.0641, 15: 0.0571, 16: 0.0508, 17: 0.0453, 18: 0.0403,
    19: 0.0359, 20: 0.0320, 21: 0.0285, 22: 0.0253, 24: 0.0201, 26: 0.0159,
}

_GAUGE_PATTERN = re.compile(r"\b(\d{1,2})\s*(?:GA|GAGE|GAUGE)\b", re.IGNORECASE)

# The shop rounds thicknesses to 2dp in places (1/8 -> 0.12, 3/8 -> 0.38) but
# keeps full precision in others (0.375). Rather than pick one convention,
# snap a measured value onto whatever the catalog actually offers. 5% is wide
# enough for .125->.12 and .1875->.18, far too tight to confuse .12 with .18.
_THICKNESS_SNAP_TOLERANCE = 0.05


def _gauge_to_inches(gauge: int, material: str) -> float | None:
    table = _ALUMINUM_GAUGES if "alum" in material.casefold() else _STEEL_GAUGES
    return table.get(int(gauge))


def snap_thickness(value: float | None, material: str) -> float | None:
    """Snap a measured thickness onto the nearest one the shop stocks.

    A drawing says .125 where the catalog says 0.12, or .1875 where it says
    0.18; those are the same sheet. Returns None when nothing is close enough,
    which correctly rejects a thickness this material isn't available in.
    """
    if value is None or value <= 0:
        return None
    available = thickness_choices(material)
    if not available:
        return None
    nearest = min(available, key=lambda option: abs(option - value))
    if abs(nearest - value) <= _THICKNESS_SNAP_TOLERANCE * max(value, nearest):
        return nearest
    return None


def _parse_fraction_or_number(text: str) -> float | None:
    cleaned = str(text or "").strip()
    if "/" in cleaned:
        parts = [part.strip() for part in cleaned.split("/", 1)]
        try:
            numerator, denominator = float(parts[0]), float(parts[1])
        except (ValueError, IndexError):
            return None
        if denominator == 0:
            return None
        # The shop's convention is the fraction rounded to 2dp
        # (1/8 -> 0.12, 3/8 -> 0.38) - Python's default round matches.
        return round(numerator / denominator, 2)
    try:
        return float(cleaned)
    except ValueError:
        return None


def extract_dxf_hints(dxf_path: Path) -> DXFHints:
    """Best-effort qty/material/thickness from a DXF's own text.

    Every step degrades to None rather than raising: a drawing with no useful
    text simply contributes nothing, which is the normal case.
    """
    values = _dxf_text_values(Path(dxf_path))
    if not values:
        return DXFHints(material=None, thickness=None, qty=None)

    joined = " | ".join(values)
    material = _match_material_in_text(joined)

    # Keep the specific line that named the material, not the whole blob, so
    # the UI can show exactly which of the drawing's words was translated.
    material_source_text = None
    if material is not None:
        for value in values:
            if _match_material_in_text(value) == material:
                material_source_text = value
                break
        if material_source_text is None:
            material_source_text = joined

    qty: int | None = None
    for pattern in (_DXF_QTY_PATTERN, _DXF_QTY_SUFFIX_PATTERN):
        match = pattern.search(joined)
        if match is not None:
            try:
                candidate = int(match.group(1))
            except ValueError:
                candidate = 0
            if 0 < candidate <= 9999:
                qty = candidate
                break

    # Thickness only means something once the material is known: the catalog
    # differs per material, and a gauge number is meaningless without it.
    thickness: float | None = None
    if material is not None:
        measured: float | None = None
        thickness_match = _DXF_THICKNESS_PATTERN.search(joined)
        if thickness_match is not None:
            measured = _parse_fraction_or_number(thickness_match.group(1))

        if measured is None:
            gauge_match = _GAUGE_PATTERN.search(joined)
            if gauge_match is not None:
                try:
                    measured = _gauge_to_inches(int(gauge_match.group(1)), material)
                except ValueError:
                    measured = None

        # Snapping is what turns the drawing's .125 into the catalog's 0.12,
        # and it rejects anything not close to a stocked thickness.
        thickness = snap_thickness(measured, material)

    contributing = tuple(
        value for value in values
        if _DXF_QTY_PATTERN.search(value)
        or _DXF_QTY_SUFFIX_PATTERN.search(value)
        or _DXF_THICKNESS_PATTERN.search(value)
        or _match_material_in_text(value) is not None
    )
    return DXFHints(
        material=material,
        thickness=thickness,
        qty=qty,
        raw_lines=contributing,
        material_source_text=material_source_text,
    )


# --- PDF drawing-print title blocks ------------------------------------------
#
# The real source of material and thickness for this shop. Verified against
# F57524's prints: the DXFs exported for laser are pure geometry with no text
# entities at all, while every print carries MATERIAL and GAUGE in its title
# block.
#
# Read positionally, not by line order. The title block is a table, and its
# value sits directly *below* its label at the same x - reading the linearised
# text stream instead gives you a run of labels followed by a run of values,
# with no way to tell which value belongs to which label, and no way to notice
# that a label has no value at all.


@dataclass(frozen=True)
class PrintHints:
    material: str | None
    thickness: float | None
    qty: int | None
    # What the title block literally said, for the user to check against.
    material_source_text: str | None = None
    thickness_source_text: str | None = None
    qty_source_text: str | None = None
    # True when the print explicitly defers quantity elsewhere ("REFER TO BOM")
    # or leaves the QTY cell empty. Distinct from "we didn't look" - a blank
    # quantity must not silently become 1.
    qty_unknown: bool = False


_PRINT_LABELS = {
    "material": ("material", "matl", "mat'l", "mtl"),
    "thickness": ("gauge", "gage", "thickness", "thk"),
    "qty": ("qty", "quantity", "qnty"),
    # Which part this page is for. A drawing set is often one multi-page PDF
    # with a different part per page, so the page has to identify itself -
    # matching the file's name to a DXF stem only works when each part happens
    # to have its own PDF, which is not the norm.
    "part": ("dwg", "dwgno", "drawing", "drawingno", "partno", "partnumber"),
}

# Values the QTY cell carries when it is deliberately not answering.
_QTY_DEFERRED = re.compile(r"refer|see|per\b|bom|as\s*req", re.IGNORECASE)

# Measured off the real prints: a value sits ~6-7pt below its label at an
# *identical* x0 (they share a column), while the next label down is ~18pt
# away. A max drop of 12 therefore separates "my value" from "the next label"
# cleanly, which is what lets an empty cell be detected as empty.
#
# Note the drop is measured from the label's y0, not its y1: the boxes overlap
# vertically (MATERIAL ends at 1088.89, its value starts at 1087.88), so
# testing against y1 rejects the value by a fraction of a point.
_LABEL_VALUE_MIN_DROP = 3.0
_LABEL_VALUE_MAX_DROP = 12.0
_LABEL_VALUE_X_TOLERANCE = 3.0
# The drawing-number cell is taller than the stacked MATERIAL/GAUGE ones - its
# value sits ~15pt below the label rather than ~7 - and nothing else shares
# that column, so it can afford a wider window without risking the next label.
_LABEL_VALUE_MAX_DROP_BY_KIND = {"part": 24.0}
# Values wrap to the right along their row; a bigger gap than this is the next
# column of the title block, not a continuation.
_LABEL_VALUE_MAX_GAP = 40.0

_ALL_LABEL_WORDS = frozenset(
    {"title", "size", "scale", "rev", "sheet", "drawn", "checked", "last", "update",
     "dwg", "no", "date", "tolerances", "unless", "otherwise", "specified"}
    | {alias for aliases in () for alias in aliases}
)


def _print_words(pdf_path: Path) -> list[list[tuple]]:
    """Per page, the word boxes (x0, y0, x1, y1, text)."""
    try:
        import fitz
    except ImportError as exc:
        raise JobIntakeError("PyMuPDF (fitz) is required to read drawing prints.") from exc

    pages: list[list[tuple]] = []
    try:
        with fitz.open(str(pdf_path)) as doc:
            for page in doc:
                pages.append([tuple(word) for word in page.get_text("words")])
    except Exception:
        return []
    return pages


def _is_label_word(word: tuple) -> bool:
    text = str(word[4]).strip().rstrip(":").casefold()
    if text in _ALL_LABEL_WORDS:
        return True
    return any(text in aliases for aliases in _PRINT_LABELS.values())


def _value_below_label(words: list[tuple], label_word: tuple, max_drop: float | None = None) -> str:
    """The text sitting directly under `label_word`, or "" if the cell is empty.

    Anchored on the shared left edge of the column, because that is the only
    thing that reliably ties a title-block value to its label.
    """
    label_x0, label_y0 = float(label_word[0]), float(label_word[1])
    drop_limit = _LABEL_VALUE_MAX_DROP if max_drop is None else max_drop

    starts = [
        word
        for word in words
        if abs(float(word[0]) - label_x0) <= _LABEL_VALUE_X_TOLERANCE
        and _LABEL_VALUE_MIN_DROP <= float(word[1]) - label_y0 <= drop_limit
        and not _is_label_word(word)
    ]
    if not starts:
        return ""

    first = min(starts, key=lambda word: float(word[1]))
    row_y = float(first[1])
    row = sorted(
        (
            word
            for word in words
            if abs(float(word[1]) - row_y) <= 3.0 and float(word[0]) >= float(first[0])
        ),
        key=lambda word: float(word[0]),
    )

    # Stop at the first big horizontal gap - that's the next column.
    parts: list[str] = []
    previous_x1: float | None = None
    for word in row:
        if previous_x1 is not None and float(word[0]) - previous_x1 > _LABEL_VALUE_MAX_GAP:
            break
        parts.append(str(word[4]))
        previous_x1 = float(word[2])
    return " ".join(parts).strip()


def _find_label_values(words: list[tuple]) -> dict[str, str]:
    """label kind -> the raw text of its cell (may be "" when the cell is empty)."""
    found: dict[str, str] = {}
    for index, word in enumerate(words):
        text = str(word[4]).strip().rstrip(":").casefold()
        # "DWG NO" arrives as two words; join with the next one so the label
        # is recognised either way.
        joined = text
        if index + 1 < len(words):
            joined = f"{text}{str(words[index + 1][4]).strip().rstrip(':').casefold()}"
        for kind, aliases in _PRINT_LABELS.items():
            if kind in found:
                continue
            if text in aliases or joined in aliases:
                found[kind] = _value_below_label(
                    words, word, _LABEL_VALUE_MAX_DROP_BY_KIND.get(kind)
                )
    return found


def extract_print_hints(pdf_path: Path, part_stem: str | None = None) -> PrintHints:
    """Best-effort material/thickness/qty from a drawing print's title block.

    When `part_stem` is given, only a page whose title block names that part is
    used. Drawing sets are routinely one multi-page PDF with a different part
    per page, so taking the first title block found would attribute page one's
    material to every DXF in the job.

    Degrades to None throughout: a PDF that isn't a print, or a print whose
    title block doesn't parse, simply contributes nothing.
    """
    material = thickness = qty = None
    material_text = thickness_text = qty_text = None
    qty_unknown = False
    wanted = _normalize_match_key(part_stem or "")

    for words in _print_words(Path(pdf_path)):
        if not words:
            continue
        cells = _find_label_values(words)
        if not cells:
            continue

        if wanted:
            stated = _normalize_match_key(cells.get("part", ""))
            # Accept a containment match: the cell often carries a revision or
            # sheet suffix the DXF's name doesn't.
            if not stated or (wanted not in stated and stated not in wanted):
                continue

        if material is None and cells.get("material"):
            material_text = cells["material"]
            material = _match_material_in_text(material_text)

        if thickness is None and cells.get("thickness"):
            thickness_text = cells["thickness"]
            measured = _parse_fraction_or_number(
                re.sub(r"(?i)\b(?:in|inch|mm|thk|thick)\b|\"", " ", thickness_text).strip()
            )
            if measured is None:
                gauge_match = _GAUGE_PATTERN.search(thickness_text)
                if gauge_match is not None and material is not None:
                    measured = _gauge_to_inches(int(gauge_match.group(1)), material)
            if material is not None:
                thickness = snap_thickness(measured, material)

        if "qty" in cells and qty is None and not qty_unknown:
            qty_text = cells["qty"]
            # An empty QTY cell, or one that defers to a BOM, is *unknown* -
            # not 1. Silently defaulting is how a 40-off part gets cut once.
            if not qty_text or _QTY_DEFERRED.search(qty_text):
                qty_unknown = True
            else:
                digits = re.search(r"\b(\d{1,4})\b", qty_text)
                if digits is not None and int(digits.group(1)) > 0:
                    qty = int(digits.group(1))
                else:
                    qty_unknown = True

        if material is not None and thickness is not None and (qty is not None or qty_unknown):
            break

    return PrintHints(
        material=material,
        thickness=thickness,
        qty=qty,
        material_source_text=material_text,
        thickness_source_text=thickness_text,
        qty_source_text=qty_text,
        qty_unknown=qty_unknown,
    )


# --- BOM parts list ----------------------------------------------------------
#
# The best source of all when present, because its DESCRIPTION column holds the
# shop's own description string verbatim - "PLATE, AL ALY, .188" THK, 5052 H32
# GENERAL" is a line in expected_laser_descriptions.csv, not a customer's
# paraphrase. That resolves material, thickness and strategy exactly, with no
# flattening or guesswork, and it carries the quantity the prints defer to
# ("QTY: AS PER BOM").


@dataclass(frozen=True)
class BomRow:
    part: str
    description: str
    qty: int | None
    material: str | None
    thickness: float | None


def _looks_like_bom(lines: list[str]) -> bool:
    joined = " ".join(lines[:40]).casefold()
    return "parts list" in joined or ("part number" in joined and "qty" in joined)


def extract_bom_rows(pdf_path: Path, part_stems: list[str]) -> dict[str, BomRow]:
    """part stem -> its BOM row, for the parts actually present.

    Anchored on the stems we hold rather than on the table's shape: the parts
    list linearises one cell per line like every other PDF table here, and the
    column order varies between templates. Finding the part number - a value we
    already know - and reading its neighbours is stable across both.
    """
    try:
        import fitz
    except ImportError as exc:
        raise JobIntakeError("PyMuPDF (fitz) is required to read a BOM.") from exc

    try:
        with fitz.open(str(pdf_path)) as doc:
            text = "\n".join(str(page.get_text("text") or "") for page in doc)
    except Exception:
        return {}

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not _looks_like_bom(lines):
        return {}

    wanted = {_normalize_match_key(stem): stem for stem in part_stems if stem}
    rules = _read_description_rules()
    rows: dict[str, BomRow] = {}

    for index, line in enumerate(lines):
        stem = wanted.get(_normalize_match_key(line))
        if stem is None or stem in rows:
            continue

        # The description sits immediately before the part number, and the
        # quantity immediately after it.
        description = lines[index - 1] if index >= 1 else ""
        qty: int | None = None
        if index + 1 < len(lines) and lines[index + 1].isdigit():
            candidate = int(lines[index + 1])
            if 0 < candidate <= 9999:
                qty = candidate

        material = thickness = None
        rule = rules.get(_normalize_match_key(description))
        if rule is not None:
            material = str(rule["material"])
            thickness = float(rule["thickness"])
        else:
            # Not a verbatim description - fall back to the usual flattening.
            material = _match_material_in_text(description)

        rows[stem] = BomRow(
            part=stem,
            description=description,
            qty=qty,
            material=material,
            thickness=thickness,
        )
    return rows


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
        # A predicted material must be confirmed by a human before it can
        # reach RADAN. The translation from the customer's wording is a guess,
        # and a wrong material is expensive - so verification is enforced here
        # rather than left to whoever remembers to look.
        if not part.get("material_confirmed"):
            source = str(part.get("material_source_text", "") or "").strip()
            detail = f' (drawing says "{source}")' if source else ""
            problems.append(
                f"{filename}: confirm the material{detail} - tick Verified once "
                f"you've checked {material} is right"
            )
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
