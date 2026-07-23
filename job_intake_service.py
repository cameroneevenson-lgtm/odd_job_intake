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
    APPROVED_SOURCE_ROOTS,
    BATTLESHIELD_ROOT,
    EXPLORER_TEMPLATE_PATH,
    INVENTOR_TO_RADAN_DIR,
    JOB_PREFIX_TO_ROOT,
    MACHINE_EIA_BATTLESHIELD_ROOT,
    PLACEHOLDER_JOB_NUMBERS,
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


def is_placeholder_job_number(job_number: str) -> bool:
    """Whether this is a stand-in for a number that hasn't been issued yet."""
    return str(job_number or "").strip().upper() in PLACEHOLDER_JOB_NUMBERS


def label_required_for(job_number: str) -> bool:
    """Whether an intake for this number must carry a label.

    True when the folder already exists - the long-standing rule that keeps a
    one-off out of an existing job's directory - and always true for a
    placeholder, where several unrelated jobs would otherwise pile into the
    same folder waiting for their real numbers.
    """
    number = str(job_number or "").strip().upper()
    return is_placeholder_job_number(number) or job_folder_exists(number)


def job_folder_exists(job_number: str) -> bool:
    number = str(job_number or "").strip().upper()
    root_name = resolve_job_root(number)
    return (BATTLESHIELD_ROOT / root_name / number).exists()


# Reserved on Windows whatever the extension, so a folder named for one of
# these is a folder no tool can reliably open again.
_WINDOWS_DEVICE_NAMES = frozenset(
    ["CON", "PRN", "AUX", "NUL"]
    + [f"COM{digit}" for digit in range(1, 10)]
    + [f"LPT{digit}" for digit in range(1, 10)]
)
_LABEL_MAX_LENGTH = 64


def validate_label(label: str | None) -> str | None:
    """Check a label is one safe path component, and return it trimmed.

    A label comes from an email subject or a typed field and is joined straight
    onto the job folder, so it has to be a name rather than a path. Validated
    here, in one place, before anything resolves a path from it - a containment
    check afterwards would catch the escape but only after the folder shape had
    already been decided by untrusted text.
    """
    text = str(label or "").strip()
    if not text:
        return None

    if len(text) > _LABEL_MAX_LENGTH:
        raise JobIntakeError(
            f"That label is {len(text)} characters; keep it under "
            f"{_LABEL_MAX_LENGTH} so it stays a usable folder name."
        )
    if any(separator in text for separator in ("/", "\\")):
        raise JobIntakeError(
            f"A label is a folder name, not a path: {text!r} contains a slash."
        )
    if ":" in text:
        raise JobIntakeError(f"A label cannot contain a drive or stream marker: {text!r}")
    # Includes "..", and any name that resolves to something other than itself.
    if text in (".", "..") or Path(text).name != text:
        raise JobIntakeError(f"{text!r} is not a usable folder name.")
    invalid = set(text) & set('<>:"|?*')
    if invalid:
        raise JobIntakeError(
            f"A label cannot contain {' '.join(sorted(invalid))} - Windows will not "
            "accept it as a folder name."
        )
    if any(ord(character) < 32 for character in text):
        raise JobIntakeError("A label cannot contain control characters.")
    # Windows silently drops these, so the folder created would not be the
    # folder recorded - and the mismatch would surface much later.
    if text != text.rstrip(". "):
        raise JobIntakeError(
            f"A label cannot end with a dot or a space: {text!r}. Windows would "
            "strip it and the folder name would stop matching what was filed."
        )
    if text.split(".")[0].upper() in _WINDOWS_DEVICE_NAMES:
        raise JobIntakeError(
            f"{text!r} is a reserved Windows device name and cannot be a folder."
        )
    return text


def resolve_job_paths(job_number: str, label: str | None = None) -> JobPaths:
    number = str(job_number or "").strip().upper()
    root_name = resolve_job_root(number)
    release_root = BATTLESHIELD_ROOT / root_name
    job_dir = release_root / number
    label_text = validate_label(label) or ""

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
    # Applied to creation as well as delete/rename. validate_label() should
    # already have made an escape impossible, but this is the check that
    # actually looks at the resolved path, and mkdir is the first thing that
    # writes anywhere - so it is worth confirming rather than assuming.
    for folder in (paths.job_dir, paths.intake_dir, paths.project_dir):
        _assert_under_shop_root(folder)

    if paths.label is None and is_placeholder_job_number(paths.job_number):
        raise JobIntakeError(
            f"{paths.job_number} is a placeholder for a number that hasn't been "
            "issued yet, so it needs a Label to park under - the customer's PO "
            "number works well (e.g. 'PFF PO-8527-001'). Rename Job once the "
            "real number exists."
        )
    if paths.label is None and paths.job_dir.exists():
        raise JobIntakeError(
            f"{paths.job_dir} already exists - give this one-off a Label so it "
            "gets its own subfolder instead of mixing into the existing job."
        )
    if paths.label is not None and paths.intake_dir.exists():
        raise JobIntakeError(f"{paths.intake_dir} already exists - pick a different Label.")
    for folder in (paths.intake_dir, paths.project_dir, paths.project_dir / "nests", paths.project_dir / "remnants"):
        folder.mkdir(parents=True, exist_ok=True)


def reference_files(source_files: list[Path]) -> list[dict[str, Any]]:
    """Record files without copying them.

    Used for work already sitting on W:. That folder is the source of truth for
    the job, so a copy on L: would only be a second version to drift out of
    date - and copying it was by far the slowest part of an intake. The
    trade-off, accepted deliberately: if engineering renames or removes
    something afterwards, these paths go stale.
    """
    referenced: list[dict[str, Any]] = []
    for source in source_files:
        source = Path(source)
        if not source.exists():
            continue
        referenced.append(
            {
                "filename": source.name,
                "saved_path": str(source),
                "size": source.stat().st_size,
                # So the UI and any later check can tell a referenced file from
                # one that was copied into the job folder.
                "in_place": True,
            }
        )
    return referenced


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


def is_approved_source_root(folder: Path) -> bool:
    """Whether a path lifted out of an email may be read from.

    An email body is untrusted text. A path found in one is followed, listed
    and its files recorded onto the job, so it is confined to the shares this
    work actually comes from - engineering's W:\\LASER and the shop's own
    L:\\BATTLESHIELD. Without this, any readable folder named anywhere in a
    message or its reply chain could be pulled into a job.
    """
    try:
        resolved = Path(folder).resolve()
    except OSError:
        return False
    for root in APPROVED_SOURCE_ROOTS:
        try:
            if resolved.is_relative_to(Path(root).resolve()):
                return True
        except OSError:
            continue
    return False


def find_email_source_folder(
    email_body: str,
) -> tuple[Path | None, list[Path], list[Path]]:
    """The folder an email points at, the work files in it, and what was skipped.

    Returns (folder, files, rejected). Shared by the listener's pre-check and
    the intake itself so the two cannot disagree about whether a message
    actually pointed at any work - the listener answering 201 for a job the
    intake then finds nothing in is a job folder on L: that someone has to go
    and clean up.

    Cheap enough to call twice: it lists one directory and copies nothing.
    """
    rejected: list[Path] = []
    for candidate in paths_in_text(email_body):
        folder = Path(candidate)
        if not folder.is_dir():
            continue
        if not is_approved_source_root(folder):
            rejected.append(folder)
            continue
        try:
            pulled = [
                item
                for item in sorted(folder.iterdir())
                if item.is_file() and item.suffix.casefold() in INGESTED_SUFFIXES
            ]
        except OSError:
            continue
        # Only the first folder that actually holds work files; the candidate
        # list contains parent paths of it too.
        return (folder, pulled, rejected) if pulled else (None, [], rejected)
    return None, [], rejected


@dataclass(frozen=True)
class EmailHints:
    """Material/qty stated in the message itself.

    Ranks above the prints and below the BOM: someone typed it deliberately
    for this job, but in prose, with no structure to lean on. Only used when
    the message names exactly one material, so a signature or a chatty
    paragraph can't decide anything.
    """

    material: str | None = None
    thickness: float | None = None
    qty: int | None = None
    material_source_text: str | None = None
    qty_source_text: str | None = None
    # True when the wording is explicitly per-part ("4 each", "all parts"), so
    # one number can safely be applied to a multi-part job.
    qty_is_per_part: bool = False


def extract_email_hints(text: str) -> EmailHints:
    body = str(text or "")
    if not body.strip():
        return EmailHints()

    # Line-by-line, so the phrase that names the material can be quoted back
    # rather than the whole message.
    material = source = None
    ambiguous = False
    for line in body.splitlines():
        found = _match_material_in_text(line)
        if found is None:
            continue
        if material is not None and found != material:
            # The message names two materials - it isn't telling us which.
            ambiguous = True
            break
        material = found
        source = line.strip()
    if ambiguous:
        material = source = None

    qty = None
    qty_source = None
    for pattern in (_DXF_QTY_PATTERN, _DXF_QTY_SUFFIX_PATTERN, _PROSE_QTY_PATTERN):
        match = pattern.search(body)
        if match is not None:
            try:
                candidate = int(match.group(1))
            except ValueError:
                continue
            if 0 < candidate <= 9999:
                qty = candidate
                qty_source = match.group(0).strip()
                break

    per_part = bool(
        qty is not None
        and re.search(r"\b(?:each|per\s+part|apiece|all\s+parts?)\b", body, re.IGNORECASE)
    )

    # Thickness stated in the message, so it can dispute the other sources
    # rather than only the print and drawing being allowed an opinion.
    #
    # Gauges are read here for the same reason the drawing and print scrapers
    # read them: an email is where shop shorthand is most likely to appear -
    # "make these in 11ga" - and this was the one source that could not
    # understand it. Only with a material in hand, because 11GA is 0.118 in
    # steel and 0.090 in aluminium.
    thickness = None
    if material is not None:
        thickness_match = _DXF_THICKNESS_PATTERN.search(body)
        if thickness_match is not None:
            thickness = snap_thickness(
                _parse_fraction_or_number(thickness_match.group(1)), material
            )
        if thickness is None:
            gauge_match = _GAUGE_PATTERN.search(body)
            if gauge_match is not None:
                thickness = snap_thickness(
                    _gauge_to_inches(int(gauge_match.group(1)), material), material
                )

    return EmailHints(
        material=material,
        thickness=thickness,
        qty=qty,
        material_source_text=source,
        qty_source_text=qty_source,
        qty_is_per_part=per_part,
    )


def intake_summary_text(entry: dict[str, Any]) -> str:
    """A plain-text summary of a finished intake, for a reply email.

    Composed here rather than in the macro: the wording will keep changing as
    the extraction does, and changing it in Python costs nothing while changing
    it in VBA means re-importing and re-signing on every machine.

    Deliberately states what is *not* known as loudly as what is - an
    unanswered quantity or a material nobody has checked is the useful part of
    this message, not the row count.
    """
    parts = list(entry.get("material_qty", []))
    lines = [
        f"Job {entry.get('job_number')}"
        + (f" / {entry['label']}" if entry.get("label") else "")
        + " has been set up for laser.",
        "",
        f"Folder: {entry.get('job_folder')}",
    ]
    if entry.get("po_number"):
        lines.append(f"PO: {entry['po_number']}")
    if entry.get("due_date"):
        lines.append(f"Due: {entry['due_date']}")

    lines += ["", f"Parts ({len(parts)}):"]
    for part in parts:
        material = str(part.get("material") or "not determined")
        thickness = part.get("thickness") or 0
        qty = part.get("qty")
        flag = ""
        if part.get("qty_unknown"):
            flag = "  <- quantity not stated"
        elif part.get("source_conflict"):
            flag = "  <- sources disagree"
        lines.append(
            f"  {part.get('filename'):<22} {material:<20} "
            f"{thickness if thickness else '?':>6}  qty {qty}{flag}"
        )

    notes = [str(note) for note in entry.get("po_unmatched", []) if str(note).strip()]
    unanswered = [p.get("filename") for p in parts if p.get("qty_unknown")]
    conflicts = [p.get("filename") for p in parts if p.get("source_conflict")]
    if notes or unanswered or conflicts:
        lines += ["", "Needs checking:"]
        lines += [f"  - {note}" for note in notes]
        if unanswered:
            lines.append(f"  - No quantity stated for: {', '.join(map(str, unanswered))}")
        if conflicts:
            lines.append(f"  - Sources disagree for: {', '.join(map(str, conflicts))}")

    lines += [
        "",
        "Nothing has been cut. Materials still need checking in the shop app "
        "before the parts are imported to RADAN.",
    ]
    return "\n".join(lines)


def begin_intake(
    job_number: str,
    label: str | None,
    *,
    source: str = "manual",
    email_subject: str = "",
    email_sender: str = "",
) -> tuple[dict[str, Any], JobPaths]:
    """The fast half of an intake: validate, make the folder, register it.

    Split out so a caller that can't afford to wait - the loopback listener,
    whose client times out after 30s while a network-to-network copy of a real
    job takes closer to a minute - can answer immediately and finish the slow
    half on a background thread.

    Everything here is decisive and quick: it resolves the paths, enforces the
    fresh-vs-label rule, creates the directories and writes the registry entry.
    Doing that up front means a second click is rejected by the existing
    already-exists guard rather than racing the first, and the job appears in
    the desktop queue straight away.
    """
    paths = resolve_job_paths(job_number, label or None)
    create_job_folders(paths)

    entry = job_intake_registry.new_entry(
        job_number=job_number,
        label=label or None,
        source=source,
        email_subject=email_subject,
        email_sender=email_sender,
    )
    entry["job_folder"] = str(paths.intake_dir)
    # Claim the registry key here too, not just the folder. Both are how a job
    # is claimed, and leaving the append to the slow half meant a duplicate
    # surfaced long after the caller had been told the job was accepted.
    job_intake_registry.append_entry(entry)
    return entry, paths


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
    entry, _paths = begin_intake(
        job_number,
        label,
        source=source,
        email_subject=email_subject,
        email_sender=email_sender,
    )
    return complete_intake(entry, files, email_body=email_body)


def complete_intake(
    entry: dict[str, Any], files: list[Path], *, email_body: str = ""
) -> dict[str, Any]:
    """The slow half: copy the work files in, read them, register the entry.

    Takes the entry begin_intake made, so the folder already exists and the
    job is already claimed. Split out so the listener can run this on a
    background thread after answering - a real job spent 53s here, past the
    30s its caller waits.
    """
    job_number = str(entry.get("job_number", ""))
    label = entry.get("label")
    paths = resolve_job_paths(job_number, label)

    # Some jobs arrive as a path on W: instead of attachments - "please
    # manufacture the parts at the following path". Those files are referenced
    # where they sit: W: is the source of truth for them, so copying would only
    # create a second version to drift out of date, and the copy was the slowest
    # part of an intake by a wide margin.
    #
    # Emailed attachments are different - they arrive in a temp directory that
    # is deleted when the request finishes, so they must be copied.
    # Warnings this code raises, kept apart from `unmatched` - which holds PO
    # lines that matched nothing and gets noise-filtered below. Mixing them let
    # that filter silently eat real warnings twice, because a message about a
    # BOM naturally contains the word "BOM". They are merged at the end.
    notes: list[str] = []

    folder, in_place, rejected = find_email_source_folder(email_body)
    ingested_from = [str(folder)] if folder is not None else []
    for skipped in rejected:
        notes.append(
            f"{skipped}: mentioned in the email but outside the folders this app "
            f"reads from, so nothing was taken from it."
        )

    attachments = copy_attachments(paths, list(files)) + reference_files(in_place)

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
    # A spreadsheet BOM outranks everything: it is structured CAM output, and
    # inventor_to_radan converts it with the shop's own rules rather than
    # anything guessed here.
    cam_rows: dict[str, dict[str, Any]] = {}
    # Spreadsheet BOMs that were recognised but could not be converted. Carried
    # onto the entry so the import gate can refuse rather than proceed on the
    # scraped sources.
    bom_blockers: list[str] = []
    for filename, path in attachment_paths.items():
        if path.suffix.casefold() not in BOM_SPREADSHEET_SUFFIXES:
            continue
        if not _looks_like_bom_spreadsheet(path):
            continue
        conversion = convert_bom_spreadsheet(path, move_output_to=paths.intake_dir)
        # A spreadsheet BOM is the authority for this job. If it exists but
        # could not be read, the scraped prints are not a substitute for it -
        # falling through to them would quietly downgrade the source of truth
        # to a guess. Recorded so the gate refuses the import and says why.
        if conversion.blocked:
            bom_blockers.append(conversion.reason)
            notes.append(f"STOP: {conversion.reason}")
            continue

        for row in conversion.rows:
            cam_rows[Path(str(row["filename"])).stem.casefold()] = row

        # inventor_to_radan's own accountability checks, passed straight
        # through. It runs these as part of converting, so reporting what it
        # found beats re-deriving them here and disagreeing later.
        for orphan in conversion.orphan_dxfs:
            notes.append(
                f"{orphan}: in the folder but not referenced by the BOM "
                f"(inventor_to_radan). Check with engineering - there may be a "
                f"reason, or the BOM may have missed it."
            )
        for missing in conversion.expected_missing_dxfs:
            notes.append(
                f"{missing}: expected by the BOM but no DXF was found "
                f"(inventor_to_radan)."
            )
        for missing_pdf in conversion.missing_pdfs:
            notes.append(f"{missing_pdf}: no PDF alongside it (inventor_to_radan).")
        if cam_rows:
            break

    bom_rows: dict[str, BomRow] = {}
    available_materials = set(material_choices())
    # Whatever the message itself said, applied to every part in the job -
    # an email names a material for the work, not per drawing.
    email_hints = extract_email_hints(email_body)
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
        bom_material = bom_thickness = bom_strategy = None
        if bom_row is not None and bom_row.material:
            if bom_row.material in available_materials:
                bom_material = bom_row.material
                bom_thickness = bom_row.thickness
                bom_strategy = bom_row.strategy
            else:
                note = (
                    f"{name}: the BOM asks for {bom_row.material} "
                    f'("{bom_row.description}"), which isn\'t in the shop\'s laser list'
                )
                if note not in notes:
                    notes.append(note)

        # Order of trust, most structured first: a BOM states it in the shop's
        # own words; the email was typed for this job but in prose; a print's
        # title block is a drawing note; the DXF's own text is the last resort.
        cam_row = cam_rows.get(Path(name).stem.casefold())

        # A DXF the PDF parts list doesn't mention. Only checked here for a PDF
        # BOM - a spreadsheet one goes through inventor_to_radan, whose own
        # orphan-DXF check is reported above rather than duplicated.
        if bom_rows and not cam_rows and Path(name).stem not in bom_rows:
            note = (
                f"{name}: this DXF isn't in the BOM. Check with engineering - "
                f"there may be a reason, or the BOM may have missed it."
            )
            if note not in notes:
                notes.append(note)
        material = (
            (cam_row["material"] if cam_row else None)
            or bom_material
            or email_hints.material
            or print_hints.material
            or dxf_hints.material
        )
        # Same precedence the material above uses, for the same reason: someone
        # typed it deliberately for this job, which outranks a drawing note but
        # not structured CAM output.
        thickness = (
            (cam_row["thickness"] if cam_row else None)
            or bom_thickness
            or email_hints.thickness
            or print_hints.thickness
            or dxf_hints.thickness
        )
        material_source = (
            (f"BOM via inventor_to_radan" if cam_row else None)
            or (bom_row.description if bom_material else None)
            or (email_hints.material_source_text if email_hints.material else None)
            or print_hints.material_source_text
            or dxf_hints.material_source_text
        )

        # Cross-check: where two sources both named a material, they should
        # agree. A disagreement is worth a human look - it usually means a
        # print was superseded by the BOM, or the email is about a different
        # job - and is far more useful surfaced than silently resolved by
        # whichever source happened to rank higher.
        # The PO's own wording for this line. It is not used to *choose* a
        # material - customer POs spell it inconsistently, which is why that
        # has always been reference-only - but it absolutely gets a say in
        # whether the sources agree. A PO asking for aluminium against a print
        # drawn in steel is the exact disagreement worth stopping for.
        po_material = _match_material_in_text(str(hint.get("raw_description", "") or ""))

        stated = {
            "the CAM BOM": cam_row["material"] if cam_row else None,
            "the BOM": bom_material,
            "the PO": po_material,
            "the email": email_hints.material,
            "the print": print_hints.material,
            "the drawing": dxf_hints.material,
        }
        named = {where: value for where, value in stated.items() if value}
        # Tracked per field, not as one blob: a quantity disagreement must not
        # be clearable by settling the material, and vice versa.
        conflicts: dict[str, str] = {}
        if len(set(named.values())) > 1:
            detail = ", ".join(f"{where} says {value}" for where, value in named.items())
            conflicts["material"] = detail
            note = f"{name}: sources disagree on material - {detail}"
            if note not in notes:
                notes.append(note)

        # Thickness was previously resolved by precedence with no check at all,
        # so two sources quietly disagreeing about how thick a part is went
        # unnoticed. Every source that can choose a value now also disputes it.
        thickness_stated = {
            "the CAM BOM": cam_row["thickness"] if cam_row else None,
            "the BOM": bom_thickness,
            "the email": email_hints.thickness,
            "the print": print_hints.thickness,
            "the drawing": dxf_hints.thickness,
        }
        thickness_named = {
            where: value for where, value in thickness_stated.items() if value
        }
        if len(set(thickness_named.values())) > 1:
            detail = ", ".join(
                f"{where} says {value:g}" for where, value in thickness_named.items()
            )
            conflicts["thickness"] = detail
            note = f"{name}: sources disagree on thickness - {detail}"
            if note not in notes:
                notes.append(note)

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
            if note not in notes:
                notes.append(note)

        # Quantity precedence: the PO is a commitment, the print is a drawing
        # note. Neither may be invented - a blank QTY or a "REFER TO BOM" is
        # recorded as unknown rather than defaulting to 1, because silently
        # cutting one of a forty-off part is the expensive failure here.
        # An email states a quantity for "the job", but the job may be thirteen
        # different parts - applying one number to all of them would cut twelve
        # of them wrong. Only trusted when it cannot be misapplied: a single-DXF
        # job, or wording that is explicitly per-part ("4 each", "all parts").
        email_qty = (
            email_hints.qty
            if email_hints.qty and (len(dxf_names) == 1 or email_hints.qty_is_per_part)
            else None
        )

        resolved_qty = (
            (cam_row["qty"] if cam_row else None)
            or po_qty
            or (bom_row.qty if bom_row is not None else None)
            or email_qty
            or print_hints.qty
            or dxf_hints.qty
        )
        qty_unknown = resolved_qty is None

        # Email and drawing are included: both can win the value above, so both
        # must be able to dispute it. Leaving them out meant a source could
        # decide an answer it was never allowed to argue about.
        qty_stated = {
            "the CAM BOM": cam_row["qty"] if cam_row else None,
            "the PO": po_qty,
            "the BOM": bom_row.qty if bom_row is not None else None,
            "the email": email_qty,
            "the print": print_hints.qty,
            "the drawing": dxf_hints.qty,
        }
        qty_named = {where: value for where, value in qty_stated.items() if value}
        if len(set(qty_named.values())) > 1:
            detail = ", ".join(f"{where} says {value}" for where, value in qty_named.items())
            conflicts["quantity"] = detail
            note = f"{name}: sources disagree on quantity - {detail}"
            if note not in notes:
                notes.append(note)

        material_qty.append(
            {
                "filename": name,
                "material": material or "",
                "thickness": thickness or 0.0,
                "qty": int(resolved_qty or 1),
                "unit": (cam_row["unit"] if cam_row else None) or "in",
                # The CAM row's own strategy, verbatim - casing included. It is
                # the converter's answer for that exact description, and
                # recomputing it from material alone threw away the one source
                # that had already decided. Only derived when nothing
                # authoritative said.
                "strategy": (
                    (cam_row["strategy"] if cam_row else None)
                    or (bom_strategy if bom_material else None)
                    or (default_strategy_for_material(material) if material else "")
                ),
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
                # Two sources gave different answers. A hard stop, not a note:
                # somebody has to go back to whoever asked for the job and find
                # out which is right, because ranking one source over another
                # would just be picking a winner silently.
                #
                # Per field, and never rewritten by the UI - this is extracted
                # evidence, not an editable decision. Rebuilding the row on save
                # used to erase it, which quietly disarmed the gate.
                "conflicts": conflicts,
                # Which fields a human has since settled. Resolving means
                # choosing a value for *that* field; settling the material says
                # nothing about a disputed quantity.
                "resolved": {},
            }
        )

    # Kept whole so later passes (material/qty scraped from the message, or a
    # W: folder named in it) can re-read it without the email being needed
    # again. The macro sends it raw precisely so this stays changeable here.
    entry["email_body"] = email_body
    # Keep only paths that exist, and drop any that is merely a parent of
    # another kept one - the candidate list deliberately includes prefixes.
    existing = [path for path in paths_in_text(email_body) if Path(path).exists()]
    # Every one-off job needs a blank project, so it is made now rather than
    # waiting for a button. Non-fatal: if the template is missing or an RPD is
    # somehow already there, the intake is still perfectly good and the desktop
    # button can retry.
    try:
        entry_rpd_path = str(clone_rpd_template(paths))
        entry_rpd_error = ""
    except JobIntakeError as exc:
        entry_rpd_path = ""
        entry_rpd_error = str(exc)

    # Flagged so it's obvious in the queue that this job is still waiting for
    # its real number, rather than looking like any other filed job.
    # PO leftovers first, then everything this code flagged. Kept separate all
    # the way to here so the noise filter above can never reach the warnings.
    unmatched = unmatched + [note for note in notes if note not in unmatched]

    entry["provisional"] = is_placeholder_job_number(job_number)
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
    entry["bom_blockers"] = bom_blockers
    if entry_rpd_path:
        entry["rpd_path"] = entry_rpd_path
        entry["status"] = job_intake_registry.STATUS_RPD_CREATED
    elif entry_rpd_error:
        # Recorded rather than raised - the job folder and parts are fine, and
        # this is retryable from the desktop page.
        entry["error"] = entry_rpd_error

    # Set last: the caller polls for these to know the slow half has finished,
    # so they must not appear until everything else has been written.
    entry["complete"] = True
    entry["state"] = job_intake_registry.STATE_SUCCEEDED

    # Updated, not appended: begin_intake already claimed the key. Written in
    # one call so the queue never shows a half-filled job.
    job_intake_registry.update_entry(
        str(entry["key"]),
        **{
            field: entry[field]
            for field in (
                "provisional", "ingested_from", "source_paths", "job_folder",
                "po_number", "due_date", "due_note", "attachments",
                "material_qty", "po_unmatched", "bom_blockers", "email_body",
                "status", "rpd_path", "error", "complete", "state",
            )
            if field in entry
        },
    )
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


# --- delete / rename ---------------------------------------------------------
#
# Both operate on real shop folders, so both refuse to touch anything outside
# BATTLESHIELD_ROOT and both report exactly what they did.

# Files whose *contents* reference the project name. Both are XML despite the
# extensions: an .rpd holds it ~40 times and a nest .drg ~27, in element text
# and in path attributes, so renaming the files alone leaves RADAN pointing at
# a project that no longer exists.
RENAMEABLE_CONTENT_SUFFIXES = (".rpd", ".drg")


def _assert_under_shop_root(target: Path) -> None:
    root = BATTLESHIELD_ROOT.resolve()
    try:
        resolved = target.resolve()
    except OSError as exc:
        raise JobIntakeError(f"Could not resolve {target}: {exc}") from exc
    if not resolved.is_relative_to(root):
        raise JobIntakeError(
            f"Refusing to touch {resolved} - it is outside {root}."
        )


def delete_job_files(entry: dict[str, Any]) -> str:
    """Remove an intake's folder from the shop drive.

    The registry entry is the caller's business; this only handles the files,
    and only inside BATTLESHIELD_ROOT.
    """
    folder_text = str(entry.get("job_folder") or "")
    if not folder_text:
        return "No job folder recorded for this intake."
    folder = Path(folder_text)
    if not folder.exists():
        return f"{folder} no longer exists."

    _assert_under_shop_root(folder)
    file_count = sum(1 for _ in folder.rglob("*") if _.is_file())
    shutil.rmtree(folder)
    return f"Deleted {folder} ({file_count} file(s))."


@dataclass(frozen=True)
class RenamePlan:
    old_job_number: str
    new_job_number: str
    old_project_name: str
    new_project_name: str
    old_intake_dir: Path
    new_intake_dir: Path
    files_to_rename: tuple[tuple[Path, str], ...]
    files_to_rewrite: tuple[Path, ...]

    def describe(self) -> str:
        lines = [
            f"Folder:  {self.old_intake_dir}",
            f"     ->  {self.new_intake_dir}",
        ]
        if self.files_to_rename:
            lines.append(f"Rename {len(self.files_to_rename)} file(s):")
            lines += [f"    {path.name}  ->  {new}" for path, new in self.files_to_rename]
        if self.files_to_rewrite:
            lines.append(
                f"Rewrite references inside {len(self.files_to_rewrite)} file(s): "
                + ", ".join(sorted({path.suffix for path in self.files_to_rewrite}))
            )
        return "\n".join(lines)


def plan_job_rename(
    entry: dict[str, Any], new_job_number: str, new_label: str | None
) -> RenamePlan:
    """Work out everything a rename touches, without changing anything.

    Separated from the doing so the user can be shown it first - this moves
    real folders on L: and edits files RADAN depends on.

    Renaming is for fixing the *project's* name - a typo in the job number, or
    a placeholder becoming a real number. Part files are never touched: their
    names are the customer's own numbering, they are what the registry rows and
    the RPD's part list refer to, and no rename of a project should change what
    a part is called.
    """
    old_number = str(entry.get("job_number", "") or "").strip().upper()
    old_label = entry.get("label")
    new_number = str(new_job_number or "").strip().upper()
    new_label_text = str(new_label or "").strip() or None

    if not new_number:
        raise JobIntakeError("Enter a new job number.")
    resolve_job_root(new_number)
    if new_number == old_number and (new_label_text or "") == (old_label or ""):
        raise JobIntakeError("That is already this job's number and label.")

    old_paths = resolve_job_paths(old_number, old_label)
    new_paths = resolve_job_paths(new_number, new_label_text)

    if not old_paths.intake_dir.exists():
        raise JobIntakeError(f"{old_paths.intake_dir} does not exist.")
    if new_paths.intake_dir.exists():
        raise JobIntakeError(f"{new_paths.intake_dir} already exists.")
    _assert_under_shop_root(old_paths.intake_dir)
    _assert_under_shop_root(new_paths.intake_dir)

    renames: list[tuple[Path, str]] = []
    rewrites: list[Path] = []
    for path in sorted(old_paths.intake_dir.rglob("*")):
        if not path.is_file():
            continue
        if path.suffix.casefold() in RENAMEABLE_CONTENT_SUFFIXES:
            rewrites.append(path)

        # Part files keep their names, always. Only the project's own files -
        # the RPD, its nests, the nest summary - carry the job number as part
        # of the project's identity and move with it.
        is_part_file = path.suffix.casefold() in (".dxf", ".sym") or (
            path.suffix.casefold() == ".pdf" and "NEST" not in path.name.upper()
        )
        if is_part_file:
            continue

        # Any file carrying the old project name in its own name, not just
        # RPD/DRG - the nests are called "P1 <project>.drg" and symbols and
        # summaries follow the same habit.
        if old_paths.project_name in path.name or old_number in path.name:
            new_name = path.name.replace(old_paths.project_name, new_paths.project_name)
            new_name = new_name.replace(old_number, new_number)
            if new_name != path.name:
                renames.append((path, new_name))

    return RenamePlan(
        old_job_number=old_number,
        new_job_number=new_number,
        old_project_name=old_paths.project_name,
        new_project_name=new_paths.project_name,
        old_intake_dir=old_paths.intake_dir,
        new_intake_dir=new_paths.intake_dir,
        files_to_rename=tuple(renames),
        files_to_rewrite=tuple(rewrites),
    )


def apply_job_rename(plan: RenamePlan) -> str:
    """Carry out a rename plan: rewrite contents, rename files, move folders.

    Contents are rewritten before anything moves, so a failure part-way leaves
    the job where it was rather than half-renamed under a new path.
    """
    rewritten = 0
    for path in plan.files_to_rewrite:
        try:
            text = _decode_template_bytes(path.read_bytes())
        except JobIntakeError:
            continue
        updated = text.replace(plan.old_project_name, plan.new_project_name)
        updated = updated.replace(plan.old_job_number, plan.new_job_number)
        # Paths embedded in the file point at the old folder too.
        updated = updated.replace(str(plan.old_intake_dir), str(plan.new_intake_dir))
        if updated != text:
            path.write_text(updated, encoding="utf-8")
            rewritten += 1

    renamed = 0
    # Deepest first, so renaming a file never depends on its parent's name.
    for path, new_name in sorted(
        plan.files_to_rename, key=lambda pair: len(pair[0].parts), reverse=True
    ):
        if path.exists():
            path.rename(path.with_name(new_name))
            renamed += 1

    # The inner project directory carries the project name as well.
    for directory in sorted(
        (item for item in plan.old_intake_dir.rglob("*") if item.is_dir()),
        key=lambda item: len(item.parts),
        reverse=True,
    ):
        new_name = directory.name.replace(plan.old_project_name, plan.new_project_name)
        new_name = new_name.replace(plan.old_job_number, plan.new_job_number)
        if new_name != directory.name:
            directory.rename(directory.with_name(new_name))

    plan.new_intake_dir.parent.mkdir(parents=True, exist_ok=True)
    plan.old_intake_dir.rename(plan.new_intake_dir)

    return (
        f"Renamed to {plan.new_project_name}: {renamed} file(s) renamed, "
        f"{rewritten} file(s) had their internal references updated, "
        f"folder now {plan.new_intake_dir}."
    )


# --- Material list -----------------------------------------------------------


DESCRIPTION_RULES_FILENAME = "description_rules.csv"


# Resolved at call time, not as a module-level constant. Binding it at import
# would freeze the directory (defeating monkeypatching in tests) and is the
# same class of bug the registry invariant warns about. The file is re-read on
# every call so materials the shop adds appear without restarting anything.
#
# expected_laser_descriptions.csv is not read: it listed a far narrower set
# than the rules table and gating on it rejected materials that turn up in real
# customer BOMs. inventor_to_radan has since deleted that file outright, so
# column A of description_rules.csv is the single list of known laser
# descriptions on both sides.
def _description_rules_path() -> Path:
    return INVENTOR_TO_RADAN_DIR / DESCRIPTION_RULES_FILENAME


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

    # No fallback. An empty catalog means description_rules.csv is missing or
    # reshaped, and it lives on C: - unreachable means something is badly
    # wrong. Substituting a hardcoded list would quietly make a stale guess the
    # authority for what gets cut, which is exactly what this file exists to
    # prevent. Callers predict nothing and the import refuses; the folder and
    # the files are still created, so no work is lost.
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
    # Title-block labels. The PO scrape also runs over drawing prints, and
    # these read like table rows to it. Harmless noise in a log; not harmless
    # in the summary that goes into a reply to whoever sent the job.
    "last update",
    "tolerances",
    "unless otherwise",
    "angles",
    "inches",
    "drawn",
    "checked",
    "material",
    "gauge",
    "sheet",
    "scale",
    "title",
    "rev",
    "size",
    "dwg",
    "qty",
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
# "2 OFF", "4 PLCS", "3 REQ'D", "10 REQUIRED" - the number before the word.
_DXF_QTY_SUFFIX_PATTERN = re.compile(
    r"\b(\d{1,4})\s*(?:OFF|PLCS?|PCS?|REQ'?D|REQUIRED)\b", re.IGNORECASE
)
# How people actually write it in an email: "can you make 19", "cut 6 of
# these", "we need 4". The negative lookahead keeps dimensions out - "cut 19
# inches", "19 x 40", "make 19mm" are not quantities.
_PROSE_QTY_PATTERN = re.compile(
    r"\b(?:make|cut|need|needs|produce|run|fabricate|do)\s+(\d{1,4})\b"
    r"(?!\s*(?:x\b|[\"']|in\b|inch|mm\b|ga\b|gauge|thk|thick))",
    re.IGNORECASE,
)
_DXF_THICKNESS_PATTERN = re.compile(
    # Leading-dot decimals first: drawings write ".125 THK" far more often than
    # "0.125 THK", and without this alternative the dot was skipped and the
    # value read as 125 inches.
    r"(\d+\s*/\s*\d+|\d*\.\d+|\d+)\s*(?:\"|IN\b|INCH)?\s*(?:THK|THICK)",
    re.IGNORECASE,
)


def _dxf_text_values(dxf_path: Path) -> list[str]:
    """Text carried by TEXT/MTEXT entities in the ENTITIES section.

    DXF is line-pairs of group code then value, and codes 1 and 3 hold entity
    text. Reading them directly avoids a CAD dependency for what is a flat
    scan - but they have to be read *in context*: the HEADER section uses the
    same codes for its own variables, so scanning the whole file returned
    thousands of values like "AC1018" (the format version) and "ANSI_1252"
    (the codepage). On a real shop DXF that was 2630 junk strings.
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

    lines = [line.strip() for line in text.splitlines()]
    values: list[str] = []
    in_entities = False
    in_text_entity = False

    for index in range(len(lines) - 1):
        code, value = lines[index], lines[index + 1]

        if code == "0":
            if value == "SECTION":
                in_text_entity = False
                continue
            if value == "ENDSEC":
                in_entities = False
                in_text_entity = False
                continue
            # A new entity ends the previous one; only TEXT-bearing kinds
            # contribute. ATTRIB/ATTDEF carry title-block values on some
            # drawings, so they count too.
            in_text_entity = in_entities and value in (
                "TEXT",
                "MTEXT",
                "ATTRIB",
                "ATTDEF",
            )
            continue

        if code == "2" and not in_entities:
            in_entities = value == "ENTITIES"
            continue

        if in_text_entity and code in ("1", "3"):
            # MTEXT carries formatting codes like \pxqc; \fArial|b0; - strip
            # them so the words survive but the markup doesn't.
            cleaned = re.sub(r"\\[A-Za-z][^;\\]*;?", " ", value)
            cleaned = re.sub(r"[{}]", " ", cleaned)
            cleaned = re.sub(r"\s+", " ", cleaned).strip()
            if cleaned:
                values.append(cleaned)
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
                # Kept as written, not normalised: matching is done on word
                # boundaries, and normalising "cold rolled" to "coldrolled"
                # would destroy the token boundary that makes that safe.
                alias = str(row.get("Alias", "") or "").strip()
                material = str(row.get("Material", "") or "").strip()
                if alias and material:
                    aliases[alias.casefold()] = material
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

    # Matched on word boundaries, never as bare substrings. Substring matching
    # meant ordinary words predicted materials - VALUE contains "alu",
    # ASSEMBLY contains "ss", ITEMS contains "ms" - and since the email body is
    # scraped, that was ordinary prose deciding what gets cut.
    #
    # Single-word aliases must match a whole token; multiword aliases must
    # match a run of consecutive tokens ("cold rolled", "checker plate").
    tokens = [word.casefold() for word in re.findall(r"[A-Za-z0-9]+", raw)]
    for alias, material in _read_material_aliases().items():
        if not alias or material not in available:
            continue
        alias_tokens = [
            word.casefold() for word in re.findall(r"[A-Za-z0-9]+", alias)
        ]
        if not alias_tokens:
            continue
        span = len(alias_tokens)
        if any(
            tokens[index : index + span] == alias_tokens
            for index in range(len(tokens) - span + 1)
        ):
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


def _page_names_part(words: list[tuple], cells: dict[str, str], wanted: str) -> bool:
    """Whether this page is about `wanted` (already normalised).

    Prefers the title block's drawing-number cell, but falls back to the part
    number appearing anywhere on the page. Prints are routinely named and
    numbered by the customer's own system rather than after the part file, and
    in that case the part number still shows up somewhere on the sheet - so
    refusing to look outside the title block would reject the whole drawing.
    """
    if not wanted:
        return True

    stated = _normalize_match_key(cells.get("part", ""))
    if stated and (wanted in stated or stated in wanted):
        return True

    anywhere = _normalize_match_key(" ".join(str(word[4]) for word in words))
    return wanted in anywhere


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

        if not _page_names_part(words, cells, wanted):
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


# --- CSV/XLSX BOM via inventor_to_radan --------------------------------------
#
# The top of the hierarchy: a spreadsheet BOM is structured output from CAM,
# not something to be scraped. inventor_to_radan already turns one into the
# exact 6-column Radan CSV this feature builds by hand, and owns the rules
# precedence (ftq_parts.csv overriding description_rules.csv) along with the
# missing-DXF classification. So it is called, not reimplemented.
#
# Its inline_runner.run_inline() is a documented cross-repo contract - the
# signature and the exceptions it raises are relied on by truck_nest_explorer
# too - so this only calls it and never reaches past it.

BOM_SPREADSHEET_SUFFIXES = (".csv", ".xlsx")
RADAN_OUTPUT_SUFFIX = "_Radan.csv"


def _load_inventor_to_radan():
    """The sibling app's inline runner, imported lazily.

    Lazily, and never at module import, so this repo stays usable when
    inventor_to_radan isn't installed - the same rule explorer_bridge follows
    for truck_nest_explorer.
    """
    import importlib
    import sys

    root = str(INVENTOR_TO_RADAN_DIR.resolve())
    if not INVENTOR_TO_RADAN_DIR.exists():
        raise JobIntakeError(f"inventor_to_radan was not found at {root}")
    if root not in sys.path:
        sys.path.insert(0, root)
    return importlib.import_module("inline_runner")


def _looks_like_bom_spreadsheet(path: Path) -> bool:
    """Cheap check that this spreadsheet is a BOM rather than any other CSV.

    inventor_to_radan raises if it can't find a header row, and a job folder
    can hold unrelated CSVs, so the obvious columns are looked for first.
    """
    if path.suffix.casefold() == ".xlsx":
        return True
    try:
        with path.open(newline="", encoding="utf-8-sig", errors="replace") as handle:
            head = " ".join(next(handle, "") for _ in range(10)).casefold()
    except OSError:
        return False
    return ("qty" in head or "quantity" in head) and (
        "part" in head or "description" in head
    )


# Why a conversion ended the way it did. The blanket "return an empty
# BomConversion" this replaces made a converter that refused to run look
# exactly like a BOM with nothing in it, so a job whose authority was
# unreadable quietly fell through to the scraped sources instead of stopping.
BOM_CONVERTED = "converted"
BOM_NOT_ATTEMPTED = "not_attempted"          # no spreadsheet to convert
BOM_NEEDS_RULES = "needs_rules"              # a description has no CAM rule yet
BOM_NEEDS_DXF_CLASSIFICATION = "needs_missing_dxf_classification"
BOM_INVALID = "invalid_bom"                  # recognised, but unreadable
BOM_CONVERTER_UNAVAILABLE = "converter_unavailable"
BOM_UNEXPECTED_FAILURE = "unexpected_failure"

# Everything except a clean conversion and "there was no BOM". A spreadsheet
# BOM is the authority for this job, so failing to read one is a stop, not a
# reason to guess from the prints.
BOM_BLOCKING_OUTCOMES = (
    BOM_NEEDS_RULES,
    BOM_NEEDS_DXF_CLASSIFICATION,
    BOM_INVALID,
    BOM_CONVERTER_UNAVAILABLE,
    BOM_UNEXPECTED_FAILURE,
)


@dataclass(frozen=True)
class BomConversion:
    """What inventor_to_radan produced, and what it noticed while doing it."""

    rows: tuple[dict[str, Any], ...] = ()
    # Its own accountability checks, reported rather than recomputed here.
    orphan_dxfs: tuple[str, ...] = ()          # in the folder, not in the BOM
    expected_missing_dxfs: tuple[str, ...] = ()  # in the BOM, no DXF found
    missing_pdfs: tuple[str, ...] = ()
    nonlaser_parts: tuple[str, ...] = ()
    outcome: str = BOM_NOT_ATTEMPTED
    # What to tell the user, in their terms, when the outcome blocks.
    reason: str = ""

    @property
    def blocked(self) -> bool:
        return self.outcome in BOM_BLOCKING_OUTCOMES


def convert_bom_spreadsheet(
    bom_path: Path, *, move_output_to: Path | None = None
) -> BomConversion:
    """Hand a spreadsheet BOM to inventor_to_radan and read back what it made.

    Returns its rows plus its own findings. Those findings are taken from the
    result object rather than recomputed here: it already checks DXF
    accountability as part of converting, and a second implementation would
    only be a second thing to disagree.

    `move_output_to` cuts the two files it writes - the Radan CSV and the audit
    report - into the job folder afterwards. The converter has to write beside
    the spreadsheet, because it resolves each DXF relative to it; when the
    spreadsheet is on W: that is the only write this app makes there, and it is
    moved rather than copied so nothing is left behind in the customer's
    folder.

    Never raises. The outcome says which of "it converted", "there was nothing
    to convert" and "it could not be converted, and here is why" happened -
    only the first two let intake carry on with the scraped sources. A BOM is
    the authority for the job, so one that exists but cannot be read has to
    stop rather than be silently demoted.
    """
    bom_path = Path(bom_path)
    if not bom_path.exists():
        return BomConversion()

    try:
        runner = _load_inventor_to_radan()
    except Exception as exc:
        return BomConversion(
            outcome=BOM_CONVERTER_UNAVAILABLE,
            reason=(
                f"{bom_path.name} is a BOM, but inventor_to_radan could not be "
                f"loaded to read it ({exc}). The parts cannot be trusted until "
                f"it converts."
            ),
        )

    try:
        result = runner.run_inline(
            INVENTOR_TO_RADAN_DIR / "inventor_to_radan.py",
            bom_path,
            allow_prompts=False,
            show_summary=False,
        )
    except Exception as exc:
        return _conversion_failure(bom_path, exc)

    findings = {
        "orphan_dxfs": tuple(getattr(result, "orphan_dxfs", ()) or ()),
        "expected_missing_dxfs": tuple(getattr(result, "expected_missing_dxfs", ()) or ()),
        "missing_pdfs": tuple(getattr(result, "missing_pdfs", ()) or ()),
        "nonlaser_parts": tuple(getattr(result, "nonlaser_parts", ()) or ()),
    }

    output = Path(getattr(result, "out_path", "") or "") or bom_path.with_name(
        f"{bom_path.stem}{RADAN_OUTPUT_SUFFIX}"
    )
    if not output.exists():
        return BomConversion(
            outcome=BOM_INVALID,
            reason=(
                f"{bom_path.name} converted without error but produced no "
                f"RADAN CSV at {output.name}."
            ),
            **findings,
        )

    rows: list[dict[str, Any]] = []
    try:
        with output.open(newline="", encoding="utf-8-sig") as handle:
            for record in csv.reader(handle):
                # The agreed headerless shape:
                # dxf_path, qty, material, thickness, unit, strategy
                if len(record) < 6:
                    continue
                try:
                    qty = int(str(record[1]).strip())
                    thickness = float(str(record[3]).strip())
                except ValueError:
                    continue
                rows.append(
                    {
                        "filename": Path(str(record[0]).strip()).name,
                        "qty": qty,
                        "material": str(record[2]).strip(),
                        "thickness": thickness,
                        "unit": str(record[4]).strip() or "in",
                        "strategy": str(record[5]).strip(),
                    }
                )
    except OSError as exc:
        return BomConversion(
            outcome=BOM_INVALID,
            reason=f"{bom_path.name} converted, but its RADAN CSV could not be read ({exc}).",
            **findings,
        )

    if move_output_to is not None:
        rows = _relocate_converter_output(
            output, Path(getattr(result, "report_path", "") or ""), Path(move_output_to), rows
        )

    return BomConversion(rows=tuple(rows), outcome=BOM_CONVERTED, **findings)


def _conversion_failure(bom_path: Path, exc: Exception) -> BomConversion:
    """Name what inventor_to_radan refused to do, in the user's terms.

    The two it raises most are recoverable in a specific, actionable way: a
    description with no CAM rule yet, and a BOM line whose DXF is missing and
    needs classifying. Both are answered by running inventor_to_radan
    interactively, which is the shop's normal way of adding a rule - so the
    message says that rather than reporting a generic failure.

    Matched on the exception's class *name*, not with isinstance. inline_runner
    execs inventor_to_radan.py under a private module name and drops it again
    afterwards, so its exception classes are rebuilt on every call and are
    never the same objects as any this process could import. isinstance would
    silently never match, which is exactly the kind of quietly-inert check this
    work exists to remove.
    """
    names = {cls.__name__ for cls in type(exc).__mro__}

    if "InventorToRadanNeedsUi" in names:
        missing_rules = list(getattr(exc, "missing_rules", ()) or ())
        missing_dxfs = list(getattr(exc, "missing_dxf_items", ()) or ())
        if missing_rules:
            shown = ", ".join(str(rule) for rule in missing_rules[:5])
            more = f" (and {len(missing_rules) - 5} more)" if len(missing_rules) > 5 else ""
            return BomConversion(
                outcome=BOM_NEEDS_RULES,
                reason=(
                    f"{bom_path.name} uses {len(missing_rules)} description(s) with no "
                    f"RADAN rule yet: {shown}{more}. Open {bom_path.name} in "
                    f"inventor_to_radan to add the rule(s), then re-apply the BOM here."
                ),
            )
        return BomConversion(
            outcome=BOM_NEEDS_DXF_CLASSIFICATION,
            reason=(
                f"{bom_path.name} has {len(missing_dxfs)} line(s) whose DXF is missing "
                f"and needs classifying. Open it in inventor_to_radan to answer that, "
                f"then re-apply the BOM here."
            ),
        )

    if names & {"InventorToRadanCancelled", "InventorToRadanReportRejected"}:
        return BomConversion(
            outcome=BOM_INVALID,
            reason=(
                f"{bom_path.name} was not converted: inventor_to_radan stopped "
                f"before finishing ({exc})."
            ),
        )

    if isinstance(exc, ImportError):
        return BomConversion(
            outcome=BOM_CONVERTER_UNAVAILABLE,
            reason=f"inventor_to_radan could not run against {bom_path.name}: {exc}",
        )

    return BomConversion(
        outcome=BOM_UNEXPECTED_FAILURE,
        reason=f"{bom_path.name} could not be converted: {type(exc).__name__}: {exc}",
    )


def _relocate_converter_output(
    csv_path: Path, report_path: Path, destination: Path, rows: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    """Move the converter's CSV and report into the job folder.

    Moved, not copied: the spreadsheet may live in a customer folder on W:, and
    leaving generated files there would be untidy at best. The rows themselves
    still reference the DXFs where they sit, which is the point - only the two
    files this app caused to be written are relocated.
    """
    destination.mkdir(parents=True, exist_ok=True)
    for source in (csv_path, report_path):
        if not source or not source.exists():
            continue
        target = destination / source.name
        try:
            if target.exists():
                target.unlink()
            shutil.move(str(source), str(target))
        except OSError:
            # Leaving it where it is beats failing the intake over a tidy-up.
            continue
    return rows


def apply_cam_bom(entry: dict[str, Any]) -> tuple[int, str]:
    """Re-run the spreadsheet BOM for an existing intake and update its rows.

    Intake already does this automatically, but a BOM is often dropped into
    the folder after the job was filed, or the conversion needed something
    that wasn't there yet. Returns (rows updated, message).

    Values it sets are treated exactly like the ones intake produced: still
    unverified, because a human confirming the material is what teaches the
    matcher, and CAM output being structured doesn't make it the right job.
    """
    attachments = {
        str(item.get("filename", "")): Path(str(item.get("saved_path", "")))
        for item in entry.get("attachments", [])
    }
    candidates = [
        path
        for name, path in attachments.items()
        if path.suffix.casefold() in BOM_SPREADSHEET_SUFFIXES
        and not name.casefold().endswith(RADAN_OUTPUT_SUFFIX.casefold())
        and path.exists()
        and _looks_like_bom_spreadsheet(path)
    ]
    if not candidates:
        return 0, "No spreadsheet BOM (.csv/.xlsx) is attached to this job."

    rows: dict[str, dict[str, Any]] = {}
    notes: list[str] = []
    blockers: list[str] = []
    for candidate in candidates:
        conversion = convert_bom_spreadsheet(candidate)
        if conversion.blocked:
            blockers.append(conversion.reason)
            continue
        for row in conversion.rows:
            rows[Path(str(row["filename"])).stem.casefold()] = row
        notes.extend(
            f"{orphan}: in the folder but not in the BOM" for orphan in conversion.orphan_dxfs
        )
        if rows:
            break

    # Recorded either way: a conversion that now succeeds must clear the block
    # left by an earlier attempt, which is the whole point of re-applying.
    entry["bom_blockers"] = blockers

    if not rows:
        if blockers:
            # Says what it wants, not just that it wanted something.
            return 0, "\n\n".join(blockers)
        return 0, (
            "inventor_to_radan converted the BOM but none of its rows matched a "
            "DXF on this job. Check the BOM is for this job."
        )

    updated = 0
    for part in entry.get("material_qty", []):
        row = rows.get(Path(str(part.get("filename", ""))).stem.casefold())
        if row is None:
            continue
        part["material"] = row["material"]
        part["thickness"] = row["thickness"]
        part["qty"] = row["qty"]
        part["unit"] = row["unit"] or "in"
        part["strategy"] = row["strategy"] or default_strategy_for_material(row["material"])
        part["material_source_text"] = "BOM via inventor_to_radan"
        part["qty_unknown"] = False
        part["material_confirmed"] = False
        updated += 1

    unmatched = [
        str(part.get("filename", ""))
        for part in entry.get("material_qty", [])
        if Path(str(part.get("filename", ""))).stem.casefold() not in rows
    ]
    message = f"BOM applied to {updated} part(s)."
    if unmatched:
        message += " Not in the BOM: " + ", ".join(unmatched)
    if notes:
        message += " " + "; ".join(notes)
    return updated, message


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
    # Carried from the rules table when the description matched verbatim, so a
    # BOM line's own strategy is used rather than one derived from material.
    strategy: str | None = None


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

        material = thickness = strategy = None
        rule = rules.get(_normalize_match_key(description))
        if rule is not None:
            material = str(rule["material"])
            thickness = float(rule["thickness"])
            strategy = str(rule.get("strategy", "") or "") or None
        else:
            # Not a verbatim description - fall back to the usual flattening.
            material = _match_material_in_text(description)

        rows[stem] = BomRow(
            part=stem,
            description=description,
            qty=qty,
            material=material,
            thickness=thickness,
            strategy=strategy,
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
    # Loaded once, here, immediately before anything is written for RADAN. The
    # catalog is the authority for what exists, so this is where that is
    # actually enforced rather than assumed from whatever the rows were given
    # when the job was filed - the file may have changed since.
    catalog = material_thickness_catalog()
    if not catalog:
        raise JobIntakeError(
            "The shop's material catalog could not be read "
            f"({_description_rules_path()}), so nothing can be validated "
            "against it. Nothing has been imported. Check that the file is "
            "readable, then try again."
        )

    # A spreadsheet BOM that was recognised but could not be converted. It is
    # the authority for this job, so the scraped prints do not stand in for it.
    blockers = [str(reason) for reason in entry.get("bom_blockers", []) if reason]
    if blockers:
        raise JobIntakeError(
            "STOP: this job has a BOM that could not be read, so the parts "
            "below are not the BOM's answer.\n\n"
            + "\n".join(f"  - {reason}" for reason in blockers)
            + "\n\nNothing has been imported. Sort the BOM out and re-apply it."
        )

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
        # A disagreement between sources is a hard stop. Whoever asked for the
        # job has to say which is right - the alternative is cutting metal on
        # a guess about whose paperwork was newer.
        #
        # Checked per field: settling the material must not release a disputed
        # quantity. Each is cleared by choosing a value for that field, which
        # is a real decision rather than a generic acknowledgement.
        conflicts = part.get("conflicts") or {}
        resolved = part.get("resolved") or {}
        open_conflicts = [
            f"{field} - {detail}"
            for field, detail in conflicts.items()
            if detail and not resolved.get(field)
        ]
        if open_conflicts:
            problems.append(
                f"{filename}: STOP - "
                + "; ".join(open_conflicts)
                + ". Confirm with whoever requested the job, then set that field"
            )
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
        # Nothing stated a quantity, so the 1 in the row is a placeholder. Let
        # it through and a forty-off part gets cut once - so someone has to put
        # a real number in, even if that number turns out to be 1.
        if part.get("qty_unknown"):
            source = str(part.get("qty_source_text", "") or "").strip()
            detail = f' (the print says "{source}")' if source else ""
            problems.append(
                f"{filename}: no quantity was stated{detail} - set the real "
                f"quantity, the 1 shown is a placeholder"
            )
            continue
        # The final tuple has to exist in the catalog as loaded right now.
        # Everything upstream is a route into that file; this is the gate that
        # makes it true rather than intended.
        if material not in catalog:
            problems.append(
                f"{filename}: {material!r} is not in the shop's material list"
            )
            continue
        if thickness not in catalog[material]:
            available = ", ".join(f"{value:g}" for value in catalog[material]) or "none"
            problems.append(
                f"{filename}: {material} isn't stocked at {thickness:g} "
                f"(available: {available})"
            )
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


def assert_radan_is_safe_to_drive(explorer_services: Any, paths: JobPaths) -> None:
    """Refuse to drive RADAN over COM while it is unsafe to do so.

    Two separate hazards, both checked by truck_nest_explorer's own helpers
    rather than reimplemented here:

    - An open RADAN window. The headless helper drives RADAN over COM, and
      attaching to a session someone is working in can corrupt the project
      they have open.
    - An import already running for this project, tracked by a PID lock file.
      Two imports writing the same RPD is the same problem twice.
    """
    check_sessions = getattr(explorer_services, "visible_radan_sessions", None)
    if callable(check_sessions):
        try:
            sessions = tuple(check_sessions())
        except Exception:
            # A failed check must not block work; the lock below still applies.
            sessions = ()
        if sessions:
            detail = ", ".join(f"PID {pid} ({title})" for pid, title in sessions)
            raise JobIntakeError(
                "RADAN is open - close it before importing parts. The import "
                "drives RADAN over COM and would be working in the same session "
                f"you have open. Found: {detail}"
            )

    check_lock = getattr(explorer_services, "radan_csv_import_lock_status", None)
    if callable(check_lock):
        try:
            running, lock_path, pid = check_lock(paths.rpd_path)
        except Exception:
            return
        if running:
            raise JobIntakeError(
                f"A RADAN import is already running for this project (PID {pid}). "
                f"Wait for it to finish. Lock: {lock_path}"
            )


def launch_radan_import(
    explorer_services: Any,
    *,
    paths: JobPaths,
    csv_path: Path,
    log_path: Path,
) -> Any:
    """Real RADAN COM conversion via the proven headless import pipeline;
    symbols land flat in the intake dir next to their DXFs (M59919 shape).
    Never uses the experimental lab symbol writer.

    refresh_project_sheets adds the stock sheets to the project, which
    truck_nest_explorer's own full flow also does (full_flow_service.py:326).
    Without it the parts import but the project has nothing to nest them on,
    so nesting cannot follow.
    """
    assert_radan_is_safe_to_drive(explorer_services, paths)
    return explorer_services.launch_radan_csv_import(
        csv_path,
        paths.intake_dir,
        project_path=paths.rpd_path,
        log_path=log_path,
        refresh_project_sheets=True,
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
