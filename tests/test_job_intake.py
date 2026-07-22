from __future__ import annotations

import csv
from datetime import date
from pathlib import Path

import pytest

import job_intake_registry
import job_intake_service
from job_intake_registry import (
    STATUS_RPD_CREATED,
    append_entry,
    delete_entry,
    entry_key,
    get_entry,
    load_entries,
    new_entry,
    update_entry,
)
from job_intake_service import (
    JobIntakeError,
    build_import_csv_rows,
    clone_rpd_template,
    create_job_folders,
    default_strategy_for_material,
    extract_po_hints,
    material_choices,
    resolve_job_paths,
    resolve_job_root,
    write_import_csv,
)


# --- registry ----------------------------------------------------------------


def test_registry_append_get_update_delete_round_trip(tmp_path: Path) -> None:
    registry_path = tmp_path / "registry.json"
    entry = new_entry(job_number="m59919", source="manual")
    append_entry(entry, registry_path)

    loaded = get_entry("M59919", registry_path)
    assert loaded is not None
    assert loaded["job_number"] == "M59919"
    assert loaded["status"] == "new"

    update_entry("M59919", registry_path, status=STATUS_RPD_CREATED, rpd_path="x.rpd")
    updated = get_entry("M59919", registry_path)
    assert updated is not None
    assert updated["status"] == STATUS_RPD_CREATED
    assert updated["rpd_path"] == "x.rpd"

    delete_entry("M59919", registry_path)
    assert get_entry("M59919", registry_path) is None


def test_registry_rejects_duplicate_keys_and_bad_status(tmp_path: Path) -> None:
    registry_path = tmp_path / "registry.json"
    append_entry(new_entry(job_number="M59919"), registry_path)
    with pytest.raises(ValueError):
        append_entry(new_entry(job_number="M59919"), registry_path)
    with pytest.raises(ValueError):
        update_entry("M59919", registry_path, status="not-a-status")
    # A labeled one-off under the same number is a distinct entry.
    append_entry(new_entry(job_number="M59919", label="Rush Plates"), registry_path)
    assert entry_key("M59919", "Rush Plates") == "M59919::rush plates"
    assert len(load_entries(registry_path)) == 2


def test_registry_load_entries_newest_first(tmp_path: Path) -> None:
    registry_path = tmp_path / "registry.json"
    first = new_entry(job_number="M50001")
    first["received_at"] = "2026-01-01T08:00:00"
    second = new_entry(job_number="M50002")
    second["received_at"] = "2026-07-01T08:00:00"
    append_entry(first, registry_path)
    append_entry(second, registry_path)
    assert [entry["job_number"] for entry in load_entries(registry_path)] == ["M50002", "M50001"]


# --- path resolution ---------------------------------------------------------


def test_resolve_job_root_maps_prefixes_and_rejects_bad_numbers() -> None:
    assert resolve_job_root("M59919") == "M-FABRICATION"
    assert resolve_job_root("w50123") == "W-WARRANTY"
    assert resolve_job_root("S123456") == "S-SERVICE"
    with pytest.raises(JobIntakeError):
        resolve_job_root("X59919")
    with pytest.raises(JobIntakeError):
        resolve_job_root("M59A19")
    with pytest.raises(JobIntakeError):
        resolve_job_root("")


def test_resolve_job_paths_fresh_job_matches_shop_convention() -> None:
    paths = resolve_job_paths("m59919")
    assert paths.intake_dir == paths.job_dir
    assert paths.job_dir.name == "M59919"
    assert paths.job_dir.parent.name == "M-FABRICATION"
    assert paths.project_dir == paths.job_dir / "M59919"
    assert paths.rpd_path == paths.project_dir / "M59919.rpd"


def test_resolve_job_paths_labeled_job_nests_under_label() -> None:
    paths = resolve_job_paths("F55334", "Rush Plates")
    assert paths.intake_dir == paths.job_dir / "Rush Plates"
    assert paths.project_name == "F55334 Rush Plates"
    assert paths.project_dir == paths.intake_dir / "F55334 Rush Plates"
    assert paths.rpd_path.name == "F55334 Rush Plates.rpd"


def test_create_job_folders_requires_label_when_job_exists(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path)
    fresh = resolve_job_paths("M59919")
    create_job_folders(fresh)
    assert (fresh.project_dir / "nests").is_dir()
    assert (fresh.project_dir / "remnants").is_dir()

    # Same number again without a label must refuse instead of mixing in.
    with pytest.raises(JobIntakeError):
        create_job_folders(resolve_job_paths("M59919"))

    labeled = resolve_job_paths("M59919", "Extra Brackets")
    create_job_folders(labeled)
    assert labeled.intake_dir.is_dir()
    with pytest.raises(JobIntakeError):
        create_job_folders(resolve_job_paths("M59919", "Extra Brackets"))


# --- RPD template clone ------------------------------------------------------


TEMPLATE_TEXT = """<?xml version="1.0" encoding="UTF-8"?>
<RadanProject xmlns="http://www.radan.com/ns/project">
  <JobName>Template</JobName>
  <NestFolder>C:\\old\\nests</NestFolder>
  <RemnantSaveFolder>C:\\old\\remnants</RemnantSaveFolder>
  <Part><Symbol>Template.rpd</Symbol></Part>
</RadanProject>
"""


def test_clone_rpd_template_substitutes_job_values(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path)
    template_path = tmp_path / "Template.rpd"
    template_path.write_text(TEMPLATE_TEXT, encoding="utf-8")

    paths = resolve_job_paths("M59919")
    create_job_folders(paths)
    rpd_path = clone_rpd_template(paths, template_path)

    text = rpd_path.read_text(encoding="utf-8")
    assert "<JobName>M59919</JobName>" in text
    assert f"<NestFolder>{paths.project_dir / 'nests'}</NestFolder>" in text
    assert f"<RemnantSaveFolder>{paths.project_dir / 'remnants'}</RemnantSaveFolder>" in text
    assert "<Symbol>M59919.rpd</Symbol>" in text
    assert "Template" not in text

    with pytest.raises(JobIntakeError):
        clone_rpd_template(paths, template_path)


# --- material list -----------------------------------------------------------


def test_material_choices_reads_rules_and_excludes_ftq(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr(job_intake_service, "INVENTOR_TO_RADAN_DIR", tmp_path)
    (tmp_path / "description_rules.csv").write_text(
        "Description,Material,Thickness,Strategy\n"
        "A,Aluminum 5052,0.12,Air\n"
        "B,Aluminum 3003 CHK FTQ,0.12,Air\n"
        "C,Mild Steel-A36,0.25,O2\n"
        "D,aluminum 5052,0.18,Air\n",
        encoding="utf-8",
    )
    choices = material_choices()
    assert "Aluminum 5052" in choices
    assert "Mild Steel-A36" in choices
    assert not any("FTQ" in choice for choice in choices)
    assert len([c for c in choices if c.casefold() == "aluminum 5052"]) == 1


def _write_shop_csvs(directory: Path, descriptions: list[str], rules: list[tuple[str, str, str, str]]) -> None:
    directory.mkdir(parents=True, exist_ok=True)
    with (directory / "expected_laser_descriptions.csv").open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Description"])
        for description in descriptions:
            writer.writerow([description])
    with (directory / "description_rules.csv").open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Description", "Material", "Thickness", "Strategy"])
        for row in rules:
            writer.writerow(list(row))


def test_catalog_offers_only_thicknesses_valid_for_each_material(tmp_path: Path, monkeypatch) -> None:
    """expected_laser_descriptions.csv is authoritative: a material/thickness
    pair the shop doesn't stock must not be selectable at all."""
    shop = tmp_path / "inventor_to_radan"
    _write_shop_csvs(
        shop,
        descriptions=[
            'SHEET, AL ALY, .125" THK, 5052 H32',
            'PLATE, AL ALY, .375" THK, 5052 H32',
            'PLATE, MS, .375" THK, 44W',
            'SHEET, AL ALY, .125" THK, 3003 H22 APT FTQ',
        ],
        rules=[
            ('SHEET, AL ALY, .125" THK, 5052 H32', "Aluminum 5052", "0.12", "Air"),
            ('PLATE, AL ALY, .375" THK, 5052 H32', "Aluminum 5052", "0.38", "Air"),
            ('PLATE, MS, .375" THK, 44W', "Mild Steel-A36", "0.375", "O2"),
            ('SHEET, AL ALY, .125" THK, 3003 H22 APT FTQ', "Aluminum 3003 CHK FTQ", "0.18", "Air"),
            # Present in the rules but NOT in the expected list - must not be offered.
            ("PLATE, SS, .25 THK, 304", "Stainless Steel", "0.25", "N2"),
        ],
    )
    monkeypatch.setattr(job_intake_service, "INVENTOR_TO_RADAN_DIR", shop)

    assert job_intake_service.material_choices() == ("Aluminum 5052", "Mild Steel-A36")
    # Stainless has a rule but no expected description, so it doesn't exist here.
    assert "Stainless Steel" not in job_intake_service.material_choices()
    # FTQ is a forced per-part override elsewhere, never a user choice.
    assert not any("FTQ" in material for material in job_intake_service.material_choices())

    assert job_intake_service.thickness_choices("Aluminum 5052") == (0.12, 0.38)
    assert job_intake_service.thickness_choices("Mild Steel-A36") == (0.375,)
    # 0.375 is valid for steel but must not be offered for aluminium.
    assert 0.375 not in job_intake_service.thickness_choices("Aluminum 5052")
    assert job_intake_service.thickness_choices("nothing like this") == ()


def test_catalog_picks_up_materials_added_after_launch(tmp_path: Path, monkeypatch) -> None:
    """The shop edits this CSV; new materials must appear without a restart,
    so nothing may be cached at import time."""
    shop = tmp_path / "inventor_to_radan"
    _write_shop_csvs(
        shop,
        descriptions=['PLATE, MS, .375" THK, 44W'],
        rules=[('PLATE, MS, .375" THK, 44W', "Mild Steel-A36", "0.375", "O2")],
    )
    monkeypatch.setattr(job_intake_service, "INVENTOR_TO_RADAN_DIR", shop)
    assert job_intake_service.material_choices() == ("Mild Steel-A36",)

    # The shop adds a material mid-session.
    _write_shop_csvs(
        shop,
        descriptions=['PLATE, MS, .375" THK, 44W', "PLATE, SS, .25 THK, 304"],
        rules=[
            ('PLATE, MS, .375" THK, 44W', "Mild Steel-A36", "0.375", "O2"),
            ("PLATE, SS, .25 THK, 304", "Stainless Steel", "0.25", "N2"),
        ],
    )
    assert job_intake_service.material_choices() == ("Mild Steel-A36", "Stainless Steel")
    assert job_intake_service.thickness_choices("Stainless Steel") == (0.25,)


def test_material_choices_falls_back_when_rules_missing(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr(job_intake_service, "INVENTOR_TO_RADAN_DIR", tmp_path / "missing")
    assert material_choices() == job_intake_service.FALLBACK_MATERIALS


def test_default_strategy_for_material() -> None:
    assert default_strategy_for_material("Aluminum 5052") == "Air"
    assert default_strategy_for_material("Mild Steel-A36") == "O2"
    assert default_strategy_for_material("Stainless Steel") == "N2"
    assert default_strategy_for_material("Stainless Steel 304") == "N2"
    assert default_strategy_for_material("Something New") == "Air"


# --- import CSV --------------------------------------------------------------


def _entry_with_parts(tmp_path: Path) -> dict:
    dxf = tmp_path / "Clip-End.DXF"
    dxf.write_text("dxf", encoding="utf-8")
    entry = new_entry(job_number="M59919")
    entry["attachments"] = [{"filename": "Clip-End.DXF", "saved_path": str(dxf), "size": 3}]
    entry["material_qty"] = [
        {"filename": "Clip-End.DXF", "material": "Mild Steel-A36", "thickness": 0.25, "unit": "in", "qty": 10, "strategy": ""}
    ]
    return entry


def test_build_import_csv_rows_happy_path_and_write(tmp_path: Path) -> None:
    entry = _entry_with_parts(tmp_path)
    rows = build_import_csv_rows(entry)
    assert rows == [[str(tmp_path / "Clip-End.DXF"), "10", "Mild Steel-A36", "0.25", "in", "O2"]]

    csv_path = write_import_csv(rows, tmp_path / "out" / "import.csv")
    with csv_path.open(newline="", encoding="utf-8") as handle:
        assert list(csv.reader(handle)) == rows


def test_build_import_csv_rows_reports_all_problems(tmp_path: Path) -> None:
    entry = _entry_with_parts(tmp_path)
    entry["material_qty"][0]["material"] = ""
    entry["material_qty"].append(
        {"filename": "Missing.DXF", "material": "Aluminum 5052", "thickness": 0.12, "unit": "in", "qty": 2, "strategy": ""}
    )
    entry["material_qty"].append(
        {"filename": "Clip-End.DXF", "material": "Aluminum 5052", "thickness": 0, "unit": "in", "qty": 2, "strategy": ""}
    )
    with pytest.raises(JobIntakeError) as excinfo:
        build_import_csv_rows(entry)
    message = str(excinfo.value)
    assert "pick a material" in message
    assert "Missing.DXF" in message
    assert "thickness" in message


# --- PO extraction -----------------------------------------------------------
# Synthetic PDFs are built with the same one-cell-per-line layout PyMuPDF
# produces for the real PFF PO template (verified against 5 real POs on L:).


def _write_po_pdf(tmp_path: Path, lines: list[str]) -> Path:
    fitz = pytest.importorskip("fitz")
    pdf_path = tmp_path / "po.pdf"
    doc = fitz.open()
    page = doc.new_page()
    y = 40.0
    for line in lines:
        page.insert_text((40, y), line)
        y += 14.0
    doc.save(str(pdf_path))
    doc.close()
    return pdf_path


PO_BODY = [
    "Date:",
    "PO Number:",
    "Date Required:",
    "Line",
    "Qty",
    "DESCRIPTION",
    "PRICING",
    "Subtotal",
    "1",
    "10",
    'Clip-End - 1/4" Mild Steel',
    "2",
    "36",
    'Clip-Mid - 1/4" Mild Steel',
    "3",
    "2",
    "End Cap D_2",
    "4",
    'ALL MATERIAL 1/8" MILD STEEL',
    "5",
    "6",
    "Sub-total",
    "0",
    "LASER ORDER",
    "July 21, 2026",
    "8665-001",
    "July 28, 2026",
]


def test_extract_po_hints_matches_lines_and_reports_unmatched(tmp_path: Path) -> None:
    pdf_path = _write_po_pdf(tmp_path, PO_BODY)
    hints = extract_po_hints(pdf_path, ["Clip-End", "Clip-Mid", "End Cap D", "End Cap D_2"])

    assert hints.po_number == "8665-001"
    assert hints.due_date == date(2026, 7, 28)
    assert hints.line_items["Clip-End"] == {"qty": 10, "raw_description": 'Clip-End - 1/4" Mild Steel'}
    assert hints.line_items["Clip-Mid"]["qty"] == 36
    # Longest-stem-first: D_2's row must not be claimed by "End Cap D".
    assert hints.line_items["End Cap D_2"]["qty"] == 2
    assert "End Cap D" not in hints.line_items
    # The order-wide material note surfaces as unmatched; footer labels don't.
    assert 'ALL MATERIAL 1/8" MILD STEEL' in hints.unmatched_lines
    assert not any("Sub-total" in line for line in hints.unmatched_lines)


def test_extract_po_hints_flags_po_lines_with_no_dxf(tmp_path: Path) -> None:
    pdf_path = _write_po_pdf(tmp_path, PO_BODY)
    hints = extract_po_hints(pdf_path, ["Clip-End"])
    assert 'Clip-Mid - 1/4" Mild Steel' in hints.unmatched_lines


def test_extract_po_hints_single_date_is_not_a_due_date(tmp_path: Path) -> None:
    # A lone date is the order date (Date Required said RUSH/ASAP or was
    # blank) - never claim it as the due date, but surface the urgency note.
    body = [line if line != "July 28, 2026" else "RUSH" for line in PO_BODY]
    hints = extract_po_hints(_write_po_pdf(tmp_path, body), ["Clip-End"])
    assert hints.due_date is None
    assert hints.due_note == "RUSH"
    assert hints.po_number == "8665-001"

    blank_body = [line for line in PO_BODY if line != "July 28, 2026"]
    blank_hints = extract_po_hints(_write_po_pdf(tmp_path, blank_body), ["Clip-End"])
    assert blank_hints.due_date is None
    assert blank_hints.due_note is None


def test_extract_po_hints_survives_non_pdf_garbage(tmp_path: Path) -> None:
    garbage = tmp_path / "not_really.pdf"
    garbage.write_bytes(b"this is not a pdf")
    hints = extract_po_hints(garbage, ["Clip-End"])
    assert hints.po_number is None
    assert hints.due_date is None
    assert hints.line_items == {}
