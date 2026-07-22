from __future__ import annotations

import os
import sys
from pathlib import Path
from types import SimpleNamespace

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PySide6.QtWidgets import QApplication

import job_intake_registry
import job_intake_service
import job_intake_page
from job_intake_registry import STATUS_RPD_CREATED
from job_intake_page import PART_MATERIAL_COL, PART_QTY_COL, PART_STRATEGY_COL, JobIntakePage


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv[:1])
    return app


@pytest.fixture()
def isolated_roots(tmp_path: Path, monkeypatch):
    registry_path = tmp_path / "registry.json"
    monkeypatch.setattr(job_intake_service, "BATTLESHIELD_ROOT", tmp_path / "L")
    monkeypatch.setattr(job_intake_registry, "JOB_INTAKE_REGISTRY_PATH", registry_path)
    template_path = tmp_path / "Template.rpd"
    template_path.write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        "<RadanProject><JobName>Template</JobName>"
        "<NestFolder>old</NestFolder><RemnantSaveFolder>old</RemnantSaveFolder></RadanProject>",
        encoding="utf-8",
    )
    monkeypatch.setattr(job_intake_service, "EXPLORER_TEMPLATE_PATH", template_path)
    return tmp_path


def _make_page(qapp) -> JobIntakePage:
    fake_api = SimpleNamespace(services=SimpleNamespace())
    return JobIntakePage(explorer_api=fake_api)


def _dxf(tmp_path: Path, name: str) -> Path:
    source = tmp_path / name
    source.write_text("dxf data", encoding="utf-8")
    return source


def test_create_intake_registers_job_and_populates_queue(qapp, isolated_roots) -> None:
    page = _make_page(qapp)
    try:
        entry = page._create_intake(
            "M59919", "", [_dxf(isolated_roots, "Clip-End.DXF"), _dxf(isolated_roots, "Clip-Mid.DXF")]
        )
        assert entry["job_number"] == "M59919"
        assert (isolated_roots / "L" / "M-FABRICATION" / "M59919" / "Clip-End.DXF").exists()
        assert len(entry["material_qty"]) == 2

        page.refresh()
        assert page.queue_table.rowCount() == 1
        assert page.queue_table.item(0, 0).text() == "M59919"
        # Parts grid shows one row per DXF with editable material/qty cells.
        assert page.parts_table.rowCount() == 2
    finally:
        page.deleteLater()


def test_create_intake_requires_label_for_existing_job(qapp, isolated_roots) -> None:
    page = _make_page(qapp)
    try:
        page._create_intake("M59919", "", [_dxf(isolated_roots, "A.DXF")])
        with pytest.raises(job_intake_service.JobIntakeError):
            page._create_intake("M59919", "", [_dxf(isolated_roots, "B.DXF")])

        labeled = page._create_intake("M59919", "Rush Plates", [_dxf(isolated_roots, "B.DXF")])
        assert labeled["label"] == "Rush Plates"
        labeled_dir = isolated_roots / "L" / "M-FABRICATION" / "M59919" / "Rush Plates"
        assert (labeled_dir / "B.DXF").exists()
        assert (labeled_dir / "M59919 Rush Plates").is_dir()
    finally:
        page.deleteLater()


def test_create_rpd_transitions_status_and_writes_file(qapp, isolated_roots) -> None:
    page = _make_page(qapp)
    try:
        page._create_intake("W50123", "", [_dxf(isolated_roots, "Bracket.DXF")])
        page.refresh()
        page.queue_table.selectRow(0)
        page._create_rpd()

        entry = job_intake_registry.get_entry("W50123")
        assert entry is not None
        assert entry["status"] == STATUS_RPD_CREATED
        rpd_path = Path(entry["rpd_path"])
        assert rpd_path.exists()
        assert rpd_path.name == "W50123.rpd"
        assert "<JobName>W50123</JobName>" in rpd_path.read_text(encoding="utf-8")
    finally:
        page.deleteLater()


def test_material_edit_autofills_strategy_and_saves(qapp, isolated_roots) -> None:
    page = _make_page(qapp)
    try:
        page._create_intake("S50001", "", [_dxf(isolated_roots, "Panel.DXF")])
        page.refresh()
        page.queue_table.selectRow(0)

        page.parts_table.item(0, PART_MATERIAL_COL).setText("Mild Steel-A36")
        assert page.parts_table.item(0, PART_STRATEGY_COL).text() == "O2"
        page.parts_table.item(0, PART_QTY_COL).setText("12")
        page.laser_hours_spin.setValue(1.5)
        page._save_details()

        entry = job_intake_registry.get_entry("S50001")
        assert entry is not None
        part = entry["material_qty"][0]
        assert part["material"] == "Mild Steel-A36"
        assert part["strategy"] == "O2"
        assert part["qty"] == 12
        assert entry["laser_hours"] == 1.5
    finally:
        page.deleteLater()
