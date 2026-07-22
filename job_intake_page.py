"""Job Intake page: one-off jobs (M/W/S/F/P + digits) that skip the truck/kit
scaffold but still get a job folder, a blank RPD, RADAN part import, and
block transfer.

The page only reads/writes the intake registry and drives job_intake_service;
RADAN work happens through the embedded truck_nest_explorer services module,
which is touched lazily inside button handlers - never at construction time
(the shell's tests build this page with a stub explorer api).
"""

from __future__ import annotations

from concurrent.futures import Future, ThreadPoolExecutor
from pathlib import Path
import time
from typing import Any

from PySide6.QtCore import QDate, Qt, QTimer
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QSplitter,
    QStyledItemDelegate,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

import job_intake_registry
import job_intake_service
from job_intake_registry import (
    STATUS_BLOCKS_SENT,
    STATUS_ERROR,
    STATUS_NEW,
    STATUS_PARTS_IMPORTED,
    STATUS_RPD_CREATED,
)
from job_intake_service import JobIntakeError, UNIT_CHOICES
from paths import APP_DIR


QUEUE_COLUMNS = ("Job #", "Label", "PO #", "Status", "Due", "Received", "Source")

PART_DXF_COL = 0
PART_MATERIAL_COL = 1
PART_THICKNESS_COL = 2
PART_QTY_COL = 3
PART_UNIT_COL = 4
PART_STRATEGY_COL = 5
PART_PO_REF_COL = 6
PART_COLUMNS = ("DXF", "Material", "Thickness", "Qty", "Unit", "Strategy", "PO Ref")

POLL_INTERVAL_MS = 4000


class MaterialComboDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        combo = QComboBox(parent)
        combo.setEditable(True)
        combo.addItem("")
        for material in job_intake_service.material_choices():
            combo.addItem(material)
        combo.setInsertPolicy(QComboBox.NoInsert)
        return combo

    def setEditorData(self, editor, index):
        editor.setCurrentText(str(index.data(Qt.EditRole) or ""))

    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText(), Qt.EditRole)


class ThicknessSpinDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        spin = QDoubleSpinBox(parent)
        spin.setDecimals(3)
        spin.setRange(0.0, 12.0)
        spin.setSingleStep(0.01)
        return spin

    def setEditorData(self, editor, index):
        try:
            editor.setValue(float(index.data(Qt.EditRole) or 0.0))
        except (TypeError, ValueError):
            editor.setValue(0.0)

    def setModelData(self, editor, model, index):
        model.setData(index, f"{editor.value():g}", Qt.EditRole)


class QtySpinDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        spin = QSpinBox(parent)
        spin.setRange(1, 9999)
        return spin

    def setEditorData(self, editor, index):
        try:
            editor.setValue(int(index.data(Qt.EditRole) or 1))
        except (TypeError, ValueError):
            editor.setValue(1)

    def setModelData(self, editor, model, index):
        model.setData(index, str(editor.value()), Qt.EditRole)


class UnitComboDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        combo = QComboBox(parent)
        for unit in UNIT_CHOICES:
            combo.addItem(unit)
        return combo

    def setEditorData(self, editor, index):
        editor.setCurrentText(str(index.data(Qt.EditRole) or "in"))

    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText(), Qt.EditRole)


class ManualIntakeDialog(QDialog):
    """Collects the job number (+ Label when the job folder already exists).
    The actual work happens in JobIntakePage._create_intake."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Manual Job Intake")
        self.setMinimumWidth(420)

        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.job_number_edit = QLineEdit()
        self.job_number_edit.setPlaceholderText("e.g. M59919")
        self.job_number_edit.editingFinished.connect(self._check_existing)
        self.label_edit = QLineEdit()
        self.label_edit.setPlaceholderText("Required when the job folder already exists")
        form.addRow("Job number", self.job_number_edit)
        form.addRow("Label", self.label_edit)
        layout.addLayout(form)

        self.hint_label = QLabel("")
        self.hint_label.setWordWrap(True)
        layout.addWidget(self.hint_label)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _check_existing(self) -> None:
        number = self.job_number_edit.text().strip().upper()
        if not number:
            self.hint_label.setText("")
            return
        try:
            if job_intake_service.job_folder_exists(number):
                self.hint_label.setText(
                    f"{number} already has a folder on L: - a Label is required so this "
                    "one-off gets its own subfolder."
                )
            else:
                self.hint_label.setText(f"{number} is a fresh job folder.")
        except JobIntakeError as exc:
            self.hint_label.setText(str(exc))

    def _validate_and_accept(self) -> None:
        number = self.job_number_edit.text().strip().upper()
        label = self.label_edit.text().strip()
        try:
            job_intake_service.resolve_job_root(number)
            if job_intake_service.job_folder_exists(number) and not label:
                raise JobIntakeError(
                    f"{number} already exists on L: - give this one-off a Label."
                )
        except JobIntakeError as exc:
            QMessageBox.warning(self, "Manual Job Intake", str(exc))
            return
        self.accept()

    def values(self) -> tuple[str, str]:
        return self.job_number_edit.text().strip().upper(), self.label_edit.text().strip()


class JobIntakePage(QWidget):
    def __init__(self, *, explorer_api: Any, parent: QWidget | None = None):
        super().__init__(parent)
        self._explorer_api = explorer_api
        self._entries: list[dict[str, Any]] = []
        self._selected_key: str | None = None
        self._loading_detail = False
        self._import_process: Any = None
        self._import_context: tuple[str, Path] | None = None
        self._blocks_executor = ThreadPoolExecutor(max_workers=1)
        self._blocks_future: Future | None = None
        self._blocks_key: str | None = None
        self._poll_timer = QTimer(self)
        self._poll_timer.setInterval(POLL_INTERVAL_MS)
        self._poll_timer.timeout.connect(self._poll_tick)
        self._build_ui()
        self.refresh()

    # --- UI assembly ---------------------------------------------------------

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(18, 16, 18, 16)
        root.setSpacing(12)

        header = QHBoxLayout()
        title_box = QVBoxLayout()
        title = QLabel("Job Intake")
        title.setObjectName("page_title")
        subtitle = QLabel(
            "One-off M/W/S/F/P jobs: folder + blank RPD, material/qty per DXF, RADAN import, blocks to machine."
        )
        subtitle.setObjectName("page_subtitle")
        subtitle.setWordWrap(True)
        title_box.addWidget(title)
        title_box.addWidget(subtitle)
        header.addLayout(title_box, 1)

        self.manual_intake_button = QPushButton("Manual Intake")
        self.manual_intake_button.clicked.connect(self._manual_intake)
        refresh_button = QPushButton("Refresh")
        refresh_button.clicked.connect(self.refresh)
        header.addWidget(self.manual_intake_button)
        header.addWidget(refresh_button)
        root.addLayout(header)

        splitter = QSplitter(Qt.Horizontal)

        self.queue_table = QTableWidget(0, len(QUEUE_COLUMNS))
        self.queue_table.setHorizontalHeaderLabels(QUEUE_COLUMNS)
        self.queue_table.verticalHeader().setVisible(False)
        self.queue_table.setAlternatingRowColors(True)
        self.queue_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.queue_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.queue_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.queue_table.itemSelectionChanged.connect(self._on_queue_selection)
        splitter.addWidget(self.queue_table)

        detail = QWidget()
        detail_layout = QVBoxLayout(detail)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        detail_layout.setSpacing(8)

        self.job_summary_label = QLabel("Select a job intake.")
        self.job_summary_label.setObjectName("panel_title")
        self.job_summary_label.setWordWrap(True)
        detail_layout.addWidget(self.job_summary_label)

        self.po_warning_label = QLabel("")
        self.po_warning_label.setWordWrap(True)
        self.po_warning_label.setVisible(False)
        self.po_warning_label.setStyleSheet(
            "color: #7C4A03; background: #FFF4D6; border: 1px solid #D7A93B; padding: 6px; border-radius: 4px;"
        )
        detail_layout.addWidget(self.po_warning_label)

        fields = QHBoxLayout()
        self.due_date_check = QCheckBox("Due")
        self.due_date_check.toggled.connect(lambda checked: self.due_date_edit.setEnabled(checked))
        self.due_date_edit = QDateEdit()
        self.due_date_edit.setCalendarPopup(True)
        self.due_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.due_date_edit.setEnabled(False)
        self.laser_hours_spin = QDoubleSpinBox()
        self.laser_hours_spin.setRange(0.0, 999.0)
        self.laser_hours_spin.setSingleStep(0.25)
        self.laser_hours_spin.setSuffix(" h")
        self.bend_hours_spin = QDoubleSpinBox()
        self.bend_hours_spin.setRange(0.0, 999.0)
        self.bend_hours_spin.setSingleStep(0.25)
        self.bend_hours_spin.setSuffix(" h")
        fields.addWidget(self.due_date_check)
        fields.addWidget(self.due_date_edit)
        fields.addSpacing(10)
        fields.addWidget(QLabel("Laser"))
        fields.addWidget(self.laser_hours_spin)
        fields.addSpacing(10)
        fields.addWidget(QLabel("Bend"))
        fields.addWidget(self.bend_hours_spin)
        fields.addStretch(1)
        detail_layout.addLayout(fields)

        self.parts_table = QTableWidget(0, len(PART_COLUMNS))
        self.parts_table.setHorizontalHeaderLabels(PART_COLUMNS)
        self.parts_table.verticalHeader().setVisible(False)
        self.parts_table.setAlternatingRowColors(True)
        self.parts_table.setItemDelegateForColumn(PART_MATERIAL_COL, MaterialComboDelegate(self.parts_table))
        self.parts_table.setItemDelegateForColumn(PART_THICKNESS_COL, ThicknessSpinDelegate(self.parts_table))
        self.parts_table.setItemDelegateForColumn(PART_QTY_COL, QtySpinDelegate(self.parts_table))
        self.parts_table.setItemDelegateForColumn(PART_UNIT_COL, UnitComboDelegate(self.parts_table))
        self.parts_table.itemChanged.connect(self._on_part_item_changed)
        detail_layout.addWidget(self.parts_table, 1)

        actions = QHBoxLayout()
        self.save_button = QPushButton("Save Details")
        self.save_button.clicked.connect(self._save_details)
        self.create_rpd_button = QPushButton("Create Blank RPD")
        self.create_rpd_button.clicked.connect(self._create_rpd)
        self.import_button = QPushButton("Import Parts to RADAN")
        self.import_button.clicked.connect(self._import_parts)
        self.send_blocks_button = QPushButton("Send Blocks to Machine")
        self.send_blocks_button.clicked.connect(self._send_blocks)
        for button in (self.save_button, self.create_rpd_button, self.import_button, self.send_blocks_button):
            actions.addWidget(button)
        actions.addStretch(1)
        detail_layout.addLayout(actions)

        self.activity_label = QLabel("")
        self.activity_label.setWordWrap(True)
        self.activity_label.setObjectName("page_subtitle")
        detail_layout.addWidget(self.activity_label)

        splitter.addWidget(detail)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        root.addWidget(splitter, 1)

    # --- lifecycle -----------------------------------------------------------

    def showEvent(self, event) -> None:  # noqa: N802 - Qt override
        super().showEvent(event)
        self._poll_timer.start()

    def hideEvent(self, event) -> None:  # noqa: N802 - Qt override
        self._poll_timer.stop()
        super().hideEvent(event)

    # --- queue ---------------------------------------------------------------

    def refresh(self) -> None:
        self._entries = job_intake_registry.load_entries()
        self.queue_table.blockSignals(True)
        try:
            self.queue_table.setRowCount(len(self._entries))
            selected_row = -1
            for row, entry in enumerate(self._entries):
                key = str(entry.get("key", ""))
                values = (
                    str(entry.get("job_number", "")),
                    str(entry.get("label") or ""),
                    str(entry.get("po_number") or ""),
                    str(entry.get("status", "")),
                    str(entry.get("due_date") or ""),
                    str(entry.get("received_at", ""))[:16].replace("T", " "),
                    str(entry.get("source", "")),
                )
                for column, value in enumerate(values):
                    item = QTableWidgetItem(value)
                    item.setData(Qt.UserRole, key)
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    self.queue_table.setItem(row, column, item)
                if key == self._selected_key:
                    selected_row = row
            if selected_row >= 0:
                self.queue_table.selectRow(selected_row)
            elif self._entries and self._selected_key is None:
                self.queue_table.selectRow(0)
        finally:
            self.queue_table.blockSignals(False)
        self.queue_table.resizeColumnsToContents()
        self._on_queue_selection()

    def _selected_entry(self) -> dict[str, Any] | None:
        row = self.queue_table.currentRow()
        if row < 0:
            return None
        item = self.queue_table.item(row, 0)
        if item is None:
            return None
        key = str(item.data(Qt.UserRole) or "")
        return next((entry for entry in self._entries if str(entry.get("key", "")) == key), None)

    def _on_queue_selection(self) -> None:
        entry = self._selected_entry()
        self._selected_key = str(entry.get("key", "")) if entry else None
        self._load_detail(entry)
        self._update_button_states(entry)

    # --- detail panel --------------------------------------------------------

    def _load_detail(self, entry: dict[str, Any] | None) -> None:
        self._loading_detail = True
        try:
            if entry is None:
                self.job_summary_label.setText("Select a job intake.")
                self.po_warning_label.setVisible(False)
                self.parts_table.setRowCount(0)
                self.due_date_check.setChecked(False)
                self.laser_hours_spin.setValue(0.0)
                self.bend_hours_spin.setValue(0.0)
                return

            summary = f"{entry.get('job_number', '')}"
            if entry.get("label"):
                summary += f" - {entry['label']}"
            if entry.get("po_number"):
                summary += f"  |  PO {entry['po_number']}"
            if entry.get("due_note"):
                summary += f"  |  {entry['due_note']}"
            summary += f"  |  {entry.get('status', '')}"
            if entry.get("job_folder"):
                summary += f"\n{entry['job_folder']}"
            if entry.get("error"):
                summary += f"\nLast error: {entry['error']}"
            self.job_summary_label.setText(summary)

            unmatched = [str(line) for line in entry.get("po_unmatched", []) if str(line).strip()]
            if unmatched:
                self.po_warning_label.setText(
                    "PO lines with no matching DXF attachment:\n- " + "\n- ".join(unmatched)
                )
                self.po_warning_label.setVisible(True)
            else:
                self.po_warning_label.setVisible(False)

            due_text = str(entry.get("due_date") or "")
            parsed_due = QDate.fromString(due_text, "yyyy-MM-dd")
            if due_text and parsed_due.isValid():
                self.due_date_check.setChecked(True)
                self.due_date_edit.setDate(parsed_due)
            else:
                self.due_date_check.setChecked(False)
                self.due_date_edit.setDate(QDate.currentDate())
            self.laser_hours_spin.setValue(float(entry.get("laser_hours") or 0.0))
            self.bend_hours_spin.setValue(float(entry.get("bend_hours") or 0.0))

            parts = list(entry.get("material_qty", []))
            self.parts_table.blockSignals(True)
            try:
                self.parts_table.setRowCount(len(parts))
                for row, part in enumerate(parts):
                    values = (
                        str(part.get("filename", "")),
                        str(part.get("material", "") or ""),
                        f"{float(part.get('thickness') or 0):g}",
                        str(int(part.get("qty") or 1)),
                        str(part.get("unit", "") or "in"),
                        str(part.get("strategy", "") or ""),
                        str(part.get("po_ref", "") or ""),
                    )
                    for column, value in enumerate(values):
                        item = QTableWidgetItem(value)
                        if column in (PART_DXF_COL, PART_STRATEGY_COL, PART_PO_REF_COL):
                            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                        self.parts_table.setItem(row, column, item)
            finally:
                self.parts_table.blockSignals(False)
            self.parts_table.resizeColumnsToContents()
        finally:
            self._loading_detail = False

    def _on_part_item_changed(self, item: QTableWidgetItem) -> None:
        if self._loading_detail or item.column() != PART_MATERIAL_COL:
            return
        strategy_item = self.parts_table.item(item.row(), PART_STRATEGY_COL)
        if strategy_item is not None:
            strategy_item.setText(job_intake_service.default_strategy_for_material(item.text()))

    def _collect_detail_fields(self) -> dict[str, Any]:
        parts: list[dict[str, Any]] = []
        for row in range(self.parts_table.rowCount()):
            def _cell(column: int) -> str:
                cell = self.parts_table.item(row, column)
                return cell.text().strip() if cell is not None else ""

            try:
                thickness = float(_cell(PART_THICKNESS_COL) or 0)
            except ValueError:
                thickness = 0.0
            try:
                qty = int(_cell(PART_QTY_COL) or 1)
            except ValueError:
                qty = 1
            parts.append(
                {
                    "filename": _cell(PART_DXF_COL),
                    "material": _cell(PART_MATERIAL_COL),
                    "thickness": thickness,
                    "qty": qty,
                    "unit": _cell(PART_UNIT_COL) or "in",
                    "strategy": _cell(PART_STRATEGY_COL),
                    "po_ref": _cell(PART_PO_REF_COL),
                }
            )
        due_date = (
            self.due_date_edit.date().toString("yyyy-MM-dd") if self.due_date_check.isChecked() else None
        )
        return {
            "material_qty": parts,
            "due_date": due_date,
            "laser_hours": self.laser_hours_spin.value() or None,
            "bend_hours": self.bend_hours_spin.value() or None,
        }

    def _update_button_states(self, entry: dict[str, Any] | None) -> None:
        status = str(entry.get("status", "")) if entry else ""
        has_entry = entry is not None
        self.save_button.setEnabled(has_entry)
        self.create_rpd_button.setEnabled(status == STATUS_NEW)
        self.import_button.setEnabled(status in (STATUS_RPD_CREATED, STATUS_ERROR) and self._import_process is None)
        self.send_blocks_button.setEnabled(
            status in (STATUS_PARTS_IMPORTED, STATUS_BLOCKS_SENT) and self._blocks_future is None
        )

    # --- actions -------------------------------------------------------------

    def _manual_intake(self) -> None:
        dialog = ManualIntakeDialog(self)
        if dialog.exec() != QDialog.Accepted:
            return
        job_number, label = dialog.values()
        files, _selected_filter = QFileDialog.getOpenFileNames(
            self,
            "Select the job's DXF files (plus any PO / reference PDFs)",
            "",
            "Job files (*.dxf *.pdf);;All files (*)",
        )
        if not files:
            return
        try:
            entry = self._create_intake(job_number, label, [Path(text) for text in files])
        except (JobIntakeError, ValueError) as exc:
            QMessageBox.warning(self, "Manual Job Intake", str(exc))
            return
        self._selected_key = str(entry.get("key", ""))
        self.refresh()

    def _create_intake(self, job_number: str, label: str, files: list[Path]) -> dict[str, Any]:
        """Testable seam the Manual Intake dialog calls. The sequence itself
        lives in job_intake_service so the 127.0.0.1 listener runs the same
        code path with source="outlook"."""
        return job_intake_service.create_intake(job_number, label or None, files, source="manual")

    def _save_details(self) -> dict[str, Any] | None:
        entry = self._selected_entry()
        if entry is None:
            return None
        fields = self._collect_detail_fields()
        updated = job_intake_registry.update_entry(str(entry["key"]), **fields)
        entry.update(updated)
        self.activity_label.setText("Details saved.")
        return updated

    def _create_rpd(self) -> None:
        entry = self._save_details()
        if entry is None:
            return
        try:
            paths = job_intake_service.resolve_job_paths(entry["job_number"], entry.get("label"))
            rpd_path = job_intake_service.clone_rpd_template(paths)
        except JobIntakeError as exc:
            QMessageBox.warning(self, "Create Blank RPD", str(exc))
            return
        job_intake_registry.update_entry(
            str(entry["key"]), status=STATUS_RPD_CREATED, rpd_path=str(rpd_path), error=None
        )
        self.activity_label.setText(f"Blank RPD created: {rpd_path}")
        self.refresh()

    def _import_parts(self) -> None:
        entry = self._save_details()
        if entry is None:
            return
        key = str(entry["key"])
        try:
            rows = job_intake_service.build_import_csv_rows(entry)
            paths = job_intake_service.resolve_job_paths(entry["job_number"], entry.get("label"))
            csv_path = job_intake_service.write_import_csv(
                rows, paths.intake_dir / f"{paths.project_name}-BOM_Radan.csv"
            )
        except JobIntakeError as exc:
            QMessageBox.warning(self, "Import Parts to RADAN", str(exc))
            return

        log_dir = APP_DIR / "_runtime"
        log_dir.mkdir(parents=True, exist_ok=True)
        safe_key = "".join(char if char.isalnum() else "_" for char in key).strip("_")
        log_path = log_dir / f"job_intake_radan_import_{safe_key}_{int(time.time())}.log"
        try:
            self._import_process = job_intake_service.launch_radan_import(
                self._explorer_api.services,
                paths=paths,
                csv_path=csv_path,
                log_path=log_path,
            )
        except Exception as exc:
            QMessageBox.critical(self, "Import Parts to RADAN", str(exc))
            return
        self._import_context = (key, log_path)
        job_intake_registry.update_entry(key, csv_log_path=str(log_path))
        self.activity_label.setText(f"RADAN import running... log: {log_path}")
        self._update_button_states(entry)

    def _send_blocks(self) -> None:
        entry = self._selected_entry()
        if entry is None:
            return
        choice = QMessageBox.question(
            self,
            "Send Blocks to Machine",
            "Copy this job's block files to the machine folder (verified, then the source is deleted)?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if choice != QMessageBox.Yes:
            return
        key = str(entry["key"])
        paths = job_intake_service.resolve_job_paths(entry["job_number"], entry.get("label"))
        services = self._explorer_api.services
        self._blocks_key = key
        self._blocks_future = self._blocks_executor.submit(
            job_intake_service.send_job_blocks_to_machine, services, paths=paths
        )
        self.activity_label.setText("Sending block files to the machine...")
        self._update_button_states(entry)

    # --- background polling --------------------------------------------------

    def _poll_tick(self) -> None:
        changed = False

        if self._import_process is not None and self._import_context is not None:
            return_code = self._import_process.poll()
            if return_code is not None:
                key, log_path = self._import_context
                self._import_process = None
                self._import_context = None
                if return_code == 0:
                    job_intake_registry.update_entry(key, status=STATUS_PARTS_IMPORTED, error=None)
                    self.activity_label.setText("RADAN import finished - parts are in the RPD.")
                else:
                    job_intake_registry.update_entry(
                        key,
                        status=STATUS_ERROR,
                        error=f"RADAN import exited with code {return_code}; see {log_path}",
                    )
                    self.activity_label.setText(f"RADAN import failed - see {log_path}")
                changed = True

        if self._blocks_future is not None and self._blocks_future.done():
            future = self._blocks_future
            key = self._blocks_key or ""
            self._blocks_future = None
            self._blocks_key = None
            try:
                result = future.result()
            except Exception as exc:
                if key:
                    job_intake_registry.update_entry(key, status=STATUS_ERROR, error=str(exc))
                self.activity_label.setText(f"Block transfer failed: {exc}")
            else:
                transferred = len(getattr(result, "transferred_paths", ()) or ())
                if key:
                    job_intake_registry.update_entry(key, status=STATUS_BLOCKS_SENT, error=None)
                self.activity_label.setText(f"Block transfer complete: {transferred} file(s) sent.")
            changed = True

        if changed:
            self.refresh()
