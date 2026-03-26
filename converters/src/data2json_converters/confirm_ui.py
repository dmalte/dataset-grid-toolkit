# Copyright (c) 2026 Malte Doerper. MIT License. See LICENSE file.

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .converter import (
    apply_review_artifact,
    build_default_review_json_output_path,
    build_review_artifact,
    save_review_artifact,
)


def json_to_excel_confirm_ui(
    json_path: Path,
    *,
    excel_path: Path,
    sheet_name: str,
    include_source_ref: bool,
    review_json_path: Path | None = None,
) -> tuple[dict[str, Any], list[str]]:
    review_artifact, warnings = build_review_artifact(
        json_path,
        excel_path=excel_path,
        sheet_name=sheet_name,
        include_source_ref=include_source_ref,
    )
    review_json_path = review_json_path or build_default_review_json_output_path(json_path)
    save_review_artifact(review_artifact, review_json_path=review_json_path)

    app = _ConfirmReviewWindow(
        review_artifact=review_artifact,
        review_json_path=review_json_path,
        excel_path=excel_path,
        sheet_name=sheet_name,
    )
    result = app.run()
    warnings.extend(result["warnings"])
    return {
        "review_json_path": str(result["review_json_path"]),
        "applied": bool(result["applied"]),
    }, warnings


class _ConfirmReviewWindow:
    def __init__(
        self,
        *,
        review_artifact: dict[str, Any],
        review_json_path: Path,
        excel_path: Path,
        sheet_name: str,
    ) -> None:
        try:
            from PyQt6.QtCore import Qt
            from PyQt6.QtWidgets import (
                QApplication,
                QAbstractItemView,
                QCheckBox,
                QFileDialog,
                QHBoxLayout,
                QHeaderView,
                QLabel,
                QLineEdit,
                QMainWindow,
                QMessageBox,
                QPushButton,
                QSplitter,
                QTableWidget,
                QTableWidgetItem,
                QTextEdit,
                QVBoxLayout,
                QWidget,
            )
        except ImportError as exc:
            raise RuntimeError("Native confirm-ui mode requires PyQt6 in the active Python environment") from exc

        self.Qt = Qt
        self.QApplication = QApplication
        self.QAbstractItemView = QAbstractItemView
        self.QCheckBox = QCheckBox
        self.QFileDialog = QFileDialog
        self.QHBoxLayout = QHBoxLayout
        self.QHeaderView = QHeaderView
        self.QLabel = QLabel
        self.QLineEdit = QLineEdit
        self.QMainWindow = QMainWindow
        self.QMessageBox = QMessageBox
        self.QPushButton = QPushButton
        self.QSplitter = QSplitter
        self.QTableWidget = QTableWidget
        self.QTableWidgetItem = QTableWidgetItem
        self.QTextEdit = QTextEdit
        self.QVBoxLayout = QVBoxLayout
        self.QWidget = QWidget

        self.review_artifact = review_artifact
        self.review_json_path = review_json_path
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.session_warnings: list[str] = []
        self.applied = False
        self.dirty = False
        self.current_index: int | None = None
        self.row_controls: list[dict[str, Any]] = []

        existing_app = QApplication.instance()
        self.app = existing_app or QApplication([])

        class _ReviewWindow(QMainWindow):
            def __init__(inner_self, owner: _ConfirmReviewWindow) -> None:
                super().__init__()
                inner_self.owner = owner

            def closeEvent(inner_self, event: Any) -> None:  # noqa: N802 - Qt API name
                inner_self.owner._on_close_event(event)

        self.window = _ReviewWindow(self)
        self.window.setWindowTitle("Review Pending Excel Changes")
        self.window.resize(1200, 760)
        self.window.setStyleSheet(
            "QMainWindow { background: #f4f7fb; }"
            "QLabel#titleLabel { font-size: 24px; font-weight: 700; color: #172033; }"
            "QLabel#summaryLabel { color: #354154; }"
            "QLineEdit, QTextEdit, QTableWidget { background: #ffffff; border: 1px solid #cbd5e1; border-radius: 6px; }"
            "QHeaderView::section { background: #e9eef5; padding: 8px; border: none; border-bottom: 1px solid #cbd5e1; font-weight: 600; }"
            "QTableWidget { gridline-color: #e2e8f0; selection-background-color: #0a74d1; selection-color: #ffffff; alternate-background-color: #f8fafc; }"
            "QPushButton { background: #ffffff; border: 1px solid #cbd5e1; border-radius: 6px; padding: 8px 14px; }"
            "QPushButton:hover { background: #f8fafc; }"
            "QPushButton:pressed { background: #e2e8f0; }"
            "QCheckBox { spacing: 8px; }"
        )

        self._build_layout()
        self._populate_rows()
        self._update_summary()

        rows = self.review_artifact.get("rows")
        if isinstance(rows, list) and rows:
            self._select_row(0)

    def run(self) -> dict[str, Any]:
        self.window.show()
        self.app.exec()
        return {
            "review_json_path": self.review_json_path,
            "applied": self.applied,
            "warnings": list(self.session_warnings),
        }

    def _build_layout(self) -> None:
        container = self.QWidget()
        root_layout = self.QVBoxLayout(container)
        root_layout.setContentsMargins(16, 16, 16, 16)
        root_layout.setSpacing(12)
        self.window.setCentralWidget(container)

        title_label = self.QLabel("Review Pending Excel Changes")
        title_label.setObjectName("titleLabel")
        root_layout.addWidget(title_label)

        self.summary_label = self.QLabel()
        self.summary_label.setObjectName("summaryLabel")
        self.summary_label.setWordWrap(True)
        root_layout.addWidget(self.summary_label)

        path_row = self.QWidget()
        path_layout = self.QHBoxLayout(path_row)
        path_layout.setContentsMargins(0, 0, 0, 0)
        path_layout.setSpacing(8)
        path_layout.addWidget(self.QLabel("Review JSON:"))
        self.path_input = self.QLineEdit(str(self.review_json_path))
        self.path_input.setReadOnly(True)
        path_layout.addWidget(self.path_input, 1)
        root_layout.addWidget(path_row)

        splitter = self.QSplitter(self.Qt.Orientation.Horizontal)
        root_layout.addWidget(splitter, 1)

        list_panel = self.QWidget()
        list_layout = self.QVBoxLayout(list_panel)
        list_layout.setContentsMargins(0, 0, 0, 0)
        list_layout.setSpacing(8)
        splitter.addWidget(list_panel)

        self.row_table = self.QTableWidget(0, 4)
        self.row_table.setHorizontalHeaderLabels(["Change", "Status", "Approved", "Item"])
        self.row_table.setAlternatingRowColors(True)
        self.row_table.setSelectionBehavior(self.QAbstractItemView.SelectionBehavior.SelectRows)
        self.row_table.setSelectionMode(self.QAbstractItemView.SelectionMode.SingleSelection)
        self.row_table.setEditTriggers(self.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.row_table.verticalHeader().setVisible(False)
        table_header = self.row_table.horizontalHeader()
        table_header.setSectionResizeMode(0, self.QHeaderView.ResizeMode.ResizeToContents)
        table_header.setSectionResizeMode(1, self.QHeaderView.ResizeMode.ResizeToContents)
        table_header.setSectionResizeMode(2, self.QHeaderView.ResizeMode.ResizeToContents)
        table_header.setSectionResizeMode(3, self.QHeaderView.ResizeMode.Stretch)
        self.row_table.itemSelectionChanged.connect(self._on_row_selection_changed)
        list_layout.addWidget(self.row_table)

        detail_panel = self.QWidget()
        detail_layout = self.QVBoxLayout(detail_panel)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        detail_layout.setSpacing(10)
        splitter.addWidget(detail_panel)
        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 6)

        self.selection_label = self.QLabel("Select a change row to review details.")
        self.selection_label.setStyleSheet("font-weight: 700; font-size: 16px; color: #172033;")
        detail_layout.addWidget(self.selection_label)

        self.approve_check = self.QCheckBox("Approve selected row")
        self.approve_check.stateChanged.connect(self._toggle_selected_approval)
        detail_layout.addWidget(self.approve_check)

        self.details_text = self.QTextEdit()
        self.details_text.setReadOnly(True)
        self.details_text.setLineWrapMode(self.QTextEdit.LineWrapMode.NoWrap)
        detail_layout.addWidget(self.details_text, 1)

        button_row = self.QWidget()
        button_layout = self.QHBoxLayout(button_row)
        button_layout.setContentsMargins(0, 0, 0, 0)
        button_layout.setSpacing(8)

        save_button = self.QPushButton("Save Review JSON")
        save_button.clicked.connect(self._save_review)
        button_layout.addWidget(save_button)

        save_as_button = self.QPushButton("Save As")
        save_as_button.clicked.connect(self._save_review_as)
        button_layout.addWidget(save_as_button)

        apply_button = self.QPushButton("Apply Approved")
        apply_button.clicked.connect(self._apply_approved)
        button_layout.addWidget(apply_button)

        button_layout.addStretch(1)

        close_button = self.QPushButton("Close")
        close_button.clicked.connect(self.window.close)
        button_layout.addWidget(close_button)

        root_layout.addWidget(button_row)

    def _rows(self) -> list[dict[str, Any]]:
        rows = self.review_artifact.get("rows")
        return rows if isinstance(rows, list) else []

    def _populate_rows(self) -> None:
        self.row_table.setRowCount(0)
        self.row_controls = []

        for index, row in enumerate(self._rows()):
            target = row.get("target") if isinstance(row.get("target"), dict) else {}
            self.row_table.insertRow(index)
            self.row_table.setItem(index, 0, self._create_table_item(str(row.get("changeId") or f"change-{index + 1}")))
            status_item = self._create_table_item(str(row.get("status") or "invalid"))
            self.row_table.setItem(index, 1, status_item)
            self.row_table.setItem(index, 2, self._create_table_item(""))
            self.row_table.setItem(index, 3, self._create_table_item(str(target.get("itemId") or "")))

            approved_checkbox = self.QCheckBox()
            approved_checkbox.setChecked(bool(row.get("approved")))
            approved_checkbox.setEnabled(str(row.get("status") or "invalid") == "ready")
            approved_checkbox.stateChanged.connect(
                lambda state, row_index=index: self._on_row_checkbox_toggled(row_index, state)
            )
            checkbox_container = self.QWidget()
            checkbox_layout = self.QHBoxLayout(checkbox_container)
            checkbox_layout.setContentsMargins(0, 0, 0, 0)
            checkbox_layout.addWidget(approved_checkbox, alignment=self.Qt.AlignmentFlag.AlignCenter)
            self.row_table.setCellWidget(index, 2, checkbox_container)

            self.row_controls.append(
                {
                    "status_item": status_item,
                    "approved_checkbox": approved_checkbox,
                }
            )

    def _create_table_item(self, text: str) -> Any:
        item = self.QTableWidgetItem(text)
        item.setFlags(item.flags() & ~self.Qt.ItemFlag.ItemIsEditable)
        return item

    def _refresh_row_item(self, index: int) -> None:
        row = self._rows()[index]
        controls = self.row_controls[index]
        controls["status_item"].setText(str(row.get("status") or "invalid"))
        checkbox = controls["approved_checkbox"]
        checkbox.blockSignals(True)
        checkbox.setChecked(bool(row.get("approved")))
        checkbox.setEnabled(str(row.get("status") or "invalid") == "ready")
        checkbox.blockSignals(False)

    def _update_summary(self, extra_message: str = "") -> None:
        approved_count = sum(1 for row in self._rows() if row.get("approved"))
        dirty_marker = " Unsaved changes." if self.dirty else ""
        base = (
            f"{len(self._rows())} pending row(s), {approved_count} approved. "
            f"Workbook: {self.excel_path}. Review file: {self.review_json_path}."
        )
        self.summary_label.setText(base + dirty_marker + (f" {extra_message}" if extra_message else ""))
        self.path_input.setText(str(self.review_json_path))

    def _set_details_text(self, content: str) -> None:
        self.details_text.setPlainText(content)

    def _on_row_selection_changed(self) -> None:
        selection_model = self.row_table.selectionModel()
        if selection_model is None:
            self._select_row(None)
            return
        selected_rows = selection_model.selectedRows()
        if not selected_rows:
            self._select_row(None)
            return
        self._select_row(selected_rows[0].row())

    def _select_row(self, index: int | None) -> None:
        if index is None:
            self.current_index = None
            self.selection_label.setText("Select a change row to review details.")
            self.approve_check.blockSignals(True)
            self.approve_check.setChecked(False)
            self.approve_check.blockSignals(False)
            self.approve_check.setEnabled(False)
            self._set_details_text("")
            return

        if self.current_index != index:
            self.row_table.blockSignals(True)
            self.row_table.selectRow(index)
            self.row_table.blockSignals(False)

        row = self._rows()[index]
        target = row.get("target") if isinstance(row.get("target"), dict) else {}

        self.current_index = index
        self.selection_label.setText(f"{row.get('changeId') or f'row-{index + 1}'} ({row.get('action') or 'unknown'})")

        self.approve_check.blockSignals(True)
        self.approve_check.setChecked(bool(row.get("approved")))
        self.approve_check.blockSignals(False)
        self.approve_check.setEnabled(str(row.get("status") or "invalid") == "ready")

        lines = [
            f"Status: {row.get('status') or 'invalid'}",
            f"Item: {target.get('itemId') or ''}",
            f"Excel row: {target.get('excelRow') or ''}",
        ]
        if target.get("sourceRef"):
            lines.append(f"SourceRef: {json.dumps(target['sourceRef'], ensure_ascii=True)}")
        lines.append("")

        for field in row.get("fields") or []:
            if not isinstance(field, dict):
                continue
            lines.extend(
                [
                    f"Field: {field.get('name') or ''}",
                    f"  Excel: {_render_value(field.get('excelCurrent'))}",
                    f"  Baseline: {_render_value(field.get('jsonBaseline'))}",
                    f"  Proposed: {_render_value(field.get('proposed'))}",
                    f"  Conflict: {'yes' if field.get('conflict') else 'no'}",
                    "",
                ]
            )

        self._set_details_text("\n".join(lines).rstrip())

    def _toggle_selected_approval(self, state: int) -> None:
        if self.current_index is None:
            return
        self._set_row_approval(self.current_index, bool(state))

    def _set_row_approval(self, index: int, approved: bool) -> None:
        row = self._rows()[index]
        if str(row.get("status") or "invalid") != "ready":
            self.approve_check.blockSignals(True)
            self.approve_check.setChecked(False)
            self.approve_check.blockSignals(False)
            checkbox = self.row_controls[index]["approved_checkbox"]
            checkbox.blockSignals(True)
            checkbox.setChecked(False)
            checkbox.blockSignals(False)
            return

        row["approved"] = approved
        self.dirty = True
        self._refresh_row_item(index)
        self._update_summary()

    def _on_row_checkbox_toggled(self, index: int, state: int) -> None:
        self._select_row(index)
        self.approve_check.blockSignals(True)
        self.approve_check.setChecked(bool(state))
        self.approve_check.blockSignals(False)
        self._set_row_approval(index, bool(state))

    def _save_review(self) -> bool:
        save_review_artifact(self.review_artifact, review_json_path=self.review_json_path)
        self.dirty = False
        self._update_summary(f"Saved review JSON to {self.review_json_path.name}.")
        return True

    def _save_review_as(self) -> bool:
        filename, _selected_filter = self.QFileDialog.getSaveFileName(
            self.window,
            "Save approved review JSON",
            str(self.review_json_path),
            "JSON files (*.json)",
        )
        if not filename:
            return False

        self.review_json_path = Path(filename)
        return self._save_review()

    def _apply_approved(self) -> None:
        if self.dirty:
            self._save_review()

        warnings = apply_review_artifact(
            self.review_artifact,
            excel_path=self.excel_path,
            sheet_name=self.sheet_name,
        )
        self.session_warnings.extend(warnings)
        self.applied = True

        if warnings:
            self.QMessageBox.warning(
                self.window,
                "Approved Changes Applied",
                "Approved rows were applied with warnings:\n\n" + "\n".join(warnings),
            )
        else:
            self.QMessageBox.information(
                self.window,
                "Approved Changes Applied",
                f"Approved rows were applied to {self.excel_path.name}.",
            )
        self._update_summary(f"Applied approved rows to {self.excel_path.name}.")

    def _on_close_event(self, event: Any) -> None:
        if self.dirty:
            answer = self.QMessageBox.question(
                self.window,
                "Save review JSON",
                "Save current review approvals before closing?",
                self.QMessageBox.StandardButton.Yes
                | self.QMessageBox.StandardButton.No
                | self.QMessageBox.StandardButton.Cancel,
                self.QMessageBox.StandardButton.Yes,
            )
            if answer == self.QMessageBox.StandardButton.Cancel:
                event.ignore()
                return
            if answer == self.QMessageBox.StandardButton.Yes and not self._save_review():
                event.ignore()
                return
        event.accept()
        self.window.hide()
        self.app.quit()


def _render_value(value: Any) -> str:
    if value is None or value == "":
        return "<empty>"
    if isinstance(value, str):
        return value
    return json.dumps(value, ensure_ascii=True)
