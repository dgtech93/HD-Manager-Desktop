"""Anteprima risultato confronto: colonne (ordine, inclusione), tabella in sola lettura."""

from __future__ import annotations

from typing import Any

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QDropEvent
from PyQt6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.excel_format import ColumnFormatSpec, format_cell_as_string, formats_or_auto

_PREVIEW_MAX_ROWS = 8000


class _ColList(QListWidget):
    """Lista colonne con notifica dopo il trascinamento."""

    rows_reordered = pyqtSignal()

    def dropEvent(self, event: QDropEvent | None) -> None:
        super().dropEvent(event)
        self.rows_reordered.emit()


class ExcelResultPreviewDialog(QDialog):
    def __init__(
        self,
        title: str,
        headers: list[str],
        rows: list[list[Any]],
        parent: QWidget | None = None,
        column_formats: list[ColumnFormatSpec] | None = None,
    ) -> None:
        super().__init__(parent)
        self._headers_in = list(headers)
        self._rows_in = [list(r) for r in rows]
        # Allinea sempre alla griglia colonne sorgente (stesso criterio di concatenazione / export).
        self._formats_in = formats_or_auto(len(self._headers_in), column_formats)
        self._out_headers: list[str] | None = None
        self._out_rows: list[list[Any]] | None = None
        self._out_formats: list[ColumnFormatSpec] | None = None

        self.setWindowTitle(title)
        self.resize(1000, 600)

        hint = QLabel(
            f"Anteprima (max {_PREVIEW_MAX_ROWS} righe). Seleziona e ordina le colonne, poi conferma."
        )
        hint.setWordWrap(True)

        self._col_list = _ColList()
        self._col_list.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        self._col_list.setDefaultDropAction(Qt.DropAction.MoveAction)

        self._table = QTableWidget()
        self._table.setAlternatingRowColors(True)
        self._table.horizontalHeader().setStretchLastSection(True)

        split = QSplitter(Qt.Orientation.Horizontal)
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.addWidget(QLabel("Colonne"))
        ll.addWidget(self._col_list, 1)
        split.addWidget(left)
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.addWidget(QLabel("Anteprima dati"))
        rl.addWidget(self._table, 1)
        split.addWidget(right)

        root = QVBoxLayout(self)
        root.addWidget(hint)
        root.addWidget(split, 1)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._on_ok)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

        self._rebuild_lists()
        self._col_list.itemChanged.connect(lambda _: self._refresh_table())
        self._col_list.rows_reordered.connect(self._refresh_table)

    def _rebuild_lists(self) -> None:
        self._col_list.clear()
        for j, h in enumerate(self._headers_in):
            it = QListWidgetItem(h)
            it.setData(Qt.ItemDataRole.UserRole, j)
            it.setFlags(
                it.flags()
                | Qt.ItemFlag.ItemIsUserCheckable
                | Qt.ItemFlag.ItemIsDragEnabled
            )
            it.setCheckState(Qt.CheckState.Checked)
            self._col_list.addItem(it)
        self._refresh_table()

    def _refresh_table(self) -> None:
        n = len(self._headers_in)
        fmts = formats_or_auto(n, self._formats_in)
        order: list[int] = []
        labels: list[str] = []
        for i in range(self._col_list.count()):
            it = self._col_list.item(i)
            if it is None or it.checkState() != Qt.CheckState.Checked:
                continue
            j = int(it.data(Qt.ItemDataRole.UserRole))
            if 0 <= j < n:
                order.append(j)
                labels.append(self._headers_in[j])
        self._table.setColumnCount(len(labels))
        self._table.setHorizontalHeaderLabels(labels)
        show_rows = min(len(self._rows_in), _PREVIEW_MAX_ROWS)
        self._table.setRowCount(show_rows)
        self._table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.ResizeToContents
        )
        for r in range(show_rows):
            row_vals = self._rows_in[r]
            for ci, j in enumerate(order):
                val = row_vals[j] if j < len(row_vals) else None
                txt = format_cell_as_string(val, fmts[j])
                cell = QTableWidgetItem(txt)
                cell.setFlags(cell.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self._table.setItem(r, ci, cell)

    def _on_ok(self) -> None:
        n = len(self._headers_in)
        order: list[int] = []
        export_headers: list[str] = []
        for i in range(self._col_list.count()):
            it = self._col_list.item(i)
            if it is None or it.checkState() != Qt.CheckState.Checked:
                continue
            j = int(it.data(Qt.ItemDataRole.UserRole))
            if 0 <= j < n:
                order.append(j)
                export_headers.append(self._headers_in[j])
        if not order:
            QMessageBox.warning(self, "Anteprima", "Seleziona almeno una colonna.")
            return
        out_rows: list[list[Any]] = []
        for row_vals in self._rows_in:
            out_rows.append([row_vals[j] if j < len(row_vals) else None for j in order])
        self._out_headers = export_headers
        self._out_rows = out_rows
        self._out_formats = [self._formats_in[j] for j in order]
        self.accept()

    def result_table(
        self,
    ) -> tuple[list[str], list[list[Any]], list[ColumnFormatSpec] | None] | None:
        if self._out_headers is None:
            return None
        return self._out_headers, self._out_rows or [], self._out_formats
