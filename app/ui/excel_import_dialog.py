"""Dialog di precaricamento Excel: colonne, ordine, righe."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QScrollArea,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.excel_format import (
    ColumnFormatSpec,
    format_cell_as_string,
    infer_column_format_spec,
    KIND_DATE,
    KIND_DATETIME,
    KIND_DECIMAL,
    KIND_INTEGER,
    KIND_STRING,
    KIND_AUTO,
)
from app.excel_io import DEFAULT_MAX_ROWS, list_sheet_names, read_sheet_rows


@dataclass
class ExcelImportOutcome:
    """Risultato confermato dall'utente."""

    source_path: str
    sheet_name: str
    headers: list[str]
    rows: list[list[Any]]
    column_formats: list[ColumnFormatSpec] | None = None
    # Per ripristino in «Modifica importazione» (indici colonna sorgente e righe dati)
    header_row_used: bool = True
    export_column_order: list[int] = field(default_factory=list)
    export_source_row_indices: list[int] = field(default_factory=list)


@dataclass
class StoredImport:
    """Import registrato (per confronti e riferimenti incrociati)."""

    id: int
    outcome: ExcelImportOutcome

    def display_label(self) -> str:
        return f"{Path(self.outcome.source_path).name} — «{self.outcome.sheet_name}»"


@dataclass
class StoredCompareResult:
    """Tabella confermata dalla vista Confronto (riusabile in Concatenazione, ecc.)."""

    id: int
    title: str
    headers: list[str]
    rows: list[list[Any]]
    column_formats: list[ColumnFormatSpec] | None = None


def _normalize_matrix(rows: list[list[Any]]) -> list[list[Any]]:
    if not rows:
        return []
    n = max((len(r) for r in rows), default=0)
    out: list[list[Any]] = []
    for r in rows:
        if len(r) < n:
            out.append(list(r) + [None] * (n - len(r)))
        else:
            out.append(list(r))
    return out


_KIND_OPTIONS = [
    ("Automatico", KIND_AUTO),
    ("Testo", KIND_STRING),
    ("Intero", KIND_INTEGER),
    ("Decimale", KIND_DECIMAL),
    ("Solo data", KIND_DATE),
    ("Data e ora", KIND_DATETIME),
]

_DATE_PRESETS = [
    ("gg/MM/yyyy (IT)", "%d/%m/%Y"),
    ("dd/MM/yy", "%d/%m/%y"),
    ("yyyy/MM/dd", "%Y/%m/%d"),
    ("yyyy-MM-dd (ISO)", "%Y-%m-%d"),
    ("dd-MM-yyyy", "%d-%m-%Y"),
]
_DATETIME_PRESETS = _DATE_PRESETS + [
    ("gg/MM/yyyy HH:mm", "%d/%m/%Y %H:%M"),
    ("ISO 8601", "%Y-%m-%d %H:%M:%S"),
]


class ExcelImportDialog(QDialog):
    """Anteprima con scelta colonne (ordine drag), righe incluse, foglio."""

    def __init__(
        self,
        file_path: str,
        parent: QWidget | None = None,
        *,
        initial_outcome: ExcelImportOutcome | None = None,
    ) -> None:
        super().__init__(parent)
        self._path = str(Path(file_path).resolve())
        self._initial_outcome = initial_outcome
        self._sheet_names: list[str] = []
        self._raw_matrix: list[list[Any]] = []
        self._total_rows_hint: int | None = None
        self._outcome: ExcelImportOutcome | None = None
        self._import_n = 0
        self._import_headers: list[str] = []
        self._import_data_rows: list[list[Any]] = []

        title = "Modifica importazione" if initial_outcome else "Importazione"
        self.setWindowTitle(f"{title} — {Path(self._path).name}")
        self.resize(1100, 780)

        self._sheet_combo = QComboBox()
        self._hdr_check = QCheckBox("Prima riga contiene le intestazioni di colonna")
        self._hdr_check.setChecked(True)
        self._col_list = QListWidget()
        self._col_list.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        self._col_list.setDefaultDropAction(Qt.DropAction.MoveAction)
        self._col_list.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self._table = QTableWidget()
        self._table.setAlternatingRowColors(True)
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.verticalHeader().setVisible(True)
        self._hint_lbl = QLabel()

        root = QVBoxLayout(self)
        top = QHBoxLayout()
        top.addWidget(QLabel("Foglio:"))
        top.addWidget(self._sheet_combo, 1)
        root.addLayout(top)
        root.addWidget(self._hdr_check)

        self._fmt_table = QTableWidget(0, 3)
        self._fmt_table.setHorizontalHeaderLabels(["Colonna", "Tipo", "Formato / opzioni"])
        self._fmt_table.horizontalHeader().setStretchLastSection(True)
        self._fmt_table.verticalHeader().setVisible(False)
        self._fmt_table.setMinimumHeight(120)

        fmt_hint = QLabel(
            "Il tipo di ogni colonna viene proposto in automatico dal campione (testo, intero, decimale, date). "
            "Decimale: cifre decimali. Data/ora: strftime. L’anteprima si aggiorna al cambio tipo."
        )
        fmt_hint.setWordWrap(True)
        fmt_hint.setStyleSheet("color: #475569; font-size: 11px;")

        main_split = QSplitter(Qt.Orientation.Horizontal)

        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QScrollArea.Shape.StyledPanel)
        left_scroll.setMinimumWidth(200)
        left_scroll.setMaximumWidth(300)
        left_wrap = QWidget()
        left_lay = QVBoxLayout(left_wrap)
        left_lay.setContentsMargins(4, 4, 4, 4)
        left_title = QLabel("Colonne")
        left_title.setStyleSheet("font-weight: 600;")
        left_lay.addWidget(left_title)
        left_sub = QLabel("Spunta e trascina per ordine nell’export.")
        left_sub.setWordWrap(True)
        left_sub.setStyleSheet("color: #475569; font-size: 11px;")
        left_lay.addWidget(left_sub)
        left_lay.addWidget(self._col_list, 1)
        left_scroll.setWidget(left_wrap)
        main_split.addWidget(left_scroll)

        center_scroll = QScrollArea()
        center_scroll.setWidgetResizable(True)
        center_scroll.setFrameShape(QScrollArea.Shape.StyledPanel)
        center_wrap = QWidget()
        center_lay = QVBoxLayout(center_wrap)
        center_lay.setContentsMargins(4, 4, 4, 4)
        center_title = QLabel("Anteprima foglio (solo colonne selezionate)")
        center_title.setStyleSheet("font-weight: 600;")
        center_lay.addWidget(center_title)
        center_sub = QLabel(
            "Prima colonna: spunta per includere la riga nell’export. "
            "L’intestazione «Seleziona tutte le righe» seleziona o deseleziona tutte le righe."
        )
        center_sub.setWordWrap(True)
        center_sub.setStyleSheet("color: #475569; font-size: 11px;")
        center_lay.addWidget(center_sub)
        center_lay.addWidget(self._table, 1)
        center_scroll.setWidget(center_wrap)
        main_split.addWidget(center_scroll)

        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)
        right_scroll.setFrameShape(QScrollArea.Shape.StyledPanel)
        right_scroll.setMinimumWidth(260)
        right_scroll.setMaximumWidth(440)
        right_wrap = QWidget()
        right_lay = QVBoxLayout(right_wrap)
        right_lay.setContentsMargins(4, 4, 4, 4)
        right_title = QLabel("Formato colonne")
        right_title.setStyleSheet("font-weight: 600;")
        right_lay.addWidget(right_title)
        right_lay.addWidget(self._fmt_table, 1)
        right_lay.addWidget(fmt_hint)
        right_scroll.setWidget(right_wrap)
        main_split.addWidget(right_scroll)

        main_split.setStretchFactor(0, 0)
        main_split.setStretchFactor(1, 1)
        main_split.setStretchFactor(2, 0)
        main_split.setSizes([240, 520, 300])
        root.addWidget(main_split, 1)

        root.addWidget(self._hint_lbl)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._on_ok)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

        self._hdr_check.toggled.connect(self._rebuild_from_matrix)
        self._sheet_combo.currentTextChanged.connect(self._on_sheet_changed)
        self._col_list.itemChanged.connect(self._on_col_list_item_changed)
        self._col_list.model().rowsMoved.connect(self._on_col_list_order_changed)
        self._table.horizontalHeader().sectionClicked.connect(self._on_include_header_clicked)
        self._table.itemChanged.connect(self._on_preview_include_item_changed)

        try:
            self._sheet_names = list_sheet_names(self._path)
        except Exception as exc:
            QMessageBox.critical(self, "Excel", str(exc))
            self._sheet_names = []

        if not self._sheet_names:
            self._hint_lbl.setText("Impossibile leggere i fogli.")
        else:
            for name in self._sheet_names:
                self._sheet_combo.addItem(name)
            if initial_outcome is not None:
                self._hdr_check.blockSignals(True)
                self._hdr_check.setChecked(initial_outcome.header_row_used)
                self._hdr_check.blockSignals(False)
                self._sheet_combo.blockSignals(True)
                idx = self._sheet_combo.findText(initial_outcome.sheet_name)
                if idx >= 0:
                    self._sheet_combo.setCurrentIndex(idx)
                self._sheet_combo.blockSignals(False)
            self._load_current_sheet()

    def _on_sheet_changed(self, _text: str) -> None:
        self._initial_outcome = None
        self._load_current_sheet()

    def _load_current_sheet(self) -> None:
        if not self._sheet_names:
            return
        name = self._sheet_combo.currentText()
        try:
            self._raw_matrix, self._total_rows_hint = read_sheet_rows(
                self._path, name, max_rows=DEFAULT_MAX_ROWS
            )
        except Exception as exc:
            QMessageBox.warning(self, "Excel", str(exc))
            self._raw_matrix = []
            self._total_rows_hint = None
        self._raw_matrix = _normalize_matrix(self._raw_matrix)
        extra = ""
        if self._total_rows_hint is not None and self._total_rows_hint > len(self._raw_matrix):
            extra = (
                f" (file con molte righe: in anteprima ne sono caricate al massimo {len(self._raw_matrix)})"
            )
        self._hint_lbl.setText(
            f"Righe lette in anteprima: {len(self._raw_matrix)}{extra}"
        )
        self._rebuild_from_matrix()
        init = self._initial_outcome
        if init is not None:
            if Path(init.source_path).resolve() == Path(self._path).resolve():
                if self._sheet_combo.currentText() == init.sheet_name and self._import_n > 0:
                    self._apply_initial_outcome(init)
            self._initial_outcome = None

    def _infer_order_from_export_headers(self, export_headers: list[str]) -> list[int]:
        """Ripristina indici colonna sorgente dai nomi esportati (import senza metadati o file cambiato)."""
        base = self._import_headers
        n = len(base)
        used: set[int] = set()
        out: list[int] = []
        for h in export_headers:
            for j in range(n):
                if j in used:
                    continue
                if base[j] == h:
                    out.append(j)
                    used.add(j)
                    break
        return out

    def _rebuild_col_list_from_order(self, order: list[int]) -> None:
        """Ordine lista colonne: prima quelle esportate (spuntate), poi le altre."""
        n = self._import_n
        wanted = set(order)
        self._col_list.blockSignals(True)
        self._col_list.clear()
        for j in order:
            if j < 0 or j >= n:
                continue
            lab = self._import_headers[j] if j < len(self._import_headers) else f"Col {j + 1}"
            it = QListWidgetItem(lab)
            it.setData(Qt.ItemDataRole.UserRole, j)
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsDragEnabled)
            it.setCheckState(Qt.CheckState.Checked)
            self._col_list.addItem(it)
        for j in range(n):
            if j in wanted:
                continue
            lab = self._import_headers[j] if j < len(self._import_headers) else f"Col {j + 1}"
            it = QListWidgetItem(lab)
            it.setData(Qt.ItemDataRole.UserRole, j)
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsDragEnabled)
            it.setCheckState(Qt.CheckState.Unchecked)
            self._col_list.addItem(it)
        self._col_list.blockSignals(False)

    def _apply_initial_outcome(self, init: ExcelImportOutcome) -> None:
        """Ripristina colonne, righe e formati da un import precedente."""
        order = [j for j in init.export_column_order if 0 <= j < self._import_n]
        if len(order) != len(init.headers):
            inferred = self._infer_order_from_export_headers(init.headers)
            if inferred:
                order = inferred
        if not order:
            order = list(range(min(len(init.headers), self._import_n)))
        order = [j for j in order if 0 <= j < self._import_n]

        self._rebuild_col_list_from_order(order)

        fmts = init.column_formats or []
        for k, src_j in enumerate(order):
            if k < len(fmts) and 0 <= src_j < self._fmt_table.rowCount():
                self._apply_spec_to_fmt_row(src_j, fmts[k])

        self._sync_format_row_visibility()
        self._refresh_preview_columns()

        dr = len(self._import_data_rows)
        indices = set(init.export_source_row_indices) if init.export_source_row_indices else set()
        if not indices:
            if len(init.rows) == dr:
                indices = set(range(dr))
            else:
                indices = set(range(min(len(init.rows), dr)))

        self._table.blockSignals(True)
        for r in range(self._table.rowCount()):
            it = self._table.item(r, 0)
            if it is not None:
                it.setCheckState(
                    Qt.CheckState.Checked if r in indices else Qt.CheckState.Unchecked
                )
        self._sync_include_header_checkbox()
        self._table.blockSignals(False)
        self._apply_preview_formats()

    def _rebuild_from_matrix(self) -> None:
        self._col_list.clear()
        self._table.clear()
        self._fmt_table.setRowCount(0)
        if not self._raw_matrix:
            self._import_n = 0
            return
        use_hdr = self._hdr_check.isChecked()
        if use_hdr:
            header_row = self._raw_matrix[0]
            data_rows = self._raw_matrix[1:]
            n = len(header_row)
            headers = [
                str(header_row[j]).strip() if header_row[j] is not None else f"Col {j + 1}"
                for j in range(n)
            ]
        else:
            n = max(len(r) for r in self._raw_matrix)
            data_rows = self._raw_matrix
            headers = [f"Col {j + 1}" for j in range(n)]

        self._col_list.blockSignals(True)
        for j in range(n):
            lab = headers[j] if j < len(headers) else f"Col {j + 1}"
            it = QListWidgetItem(lab)
            it.setData(Qt.ItemDataRole.UserRole, j)
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsDragEnabled)
            it.setCheckState(Qt.CheckState.Checked)
            self._col_list.addItem(it)
        self._col_list.blockSignals(False)

        self._import_n = n
        self._import_headers = headers[:]
        self._import_data_rows = data_rows
        self._rebuild_format_table(headers, n)
        self._sync_format_row_visibility()
        self._refresh_preview_columns()

    def _visible_column_order(self) -> list[int]:
        """Indici colonna sorgente nell’ordine della lista (solo spuntate)."""
        out: list[int] = []
        for i in range(self._col_list.count()):
            it = self._col_list.item(i)
            if it is None or it.checkState() != Qt.CheckState.Checked:
                continue
            j = int(it.data(Qt.ItemDataRole.UserRole))
            out.append(j)
        return out

    def _is_source_col_checked(self, j: int) -> bool:
        for i in range(self._col_list.count()):
            it = self._col_list.item(i)
            if it is None:
                continue
            if int(it.data(Qt.ItemDataRole.UserRole)) == j:
                return it.checkState() == Qt.CheckState.Checked
        return False

    def _on_col_list_item_changed(self, _item: QListWidgetItem) -> None:
        self._sync_format_row_visibility()
        self._refresh_preview_columns()

    def _on_col_list_order_changed(self, *_args: object) -> None:
        self._refresh_preview_columns()

    def _sync_format_row_visibility(self) -> None:
        for j in range(self._fmt_table.rowCount()):
            self._fmt_table.setRowHidden(j, not self._is_source_col_checked(j))

    def _apply_include_header_item(self) -> None:
        """Intestazione colonna 0: testo e checkbox «seleziona tutte le righe»."""
        h0 = self._table.horizontalHeaderItem(0)
        if h0 is None:
            h0 = QTableWidgetItem()
            self._table.setHorizontalHeaderItem(0, h0)
        h0.setText("Seleziona tutte le righe")
        h0.setToolTip(
            "Seleziona tutte le righe: clic per attivare o disattivare la spunta su tutte le righe "
            "dell’anteprima."
        )
        h0.setFlags(
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsUserCheckable
            | Qt.ItemFlag.ItemIsSelectable
        )

    def _all_include_rows_checked(self) -> bool:
        n = self._table.rowCount()
        if n == 0:
            return False
        for r in range(n):
            it = self._table.item(r, 0)
            if it is None or it.checkState() != Qt.CheckState.Checked:
                return False
        return True

    def _sync_include_header_checkbox(self) -> None:
        h0 = self._table.horizontalHeaderItem(0)
        if h0 is None:
            return
        h0.setCheckState(
            Qt.CheckState.Checked if self._all_include_rows_checked() else Qt.CheckState.Unchecked
        )

    def _on_include_header_clicked(self, logical_index: int) -> None:
        """Click sull’intestazione colonna 0: seleziona / deseleziona tutte le righe."""
        if logical_index != 0:
            return
        n = self._table.rowCount()
        if n == 0:
            return
        all_on = self._all_include_rows_checked()
        target = Qt.CheckState.Unchecked if all_on else Qt.CheckState.Checked
        self._table.blockSignals(True)
        try:
            for r in range(n):
                it = self._table.item(r, 0)
                if it is not None:
                    it.setCheckState(target)
            h0 = self._table.horizontalHeaderItem(0)
            if h0 is not None:
                h0.setCheckState(target)
        finally:
            self._table.blockSignals(False)

    def _on_preview_include_item_changed(self, item: QTableWidgetItem) -> None:
        if item.column() != 0:
            return
        self._table.blockSignals(True)
        try:
            self._sync_include_header_checkbox()
        finally:
            self._table.blockSignals(False)

    def _refresh_preview_columns(self) -> None:
        """Anteprima: colonne dati = solo quelle spuntate, ordine = lista sinistra."""
        if self._import_n <= 0:
            return
        self._table.blockSignals(True)
        try:
            data_rows = self._import_data_rows
            order = self._visible_column_order()

            n_prev = self._table.rowCount()
            prev_include: list[bool] = []
            for r in range(n_prev):
                inc = self._table.item(r, 0)
                prev_include.append(
                    inc is not None and inc.checkState() == Qt.CheckState.Checked
                )

            n_data = len(order)
            self._table.setColumnCount(1 + max(0, n_data))
            hdrs = ["Seleziona tutte le righe"] + [self._import_headers[j] for j in order]
            self._table.setHorizontalHeaderLabels(hdrs)
            self._apply_include_header_item()
            self._table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)

            self._table.setRowCount(len(data_rows))
            for r, row_vals in enumerate(data_rows):
                cb = QTableWidgetItem()
                cb.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
                if r < len(prev_include):
                    cb.setCheckState(
                        Qt.CheckState.Checked if prev_include[r] else Qt.CheckState.Unchecked
                    )
                else:
                    cb.setCheckState(Qt.CheckState.Checked)
                self._table.setItem(r, 0, cb)
                for k, j in enumerate(order):
                    val = row_vals[j] if j < len(row_vals) else None
                    spec = self._spec_from_row(j)
                    txt = format_cell_as_string(val, spec)
                    cell = QTableWidgetItem(txt)
                    cell.setFlags(cell.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    self._table.setItem(r, k + 1, cell)

            self._sync_include_header_checkbox()
        finally:
            self._table.blockSignals(False)

    def _rebuild_format_table(self, headers: list[str], n: int) -> None:
        self._fmt_table.blockSignals(True)
        self._fmt_table.setRowCount(n)
        for j in range(n):
            lab = headers[j] if j < len(headers) else f"Col {j + 1}"
            it = QTableWidgetItem(lab)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self._fmt_table.setItem(j, 0, it)
            cb_k = QComboBox()
            for tlab, key in _KIND_OPTIONS:
                cb_k.addItem(tlab, key)
            self._fmt_table.setCellWidget(j, 1, cb_k)
            cb_o = QComboBox()
            self._fmt_table.setCellWidget(j, 2, cb_o)
            self._sync_format_options(j)
            self._connect_fmt_row(j)
        self._fmt_table.blockSignals(False)
        data_rows = self._import_data_rows
        if self._initial_outcome is None:
            for j in range(n):
                col_vals = [row[j] if j < len(row) else None for row in data_rows]
                self._apply_spec_to_fmt_row(j, infer_column_format_spec(col_vals))
        self._apply_preview_formats()

    def _apply_spec_to_fmt_row(self, row: int, spec: ColumnFormatSpec) -> None:
        kcb = self._fmt_table.cellWidget(row, 1)
        ocb = self._fmt_table.cellWidget(row, 2)
        if not isinstance(kcb, QComboBox):
            return
        kcb.blockSignals(True)
        for i in range(kcb.count()):
            if kcb.itemData(i) == spec.kind:
                kcb.setCurrentIndex(i)
                break
        kcb.blockSignals(False)
        self._sync_format_options(row)
        if not isinstance(ocb, QComboBox):
            return
        ocb.blockSignals(True)
        kind = spec.kind or KIND_AUTO
        idx = 0
        if kind == KIND_DECIMAL:
            want = int(spec.decimal_places)
            for i in range(ocb.count()):
                d = ocb.itemData(i)
                if d is not None and int(d) == want:
                    idx = i
                    break
        elif kind in (KIND_DATE, KIND_DATETIME):
            pat = spec.date_pattern or "%d/%m/%Y"
            for i in range(ocb.count()):
                if ocb.itemData(i) == pat:
                    idx = i
                    break
        ocb.setCurrentIndex(idx)
        ocb.blockSignals(False)

    def _sync_format_options(self, row: int) -> None:
        kcb = self._fmt_table.cellWidget(row, 1)
        ocb = self._fmt_table.cellWidget(row, 2)
        if not isinstance(kcb, QComboBox) or not isinstance(ocb, QComboBox):
            return
        kind = str(kcb.currentData() or KIND_AUTO)
        ocb.blockSignals(True)
        ocb.clear()
        if kind in (KIND_AUTO, KIND_STRING):
            ocb.addItem("—", "")
        elif kind == KIND_INTEGER:
            ocb.addItem("Arrotondato all’intero", "round")
        elif kind == KIND_DECIMAL:
            for d in range(0, 9):
                ocb.addItem(f"{d} decimali", d)
        elif kind == KIND_DATE:
            for tlab, pat in _DATE_PRESETS:
                ocb.addItem(tlab, pat)
        elif kind == KIND_DATETIME:
            for tlab, pat in _DATETIME_PRESETS:
                ocb.addItem(tlab, pat)
        ocb.blockSignals(False)

    def _connect_fmt_row(self, row: int) -> None:
        kcb = self._fmt_table.cellWidget(row, 1)
        ocb = self._fmt_table.cellWidget(row, 2)
        if isinstance(kcb, QComboBox):
            kcb.currentIndexChanged.connect(lambda _i, r=row: self._on_kind_changed(r))
        if isinstance(ocb, QComboBox):
            ocb.currentIndexChanged.connect(lambda _i: self._apply_preview_formats())

    def _on_kind_changed(self, row: int) -> None:
        self._sync_format_options(row)
        self._apply_preview_formats()

    def _spec_from_row(self, row: int) -> ColumnFormatSpec:
        kcb = self._fmt_table.cellWidget(row, 1)
        ocb = self._fmt_table.cellWidget(row, 2)
        if not isinstance(kcb, QComboBox) or not isinstance(ocb, QComboBox):
            return ColumnFormatSpec()
        kind = str(kcb.currentData() or KIND_AUTO)
        spec = ColumnFormatSpec(kind=kind)
        if kind == KIND_DECIMAL:
            raw = ocb.currentData()
            spec.decimal_places = int(raw) if raw is not None else 2
        elif kind in (KIND_DATE, KIND_DATETIME):
            pat = ocb.currentData()
            spec.date_pattern = str(pat) if pat else "%d/%m/%Y"
        return spec

    def _apply_preview_formats(self) -> None:
        """Aggiorna solo i testi delle celle visibili (stesso ordine colonne dell’anteprima)."""
        if not self._raw_matrix or self._import_n <= 0:
            return
        order = self._visible_column_order()
        if not order:
            return
        data_rows = self._import_data_rows
        for r in range(self._table.rowCount()):
            if r >= len(data_rows):
                break
            row_vals = data_rows[r]
            for k, j in enumerate(order):
                val = row_vals[j] if j < len(row_vals) else None
                spec = self._spec_from_row(j)
                txt = format_cell_as_string(val, spec)
                item = self._table.item(r, k + 1)
                if item is not None:
                    item.setText(txt)

    def _on_ok(self) -> None:
        if not self._raw_matrix:
            QMessageBox.warning(self, "Importazione", "Nessun dato da importare.")
            return
        use_hdr = self._hdr_check.isChecked()
        if use_hdr:
            header_row = self._raw_matrix[0]
            data_rows = self._raw_matrix[1:]
            n = len(header_row)
            base_headers = [
                str(header_row[j]).strip() if header_row[j] is not None else f"Col {j + 1}"
                for j in range(n)
            ]
        else:
            n = max(len(r) for r in self._raw_matrix)
            data_rows = self._raw_matrix
            base_headers = [f"Col {j + 1}" for j in range(n)]

        col_order: list[int] = []
        export_headers: list[str] = []
        for i in range(self._col_list.count()):
            it = self._col_list.item(i)
            if it is None:
                continue
            if it.checkState() != Qt.CheckState.Checked:
                continue
            j = int(it.data(Qt.ItemDataRole.UserRole))
            if 0 <= j < n:
                col_order.append(j)
                export_headers.append(base_headers[j])

        if not col_order:
            QMessageBox.warning(self, "Importazione", "Seleziona almeno una colonna.")
            return

        out_rows: list[list[Any]] = []
        for r in range(self._table.rowCount()):
            inc = self._table.item(r, 0)
            if inc is None or inc.checkState() != Qt.CheckState.Checked:
                continue
            if r >= len(data_rows):
                break
            row_vals = data_rows[r]
            out_rows.append(
                [row_vals[j] if j < len(row_vals) else None for j in col_order]
            )

        if not out_rows:
            QMessageBox.warning(self, "Importazione", "Seleziona almeno una riga.")
            return

        out_formats: list[ColumnFormatSpec] = []
        for j in col_order:
            out_formats.append(self._spec_from_row(j))

        export_row_indices: list[int] = []
        for r in range(self._table.rowCount()):
            inc = self._table.item(r, 0)
            if inc is None or inc.checkState() != Qt.CheckState.Checked:
                continue
            if r >= len(data_rows):
                break
            export_row_indices.append(r)

        self._outcome = ExcelImportOutcome(
            source_path=self._path,
            sheet_name=self._sheet_combo.currentText(),
            headers=export_headers,
            rows=out_rows,
            column_formats=out_formats,
            header_row_used=use_hdr,
            export_column_order=list(col_order),
            export_source_row_indices=export_row_indices,
        )
        self.accept()

    def outcome(self) -> ExcelImportOutcome | None:
        return self._outcome
