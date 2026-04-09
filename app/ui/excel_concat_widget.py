"""Linguetta Concatenazione: testo da colonna (anche IN SQL) o unione colonne per riga."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, Any, cast

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from app.excel_concat import (
    concat_column_sql_in,
    concat_column_with_separator,
    concat_rows_merge_columns,
    filter_data_rows,
)
from app.excel_format import format_cell_as_string, formats_or_auto, sort_unique_display_values_for_spec
from app.excel_export import export_table_xlsx

if TYPE_CHECKING:
    from app.ui.excel_management_widget import ExcelManagementWidget

_RULES_ACCENT = "#0f766e"

_FILTER_OPS: list[tuple[str, str]] = [
    ("Uguale a", "eq"),
    ("Diverso da", "ne"),
    ("Contiene", "contains"),
    ("Non contiene", "not_contains"),
    ("Inizia con", "starts"),
    ("Finisce con", "ends"),
    ("Vuoto", "empty"),
    ("Non vuoto", "not_empty"),
]

_CONCAT_STYLESHEET = f"""
QWidget#concatColumn {{
    background-color: transparent;
}}
QWidget#concatColumn QLabel#sectionTitle {{
    color: #0f172a;
    margin: 0;
}}
QWidget#concatColumn QGroupBox {{
    background-color: #f0fdfa;
    border: 1px solid #a7f3d0;
    border-left: 4px solid {_RULES_ACCENT};
    border-radius: 8px;
    margin-top: 10px;
    padding: 12px 10px 14px 10px;
}}
QWidget#concatColumn QGroupBox:first-of-type {{
    margin-top: 0px;
}}
QWidget#concatColumn QGroupBox QLabel {{
    color: #334155;
}}
QWidget#concatColumn QGroupBox QLabel#subText {{
    color: #475569;
    font-weight: 600;
}}
QWidget#concatColumn QComboBox, QWidget#concatColumn QLineEdit {{
    background-color: #ffffff;
    color: #0f172a;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 5px 10px;
}}
"""


def _apply_gb_title_font(gb: QGroupBox) -> None:
    tf = QFont(gb.font())
    tf.setPointSize(15)
    tf.setBold(True)
    tf.setWeight(QFont.Weight.Bold)
    gb.setFont(tf)
    bf = QFont()
    bf.setBold(False)
    bf.setWeight(QFont.Weight.Normal)
    for w in gb.findChildren(QWidget):
        if isinstance(w, QLabel):
            w.setFont(bf)


class ExcelConcatWidget(QWidget):
    def __init__(self, manager: ExcelManagementWidget) -> None:
        super().__init__()
        self._manager = manager
        self._filter_headers: list[str] = []

        root = QHBoxLayout(self)
        root.setContentsMargins(8, 4, 8, 8)
        root.setSpacing(16)

        left_wrap = QWidget()
        left_wrap.setObjectName("concatColumn")
        left_wrap.setStyleSheet(_CONCAT_STYLESHEET)
        left_wrap.setMinimumWidth(380)
        left_wrap.setMaximumWidth(560)
        left_lay = QVBoxLayout(left_wrap)
        left_lay.setContentsMargins(4, 0, 8, 0)
        left_lay.setSpacing(8)
        left_lay.setAlignment(Qt.AlignmentFlag.AlignTop)

        t = QLabel("Opzioni")
        t.setObjectName("sectionTitle")
        tf = QFont()
        tf.setPointSize(16)
        tf.setBold(True)
        tf.setWeight(QFont.Weight.Bold)
        t.setFont(tf)
        left_lay.addWidget(t)

        gb_src = QGroupBox("1. Fonte dati")
        sl = QVBoxLayout(gb_src)
        sl.setSpacing(8)
        sl.setContentsMargins(4, 8, 4, 4)
        h0 = QLabel("Import dalla vista Importazione o risultati confermati dalla vista Confronto.")
        h0.setObjectName("subText")
        h0.setWordWrap(True)
        sl.addWidget(h0)
        self._source_combo = QComboBox()
        self._source_combo.setMinimumContentsLength(40)
        sl.addWidget(self._source_combo)
        left_lay.addWidget(gb_src)
        _apply_gb_title_font(gb_src)

        gb_op = QGroupBox("2. Operazione")
        ol = QVBoxLayout(gb_op)
        ol.setSpacing(8)
        ol.setContentsMargins(4, 8, 4, 4)
        self._op_combo = QComboBox()
        self._op_combo.addItem("Concatenazione colonna", "col")
        self._op_combo.addItem("Concatenazione righe", "row")
        ol.addWidget(self._op_combo)
        left_lay.addWidget(gb_op)
        _apply_gb_title_font(gb_op)

        gb_filt = QGroupBox("3. Filtro righe (opzionale)")
        fl = QVBoxLayout(gb_filt)
        fl.setSpacing(8)
        fl.setContentsMargins(4, 8, 4, 4)
        hf = QLabel(
            "Solo le righe che passano tutte le condizioni (AND) entrano nel risultato. "
            "Utile anche per concatenazioni SQL IN o testo su sottoinsiemi. "
            "Il confronto usa il valore formattato come colonna (tipo/formato). "
            "Le condizioni vuote sono ignorate."
        )
        hf.setObjectName("subText")
        hf.setWordWrap(True)
        fl.addWidget(hf)
        self._filter_enable = QCheckBox("Applica filtri alle righe")
        fl.addWidget(self._filter_enable)
        self._filter_add_btn = QPushButton("Aggiungi condizione")
        self._filter_add_btn.clicked.connect(self._on_filter_add_row)
        fl.addWidget(self._filter_add_btn)
        self._filter_wrap = QWidget()
        self._filter_wrap.setSizePolicy(
            QSizePolicy.Policy.Preferred,
            QSizePolicy.Policy.Maximum,
        )
        self._filter_lay = QVBoxLayout(self._filter_wrap)
        self._filter_lay.setContentsMargins(0, 0, 0, 0)
        self._filter_lay.setSpacing(6)
        self._filter_row_widgets: list[tuple[QWidget, QComboBox, QComboBox, QComboBox]] = []
        fl.addWidget(self._filter_wrap)
        self._filter_wrap.setVisible(False)
        self._filter_add_btn.setVisible(False)
        self._filter_enable.toggled.connect(self._on_filter_toggled)
        left_lay.addWidget(gb_filt)
        _apply_gb_title_font(gb_filt)

        self._opt_col = QGroupBox("4. Opzioni — colonna")
        ocl = QVBoxLayout(self._opt_col)
        ocl.setSpacing(8)
        ocl.setContentsMargins(4, 8, 4, 4)
        ocl.addWidget(QLabel("Colonna:"))
        self._col_combo = QComboBox()
        ocl.addWidget(self._col_combo)
        self._sql_cb = QCheckBox("Filtro SQL (formato IN (…))")
        ocl.addWidget(self._sql_cb)
        self._plain_sep_row = QWidget()
        ps = QHBoxLayout(self._plain_sep_row)
        ps.setContentsMargins(0, 0, 0, 0)
        ps.addWidget(QLabel("Separatore:"))
        self._sep_edit = QLineEdit()
        self._sep_edit.setPlaceholderText("es. ,  |  ;  (spazio)")
        self._sep_edit.setText(", ")
        ps.addWidget(self._sep_edit, 1)
        ocl.addWidget(self._plain_sep_row)
        self._sql_opts = QWidget()
        sol = QVBoxLayout(self._sql_opts)
        sol.setContentsMargins(8, 4, 0, 0)
        sol.setSpacing(6)
        sol.addWidget(QLabel("Disposizione IN:"))
        horiz = QHBoxLayout()
        self._sql_h = QRadioButton("In orizzontale (una riga)")
        self._sql_v = QRadioButton("In verticale (più righe)")
        self._sql_h.setChecked(True)
        horiz.addWidget(self._sql_h)
        horiz.addWidget(self._sql_v)
        horiz.addStretch(1)
        sol.addLayout(horiz)
        self._quotes_cb = QCheckBox("Valori tra apici singoli '…'")
        self._quotes_cb.setChecked(True)
        sol.addWidget(self._quotes_cb)
        ocl.addWidget(self._sql_opts)
        left_lay.addWidget(self._opt_col)
        _apply_gb_title_font(self._opt_col)

        self._opt_row = QGroupBox("4. Opzioni — righe")
        orl = QVBoxLayout(self._opt_row)
        orl.setSpacing(8)
        orl.setContentsMargins(4, 8, 4, 4)
        h1 = QLabel(
            "Spunta le colonne da unire; trascina le righe per l’ordine nella stringa. "
            "Valori vuoti o null non compaiono (nessun separatore in più). "
            "«Tutte le colonne» spunta tutte (puoi poi deselezionarne alcune)."
        )
        h1.setObjectName("subText")
        h1.setWordWrap(True)
        orl.addWidget(h1)
        self._all_cols_cb = QCheckBox("Tutte le colonne")
        orl.addWidget(self._all_cols_cb)
        orl.addWidget(QLabel("Colonne e ordine:"))
        self._cols_list = QListWidget()
        self._cols_list.setMinimumHeight(160)
        self._cols_list.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        self._cols_list.setDefaultDropAction(Qt.DropAction.MoveAction)
        self._cols_list.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        orl.addWidget(self._cols_list)
        rs = QHBoxLayout()
        rs.addWidget(QLabel("Separatore tra colonne:"))
        self._row_sep_edit = QLineEdit()
        self._row_sep_edit.setPlaceholderText("es. |  ;  spazio")
        self._row_sep_edit.setText(" | ")
        rs.addWidget(self._row_sep_edit, 1)
        orl.addLayout(rs)
        left_lay.addWidget(self._opt_row)
        _apply_gb_title_font(self._opt_row)

        run = QPushButton("Genera testo…")
        run.setMinimumHeight(36)
        run.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        run.clicked.connect(self._on_run)
        left_lay.addWidget(run)

        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        left_scroll.setFrameShape(QFrame.Shape.NoFrame)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        left_scroll.setWidget(left_wrap)

        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(8)
        ot = QLabel("Risultato testo")
        ot.setObjectName("sectionTitle")
        ot.setFont(tf)
        rl.addWidget(ot)
        self._out = QPlainTextEdit()
        self._out.setReadOnly(True)
        self._out.setPlaceholderText("Il testo generato compare qui.")
        self._out.setMinimumHeight(280)
        fmono = QFont("Consolas")
        if not fmono.exactMatch():
            fmono = QFont("Courier New")
        fmono.setPointSize(10)
        self._out.setFont(fmono)
        rl.addWidget(self._out, 1)
        bar = QHBoxLayout()
        cp = QPushButton("Copia negli appunti")
        cp.clicked.connect(self._on_copy)
        ex = QPushButton("Esporta…")
        ex.clicked.connect(self._on_export)
        bar.addWidget(cp)
        bar.addWidget(ex)
        bar.addStretch(1)
        rl.addLayout(bar)

        vsep = QFrame()
        vsep.setFrameShape(QFrame.Shape.VLine)
        vsep.setFrameShadow(QFrame.Shadow.Sunken)

        root.addWidget(left_scroll, 2)
        root.addWidget(vsep, 0)
        root.addWidget(right, 3)

        self._source_combo.currentIndexChanged.connect(self._on_source_changed)
        self._op_combo.currentIndexChanged.connect(self._on_op_changed)
        self._sql_cb.toggled.connect(self._on_sql_toggled)
        self._all_cols_cb.toggled.connect(self._on_all_cols_toggled)
        self._cols_list.itemChanged.connect(self._on_cols_item_changed)

        self._on_sql_toggled(self._sql_cb.isChecked())
        self._on_op_changed()
        self._refresh_sources()

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self._refresh_sources()

    def refresh_imports_from_manager(self) -> None:
        """Aggiorna fonti dati (chiamato da ExcelManagementWidget)."""
        self._refresh_sources()

    def _refresh_sources(self) -> None:
        cur = self._source_combo.currentData()
        self._source_combo.blockSignals(True)
        self._source_combo.clear()
        for s in self._manager.stored_imports():
            self._source_combo.addItem(f"Import: {s.display_label()}", ("import", s.id))
        for r in self._manager.compare_results():
            self._source_combo.addItem(f"Confronto: {r.title}", ("compare", r.id))
        self._source_combo.blockSignals(False)
        if cur is not None:
            for i in range(self._source_combo.count()):
                if self._source_combo.itemData(i) == cur:
                    self._source_combo.setCurrentIndex(i)
                    break
        self._on_source_changed()

    def _current_table(self) -> tuple[list[str], list[list[Any]]] | None:
        raw = self._source_combo.currentData()
        if raw is None:
            return None
        kind, sid = cast(tuple[str, int], raw)
        return self._manager.resolve_table_source(kind, sid)

    def _on_source_changed(self) -> None:
        self._col_combo.clear()
        self._cols_list.clear()
        t = self._current_table()
        if t is None:
            return
        headers, _rows = t
        self._cols_list.blockSignals(True)
        for j, h in enumerate(headers):
            self._col_combo.addItem(h, j)
            it = QListWidgetItem(h)
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            it.setCheckState(Qt.CheckState.Unchecked)
            it.setData(Qt.ItemDataRole.UserRole, j)
            self._cols_list.addItem(it)
        self._cols_list.blockSignals(False)
        self._all_cols_cb.blockSignals(True)
        self._all_cols_cb.setChecked(False)
        self._all_cols_cb.blockSignals(False)

    def _clear_filter_rows(self) -> None:
        while self._filter_lay.count():
            item = self._filter_lay.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        self._filter_row_widgets.clear()

    def _fill_filter_col_combo(self, cb: QComboBox) -> None:
        cb.clear()
        for j, h in enumerate(self._filter_headers):
            cb.addItem(h, j)

    def _unique_formatted_values_for_column(self, col_idx: int) -> list[str]:
        """Valori distinti come stringa formattata (stesso criterio del filtro), ordine di apparizione."""
        t = self._current_table()
        raw_src = self._source_combo.currentData()
        if t is None or raw_src is None:
            return []
        headers, rows = t
        hl = len(headers)
        if col_idx < 0 or col_idx >= hl:
            return []
        kind, sid = cast(tuple[str, int], raw_src)
        fmts = self._manager.resolve_column_formats(kind, sid)
        specs = formats_or_auto(hl, fmts)
        spec = specs[col_idx]
        pairs: list[tuple[Any, str]] = []
        seen: set[str] = set()
        for row in rows:
            v = row[col_idx] if col_idx < len(row) else None
            s = format_cell_as_string(v, spec)
            if s not in seen:
                seen.add(s)
                pairs.append((v, s))
        return sort_unique_display_values_for_spec(pairs, spec)

    def _populate_filter_value_combo(self, col_cb: QComboBox, val_cb: QComboBox) -> None:
        jd = col_cb.currentData()
        prev = val_cb.currentText().strip()
        val_cb.blockSignals(True)
        val_cb.clear()
        if jd is None:
            val_cb.blockSignals(False)
            return
        col_idx = int(jd)
        for s in self._unique_formatted_values_for_column(col_idx):
            val_cb.addItem(s)
        if prev:
            ix = val_cb.findText(prev, Qt.MatchFlag.MatchExactly)
            if ix >= 0:
                val_cb.setCurrentIndex(ix)
            else:
                val_cb.setEditText(prev)
        else:
            val_cb.setCurrentIndex(-1)
            val_cb.setEditText("")
        val_cb.blockSignals(False)

    def _on_filter_toggled(self, checked: bool) -> None:
        self._filter_wrap.setVisible(checked)
        self._filter_add_btn.setVisible(checked)
        if checked and not self._filter_row_widgets:
            self._on_filter_add_row()

    def _on_filter_add_row(self) -> None:
        t = self._current_table()
        if t is None:
            QMessageBox.warning(
                self,
                "Concatenazione",
                "Seleziona prima una fonte dati nella sezione «1. Fonte dati».",
            )
            return
        headers, _rows = t
        self._filter_headers = list(headers)
        if not self._filter_headers:
            QMessageBox.information(
                self,
                "Concatenazione",
                "Questa fonte non ha colonne: non è possibile aggiungere condizioni di filtro.",
            )
            return
        row_w = QWidget()
        h = QHBoxLayout(row_w)
        h.setContentsMargins(0, 0, 0, 0)
        col_cb = QComboBox()
        self._fill_filter_col_combo(col_cb)
        op_cb = QComboBox()
        for lab, key in _FILTER_OPS:
            op_cb.addItem(lab, key)
        val_cb = QComboBox()
        val_cb.setEditable(True)
        val_cb.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        val_cb.setMinimumWidth(160)
        val_cb.setPlaceholderText("Valore o scegli dall’elenco")

        def _sync_val_enabled() -> None:
            op = str(op_cb.currentData() or "")
            en = op not in ("empty", "not_empty")
            val_cb.setEnabled(en)

        op_cb.currentIndexChanged.connect(lambda _i: _sync_val_enabled())
        _sync_val_enabled()

        col_cb.currentIndexChanged.connect(lambda _i: self._populate_filter_value_combo(col_cb, val_cb))
        self._populate_filter_value_combo(col_cb, val_cb)

        rm = QPushButton("Rimuovi")
        rm.clicked.connect(lambda _c=False, rw=row_w: self._on_filter_remove_row(rw))
        h.addWidget(col_cb, 2)
        h.addWidget(op_cb, 2)
        h.addWidget(val_cb, 2)
        h.addWidget(rm)
        self._filter_lay.addWidget(row_w)
        self._filter_row_widgets.append((row_w, col_cb, op_cb, val_cb))

    def _on_filter_remove_row(self, row_w: QWidget) -> None:
        for i, (w, _c, _o, _v) in enumerate(self._filter_row_widgets):
            if w is row_w:
                self._filter_row_widgets.pop(i)
                self._filter_lay.removeWidget(row_w)
                row_w.deleteLater()
                break

    def _filter_conditions_from_ui(self) -> list[tuple[int, str, str]]:
        out: list[tuple[int, str, str]] = []
        for _w, col_cb, op_cb, val_cb in self._filter_row_widgets:
            jd = col_cb.currentData()
            if jd is None:
                continue
            op = str(op_cb.currentData() or "eq")
            needle = val_cb.currentText()
            if op in ("contains", "not_contains", "starts", "ends") and needle.strip() == "":
                continue
            out.append((int(jd), op, needle))
        return out

    def _on_op_changed(self, *_args: object) -> None:
        is_col = self._op_combo.currentData() == "col"
        self._opt_col.setVisible(is_col)
        self._opt_row.setVisible(not is_col)

    def _on_sql_toggled(self, checked: bool) -> None:
        self._plain_sep_row.setVisible(not checked)
        self._sql_opts.setVisible(checked)

    def _on_all_cols_toggled(self, checked: bool) -> None:
        if not checked:
            return
        self._cols_list.blockSignals(True)
        for i in range(self._cols_list.count()):
            it = self._cols_list.item(i)
            if it is not None:
                it.setCheckState(Qt.CheckState.Checked)
        self._cols_list.blockSignals(False)

    def _on_cols_item_changed(self, _item: QListWidgetItem) -> None:
        n = self._cols_list.count()
        if n == 0:
            return
        all_on = True
        for i in range(n):
            it = self._cols_list.item(i)
            if it is None or it.checkState() != Qt.CheckState.Checked:
                all_on = False
                break
        self._all_cols_cb.blockSignals(True)
        self._all_cols_cb.setChecked(all_on)
        self._all_cols_cb.blockSignals(False)

    def _on_run(self) -> None:
        t = self._current_table()
        if t is None:
            QMessageBox.warning(self, "Concatenazione", "Seleziona una fonte dati.")
            return
        headers, rows = t
        if not rows:
            QMessageBox.information(self, "Concatenazione", "Nessuna riga nella tabella selezionata.")
            return

        raw_src = self._source_combo.currentData()
        if raw_src is None:
            return
        kind, sid = cast(tuple[str, int], raw_src)
        fmts = self._manager.resolve_column_formats(kind, sid)
        hl = len(headers)

        work_rows = rows
        if self._filter_enable.isChecked():
            cond = self._filter_conditions_from_ui()
            if cond:
                work_rows = filter_data_rows(
                    rows, cond, column_formats=fmts, headers_len=hl
                )
                if not work_rows:
                    QMessageBox.information(
                        self,
                        "Concatenazione",
                        "Nessuna riga corrisponde ai filtri impostati.",
                    )
                    return

        if self._op_combo.currentData() == "col":
            cix = int(self._col_combo.currentData()) if self._col_combo.currentData() is not None else 0
            if cix >= len(headers):
                QMessageBox.warning(self, "Concatenazione", "Colonna non valida.")
                return
            if self._sql_cb.isChecked():
                vertical = self._sql_v.isChecked()
                quotes = self._quotes_cb.isChecked()
                text = concat_column_sql_in(
                    work_rows, cix, vertical, quotes, column_formats=fmts, headers_len=hl
                )
            else:
                sep = self._sep_edit.text()
                text = concat_column_with_separator(
                    work_rows, cix, sep, column_formats=fmts, headers_len=hl
                )
            self._out.setPlainText(text)
            self._last_export_kind = "text"
            return

        # row mode: ordine = ordine visuale nella lista; solo colonne spuntate
        indices: list[int] = []
        for i in range(self._cols_list.count()):
            it = self._cols_list.item(i)
            if it is not None and it.checkState() == Qt.CheckState.Checked:
                indices.append(int(it.data(Qt.ItemDataRole.UserRole)))
        if not indices:
            QMessageBox.warning(
                self,
                "Concatenazione",
                "Spunta almeno una colonna (o usa «Tutte le colonne» per selezionarle tutte).",
            )
            return
        sep = self._row_sep_edit.text()
        lines = concat_rows_merge_columns(
            work_rows, indices, sep, column_formats=fmts, headers_len=hl
        )
        self._out.setPlainText("\n".join(lines))
        self._last_export_kind = "lines"
        self._last_lines = lines

    def _on_copy(self) -> None:
        txt = self._out.toPlainText()
        if not txt.strip():
            return
        QApplication.clipboard().setText(txt)

    def _on_export(self) -> None:
        txt = self._out.toPlainText()
        if not txt.strip():
            QMessageBox.information(self, "Concatenazione", "Niente da esportare.")
            return
        path, sel = QFileDialog.getSaveFileName(
            self,
            "Salva",
            "",
            "Excel (*.xlsx);;Testo (*.txt);;Tutti i file (*.*)",
        )
        if not path:
            return
        low = path.lower()
        kind = getattr(self, "_last_export_kind", "text")
        lines = getattr(self, "_last_lines", [])
        want_xlsx = low.endswith(".xlsx") or ("xlsx" in sel.lower() and not low.endswith(".txt"))
        if want_xlsx:
            p = path if low.endswith(".xlsx") else f"{path}.xlsx"
            try:
                if kind == "lines" and isinstance(lines, list) and lines:
                    export_table_xlsx(p, ["Risultato"], [[x] for x in lines])
                else:
                    export_table_xlsx(p, ["Testo"], [[txt]])
            except Exception as exc:
                QMessageBox.critical(self, "Esportazione", str(exc))
            return
        p = path if low.endswith(".txt") else f"{path}.txt"
        try:
            Path(p).write_text(txt, encoding="utf-8")
        except Exception as exc:
            QMessageBox.critical(self, "Esportazione", str(exc))
