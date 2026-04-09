"""Linguetta Confronto: verifica su un foglio o tra più import."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, Any

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QStackedLayout,
    QSizePolicy,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.excel_export import export_table_xlsx
from app.excel_format import ColumnFormatSpec, format_cell_as_string, format_matrix_strings, formats_or_auto
from app.excel_compare import (
    multi_anti_join_wide,
    multi_inner_join_wide,
    unique_sorted_values_for_column_after_filters,
    unique_sorted_values_for_column_joined,
    single_duplicates_by_column,
    single_duplicates_exact,
    single_filter_fields,
    single_unique_by_column,
    single_unique_exact,
    wide_column_formats_for_tables,
)
from app.ui.excel_import_dialog import ExcelImportOutcome, StoredImport
from app.ui.excel_result_preview_dialog import ExcelResultPreviewDialog

if TYPE_CHECKING:
    from app.ui.excel_management_widget import ExcelManagementWidget

# Accento = menu laterale Strumenti (#toolsSideMenu::item:selected) — main_window.py
_RULES_ACCENT = "#0f766e"


def _apply_rules_groupbox_title_font(gb: QGroupBox) -> None:
    """Grassetto reale sul titolo: QSS su QGroupBox::title non è sempre rispettato (Windows)."""
    title_font = QFont(gb.font())
    title_font.setPointSize(15)
    title_font.setBold(True)
    title_font.setWeight(QFont.Weight.Bold)
    gb.setFont(title_font)
    body_font = QFont()
    body_font.setBold(False)
    body_font.setWeight(QFont.Weight.Normal)
    for w in gb.findChildren(QWidget):
        if isinstance(w, QLabel):
            on = w.objectName()
            if on == "subText":
                f = QFont(body_font)
                f.setWeight(QFont.Weight.DemiBold)
                w.setFont(f)
            elif on == "keyPartTitle":
                f = QFont(body_font)
                f.setPointSize(14)
                f.setBold(True)
                f.setWeight(QFont.Weight.Bold)
                w.setFont(f)
            elif on == "fileTag":
                f = QFont(body_font)
                f.setWeight(QFont.Weight.DemiBold)
                w.setFont(f)
            else:
                w.setFont(body_font)
        else:
            w.setFont(body_font)

# Superficie sezione (area) ≠ campo (input bianco): gerarchia chiara, testi sempre scuri su chiaro
_RULES_COLUMN_STYLESHEET = f"""
QWidget#rulesColumn {{
    background-color: transparent;
}}
QWidget#rulesColumn QLabel#sectionTitle {{
    color: #0f172a;
    margin-bottom: 0px;
    margin-top: 0px;
}}
/* Area sezione: mint tenue; bordo sinistro accento teal */
QWidget#rulesColumn QGroupBox {{
    background-color: #f0fdfa;
    border: 1px solid #a7f3d0;
    border-left: 4px solid {_RULES_ACCENT};
    border-radius: 8px;
    margin-top: 14px;
    padding: 14px 12px 18px 12px;
}}
QWidget#rulesColumn QGroupBox:first-of-type {{
    margin-top: 0px;
}}
QWidget#rulesColumn QGroupBox::title {{
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
    color: #0f172a;
    font-weight: 700;
    font-size: 15px;
}}
QWidget#rulesColumn QGroupBox QLabel {{
    color: #334155;
}}
QWidget#rulesColumn QGroupBox QLabel#subText {{
    color: #475569;
    background-color: transparent;
    font-weight: 600;
}}
QWidget#rulesColumn QRadioButton {{
    color: #0f172a;
    background-color: transparent;
    spacing: 10px;
}}
QWidget#rulesColumn QRadioButton::indicator {{
    width: 18px;
    height: 18px;
}}
QWidget#rulesColumn QRadioButton::indicator:unchecked {{
    background-color: #ffffff;
    border: 2px solid #94a3b8;
    border-radius: 9px;
}}
QWidget#rulesColumn QRadioButton::indicator:unchecked:hover {{
    border-color: #64748b;
    background-color: #f8fafc;
}}
QWidget#rulesColumn QRadioButton::indicator:checked {{
    background-color: {_RULES_ACCENT};
    border: 2px solid {_RULES_ACCENT};
    border-radius: 9px;
}}
QWidget#rulesColumn QRadioButton::indicator:checked:hover {{
    background-color: #0d9488;
    border-color: #0d9488;
}}
QWidget#rulesColumn QRadioButton:focus {{
    outline: none;
}}
/* Lista: come i campi, bianca */
QWidget#rulesColumn QListWidget {{
    background-color: #ffffff;
    color: #0f172a;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    outline: none;
    padding: 4px;
}}
QWidget#rulesColumn QListWidget::item {{
    padding: 6px 8px;
    border-radius: 4px;
}}
QWidget#rulesColumn QListWidget::item:selected {{
    background-color: #ccfbf1;
    color: #0f172a;
}}
QWidget#rulesColumn QListWidget::item:hover {{
    background-color: #f0fdfa;
}}
/* Campi: sempre bianchi, separati dall’area sezione */
QWidget#rulesColumn QComboBox {{
    background-color: #ffffff;
    color: #0f172a;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 5px 10px;
    min-height: 1.15em;
}}
QWidget#rulesColumn QComboBox:hover {{
    border-color: #94a3b8;
}}
QWidget#rulesColumn QComboBox:focus {{
    border-color: {_RULES_ACCENT};
}}
QWidget#rulesColumn QComboBox::drop-down {{
    border: none;
    width: 22px;
}}
QWidget#rulesColumn QComboBox::down-arrow {{
    width: 0;
    height: 0;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid #334155;
    margin-right: 8px;
}}
QWidget#rulesColumn QComboBox QAbstractItemView {{
    background-color: #ffffff;
    color: #0f172a;
    selection-background-color: #ccfbf1;
    selection-color: #0f172a;
    border: 1px solid #cbd5e1;
    outline: none;
}}
QWidget#rulesColumn QComboBox QLineEdit {{
    color: #0f172a;
    background: transparent;
    selection-background-color: #99f6e4;
    selection-color: #0f172a;
}}
QWidget#rulesColumn QSpinBox {{
    background-color: #ffffff;
    color: #0f172a;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 4px 8px;
    min-height: 1.15em;
}}
QWidget#rulesColumn QSpinBox:focus {{
    border-color: {_RULES_ACCENT};
}}
QWidget#rulesColumn QSpinBox::up-button, QWidget#rulesColumn QSpinBox::down-button {{
    background: #f1f5f9;
    border: none;
    width: 18px;
}}
QWidget#rulesColumn QSpinBox::up-button:hover, QWidget#rulesColumn QSpinBox::down-button:hover {{
    background: #e2e8f0;
}}
QWidget#rulesColumn QPushButton {{
    background-color: {_RULES_ACCENT};
    color: #ffffff;
    border: none;
    border-radius: 6px;
    padding: 8px 12px;
    font-weight: 600;
}}
QWidget#rulesColumn QPushButton:hover {{
    background-color: #0d9488;
}}
QWidget#rulesColumn QPushButton:pressed {{
    background-color: #0b7f78;
}}
/* Card interne (chiave / filtri): bianco su mint, titoli mai bianchi */
QWidget#rulesColumn QFrame#keyPartBlock {{
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
}}
QWidget#rulesColumn QLabel#keyPartTitle {{
    color: #0f172a;
    font-weight: 700;
    font-size: 14px;
}}
QWidget#rulesColumn QLabel#fileTag {{
    color: #475569;
    font-size: 11px;
    font-weight: 600;
}}
/* Blocco “Numero campi” come card bianca */
QWidget#rulesColumn QWidget#rulesInnerCard {{
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
}}
"""


def _short_import_caption(s: StoredImport) -> str:
    """Etichetta breve per non affollare le righe (solo nome file + foglio)."""
    name = Path(s.outcome.source_path).name
    return f"{name} — «{s.outcome.sheet_name}»"


def _fill_table(
    table: QTableWidget,
    headers: list[str],
    rows: list[list[Any]],
    max_rows: int = 5000,
    column_formats: list[ColumnFormatSpec] | None = None,
) -> None:
    table.clear()
    n = len(headers)
    fmts = formats_or_auto(n, column_formats)
    table.setColumnCount(n)
    table.setHorizontalHeaderLabels(headers)
    show = min(len(rows), max_rows)
    table.setRowCount(show)
    for r in range(show):
        for c in range(n):
            val = rows[r][c] if c < len(rows[r]) else None
            txt = format_cell_as_string(val, fmts[c])
            it = QTableWidgetItem(txt)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(r, c, it)


class _FilterConditionsWidget(QWidget):
    """Righe (colonna + valore combobox editabile con valori dal file)."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._root = QVBoxLayout(self)
        self._root.setContentsMargins(0, 0, 0, 0)
        self._rows: list[tuple[QComboBox, QComboBox]] = []
        self._outcome: ExcelImportOutcome | None = None

    @staticmethod
    def _make_value_combo() -> QComboBox:
        c = QComboBox()
        c.setEditable(True)
        c.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        c.setMaxVisibleItems(30)
        le = c.lineEdit()
        if le is not None:
            le.setPlaceholderText("Scegli o digita un valore…")
        return c

    def _row_index(self, *, col_cb: QComboBox | None = None, val_cb: QComboBox | None = None) -> int:
        for i, (cb, vc) in enumerate(self._rows):
            if col_cb is not None and cb is col_cb:
                return i
            if val_cb is not None and vc is val_cb:
                return i
        return -1

    def _prior_filters_for_row(self, row_index: int) -> list[tuple[int, str]]:
        out: list[tuple[int, str]] = []
        for j in range(row_index):
            ccb, vcb = self._rows[j]
            txt = vcb.currentText().strip()
            if not txt:
                continue
            raw = ccb.currentData()
            if raw is None:
                continue
            out.append((int(raw), txt))
        return out

    def _refresh_value_combos_from(self, start: int = 0) -> None:
        for i in range(start, len(self._rows)):
            col_cb, val_cb = self._rows[i]
            self._refill_row_value(col_cb, val_cb, self._prior_filters_for_row(i))

    def _refill_row_value(
        self, col_cb: QComboBox, val_cb: QComboBox, prior_filters: list[tuple[int, str]]
    ) -> None:
        val_cb.blockSignals(True)
        saved = val_cb.currentText()
        val_cb.clear()
        if self._outcome is None:
            val_cb.setEnabled(False)
            val_cb.clearEditText()
            val_cb.blockSignals(False)
            return
        raw = col_cb.currentData()
        if raw is None:
            val_cb.setEnabled(False)
            val_cb.clearEditText()
            val_cb.blockSignals(False)
            return
        val_cb.setEnabled(True)
        col_idx = int(raw)
        for v in unique_sorted_values_for_column_after_filters(
            self._outcome, col_idx, prior_filters
        ):
            val_cb.addItem(v)
        val_cb.setEditText(saved)
        val_cb.blockSignals(False)

    def set_import_outcome(self, outcome: ExcelImportOutcome | None) -> None:
        self._outcome = outcome
        while self._root.count():
            it = self._root.takeAt(0)
            w = it.widget()
            if w is not None:
                w.deleteLater()
        self._rows.clear()
        if outcome is not None and outcome.headers:
            self._add_row()

    def _add_row(self) -> None:
        row_w = QWidget()
        hl = QHBoxLayout(row_w)
        hl.setContentsMargins(0, 2, 0, 2)
        hl.setSpacing(8)
        cb = QComboBox()
        headers = self._outcome.headers if self._outcome else []
        for j, h in enumerate(headers):
            cb.addItem(h, j)
        val_cb = self._make_value_combo()
        rm = QPushButton("Rimuovi")
        hl.addWidget(cb, 1)
        hl.addWidget(val_cb, 1)
        hl.addWidget(rm)
        self._rows.append((cb, val_cb))
        self._root.addWidget(row_w)

        def on_col_changed(_i: int) -> None:
            idx = self._row_index(col_cb=cb)
            if idx >= 0:
                self._refresh_value_combos_from(idx)

        def on_val_changed() -> None:
            idx = self._row_index(val_cb=val_cb)
            if idx >= 0:
                self._refresh_value_combos_from(idx + 1)

        cb.currentIndexChanged.connect(on_col_changed)
        val_cb.currentIndexChanged.connect(on_val_changed)
        val_cb.currentTextChanged.connect(lambda _t: on_val_changed())

        self._refresh_value_combos_from(self._row_index(col_cb=cb))

        def remove() -> None:
            self._rows.remove((cb, val_cb))
            row_w.deleteLater()
            if not self._rows:
                self._add_row()
            else:
                self._refresh_value_combos_from(0)

        rm.clicked.connect(remove)

    def add_condition_clicked(self) -> None:
        self._add_row()

    def active_filters(self) -> list[tuple[int, str]]:
        out: list[tuple[int, str]] = []
        for col_cb, val_cb in self._rows:
            txt = val_cb.currentText().strip()
            if not txt:
                continue
            j = int(col_cb.currentData())
            out.append((j, txt))
        return out


class ExcelCompareWidget(QWidget):
    def __init__(self, manager: ExcelManagementWidget) -> None:
        super().__init__()
        self._manager = manager

        root = QHBoxLayout(self)
        root.setContentsMargins(8, 4, 8, 8)
        root.setSpacing(16)

        # --- Colonna Regole ---
        left_wrap = QWidget()
        left_wrap.setObjectName("rulesColumn")
        left_wrap.setStyleSheet(_RULES_COLUMN_STYLESHEET)
        left_wrap.setMinimumWidth(420)
        left_wrap.setMaximumWidth(580)
        left_lay = QVBoxLayout(left_wrap)
        left_lay.setContentsMargins(4, 0, 8, 0)
        left_lay.setSpacing(10)
        left_lay.setAlignment(Qt.AlignmentFlag.AlignTop)

        rules_title = QLabel("Regole")
        rules_title.setObjectName("sectionTitle")
        _rt = QFont()
        _rt.setPointSize(16)
        _rt.setBold(True)
        _rt.setWeight(QFont.Weight.Bold)
        rules_title.setFont(_rt)
        left_lay.addWidget(rules_title)

        gb_choice = QGroupBox("1. Scelta del confronto")
        choice_lay = QVBoxLayout(gb_choice)
        choice_lay.setSpacing(10)
        choice_lay.setContentsMargins(4, 8, 4, 4)
        self._mode_single = QRadioButton("Verifica su un singolo foglio importato")
        self._mode_multi = QRadioButton("Verifica tra due o più fogli importati")
        self._mode_single.setChecked(True)
        choice_lay.addWidget(self._mode_single)
        choice_lay.addWidget(self._mode_multi)
        left_lay.addWidget(gb_choice)
        _apply_rules_groupbox_title_font(gb_choice)

        # Due pannelli alternati: QStackedLayout (non QStackedWidget) size = solo pagina corrente.
        self._single_page_widget = self._build_single_page()
        self._multi_page_widget = self._build_multi_page()
        mode_pages = QWidget()
        self._mode_stack = QStackedLayout(mode_pages)
        self._mode_stack.setContentsMargins(0, 0, 0, 0)
        self._mode_stack.addWidget(self._single_page_widget)
        self._mode_stack.addWidget(self._multi_page_widget)
        left_lay.addWidget(mode_pages, 0)

        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        left_scroll.setFrameShape(QFrame.Shape.NoFrame)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        left_scroll.setWidget(left_wrap)

        self._mode_single.toggled.connect(self._on_compare_mode_toggled)
        self._mode_multi.toggled.connect(self._on_compare_mode_toggled)

        # --- Colonna Risultati ---
        right_col = QWidget()
        right_lay = QVBoxLayout(right_col)
        right_lay.setContentsMargins(0, 0, 0, 0)
        right_lay.setSpacing(8)

        res_title = QLabel("Risultati")
        res_title.setObjectName("sectionTitle")
        right_lay.addWidget(res_title)

        hint = QLabel(
            "Dopo il confronto si apre un’anteprima per colonne; qui compaiono i risultati confermati "
            "(esportabili in Excel)."
        )
        hint.setObjectName("subText")
        hint.setWordWrap(True)
        right_lay.addWidget(hint)

        result_scroll = QScrollArea()
        result_scroll.setWidgetResizable(True)
        result_scroll.setFrameShape(QFrame.Shape.StyledPanel)
        inner = QWidget()
        self._result_host = QVBoxLayout(inner)
        self._result_host.setContentsMargins(8, 8, 8, 8)
        self._result_host.addStretch(1)
        result_scroll.setWidget(inner)
        right_lay.addWidget(result_scroll, 1)

        vsep = QFrame()
        vsep.setFrameShape(QFrame.Shape.VLine)
        vsep.setFrameShadow(QFrame.Shadow.Sunken)

        root.addWidget(left_scroll, 2)
        root.addWidget(vsep, 0)
        root.addWidget(right_col, 3)

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self._refresh_import_lists()

    def refresh_imports_from_manager(self) -> None:
        """Aggiorna elenco fonti importate (chiamato da ExcelManagementWidget)."""
        self._refresh_import_lists()

    def _on_compare_mode_toggled(self) -> None:
        single = self._mode_single.isChecked()
        self._mode_stack.setCurrentIndex(0 if single else 1)

    def _build_single_page(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        lay.setAlignment(Qt.AlignmentFlag.AlignTop)

        gb_imp = QGroupBox("2. Importazione")
        il = QVBoxLayout(gb_imp)
        il.setSpacing(10)
        il.setContentsMargins(4, 8, 4, 4)
        imp_hint = QLabel("Scegli un import già caricato nella vista Importazione.")
        imp_hint.setObjectName("subText")
        imp_hint.setWordWrap(True)
        il.addWidget(imp_hint)
        self._single_import_combo = QComboBox()
        il.addWidget(self._single_import_combo)
        lay.addWidget(gb_imp)

        gb_op = QGroupBox("3. Operazione e filtri")
        ol = QVBoxLayout(gb_op)
        ol.setSpacing(10)
        ol.setContentsMargins(4, 8, 4, 4)
        self._single_op = QComboBox()
        self._single_op.addItem(
            "Estrai righe doppie — tutti i campi perfettamente uguali", "dup_all"
        )
        self._single_op.addItem(
            "Estrai righe doppie — uguali in base al campo selezionato", "dup_col"
        )
        self._single_op.addItem(
            "Estrai righe singole (uniche) — nessun duplicato su tutti i campi", "uniq_all"
        )
        self._single_op.addItem(
            "Estrai righe singole — uniche in base al campo selezionato", "uniq_col"
        )
        self._single_op.addItem(
            "Filtra per campo — una o più condizioni (AND)", "filter"
        )
        ol.addWidget(self._single_op)

        self._single_col_combo = QComboBox()
        self._single_col_combo.setVisible(False)
        ol.addWidget(self._single_col_combo)

        self._filter_block = _FilterConditionsWidget()
        self._filter_block.setVisible(False)
        ol.addWidget(self._filter_block)
        filt_btns = QHBoxLayout()
        add_f = QPushButton("Aggiungi condizione di filtro")
        add_f.clicked.connect(self._filter_block.add_condition_clicked)
        filt_btns.addWidget(add_f)
        filt_btns.addStretch(1)
        ol.addLayout(filt_btns)

        lay.addWidget(gb_op)

        lay.addSpacing(20)
        self._single_run_btn = QPushButton("Esegui e apri anteprima…")
        self._single_run_btn.setMinimumHeight(32)
        self._single_run_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._single_run_btn.clicked.connect(self._run_single)
        lay.addWidget(self._single_run_btn)

        _apply_rules_groupbox_title_font(gb_imp)
        _apply_rules_groupbox_title_font(gb_op)

        self._single_import_combo.currentIndexChanged.connect(self._on_single_import_changed)
        self._single_op.currentIndexChanged.connect(self._on_single_op_changed)
        self._on_single_op_changed()
        return w

    def _make_filter_value_combo(self) -> QComboBox:
        c = QComboBox()
        c.setEditable(True)
        c.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        c.setMaxVisibleItems(30)
        c.setMinimumWidth(200)
        le = c.lineEdit()
        if le is not None:
            le.setPlaceholderText("Scegli o digita un valore…")
        return c

    def _refresh_all_opt_filter_values(self) -> None:
        if not self._opt_filt_rows_per_store:
            return
        combos: list[QComboBox] = []
        for rows in self._opt_filt_rows_per_store:
            for ccol, val_cb in rows:
                combos.append(val_cb)
        for c in combos:
            c.blockSignals(True)
        try:
            for si, rows in enumerate(self._opt_filt_rows_per_store):
                for ri, _ in enumerate(rows):
                    self._refill_opt_filter_value_at(si, ri)
        finally:
            for c in combos:
                c.blockSignals(False)
        for si in range(len(self._opt_filt_rows_per_store)):
            self._update_opt_filter_remove_buttons_for_store(si)

    def _update_opt_filter_remove_buttons_for_store(self, store_idx: int) -> None:
        if store_idx >= len(self._opt_filt_rows_layouts):
            return
        lv = self._opt_filt_rows_layouts[store_idx]
        n = len(self._opt_filt_rows_per_store[store_idx])
        for i in range(lv.count()):
            it = lv.itemAt(i)
            if it is None:
                continue
            row_w = it.widget()
            if row_w is None:
                continue
            lay = row_w.layout()
            if lay is None or lay.count() < 1:
                continue
            rm_w = lay.itemAt(lay.count() - 1).widget()
            if isinstance(rm_w, QPushButton):
                rm_w.setEnabled(n > 1)

    def _add_opt_filter_row_for_store(self, store_idx: int, *, run_refresh: bool = True) -> None:
        stores = self._stores_for_key_mapping()
        if store_idx >= len(stores) or store_idx >= len(self._opt_filt_rows_layouts):
            return
        s = stores[store_idx]
        rvl = self._opt_filt_rows_layouts[store_idx]
        row_w = QWidget()
        hl = QHBoxLayout(row_w)
        hl.setContentsMargins(0, 2, 0, 2)
        hl.setSpacing(8)
        ccol = QComboBox()
        ccol.addItem("— Nessun filtro —", None)
        for j, h in enumerate(s.outcome.headers):
            ccol.addItem(h, j)
        val_cb = self._make_filter_value_combo()
        rm = QPushButton("Rimuovi")
        hl.addWidget(ccol, 1)
        hl.addWidget(val_cb, 1)
        hl.addWidget(rm)
        self._opt_filt_rows_per_store[store_idx].append((ccol, val_cb))
        rvl.addWidget(row_w)

        def on_remove() -> None:
            self._remove_opt_filter_row_for_store(store_idx, row_w)

        rm.clicked.connect(on_remove)
        ccol.currentIndexChanged.connect(lambda _i: self._refresh_all_opt_filter_values())
        val_cb.currentIndexChanged.connect(lambda _i: self._refresh_all_opt_filter_values())
        val_cb.currentTextChanged.connect(lambda _t: self._refresh_all_opt_filter_values())
        if run_refresh:
            self._refresh_all_opt_filter_values()

    def _remove_opt_filter_row_for_store(self, store_idx: int, row_w: QWidget) -> None:
        if store_idx >= len(self._opt_filt_rows_per_store):
            return
        if len(self._opt_filt_rows_per_store[store_idx]) <= 1:
            return
        idx = None
        for i, (ccol, val_cb) in enumerate(self._opt_filt_rows_per_store[store_idx]):
            if ccol.parentWidget() is row_w:
                idx = i
                break
        if idx is None:
            return
        lv = self._opt_filt_rows_layouts[store_idx]
        for i in range(lv.count()):
            it = lv.itemAt(i)
            if it is not None and it.widget() is row_w:
                lv.takeAt(i)
                break
        self._opt_filt_rows_per_store[store_idx].pop(idx)
        row_w.deleteLater()
        self._refresh_all_opt_filter_values()

    def _refill_opt_filter_value_at(self, store_idx: int, row_idx: int) -> None:
        if store_idx >= len(self._opt_filt_rows_per_store):
            return
        rows = self._opt_filt_rows_per_store[store_idx]
        if row_idx >= len(rows):
            return
        col_combo, val_combo = rows[row_idx]
        saved = val_combo.currentText()
        val_combo.clear()
        raw = col_combo.currentData()
        if raw is None:
            val_combo.setEnabled(False)
            val_combo.clearEditText()
            return
        stores = self._stores_for_key_mapping()
        if len(stores) < 2 or store_idx >= len(stores):
            val_combo.setEnabled(False)
            val_combo.clearEditText()
            return
        key_cols = self._collect_key_cols()
        if not key_cols or any(len(x) == 0 for x in key_cols):
            val_combo.setEnabled(False)
            val_combo.clearEditText()
            return
        val_combo.setEnabled(True)
        col_idx = int(raw)
        opts = self._collect_optional_filters()
        if len(opts) != len(stores):
            val_combo.setEditText(saved)
            return
        outcomes = [s.outcome for s in stores]
        vals = unique_sorted_values_for_column_joined(
            outcomes,
            key_cols,
            opts,
            store_idx,
            col_idx,
            target_skip_row=row_idx,
        )
        for v in vals:
            val_combo.addItem(v)
        val_combo.setEditText(saved)

    def _build_multi_page(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        lay.setAlignment(Qt.AlignmentFlag.AlignTop)

        gb_files = QGroupBox("2. Import da includere")
        gl = QVBoxLayout(gb_files)
        gl.setSpacing(10)
        gl.setContentsMargins(4, 8, 4, 4)
        self._hint_files = QLabel()
        self._hint_files.setObjectName("subText")
        self._hint_files.setWordWrap(True)
        gl.addWidget(self._hint_files)
        self._multi_list = QListWidget()
        self._multi_list.setMinimumHeight(100)
        self._multi_list.setMaximumHeight(200)
        self._multi_list.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._multi_list.itemChanged.connect(self._on_multi_list_item_changed)
        gl.addWidget(self._multi_list)
        lay.addWidget(gb_files)
        _apply_rules_groupbox_title_font(gb_files)

        gb_op = QGroupBox("3. Cosa cercare")
        ol = QVBoxLayout(gb_op)
        ol.setSpacing(12)
        ol.setContentsMargins(4, 8, 4, 4)
        self._multi_op = QComboBox()
        self._multi_op.addItem(
            "Righe comuni in tutti i fogli selezionati (join sulla chiave sotto)",
            "common",
        )
        self._multi_op.addItem(
            "Righe del primo file che non compaiono nel secondo (anti-join sulla chiave)",
            "missing",
        )
        ol.addWidget(self._multi_op)

        self._missing_row = QWidget()
        mlay = QVBoxLayout(self._missing_row)
        mlay.setSpacing(10)
        mlay.setContentsMargins(0, 4, 0, 0)
        ra = QHBoxLayout()
        ra.addWidget(QLabel("File A (origine):"), 0)
        self._miss_a = QComboBox()
        self._miss_a.setMinimumContentsLength(28)
        ra.addWidget(self._miss_a, 1)
        mlay.addLayout(ra)
        rb = QHBoxLayout()
        rb.addWidget(QLabel("File B (confronto):"), 0)
        self._miss_b = QComboBox()
        self._miss_b.setMinimumContentsLength(28)
        rb.addWidget(self._miss_b, 1)
        mlay.addLayout(rb)
        ol.addWidget(self._missing_row)
        lay.addWidget(gb_op)
        _apply_rules_groupbox_title_font(gb_op)

        self._key_map_group = QGroupBox("4. Chiave di associazione")
        km = QVBoxLayout(self._key_map_group)
        km.setSpacing(12)
        km.setContentsMargins(4, 8, 4, 4)
        spin_wrap = QWidget()
        spin_wrap.setObjectName("rulesInnerCard")
        spin_lay = QVBoxLayout(spin_wrap)
        spin_lay.setSpacing(8)
        spin_lay.setContentsMargins(10, 10, 10, 10)
        spin_hint = QLabel("Definisci numero chiavi associazioni")
        spin_hint.setObjectName("subText")
        spin_hint.setWordWrap(True)
        spin_lay.addWidget(spin_hint)
        spin_row = QHBoxLayout()
        spin_row.addWidget(QLabel("Numero campi:"), 0)
        self._key_spin = QSpinBox()
        self._key_spin.setMinimum(1)
        self._key_spin.setMaximum(10)
        self._key_spin.setValue(1)
        self._key_spin.setMinimumWidth(72)
        self._key_spin.setToolTip(
            "Esempio: 1 = un solo campo; 2+ = chiave composta (es. Cognome + Nome)."
        )
        spin_row.addWidget(self._key_spin)
        spin_row.addStretch(1)
        spin_lay.addLayout(spin_row)
        km.addWidget(spin_wrap)
        self._key_map_dynamic = QWidget()
        self._key_map_layout = QVBoxLayout(self._key_map_dynamic)
        self._key_map_layout.setContentsMargins(0, 6, 0, 0)
        self._key_map_layout.setSpacing(12)
        km.addWidget(self._key_map_dynamic)
        lay.addWidget(self._key_map_group)

        self._opt_filter_group = QGroupBox("5. Filtri opzionali")
        of = QVBoxLayout(self._opt_filter_group)
        of.setSpacing(10)
        of.setContentsMargins(4, 8, 4, 4)
        filt_hint = QLabel("Seleziona filtri per ogni file")
        filt_hint.setObjectName("subText")
        filt_hint.setWordWrap(True)
        of.addWidget(filt_hint)
        self._opt_filter_layout = QVBoxLayout()
        of.addLayout(self._opt_filter_layout)
        lay.addWidget(self._opt_filter_group)

        self._key_combos: list[list[QComboBox]] = []
        self._opt_filt_rows_per_store: list[list[tuple[QComboBox, QComboBox]]] = []
        self._opt_filt_rows_layouts: list[QVBoxLayout] = []

        self._multi_op.currentIndexChanged.connect(self._on_multi_op_changed)
        self._key_spin.valueChanged.connect(lambda _v: self._rebuild_key_mapping())
        self._miss_a.currentIndexChanged.connect(lambda _i: self._rebuild_key_mapping())
        self._miss_b.currentIndexChanged.connect(lambda _i: self._rebuild_key_mapping())
        self._on_multi_op_changed()

        lay.addSpacing(16)
        btn = QPushButton("Esegui confronto e apri anteprima…")
        btn.setMinimumHeight(40)
        btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        btn.clicked.connect(self._run_multi)
        lay.addWidget(btn)
        return page

    def _on_single_import_changed(self) -> None:
        imp = self._current_stored_import()
        if imp is None:
            self._filter_block.set_import_outcome(None)
            return
        self._filter_block.set_import_outcome(imp.outcome)
        self._single_col_combo.clear()
        for j, h in enumerate(imp.outcome.headers):
            self._single_col_combo.addItem(h, j)

    def _on_single_op_changed(self) -> None:
        key = self._single_op.currentData()
        need_col = key in ("dup_col", "uniq_col")
        self._single_col_combo.setVisible(need_col)
        self._filter_block.setVisible(key == "filter")

    def _on_multi_op_changed(self) -> None:
        key = self._multi_op.currentData()
        self._missing_row.setVisible(key == "missing")
        self._opt_filter_group.setVisible(key in ("common", "missing"))
        self._update_files_hint()
        if key == "missing":
            self._enforce_max_two_checked_for_missing()
        self._rebuild_key_mapping()

    def _update_files_hint(self) -> None:
        if self._multi_op.currentData() == "missing":
            self._hint_files.setText(
                "Modalità «mancanti»: nell’elenco puoi spuntare al massimo due import. "
                "Il confronto tra due file usa «File A» e «File B» nel punto 3."
            )
        else:
            self._hint_files.setText(
                "Righe comuni: spunta almeno due voci da confrontare insieme."
            )

    def _enforce_max_two_checked_for_missing(self) -> None:
        """In modalità mancanti, non più di due spunte nella lista."""
        if self._multi_op.currentData() != "missing":
            return
        checked: list[QListWidgetItem] = []
        for i in range(self._multi_list.count()):
            it = self._multi_list.item(i)
            if it is not None and it.checkState() == Qt.CheckState.Checked:
                checked.append(it)
        if len(checked) <= 2:
            return
        self._multi_list.blockSignals(True)
        for it in checked[2:]:
            it.setCheckState(Qt.CheckState.Unchecked)
        self._multi_list.blockSignals(False)
        QMessageBox.information(
            self,
            "Confronto",
            "In modalità «mancanti» puoi selezionare al massimo due import nell’elenco: "
            "le spunte in eccesso sono state rimosse.",
        )

    def _on_multi_list_item_changed(self, item: QListWidgetItem) -> None:
        if self._multi_op.currentData() == "missing" and item.checkState() == Qt.CheckState.Checked:
            n = 0
            for i in range(self._multi_list.count()):
                it = self._multi_list.item(i)
                if it is not None and it.checkState() == Qt.CheckState.Checked:
                    n += 1
            if n > 2:
                self._multi_list.blockSignals(True)
                item.setCheckState(Qt.CheckState.Unchecked)
                self._multi_list.blockSignals(False)
                QMessageBox.warning(
                    self,
                    "Confronto",
                    "In modalità «mancanti» puoi selezionare al massimo due import.",
                )
        self._rebuild_key_mapping()

    def _stores_for_key_mapping(self) -> list[StoredImport]:
        mode = self._multi_op.currentData()
        if mode == "missing":
            ia = self._miss_a.currentData()
            ib = self._miss_b.currentData()
            if ia is None or ib is None:
                return []
            a = self._manager.import_by_id(int(ia))
            b = self._manager.import_by_id(int(ib))
            if not a or not b:
                return []
            return [a, b]
        ids = self._selected_multi_ids()
        return [x for x in (self._manager.import_by_id(i) for i in ids) if x is not None]

    def _rebuild_key_mapping(self) -> None:
        while self._key_map_layout.count():
            it = self._key_map_layout.takeAt(0)
            w = it.widget()
            if w is not None:
                w.deleteLater()
        while self._opt_filter_layout.count():
            it = self._opt_filter_layout.takeAt(0)
            w = it.widget()
            if w is not None:
                w.deleteLater()
        self._key_combos = []
        self._opt_filt_rows_per_store = []
        self._opt_filt_rows_layouts = []

        stores = self._stores_for_key_mapping()
        if len(stores) < 2:
            hint = QLabel(
                "Seleziona almeno due import con la spunta, oppure (per «mancanti») due file diversi nelle liste sopra."
            )
            hint.setObjectName("subText")
            hint.setWordWrap(True)
            self._key_map_layout.addWidget(hint)
            _apply_rules_groupbox_title_font(self._key_map_group)
            _apply_rules_groupbox_title_font(self._opt_filter_group)
            return

        K = int(self._key_spin.value())
        for p in range(K):
            part_frame = QFrame()
            part_frame.setObjectName("keyPartBlock")
            pf = QVBoxLayout(part_frame)
            pf.setContentsMargins(10, 10, 10, 10)
            pf.setSpacing(10)
            part_title = QLabel(f"Parte {p + 1} — colonna equivalente per ogni foglio")
            part_title.setObjectName("keyPartTitle")
            part_title.setWordWrap(True)
            pf.addWidget(part_title)
            part_row: list[QComboBox] = []
            for s in stores:
                row_block = QWidget()
                rb = QVBoxLayout(row_block)
                rb.setSpacing(4)
                rb.setContentsMargins(0, 0, 0, 0)
                tag = QLabel(_short_import_caption(s))
                tag.setObjectName("fileTag")
                tag.setWordWrap(True)
                cb = QComboBox()
                cb.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
                for j, h in enumerate(s.outcome.headers):
                    cb.addItem(h, j)
                rb.addWidget(tag)
                rb.addWidget(cb)
                pf.addWidget(row_block)
                part_row.append(cb)
                cb.currentIndexChanged.connect(lambda _i: self._refresh_all_opt_filter_values())
            self._key_combos.append(part_row)
            self._key_map_layout.addWidget(part_frame)

        mode = self._multi_op.currentData()
        if mode in ("common", "missing"):
            for store_idx, s in enumerate(stores):
                block = QFrame()
                block.setObjectName("keyPartBlock")
                bl = QVBoxLayout(block)
                bl.setContentsMargins(10, 10, 10, 10)
                bl.setSpacing(8)
                title = QLabel(_short_import_caption(s))
                title.setObjectName("keyPartTitle")
                title.setWordWrap(True)
                bl.addWidget(title)
                lab_row = QHBoxLayout()
                lc = QLabel("Colonna")
                lc.setObjectName("subText")
                lv = QLabel("Valore")
                lv.setObjectName("subText")
                lab_row.addWidget(lc, 1)
                lab_row.addWidget(lv, 1)
                lab_row.addWidget(QLabel(""))
                bl.addLayout(lab_row)

                rows_host = QWidget()
                rvl = QVBoxLayout(rows_host)
                rvl.setSpacing(8)
                rvl.setContentsMargins(0, 0, 0, 0)
                self._opt_filt_rows_layouts.append(rvl)
                self._opt_filt_rows_per_store.append([])
                bl.addWidget(rows_host)

                add_f = QPushButton("Aggiungi filtro")
                add_f.clicked.connect(lambda _c=False, si=store_idx: self._add_opt_filter_row_for_store(si))
                bl.addWidget(add_f)

                self._opt_filter_layout.addWidget(block)
                self._add_opt_filter_row_for_store(store_idx, run_refresh=False)

            self._refresh_all_opt_filter_values()

        _apply_rules_groupbox_title_font(self._key_map_group)
        _apply_rules_groupbox_title_font(self._opt_filter_group)

    def _collect_key_cols(self) -> list[list[int]]:
        if not self._key_combos:
            return []
        n_stores = len(self._key_combos[0])
        K = len(self._key_combos)
        out: list[list[int]] = []
        for s in range(n_stores):
            cols = [int(self._key_combos[p][s].currentData()) for p in range(K)]
            out.append(cols)
        return out

    def _collect_optional_filters(self) -> list[list[tuple[int, str]]]:
        stores = self._stores_for_key_mapping()
        n = len(stores)
        out: list[list[tuple[int, str]]] = []
        for i in range(n):
            if i >= len(self._opt_filt_rows_per_store):
                out.append([])
                continue
            part: list[tuple[int, str]] = []
            for ccol, val_cb in self._opt_filt_rows_per_store[i]:
                raw = ccol.currentData()
                if raw is None:
                    continue
                v = val_cb.currentText().strip()
                if not v:
                    continue
                part.append((int(raw), v))
            out.append(part)
        return out

    def _refresh_import_lists(self) -> None:
        imports = self._manager.stored_imports()
        self._single_import_combo.clear()
        for s in imports:
            self._single_import_combo.addItem(s.display_label(), s.id)
        self._on_single_import_changed()

        self._multi_list.clear()
        for s in imports:
            it = QListWidgetItem(s.display_label())
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            it.setCheckState(Qt.CheckState.Unchecked)
            it.setData(Qt.ItemDataRole.UserRole, s.id)
            self._multi_list.addItem(it)

        self._miss_a.clear()
        self._miss_b.clear()
        for s in imports:
            self._miss_a.addItem(s.display_label(), s.id)
            self._miss_b.addItem(s.display_label(), s.id)

        self._rebuild_key_mapping()

    def _current_stored_import(self) -> StoredImport | None:
        sid = self._single_import_combo.currentData()
        if sid is None:
            return None
        return self._manager.import_by_id(int(sid))

    def _run_single(self) -> None:
        imp = self._current_stored_import()
        if imp is None:
            QMessageBox.warning(self, "Confronto", "Aggiungi prima un’importazione nella vista Importazione.")
            return
        o = imp.outcome
        headers = list(o.headers)
        rows = [list(r) for r in o.rows]
        key = self._single_op.currentData()
        cf = o.column_formats
        if cf is not None and len(cf) != len(headers):
            cf = None

        try:
            if key == "dup_all":
                h, r = single_duplicates_exact(headers, rows, cf)
            elif key == "dup_col":
                cix = int(self._single_col_combo.currentData())
                h, r = single_duplicates_by_column(headers, rows, cix, cf)
            elif key == "uniq_all":
                h, r = single_unique_exact(headers, rows, cf)
            elif key == "uniq_col":
                cix = int(self._single_col_combo.currentData())
                h, r = single_unique_by_column(headers, rows, cix, cf)
            elif key == "filter":
                flt = self._filter_block.active_filters()
                if not flt:
                    QMessageBox.warning(self, "Confronto", "Imposta almeno un valore di filtro.")
                    return
                h, r = single_filter_fields(headers, rows, flt, cf)
            else:
                return
        except Exception as exc:
            QMessageBox.critical(self, "Confronto", str(exc))
            return

        self._open_preview_and_commit(f"Confronto — {imp.display_label()}", h, r, cf)

    def _selected_multi_ids(self) -> list[int]:
        out: list[int] = []
        for i in range(self._multi_list.count()):
            it = self._multi_list.item(i)
            if it is None or it.checkState() != Qt.CheckState.Checked:
                continue
            out.append(int(it.data(Qt.ItemDataRole.UserRole)))
        return out

    def _run_multi(self) -> None:
        mode = self._multi_op.currentData()
        stores = self._stores_for_key_mapping()
        if len(stores) < 2:
            QMessageBox.warning(
                self,
                "Confronto",
                "Per «righe comuni» spunta almeno due import; per «mancanti» scegli due file diversi nelle combo.",
            )
            return
        if mode == "missing":
            ia = int(self._miss_a.currentData())
            ib = int(self._miss_b.currentData())
            if ia == ib:
                QMessageBox.warning(self, "Confronto", "Scegli due file diversi.")
                return

        key_cols = self._collect_key_cols()
        if not key_cols or any(len(x) == 0 for x in key_cols):
            QMessageBox.warning(self, "Confronto", "Definisci l’associazione dei campi (chiave).")
            return

        labels = [s.display_label() for s in stores]
        outcomes = [s.outcome for s in stores]

        try:
            if mode == "common":
                opt = self._collect_optional_filters()
                if len(opt) != len(stores):
                    opt = [[] for _ in stores]
                h, r = multi_inner_join_wide(labels, outcomes, key_cols, opt)
            elif mode == "missing":
                a, b = stores[0], stores[1]
                ka, kb = key_cols[0], key_cols[1]
                fa_list = self._collect_optional_filters()
                if len(fa_list) >= 2:
                    f_a, f_b = fa_list[0], fa_list[1]
                else:
                    f_a = f_b = []
                h, r = multi_anti_join_wide(
                    a.display_label(),
                    a.outcome,
                    ka,
                    b.display_label(),
                    b.outcome,
                    kb,
                    f_a,
                    f_b,
                )
            else:
                return
        except Exception as exc:
            QMessageBox.critical(self, "Confronto", str(exc))
            return

        title = (
            "Righe comuni (join)"
            if mode == "common"
            else f"Mancanti: {stores[0].display_label()} → {stores[1].display_label()}"
        )
        cf_wide = wide_column_formats_for_tables(outcomes)
        self._open_preview_and_commit(title, h, r, cf_wide)

    def _open_preview_and_commit(
        self,
        title: str,
        headers: list[str],
        rows: list[list[Any]],
        column_formats: list[ColumnFormatSpec] | None = None,
    ) -> None:
        if not rows:
            QMessageBox.information(self, "Confronto", "Nessuna riga nel risultato.")
            return
        dlg = ExcelResultPreviewDialog(
            f"Anteprima — {title}", headers, rows, parent=self._manager, column_formats=column_formats
        )
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        res = dlg.result_table()
        if not res:
            return
        h, r, fmts = res
        self.append_result_block(title, h, r, fmts)

    def append_result_block(
        self,
        title: str,
        headers: list[str],
        rows: list[list[Any]],
        column_formats: list[ColumnFormatSpec] | None = None,
    ) -> None:
        assert self._result_host is not None
        if self._result_host.count() > 0:
            last_item = self._result_host.itemAt(self._result_host.count() - 1)
            if last_item is not None and last_item.spacerItem() is not None:
                self._result_host.takeAt(self._result_host.count() - 1)

        box = QGroupBox(f"Risultato — {title}")
        bl = QVBoxLayout(box)
        head = QHBoxLayout()
        head.addStretch(1)
        exp = QPushButton("Esporta in Excel…")
        exp.clicked.connect(
            lambda _c, h=list(headers), r=[list(x) for x in rows]: self._export_rows(h, r)
        )
        rm = QPushButton("Rimuovi")
        rm.clicked.connect(lambda _checked, w=box: self._remove_result_block(w))
        head.addWidget(exp)
        head.addWidget(rm)
        bl.addLayout(head)
        sub = QLabel(f"{len(rows)} righe × {len(headers)} colonne")
        sub.setObjectName("subText")
        bl.addWidget(sub)
        table = QTableWidget()
        table.setMinimumHeight(220)
        _fill_table(table, headers, rows, column_formats=column_formats)
        bl.addWidget(table)

        rid = self._manager.register_compare_result(title, headers, rows, column_formats)
        setattr(box, "_compare_result_id", rid)

        self._result_host.addWidget(box)
        self._result_host.addStretch(1)

    def _export_rows(
        self,
        headers: list[str],
        rows: list[list[Any]],
        column_formats: list[ColumnFormatSpec] | None = None,
    ) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self._manager,
            "Salva Excel",
            "",
            "Excel (*.xlsx);;Tutti i file (*.*)",
        )
        if not path:
            return
        p = path if str(path).lower().endswith(".xlsx") else f"{path}.xlsx"
        try:
            out_rows: list[list[Any]] = rows
            if column_formats is not None:
                out_rows = format_matrix_strings(rows, len(headers), column_formats)
            export_table_xlsx(p, headers, out_rows)
        except Exception as exc:
            QMessageBox.critical(self, "Esportazione", str(exc))

    def _remove_result_block(self, box: QWidget) -> None:
        assert self._result_host is not None
        rid = getattr(box, "_compare_result_id", None)
        if rid is not None:
            self._manager.unregister_compare_result(int(rid))
        self._result_host.removeWidget(box)
        box.deleteLater()
