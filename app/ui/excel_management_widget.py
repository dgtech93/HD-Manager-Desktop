"""Strumenti Excel (linguetta principale): menu laterale + Importazione / Confronto / Concatenazione."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QAbstractItemView,
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
    QScrollArea,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.excel_format import ColumnFormatSpec, formats_or_auto
from app.excel_io import excel_extensions_filter
from app.ui.excel_compare_widget import ExcelCompareWidget
from app.ui.excel_concat_widget import ExcelConcatWidget
from app.ui.excel_import_dialog import (
    ExcelImportDialog,
    ExcelImportOutcome,
    StoredCompareResult,
    StoredImport,
)
from app.ui.ui_constants import SIDE_MENU_TAB_CONTENT_MARGINS, SIDE_MENU_WIDTH_PX


def _fill_preview_table(
    table: QTableWidget,
    headers: list[str],
    rows: list[list],
    column_formats: list[ColumnFormatSpec] | None = None,
) -> None:
    from app.excel_format import format_cell_as_string, formats_or_auto

    table.clear()
    ncols = len(headers)
    fmts = formats_or_auto(ncols, column_formats)
    table.setColumnCount(ncols)
    table.setHorizontalHeaderLabels(headers)
    table.setRowCount(len(rows))
    table.setAlternatingRowColors(True)
    for r, row in enumerate(rows):
        for c in range(ncols):
            val = row[c] if c < len(row) else None
            txt = format_cell_as_string(val, fmts[c])
            item = QTableWidgetItem(txt)
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(r, c, item)


class ImportPreviewBox(QGroupBox):
    """Riquadro anteprima import con id stabile."""

    def __init__(self, import_id: int, title: str, parent: QWidget | None = None) -> None:
        super().__init__(title, parent)
        self.setObjectName("excelImportPreviewGroup")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.import_id = import_id
        self._sub_label: QLabel | None = None
        self._table: QTableWidget | None = None

    def set_outcome(self, outcome: ExcelImportOutcome) -> None:
        name = Path(outcome.source_path).name
        self.setTitle(f"{name} — foglio «{outcome.sheet_name}»")
        if self._sub_label is not None:
            self._sub_label.setText(f"{len(outcome.rows)} righe × {len(outcome.headers)} colonne")
        if self._table is not None:
            _fill_preview_table(self._table, outcome.headers, outcome.rows, outcome.column_formats)


class ExcelManagementWidget(QWidget):
    """Area principale Strumenti Excel: menu laterale e stack di viste."""

    def __init__(self, main_window: QWidget | None = None) -> None:
        super().__init__()
        self.setObjectName("excelTabPage")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self._main_window = main_window
        self._preview_host: QVBoxLayout | None = None
        self._imports: list[StoredImport] = []
        self._next_import_id = 1
        self._compare_results: list[StoredCompareResult] = []
        self._next_compare_result_id = 1

        root = QHBoxLayout(self)
        root.setContentsMargins(*SIDE_MENU_TAB_CONTENT_MARGINS)
        root.setSpacing(0)

        self._excel_side_menu = QListWidget()
        self._excel_side_menu.setObjectName("excelSideMenu")
        self._excel_side_menu.setFixedWidth(SIDE_MENU_WIDTH_PX)
        self._excel_side_menu.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._excel_side_menu.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)

        self._excel_stack = QStackedWidget()
        self._excel_stack.setObjectName("excelToolsStack")

        self._excel_stack.addWidget(self._build_import_tab())
        self._compare_widget = ExcelCompareWidget(self)
        self._excel_stack.addWidget(self._compare_widget)
        self._concat_widget = ExcelConcatWidget(self)
        self._excel_stack.addWidget(self._concat_widget)

        it_imp = QListWidgetItem("Importazione")
        it_cmp = QListWidgetItem("Confronto")
        it_cat = QListWidgetItem("Concatenazione")
        self._excel_side_menu.addItem(it_imp)
        self._excel_side_menu.addItem(it_cmp)
        self._excel_side_menu.addItem(it_cat)
        it_cmp.setHidden(True)
        it_cat.setHidden(True)

        self._excel_side_menu.currentRowChanged.connect(self._on_excel_menu_row_changed)

        root.addWidget(self._excel_side_menu)
        root.addWidget(self._excel_stack, 1)

        self._excel_side_menu.setCurrentRow(0)
        self._sync_excel_submenu_state()

    def stored_imports(self) -> list[StoredImport]:
        return list(self._imports)

    def import_by_id(self, sid: int) -> StoredImport | None:
        for s in self._imports:
            if s.id == sid:
                return s
        return None

    def compare_results(self) -> list[StoredCompareResult]:
        return list(self._compare_results)

    def register_compare_result(
        self,
        title: str,
        headers: list[str],
        rows: list[list[Any]],
        column_formats: list[ColumnFormatSpec] | None = None,
    ) -> int:
        rid = self._next_compare_result_id
        self._next_compare_result_id += 1
        n = len(headers)
        cf_aligned = formats_or_auto(n, column_formats)
        self._compare_results.append(
            StoredCompareResult(
                id=rid,
                title=title,
                headers=list(headers),
                rows=[list(r) for r in rows],
                column_formats=cf_aligned,
            )
        )
        return rid

    def unregister_compare_result(self, result_id: int) -> None:
        self._compare_results = [r for r in self._compare_results if r.id != result_id]

    def resolve_table_source(self, kind: str, source_id: int) -> tuple[list[str], list[list[Any]]] | None:
        if kind == "import":
            s = self.import_by_id(source_id)
            if s is None:
                return None
            o = s.outcome
            return o.headers, o.rows
        if kind == "compare":
            for r in self._compare_results:
                if r.id == source_id:
                    return r.headers, r.rows
        return None

    def resolve_column_formats(self, kind: str, source_id: int) -> list[ColumnFormatSpec]:
        """Formati colonna allineati a ``len(headers)`` della stessa fonte (mai lista «spezzata»)."""
        t = self.resolve_table_source(kind, source_id)
        if t is None:
            return []
        headers, _ = t
        n = len(headers)
        raw: list[ColumnFormatSpec] | None = None
        if kind == "import":
            s = self.import_by_id(source_id)
            if s is not None:
                raw = s.outcome.column_formats
        elif kind == "compare":
            for r in self._compare_results:
                if r.id == source_id:
                    raw = r.column_formats
                    break
        return formats_or_auto(n, raw)

    def _placeholder(self, title: str, hint: str) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(16, 16, 16, 16)
        t = QLabel(title)
        t.setObjectName("sectionTitle")
        h = QLabel(hint)
        h.setObjectName("subText")
        h.setWordWrap(True)
        lay.addWidget(t)
        lay.addWidget(h)
        lay.addStretch(1)
        return w

    def _build_import_tab(self) -> QWidget:
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(16, 16, 16, 16)
        lay.setSpacing(14)

        page_title = QLabel("Importazione")
        page_title.setObjectName("subSectionTitle")
        lay.addWidget(page_title)

        card = QFrame()
        card.setObjectName("clientDashboardCard")
        card.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        cl = QVBoxLayout(card)
        cl.setContentsMargins(14, 14, 14, 14)
        cl.setSpacing(12)

        bar = QHBoxLayout()
        self._import_btn = QPushButton("Importa…")
        self._import_btn.setObjectName("archiveActionButton")
        self._import_btn.setToolTip("Seleziona uno o più file Excel e configura colonne e righe")
        self._import_btn.clicked.connect(self._on_import_clicked)
        bar.addWidget(self._import_btn)
        bar.addStretch(1)
        cl.addLayout(bar)

        hint = QLabel(
            "Ogni file si apre in una finestra di precaricamento: scegli foglio, colonne, ordine e righe. "
            "Dopo l’OK puoi usare «Modifica…» sulla singola anteprima per cambiare scelte e formati."
        )
        hint.setObjectName("accessProductHint")
        hint.setWordWrap(True)
        cl.addWidget(hint)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        inner = QWidget()
        self._preview_host = QVBoxLayout(inner)
        self._preview_host.setContentsMargins(0, 0, 4, 0)
        self._preview_host.addStretch(1)
        scroll.setWidget(inner)
        cl.addWidget(scroll, 1)

        lay.addWidget(card, 1)

        return page

    def _on_excel_menu_row_changed(self, row: int) -> None:
        if row < 0 or row >= self._excel_stack.count():
            return
        self._excel_stack.setCurrentIndex(row)

    def _sync_excel_submenu_state(self) -> None:
        has_imports = len(self._imports) > 0
        if not has_imports and self._excel_stack.currentIndex() != 0:
            self._excel_side_menu.blockSignals(True)
            self._excel_side_menu.setCurrentRow(0)
            self._excel_stack.setCurrentIndex(0)
            self._excel_side_menu.blockSignals(False)
        it_cmp = self._excel_side_menu.item(1)
        it_cat = self._excel_side_menu.item(2)
        if it_cmp is not None:
            it_cmp.setHidden(not has_imports)
        if it_cat is not None:
            it_cat.setHidden(not has_imports)
        self._compare_widget.refresh_imports_from_manager()
        self._concat_widget.refresh_imports_from_manager()

    def _on_import_clicked(self) -> None:
        parent = self._main_window or self
        paths, _ = QFileDialog.getOpenFileNames(
            parent,
            "Seleziona file Excel",
            "",
            excel_extensions_filter(),
        )
        if not paths:
            return
        for path in paths:
            dlg = ExcelImportDialog(path, parent=self._main_window or self)
            if dlg.exec() != QDialog.DialogCode.Accepted:
                continue
            out = dlg.outcome()
            if out is None:
                continue
            self._append_preview(out)

    def _append_preview(self, outcome: ExcelImportOutcome) -> None:
        assert self._preview_host is not None
        if self._preview_host.count() > 0:
            last_item = self._preview_host.itemAt(self._preview_host.count() - 1)
            if last_item is not None and last_item.spacerItem() is not None:
                self._preview_host.takeAt(self._preview_host.count() - 1)

        sid = self._next_import_id
        self._next_import_id += 1
        self._imports.append(StoredImport(id=sid, outcome=outcome))

        name = Path(outcome.source_path).name
        box = ImportPreviewBox(sid, f"{name} — foglio «{outcome.sheet_name}»")
        bl = QVBoxLayout(box)
        head = QHBoxLayout()
        head.addStretch(1)
        ed = QPushButton("Modifica…")
        ed.setObjectName("archiveActionButton")
        ed.setToolTip("Riapri la finestra di importazione per cambiare foglio, colonne, righe e formati")
        ed.clicked.connect(lambda _checked=False, w=box: self._on_edit_import(w))
        head.addWidget(ed)
        rm = QPushButton("Rimuovi")
        rm.setObjectName("archiveActionButton")
        rm.setToolTip("Elimina questa anteprima dall’elenco")
        rm.clicked.connect(lambda _checked, w=box: self._remove_preview(w))
        head.addWidget(rm)
        bl.addLayout(head)
        sub = QLabel(f"{len(outcome.rows)} righe × {len(outcome.headers)} colonne")
        sub.setObjectName("subText")
        box._sub_label = sub
        bl.addWidget(sub)
        table = QTableWidget()
        table.setMinimumHeight(220)
        box._table = table
        _fill_preview_table(table, outcome.headers, outcome.rows, outcome.column_formats)
        bl.addWidget(table)

        self._preview_host.addWidget(box)
        self._preview_host.addStretch(1)
        self._sync_excel_submenu_state()

    def _on_edit_import(self, box: ImportPreviewBox) -> None:
        imp = self.import_by_id(box.import_id)
        if imp is None:
            return
        path = imp.outcome.source_path
        if not Path(path).exists():
            QMessageBox.warning(
                self,
                "Importazione",
                "Il file originale non è più nel percorso salvato. Usa «Importa…» per selezionarlo di nuovo.",
            )
            return
        dlg = ExcelImportDialog(
            path,
            parent=self._main_window or self,
            initial_outcome=imp.outcome,
        )
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        new_out = dlg.outcome()
        if new_out is None:
            return
        self._imports = [
            StoredImport(id=s.id, outcome=new_out) if s.id == box.import_id else s for s in self._imports
        ]
        box.set_outcome(new_out)

    def _remove_preview(self, box: ImportPreviewBox) -> None:
        assert self._preview_host is not None
        self._imports = [s for s in self._imports if s.id != box.import_id]
        self._preview_host.removeWidget(box)
        box.deleteLater()
        self._sync_excel_submenu_state()
