from __future__ import annotations

import json
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from PyQt6.QtCore import QDate, QEvent, Qt, QThread, QTime, pyqtSignal
from PyQt6.QtGui import QFont, QGuiApplication, QIcon, QKeySequence
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QStyledItemDelegate,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
    QComboBox,
    QTimeEdit,
    QColorDialog,
    QFileDialog,
    QScrollArea,
    QSpinBox,
)

from app.db_export import (
    EXPORT_FORMAT_CSV,
    EXPORT_FORMAT_JSON,
    EXPORT_FORMAT_XLSX,
    EXPORT_FORMAT_XML,
    EXPORT_JSON_MAGIC,
    import_tables_from_bundle,
    import_tables_from_csv_file,
    import_tables_from_xml,
    list_exportable_tables,
    run_export,
)
from app.italian_holidays import fixed_recurring_italian_holidays
from app.services.packaging_service import PackagingService
from app.ui.package_export_dialog import PackageExportDialog


_DAY_NAMES_FULL = [
    "Lunedì",
    "Martedì",
    "Mercoledì",
    "Giovedì",
    "Venerdì",
    "Sabato",
    "Domenica",
]


def _validate_work_schedule_dict(schedule: dict[int, dict[str, Any]]) -> str | None:
    """Restituisce messaggio errore o None se valido."""
    any_work = any((schedule.get(d) or {}).get("lavorativo") for d in range(7))
    if not any_work:
        return "Imposta almeno un giorno come «Giorno lavorativo»."
    for d in range(7):
        entry = schedule.get(d) or {}
        if not entry.get("lavorativo"):
            continue
        ti = QTime.fromString(str(entry.get("inizio", "")), "HH:mm")
        tf = QTime.fromString(str(entry.get("fine", "")), "HH:mm")
        if not ti.isValid() or not tf.isValid() or ti >= tf:
            return f"{_DAY_NAMES_FULL[d]}: la fine turno deve essere dopo l’inizio turno."
        if entry.get("pausa_abilitata"):
            pi_t = QTime.fromString(str(entry.get("pausa_inizio", "")), "HH:mm")
            pf_t = QTime.fromString(str(entry.get("pausa_fine", "")), "HH:mm")
            if not pi_t.isValid() or not pf_t.isValid() or pi_t >= pf_t:
                return f"{_DAY_NAMES_FULL[d]}: la fine pausa deve essere dopo l’inizio pausa."
    return None


class ToolsWorkScheduleDialog(QDialog):
    """Configurazione per ogni giorno (lun–dom): turno, pausa, giorno lavorativo."""

    def __init__(self, parent: QWidget | None, schedule: dict[int, dict[str, Any]]) -> None:
        super().__init__(parent)
        self.setWindowTitle("Orari e giorni lavorativi")
        self.setModal(True)
        self.setObjectName("workScheduleDialog")
        self.resize(620, 560)
        self._rows: list[dict[str, object]] = []
        self._result_schedule: dict[int, dict[str, Any]] = {}
        self._schedule_in = deepcopy(schedule)

        root = QVBoxLayout(self)
        root.setContentsMargins(18, 18, 18, 16)
        root.setSpacing(14)

        header = QFrame()
        header.setObjectName("workScheduleDialogHeader")
        header_l = QVBoxLayout(header)
        header_l.setContentsMargins(18, 16, 18, 16)
        header_l.setSpacing(6)
        head_title = QLabel("Settimana lavorativa")
        head_title.setObjectName("workScheduleDialogTitle")
        hint = QLabel(
            "Per ogni giorno indica se è lavorativo e, in caso affermativo, orario di turno e pausa pranzo."
        )
        hint.setWordWrap(True)
        hint.setObjectName("workScheduleDialogSubtitle")
        header_l.addWidget(head_title)
        header_l.addWidget(hint)
        root.addWidget(header)

        scroll = QScrollArea()
        scroll.setObjectName("workScheduleDialogScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        inner = QWidget()
        inner.setObjectName("workScheduleDialogInner")
        inner_l = QVBoxLayout(inner)
        inner_l.setSpacing(12)
        inner_l.setContentsMargins(2, 0, 6, 0)

        for i, day_title in enumerate(_DAY_NAMES_FULL):
            card = QFrame()
            card.setObjectName(
                "workScheduleDayCardWeekend" if i >= 5 else "workScheduleDayCard"
            )
            card_l = QVBoxLayout(card)
            card_l.setContentsMargins(16, 14, 16, 14)
            card_l.setSpacing(12)

            top = QHBoxLayout()
            day_lbl = QLabel(day_title)
            day_lbl.setObjectName("workScheduleDayTitle")
            lav = QCheckBox("Giorno lavorativo")
            lav.setObjectName("workScheduleWorkdayCheck")
            lav.setToolTip("Se disattivi, il giorno è considerato non lavorativo (nessun orario applicato).")
            top.addWidget(day_lbl, 0, Qt.AlignmentFlag.AlignLeft)
            top.addStretch(1)
            top.addWidget(lav, 0, Qt.AlignmentFlag.AlignRight)
            card_l.addLayout(top)

            sec_turn = QLabel("Turno")
            sec_turn.setObjectName("workScheduleSectionLabel")
            card_l.addWidget(sec_turn)
            turn_grid = QGridLayout()
            turn_grid.setHorizontalSpacing(14)
            turn_grid.setVerticalSpacing(4)
            li_turn = QLabel("Inizio")
            li_turn.setObjectName("workScheduleFieldLabel")
            lf_turn = QLabel("Fine")
            lf_turn.setObjectName("workScheduleFieldLabel")
            te_in = QTimeEdit()
            te_out = QTimeEdit()
            for te in (te_in, te_out):
                te.setDisplayFormat("HH:mm")
                te.setCalendarPopup(False)
                te.setObjectName("workScheduleTimeEdit")
            turn_grid.addWidget(li_turn, 0, 0)
            turn_grid.addWidget(lf_turn, 0, 1)
            turn_grid.addWidget(te_in, 1, 0)
            turn_grid.addWidget(te_out, 1, 1)
            card_l.addLayout(turn_grid)

            sep = QFrame()
            sep.setObjectName("workScheduleCardSep")
            sep.setFrameShape(QFrame.Shape.HLine)
            sep.setFixedHeight(1)
            card_l.addWidget(sep)

            pause_cb = QCheckBox("Pausa pranzo")
            pause_cb.setObjectName("workSchedulePauseCheck")
            pause_cb.setToolTip("Se attiva, indica fascia oraria della pausa (es. 13:00–14:00).")
            card_l.addWidget(pause_cb)

            pause_grid = QGridLayout()
            pause_grid.setHorizontalSpacing(14)
            pause_grid.setVerticalSpacing(4)
            li_p = QLabel("Inizio pausa")
            li_p.setObjectName("workScheduleFieldLabel")
            lf_p = QLabel("Fine pausa")
            lf_p.setObjectName("workScheduleFieldLabel")
            pte_in = QTimeEdit()
            pte_out = QTimeEdit()
            for te in (pte_in, pte_out):
                te.setDisplayFormat("HH:mm")
                te.setCalendarPopup(False)
                te.setObjectName("workScheduleTimeEdit")
            pause_grid.addWidget(li_p, 0, 0)
            pause_grid.addWidget(lf_p, 0, 1)
            pause_grid.addWidget(pte_in, 1, 0)
            pause_grid.addWidget(pte_out, 1, 1)
            card_l.addLayout(pause_grid)

            inner_l.addWidget(card)
            row: dict[str, object] = {
                "lavorativo": lav,
                "inizio": te_in,
                "fine": te_out,
                "pausa_abilitata": pause_cb,
                "pausa_inizio": pte_in,
                "pausa_fine": pte_out,
            }
            self._rows.append(row)
            lav.toggled.connect(lambda _c, r=row: self._row_update_enabled(r))
            pause_cb.toggled.connect(lambda _c, r=row: self._row_update_enabled(r))

        inner_l.addStretch(1)
        scroll.setWidget(inner)
        root.addWidget(scroll, 1)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.setObjectName("workScheduleDialogButtons")
        ok_btn = buttons.button(QDialogButtonBox.StandardButton.Ok)
        cancel_btn = buttons.button(QDialogButtonBox.StandardButton.Cancel)
        if ok_btn is not None:
            ok_btn.setObjectName("primaryActionButton")
            ok_btn.setText("Salva")
        if cancel_btn is not None:
            cancel_btn.setObjectName("secondaryActionButton")
            cancel_btn.setText("Annulla")
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

        self._apply_work_schedule_dialog_style()
        self._apply_schedule_to_widgets(self._schedule_in)

    def _apply_work_schedule_dialog_style(self) -> None:
        self.setStyleSheet(
            """
            #workScheduleDialog {
                font-family: "Segoe UI", "SF Pro Text", system-ui, sans-serif;
                font-size: 12px;
                background: #f1f5f9;
            }
            #workScheduleDialogHeader {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #0f172a, stop:1 #1e3a8a);
                border-radius: 14px;
                border: none;
            }
            #workScheduleDialogTitle {
                color: #f8fafc;
                font-size: 17px;
                font-weight: 700;
                background: transparent;
            }
            #workScheduleDialogSubtitle {
                color: #cbd5e1;
                font-size: 12px;
                background: transparent;
                line-height: 1.35;
            }
            #workScheduleDialogScroll {
                background: transparent;
                border: none;
            }
            #workScheduleDialogInner {
                background: transparent;
            }
            QFrame#workScheduleDayCard {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
            }
            QFrame#workScheduleDayCardWeekend {
                background: #fffbeb;
                border: 1px solid #fcd34d;
                border-radius: 12px;
            }
            QLabel#workScheduleDayTitle {
                font-size: 14px;
                font-weight: 700;
                color: #0f172a;
                background: transparent;
            }
            QLabel#workScheduleSectionLabel {
                font-size: 10px;
                font-weight: 700;
                color: #64748b;
                letter-spacing: 0.08em;
                text-transform: uppercase;
                background: transparent;
            }
            QLabel#workScheduleFieldLabel {
                font-size: 11px;
                color: #64748b;
                background: transparent;
            }
            QFrame#workScheduleCardSep {
                background: #e2e8f0;
                border: none;
                min-height: 1px;
                max-height: 1px;
            }
            QFrame#workScheduleDayCardWeekend QFrame#workScheduleCardSep {
                background: #fde68a;
            }
            QCheckBox#workScheduleWorkdayCheck {
                font-weight: 600;
                color: #334155;
                spacing: 8px;
            }
            QCheckBox#workSchedulePauseCheck {
                font-weight: 600;
                color: #475569;
                spacing: 8px;
            }
            QTimeEdit#workScheduleTimeEdit {
                padding: 7px 12px;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                background: #ffffff;
                min-width: 96px;
                font-weight: 600;
                color: #0f172a;
            }
            QTimeEdit#workScheduleTimeEdit:focus {
                border-color: #2563eb;
                background: #f8fafc;
            }
            QTimeEdit#workScheduleTimeEdit:disabled {
                background: #f1f5f9;
                color: #94a3b8;
                border-color: #e2e8f0;
            }
            QFrame#workScheduleDayCardWeekend QTimeEdit#workScheduleTimeEdit {
                background: #fffbeb;
            }
            QFrame#workScheduleDayCardWeekend QTimeEdit#workScheduleTimeEdit:focus {
                background: #ffffff;
            }
            QFrame#workScheduleDayCardWeekend QTimeEdit#workScheduleTimeEdit:disabled {
                background: #fef3c7;
            }
            #workScheduleDialog QPushButton#primaryActionButton {
                background: #2563eb;
                color: #ffffff;
                border: 1px solid #1d4ed8;
                border-radius: 8px;
                padding: 8px 20px;
                font-weight: 700;
                min-width: 100px;
            }
            #workScheduleDialog QPushButton#primaryActionButton:hover {
                background: #1d4ed8;
            }
            #workScheduleDialog QPushButton#secondaryActionButton {
                background: #ffffff;
                color: #0f172a;
                border: 1px solid #94a3b8;
                border-radius: 8px;
                padding: 8px 20px;
                font-weight: 600;
                min-width: 100px;
            }
            #workScheduleDialog QPushButton#secondaryActionButton:hover {
                background: #e0f2fe;
                border-color: #0284c7;
            }
            """
        )

    def _row_update_enabled(self, row: dict[str, object]) -> None:
        lav = row["lavorativo"]
        assert isinstance(lav, QCheckBox)
        on = lav.isChecked()
        for key in ("inizio", "fine", "pausa_abilitata"):
            w = row[key]
            assert isinstance(w, (QTimeEdit, QCheckBox))
            w.setEnabled(on)
        pause_on = on
        if on:
            p = row["pausa_abilitata"]
            assert isinstance(p, QCheckBox)
            pause_on = p.isChecked()
        for key in ("pausa_inizio", "pausa_fine"):
            w = row[key]
            assert isinstance(w, QTimeEdit)
            w.setEnabled(pause_on)

    def _apply_schedule_to_widgets(self, sch: dict[int, dict[str, Any]]) -> None:
        for d, row in enumerate(self._rows):
            data = sch.get(d) or {}
            lav = row["lavorativo"]
            assert isinstance(lav, QCheckBox)
            lav.setChecked(bool(data.get("lavorativo", False)))
            te_in = row["inizio"]
            te_out = row["fine"]
            assert isinstance(te_in, QTimeEdit) and isinstance(te_out, QTimeEdit)
            s_in = QTime.fromString(str(data.get("inizio", "09:00")), "HH:mm")
            s_out = QTime.fromString(str(data.get("fine", "18:00")), "HH:mm")
            if not s_in.isValid():
                s_in = QTime(9, 0)
            if not s_out.isValid():
                s_out = QTime(18, 0)
            te_in.setTime(s_in)
            te_out.setTime(s_out)
            p_cb = row["pausa_abilitata"]
            assert isinstance(p_cb, QCheckBox)
            p_cb.setChecked(bool(data.get("pausa_abilitata", True)))
            pi = row["pausa_inizio"]
            pf = row["pausa_fine"]
            assert isinstance(pi, QTimeEdit) and isinstance(pf, QTimeEdit)
            tpi = QTime.fromString(str(data.get("pausa_inizio", "13:00")), "HH:mm")
            tpf = QTime.fromString(str(data.get("pausa_fine", "14:00")), "HH:mm")
            if not tpi.isValid():
                tpi = QTime(13, 0)
            if not tpf.isValid():
                tpf = QTime(14, 0)
            pi.setTime(tpi)
            pf.setTime(tpf)
            self._row_update_enabled(row)

    def _widgets_to_schedule(self) -> dict[int, dict[str, Any]]:
        schedule: dict[int, dict[str, Any]] = {}
        for d, row in enumerate(self._rows):
            lav = row["lavorativo"]
            te_in = row["inizio"]
            te_out = row["fine"]
            p_cb = row["pausa_abilitata"]
            pi = row["pausa_inizio"]
            pf = row["pausa_fine"]
            assert isinstance(lav, QCheckBox)
            assert isinstance(te_in, QTimeEdit) and isinstance(te_out, QTimeEdit)
            assert isinstance(p_cb, QCheckBox)
            assert isinstance(pi, QTimeEdit) and isinstance(pf, QTimeEdit)
            schedule[d] = {
                "lavorativo": lav.isChecked(),
                "inizio": te_in.time().toString("HH:mm"),
                "fine": te_out.time().toString("HH:mm"),
                "pausa_abilitata": p_cb.isChecked(),
                "pausa_inizio": pi.time().toString("HH:mm"),
                "pausa_fine": pf.time().toString("HH:mm"),
            }
        return schedule

    def _on_accept(self) -> None:
        sch = self._widgets_to_schedule()
        err = _validate_work_schedule_dict(sch)
        if err:
            QMessageBox.warning(self, "Orari e giorni lavorativi", err)
            return
        self._result_schedule = sch
        self.accept()

    def result_schedule(self) -> dict[int, dict[str, Any]]:
        return deepcopy(self._result_schedule)


@dataclass
class TableColumn:
    key: str
    label: str
    width: int = 140
    required: bool = False
    editor: str = "text"  # text, combo, multi, bool
    options_loader: Callable[[], list[str]] | None = None
    visible: bool = True


class ComboBoxDelegate(QStyledItemDelegate):
    def __init__(self, page: "EditableTablePage", parent=None) -> None:
        super().__init__(parent)
        self.page = page

    def createEditor(self, parent, option, index):
        col_cfg = self.page.columns[index.column()]
        if col_cfg.editor != "combo":
            return super().createEditor(parent, option, index)

        combo = QComboBox(parent)
        combo.addItem("")
        combo.addItems(self.page.column_options.get(index.column(), []))
        combo.setEditable(False)
        return combo

    def setEditorData(self, editor, index):
        if isinstance(editor, QComboBox):
            value = index.model().data(index, Qt.ItemDataRole.EditRole) or ""
            editor.setCurrentText(str(value))
            return
        super().setEditorData(editor, index)

    def setModelData(self, editor, model, index):
        if isinstance(editor, QComboBox):
            model.setData(index, editor.currentText(), Qt.ItemDataRole.EditRole)
            return
        super().setModelData(editor, model, index)


class MultiSelectDialog(QDialog):
    def __init__(self, title: str, options: list[str], selected: list[str], parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(380, 420)

        layout = QVBoxLayout(self)
        info = QLabel("Seleziona uno o piu valori:")
        layout.addWidget(info)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        layout.addWidget(self.list_widget, 1)

        selected_set = {value.strip() for value in selected if value.strip()}
        for option in options:
            item = QListWidgetItem(option)
            self.list_widget.addItem(item)
            if option in selected_set:
                item.setSelected(True)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def selected_values(self) -> list[str]:
        return [item.text().strip() for item in self.list_widget.selectedItems()]


class EditableTablePage(QWidget):
    def __init__(
        self,
        title: str,
        columns: list[TableColumn],
        load_rows: Callable[[], list[dict]],
        save_row: Callable[[int | None, dict], int],
        delete_row: Callable[[int], None],
        get_help_text: Callable[[], str] | None = None,
        on_change: Callable[[], None] | None = None,
        *,
        embedded: bool = False,
        embedded_toolbar: bool = True,
        show_help: bool = True,
        catalog_context_menu: bool = False,
        catalog_enable_all_edit: Callable[[], None] | None = None,
        use_catalog_styling: bool = False,
    ) -> None:
        super().__init__()
        self.title = title
        self.embedded = embedded
        self.embedded_toolbar = embedded_toolbar
        self.show_help = show_help
        self.catalog_context_menu = catalog_context_menu
        self._catalog_enable_all_edit = catalog_enable_all_edit
        self.use_catalog_styling = use_catalog_styling
        self.columns = columns
        self.load_rows = load_rows
        self.save_row = save_row
        self.delete_row = delete_row
        self.get_help_text = get_help_text
        self.on_change = on_change
        self.edit_mode = False
        self.quick_insert_mode = False
        self.column_options: dict[int, list[str]] = {}

        self._build_ui()
        self.refresh_page()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        if self.embedded:
            layout.setContentsMargins(0, 0, 0, 0)
        else:
            layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        title_lbl = QLabel(self.title)
        if self.embedded:
            title_lbl.setObjectName("subSectionTitle")
            title_lbl.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        else:
            title_lbl.setObjectName("pageTitle")
            title_lbl.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        layout.addWidget(title_lbl)

        self.help_lbl = QLabel("")
        self.help_lbl.setObjectName("helpLabel")
        self.help_lbl.setWordWrap(True)
        if not self.show_help:
            self.help_lbl.hide()
        layout.addWidget(self.help_lbl)

        self.new_btn = None
        self.edit_btn = None
        self.save_btn = None
        self.delete_btn = None
        self.refresh_btn = None
        if self.embedded_toolbar:
            actions = QHBoxLayout()
            actions.setSpacing(8)
            self.new_btn = QPushButton("Nuova riga")
            self.edit_btn = QPushButton("Modifica")
            self.edit_btn.setCheckable(True)
            self.save_btn = QPushButton("Salva")
            self.delete_btn = QPushButton("Elimina")
            self.refresh_btn = QPushButton("Aggiorna")
            self.new_btn.setObjectName("primaryActionButton")
            self.save_btn.setObjectName("primaryActionButton")
            self.delete_btn.setObjectName("dangerActionButton")
            self.refresh_btn.setObjectName("secondaryActionButton")

            actions.addWidget(self.new_btn)
            actions.addWidget(self.edit_btn)
            actions.addWidget(self.save_btn)
            actions.addWidget(self.delete_btn)
            actions.addWidget(self.refresh_btn)
            actions.addStretch()
            layout.addLayout(actions)

        self.table = QTableWidget()
        self.table.setColumnCount(len(self.columns))
        self.table.setHorizontalHeaderLabels([c.label for c in self.columns])
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setStretchLastSection(True)
        for idx, col in enumerate(self.columns):
            self.table.setColumnWidth(idx, col.width)
            if not col.visible:
                self.table.setColumnHidden(idx, True)
        self.table.setItemDelegate(ComboBoxDelegate(self, self.table))
        self.table.installEventFilter(self)
        if self.use_catalog_styling:
            self.table.setObjectName("settingsCatalogDataTable")
            self.table.setShowGrid(False)
            wrap = QFrame()
            wrap.setObjectName("clientDashboardTableWrap")
            wrap.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
            wl = QVBoxLayout(wrap)
            wl.setContentsMargins(0, 0, 0, 0)
            wl.setSpacing(0)
            wl.addWidget(self.table)
            layout.addWidget(wrap, 1)
        else:
            layout.addWidget(self.table, 1)

        if self.catalog_context_menu:
            self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            self.table.customContextMenuRequested.connect(self._on_catalog_context_menu)

        self.status_lbl = QLabel("")
        self.status_lbl.setObjectName("statusLabel")
        layout.addWidget(self.status_lbl)

        if self.embedded_toolbar:
            assert self.new_btn is not None and self.edit_btn is not None
            self.new_btn.clicked.connect(self.add_row)
            self.edit_btn.toggled.connect(self.toggle_edit_mode)
            assert self.save_btn is not None and self.delete_btn is not None and self.refresh_btn is not None
            self.save_btn.clicked.connect(self.save_changes)
            self.delete_btn.clicked.connect(self.delete_selected)
            self.refresh_btn.clicked.connect(self.refresh_page)
        self.table.cellDoubleClicked.connect(self._handle_multi_selector)

    def _first_input_column(self) -> int:
        for idx, col in enumerate(self.columns):
            if col.visible and col.key not in {"id", "code"}:
                return idx
        return 0

    @staticmethod
    def _is_truthy(value: str | int | bool | None) -> bool:
        if isinstance(value, bool):
            return value
        if isinstance(value, int):
            return value == 1
        text = ("" if value is None else str(value)).strip().lower()
        return text in {"1", "true", "si", "yes", "y"}

    def _refresh_lookup_options(self) -> None:
        self.column_options = {}
        for idx, col in enumerate(self.columns):
            if col.options_loader is not None:
                self.column_options[idx] = col.options_loader()

    def refresh_page(self) -> None:
        self._refresh_lookup_options()
        self.table.setRowCount(0)
        rows = self.load_rows()
        for row in rows:
            self._append_row(row)
        self.quick_insert_mode = False
        self._update_help()
        self.status_lbl.setText(f"Righe: {len(rows)}")

    def _append_row(self, row_data: dict) -> None:
        row = self.table.rowCount()
        self.table.insertRow(row)
        for col_idx, col in enumerate(self.columns):
            value = row_data.get(col.key, "")
            if col.editor == "bool":
                item = QTableWidgetItem("")
                item.setCheckState(
                    Qt.CheckState.Checked
                    if self._is_truthy(value)
                    else Qt.CheckState.Unchecked
                )
            else:
                item = QTableWidgetItem("" if value is None else str(value))
            self.table.setItem(row, col_idx, item)
        self._apply_row_flags(row)

    def _apply_row_flags(self, row: int) -> None:
        for col_idx, col in enumerate(self.columns):
            item = self.table.item(row, col_idx)
            if item is None:
                continue

            flags = item.flags()
            flags |= Qt.ItemFlag.ItemIsSelectable

            if col.key in {"id", "code"}:
                flags &= ~Qt.ItemFlag.ItemIsEditable
                flags &= ~Qt.ItemFlag.ItemIsUserCheckable
                item.setFlags(flags)
                continue

            if col.editor == "bool":
                flags &= ~Qt.ItemFlag.ItemIsEditable
                if self.edit_mode:
                    flags |= Qt.ItemFlag.ItemIsUserCheckable
                else:
                    flags &= ~Qt.ItemFlag.ItemIsUserCheckable
                item.setFlags(flags)
                continue

            if col.editor == "multi":
                flags &= ~Qt.ItemFlag.ItemIsEditable
                flags &= ~Qt.ItemFlag.ItemIsUserCheckable
                item.setFlags(flags)
                continue

            if col.editor == "color":
                flags &= ~Qt.ItemFlag.ItemIsEditable
                flags &= ~Qt.ItemFlag.ItemIsUserCheckable
                item.setFlags(flags)
                continue

            flags &= ~Qt.ItemFlag.ItemIsUserCheckable
            if self.edit_mode:
                flags |= Qt.ItemFlag.ItemIsEditable
            else:
                flags &= ~Qt.ItemFlag.ItemIsEditable
            item.setFlags(flags)

    def _update_editability(self) -> None:
        triggers = (
            QAbstractItemView.EditTrigger.AllEditTriggers
            if self.edit_mode
            else QAbstractItemView.EditTrigger.NoEditTriggers
        )
        self.table.setEditTriggers(triggers)
        for row in range(self.table.rowCount()):
            self._apply_row_flags(row)

    def set_edit_mode(self, enabled: bool) -> None:
        self.edit_mode = enabled
        if not enabled:
            self.quick_insert_mode = False
        if self.edit_btn is not None:
            self.edit_btn.blockSignals(True)
            self.edit_btn.setChecked(enabled)
            self.edit_btn.blockSignals(False)
        self._update_editability()
        state = "attiva" if enabled else "disattiva"
        self.status_lbl.setText(f"Modalita modifica {state}.")

    def toggle_edit_mode(self, enabled: bool) -> None:
        self.set_edit_mode(enabled)

    def _on_catalog_context_menu(self, pos) -> None:
        menu = QMenu(self)
        act = menu.addAction("Nuovo")
        act.triggered.connect(self.add_row)
        menu.exec(self.table.mapToGlobal(pos))

    def _handle_multi_selector(self, row: int, column: int) -> None:
        if not self.edit_mode:
            return
        col_cfg = self.columns[column]
        if col_cfg.editor == "color":
            item = self.table.item(row, column)
            current = item.text().strip() if item else ""
            color = QColorDialog.getColor(parent=self)
            if color.isValid():
                value = color.name().lower()
                if item is None:
                    item = QTableWidgetItem(value)
                    self.table.setItem(row, column, item)
                else:
                    item.setText(value)
            return
        if col_cfg.editor != "multi":
            return

        options = self.column_options.get(column, [])
        item = self.table.item(row, column)
        current_text = item.text().strip() if item else ""
        current_values = [part.strip() for part in current_text.split(",") if part.strip()]

        dialog = MultiSelectDialog(col_cfg.label, options, current_values, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            selected_text = ", ".join(dialog.selected_values())
            if item is None:
                item = QTableWidgetItem(selected_text)
                self.table.setItem(row, column, item)
            else:
                item.setText(selected_text)

    def add_row(self, from_arrow: bool = False) -> None:
        if not self.edit_mode:
            if self._catalog_enable_all_edit is not None:
                self._catalog_enable_all_edit()
            else:
                self.set_edit_mode(True)

        if not from_arrow:
            self.quick_insert_mode = True

        self._append_row({column.key: "" for column in self.columns})
        new_row = self.table.rowCount() - 1
        target_col = self.table.currentColumn()
        if (
            target_col < 0
            or target_col >= len(self.columns)
            or self.columns[target_col].key in {"id", "code"}
            or not self.columns[target_col].visible
        ):
            target_col = self._first_input_column()
        self.table.setCurrentCell(new_row, target_col)
        self.table.scrollToBottom()
        if from_arrow:
            self.status_lbl.setText("Riga successiva creata (inserimento rapido).")
        else:
            self.status_lbl.setText("Nuova riga aggiunta. Compila i campi in tabella e premi Salva.")

    def eventFilter(self, watched, event):
        if (
            watched is self.table
            and event.type() == QEvent.Type.KeyPress
            and self.quick_insert_mode
            and self.edit_mode
            and event.key() == Qt.Key.Key_Down
        ):
            current_row = self.table.currentRow()
            if current_row >= 0 and current_row == self.table.rowCount() - 1:
                current_col = self.table.currentColumn()
                self.add_row(from_arrow=True)
                if current_col >= 0 and current_col < len(self.columns):
                    if self.columns[current_col].key not in {"id", "code"} and self.columns[current_col].visible:
                        self.table.setCurrentCell(self.table.rowCount() - 1, current_col)
                return True
        if watched is self.table and event.type() == QEvent.Type.KeyPress:
            if event.matches(QKeySequence.StandardKey.Paste):
                self._paste_from_clipboard()
                return True
        return super().eventFilter(watched, event)

    def _paste_from_clipboard(self) -> None:
        text = QGuiApplication.clipboard().text()
        if not text:
            return
        if not self.edit_mode:
            if self._catalog_enable_all_edit is not None:
                self._catalog_enable_all_edit()
            else:
                self.set_edit_mode(True)

        lines = [line for line in text.splitlines() if line.strip() != ""]
        if not lines:
            return

        editable_cols = [
            idx
            for idx, col in enumerate(self.columns)
            if col.visible and col.key not in {"id", "code"}
        ]
        if not editable_cols:
            return

        start_row = self.table.currentRow()
        if start_row < 0:
            start_row = self.table.rowCount()

        start_col = self.table.currentColumn()
        if start_col < 0 or start_col not in editable_cols:
            start_col = editable_cols[0]

        start_col_index = editable_cols.index(start_col)

        for row_offset, line in enumerate(lines):
            values = line.split("\t")
            row_index = start_row + row_offset
            if row_index >= self.table.rowCount():
                self._append_row({column.key: "" for column in self.columns})

            for col_offset, value in enumerate(values):
                if start_col_index + col_offset >= len(editable_cols):
                    break
                col_idx = editable_cols[start_col_index + col_offset]
                col_cfg = self.columns[col_idx]
                item = self.table.item(row_index, col_idx)
                if col_cfg.editor == "bool":
                    if item is None:
                        item = QTableWidgetItem("")
                        self.table.setItem(row_index, col_idx, item)
                    item.setCheckState(
                        Qt.CheckState.Checked
                        if str(value).strip() in {"1", "si", "true", "yes", "y"}
                        else Qt.CheckState.Unchecked
                    )
                else:
                    if item is None:
                        item = QTableWidgetItem(str(value).strip())
                        self.table.setItem(row_index, col_idx, item)
                    else:
                        item.setText(str(value).strip())

        self.status_lbl.setText("Dati incollati dalla clipboard. Premi Salva.")

    def _validate_lookup_value(self, col_idx: int, value: str) -> None:
        col_cfg = self.columns[col_idx]
        if col_cfg.editor not in {"combo", "multi"}:
            return
        allowed = self.column_options.get(col_idx, [])
        if not allowed:
            return
        if col_cfg.editor == "combo":
            if value and value not in allowed:
                raise ValueError(
                    f"Campo '{col_cfg.label}': valore '{value}' non valido."
                )
            return
        for name in [part.strip() for part in value.split(",") if part.strip()]:
            if name not in allowed:
                raise ValueError(
                    f"Campo '{col_cfg.label}': valore '{name}' non valido."
                )

    def _row_to_payload(self, row_index: int) -> tuple[int | None, dict] | None:
        payload: dict[str, str] = {}
        row_id: int | None = None
        has_data = False

        for col_idx, col in enumerate(self.columns):
            item = self.table.item(row_index, col_idx)
            if col.editor == "bool":
                if item is None:
                    text = "0"
                else:
                    text = "1" if item.checkState() == Qt.CheckState.Checked else "0"
            else:
                text = (item.text() if item else "").strip()
            payload[col.key] = text
            self._validate_lookup_value(col_idx, text)
            if col.key == "id":
                if text:
                    row_id = int(text)
            elif text:
                has_data = True

        if row_id is None and not has_data:
            return None

        for col in self.columns:
            if col.required and col.key != "id" and not payload.get(col.key, "").strip():
                raise ValueError(f"Riga {row_index + 1}: campo '{col.label}' obbligatorio.")

        return row_id, payload

    def save_changes(self) -> None:
        if not self.edit_mode:
            QMessageBox.information(
                self,
                "Modifica non attiva",
                "Attiva Modifica per salvare le variazioni in tabella.",
            )
            return

        saved = 0
        for row in range(self.table.rowCount()):
            try:
                converted = self._row_to_payload(row)
                if converted is None:
                    continue
                row_id, payload = converted
                self.save_row(row_id, payload)
                saved += 1
            except ValueError as exc:
                QMessageBox.warning(self, "Errore salvataggio", str(exc))
                return

        self.refresh_page()
        self.status_lbl.setText(f"Salvate {saved} righe.")
        if self.on_change is not None and saved:
            self.on_change()

    def try_save_all_rows(self) -> tuple[int, int, str | None]:
        """Salva tutte le righe; ritorna (inseriti, aggiornati, errore)."""
        inserted = 0
        updated = 0
        for row in range(self.table.rowCount()):
            try:
                converted = self._row_to_payload(row)
                if converted is None:
                    continue
                row_id, payload = converted
                is_new = row_id is None
                self.save_row(row_id, payload)
                if is_new:
                    inserted += 1
                else:
                    updated += 1
            except ValueError as exc:
                return inserted, updated, str(exc)
        return inserted, updated, None

    def delete_selected(self) -> None:
        selected_rows = sorted(
            {idx.row() for idx in self.table.selectedIndexes()}, reverse=True
        )
        if not selected_rows:
            QMessageBox.information(self, "Elimina", "Seleziona almeno una riga.")
            return

        if (
            QMessageBox.question(
                self,
                "Conferma eliminazione",
                f"Eliminare {len(selected_rows)} righe selezionate?",
            )
            != QMessageBox.StandardButton.Yes
        ):
            return

        deleted = 0
        id_col_index = next(
            (index for index, col in enumerate(self.columns) if col.key == "id"), 0
        )
        for row in selected_rows:
            id_item = self.table.item(row, id_col_index)
            id_text = id_item.text().strip() if id_item else ""
            if id_text:
                try:
                    self.delete_row(int(id_text))
                    deleted += 1
                except ValueError as exc:
                    QMessageBox.warning(self, "Errore eliminazione", str(exc))
                    return
            else:
                self.table.removeRow(row)

        self.refresh_page()
        self.status_lbl.setText(f"Eliminate {deleted} righe.")
        if self.on_change is not None and deleted:
            self.on_change()

    def _update_help(self) -> None:
        if not self.show_help:
            return
        if self.get_help_text is None:
            self.help_lbl.setText("")
            return
        self.help_lbl.setText(self.get_help_text())


class SettingsWindow(QWidget):
    data_changed = pyqtSignal()

    def __init__(self, repository, parent=None) -> None:
        super().__init__()
        self.repository = repository
        # Imposta l'icona anche per la finestra impostazioni (non solo QApplication).
        try:
            from pathlib import Path

            icon_path = Path(__file__).resolve().parent.parent / "assets" / "image.ico"
            if icon_path.exists():
                self.setWindowIcon(QIcon(str(icon_path)))
        except Exception:
            pass
        self.setWindowFlag(Qt.WindowType.Window, True)
        self.setWindowModality(Qt.WindowModality.NonModal)
        self._windows_vpn_names: list[str] = []
        self._vpn_worker: VpnListWorker | None = None
        self._vpn_table_page: EditableTablePage | None = None
        self._vpn_page_index: int | None = None
        self._tools_setup_index: int | None = None
        self._tools_work_schedule: dict[int, dict[str, Any]] = {}
        self._catalog_page_widget: QWidget | None = None
        self._catalog_focused_page: EditableTablePage | None = None
        self._rr_page_widget: QWidget | None = None
        self._rr_focused_page: EditableTablePage | None = None
        self._tag_page_widget: QWidget | None = None
        self._tag_table_page: EditableTablePage | None = None
        self._clients_page_widget: QWidget | None = None
        self._clients_table_page: EditableTablePage | None = None
        self._vpn_page_widget: QWidget | None = None
        self._tag_modify_cb: QCheckBox | None = None
        self._clients_modify_cb: QCheckBox | None = None
        self._vpn_modify_cb: QCheckBox | None = None
        self.setWindowTitle("Impostazioni - HD Manager")
        self.resize(1450, 900)

        self._build_ui()
        self._apply_style()

    def _build_ui(self) -> None:
        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        self.menu = QListWidget()
        self.menu.setObjectName("sideMenu")
        self.menu.setFixedWidth(250)

        self.stack = QStackedWidget()
        self.stack.setObjectName("contentStack")

        root.addWidget(self.menu)
        root.addWidget(self.stack, 1)

        pages: list[tuple[str, QWidget]] = [
            ("Clienti", self._build_clients_page()),
            ("Catalogo prodotti", self._build_product_catalog_page()),
            ("Ruoli, risorse e competenze", self._build_roles_resources_page()),
            ("VPN", self._build_vpns_page()),
            ("Tag", self._build_tags_page()),
            ("Setup Strumenti", self._build_tools_setup_page()),
            ("Pacchetti", self._build_packages_page()),
            ("Informazione Prodotto", self._build_product_info_page()),
        ]

        for index, (title, page) in enumerate(pages):
            self.menu.addItem(QListWidgetItem(title))
            self.stack.addWidget(page)
            if title == "Setup Strumenti":
                self._tools_setup_index = index
            if title == "VPN":
                self._vpn_page_index = index

        self.menu.currentRowChanged.connect(self._show_page)
        self.menu.setCurrentRow(0)
        self._refresh_windows_vpn_async()

    def _show_page(self, index: int) -> None:
        self.stack.setCurrentIndex(index)
        current = self.stack.currentWidget()
        if (
            self._catalog_page_widget is not None
            and current is self._catalog_page_widget
        ):
            self._catalog_env_page.refresh_page()
            self._catalog_rel_page.refresh_page()
            self._catalog_pt_page.refresh_page()
            self._catalog_prod_page.refresh_page()
        elif (
            self._rr_page_widget is not None
            and current is self._rr_page_widget
        ):
            self._rr_comp_page.refresh_page()
            self._rr_roles_page.refresh_page()
            self._rr_res_page.refresh_page()
        elif (
            self._tag_page_widget is not None
            and current is self._tag_page_widget
            and self._tag_table_page is not None
        ):
            self._tag_table_page.refresh_page()
        elif (
            self._clients_page_widget is not None
            and current is self._clients_page_widget
            and self._clients_table_page is not None
        ):
            self._clients_table_page.refresh_page()
        elif (
            self._vpn_page_widget is not None
            and current is self._vpn_page_widget
            and self._vpn_table_page is not None
        ):
            self._vpn_table_page.refresh_page()
        if self._tools_setup_index is not None and index == self._tools_setup_index:
            self._load_tools_setup_page()
        if self._vpn_page_index is not None and index == self._vpn_page_index:
            self._refresh_windows_vpn_async()

    def _settings_service(self):
        repo = self.repository
        return repo.settings if hasattr(repo, "settings") else None

    def _catalog_table_pages(self) -> list[EditableTablePage]:
        return [
            self._catalog_env_page,
            self._catalog_rel_page,
            self._catalog_pt_page,
            self._catalog_prod_page,
        ]

    def _rr_table_pages(self) -> list[EditableTablePage]:
        return [self._rr_comp_page, self._rr_roles_page, self._rr_res_page]

    def _rr_enable_all_edit_mode(self) -> None:
        cb = getattr(self, "_rr_modify_cb", None)
        if cb is None:
            return
        self._catalog_enable_modify_and_edit(cb, self._rr_table_pages())

    def _catalog_enable_all_edit_mode(self) -> None:
        cb = getattr(self, "_catalog_modify_cb", None)
        if cb is None:
            return
        self._catalog_enable_modify_and_edit(cb, self._catalog_table_pages())

    def _on_catalog_edit_toggled(self, checked: bool) -> None:
        self._catalog_edit_pages_toggled(checked, self._catalog_table_pages())

    @staticmethod
    def _format_catalog_section_summary(title: str, inserted: int, updated: int) -> str | None:
        """Una riga per sezione: aggiornati vs aggiunti (stesso schema della vista Clienti)."""
        if inserted == 0 and updated == 0:
            return None
        parts: list[str] = []
        if updated > 0:
            agg = "aggiornato" if updated == 1 else "aggiornati"
            parts.append(f"{updated} record {agg}")
        if inserted > 0:
            ag = "aggiunto" if inserted == 1 else "aggiunti"
            parts.append(f"{inserted} record {ag}")
        return f"{title}: " + ", ".join(parts) + "."

    def _catalog_enable_modify_and_edit(
        self, modify_cb: QCheckBox, pages: list[EditableTablePage]
    ) -> None:
        modify_cb.blockSignals(True)
        modify_cb.setChecked(True)
        modify_cb.blockSignals(False)
        for p in pages:
            p.set_edit_mode(True)

    def _catalog_edit_pages_toggled(
        self, checked: bool, pages: list[EditableTablePage]
    ) -> None:
        for p in pages:
            p.set_edit_mode(checked)

    def _catalog_refresh_pages(self, pages: list[EditableTablePage]) -> None:
        for p in pages:
            p.refresh_page()

    def _catalog_save_pages_generic(
        self,
        modify_cb: QCheckBox,
        pages: list[EditableTablePage],
        dialog_title: str,
    ) -> None:
        if not modify_cb.isChecked():
            QMessageBox.information(
                self,
                dialog_title,
                "Attiva «Modifica unica per tutte le tabelle» per salvare.",
            )
            return
        lines: list[str] = []
        any_change = False
        for p in pages:
            p._refresh_lookup_options()
            ins, upd, err = p.try_save_all_rows()
            if err:
                QMessageBox.warning(
                    self,
                    dialog_title,
                    f"{p.title}: {err}",
                )
                return
            if ins > 0 or upd > 0:
                any_change = True
            row = self._format_catalog_section_summary(p.title, ins, upd)
            if row:
                lines.append(row)
        for p in pages:
            p.refresh_page()
        if any_change:
            self._notify_data_changed()
        msg = "\n".join(lines) if lines else "Nessuna modifica da salvare."
        QMessageBox.information(self, dialog_title, msg)

    def _catalog_delete_single_table(
        self,
        modify_cb: QCheckBox,
        table: EditableTablePage,
        dialog_title: str,
    ) -> None:
        if not modify_cb.isChecked():
            QMessageBox.information(
                self,
                dialog_title,
                "Attiva «Modifica unica per tutte le tabelle» per eliminare righe.",
            )
            return
        table.delete_selected()

    def _on_catalog_save_all(self) -> None:
        self._catalog_save_pages_generic(
            self._catalog_modify_cb,
            self._catalog_table_pages(),
            "Catalogo prodotti",
        )

    def _on_catalog_refresh_all(self) -> None:
        self._catalog_refresh_pages(self._catalog_table_pages())

    def _on_catalog_delete_focused(self) -> None:
        if not self._catalog_modify_cb.isChecked():
            QMessageBox.information(
                self,
                "Catalogo prodotti",
                "Attiva «Modifica unica per tutte le tabelle» per eliminare righe.",
            )
            return
        pg = self._catalog_focused_page
        if pg is None:
            pg = self._catalog_env_page
        pg.delete_selected()

    def _on_rr_edit_toggled(self, checked: bool) -> None:
        self._catalog_edit_pages_toggled(checked, self._rr_table_pages())

    def _on_rr_save_all(self) -> None:
        self._catalog_save_pages_generic(
            self._rr_modify_cb,
            self._rr_table_pages(),
            "Ruoli, risorse e competenze",
        )

    def _on_rr_refresh_all(self) -> None:
        self._catalog_refresh_pages(self._rr_table_pages())

    def _on_rr_delete_focused(self) -> None:
        if not self._rr_modify_cb.isChecked():
            QMessageBox.information(
                self,
                "Ruoli, risorse e competenze",
                "Attiva «Modifica unica per tutte le tabelle» per eliminare righe.",
            )
            return
        pg = self._rr_focused_page
        if pg is None:
            pg = self._rr_comp_page
        pg.delete_selected()

    def _tag_enable_all_edit_mode(self) -> None:
        cb = self._tag_modify_cb
        tp = self._tag_table_page
        if cb is None or tp is None:
            return
        self._catalog_enable_modify_and_edit(cb, [tp])

    def _clients_enable_all_edit_mode(self) -> None:
        cb = self._clients_modify_cb
        tp = self._clients_table_page
        if cb is None or tp is None:
            return
        self._catalog_enable_modify_and_edit(cb, [tp])

    def _vpn_enable_all_edit_mode(self) -> None:
        cb = self._vpn_modify_cb
        tp = self._vpn_table_page
        if cb is None or tp is None:
            return
        self._catalog_enable_modify_and_edit(cb, [tp])

    def _assemble_catalog_toolbar_page(
        self,
        *,
        dialog_title: str,
        main_heading: str,
        intro: str,
        table_page: EditableTablePage,
    ) -> tuple[QWidget, QCheckBox]:
        page = QWidget()
        outer = QVBoxLayout(page)
        outer.setContentsMargins(14, 14, 14, 14)
        outer.setSpacing(12)

        main_title = QLabel(main_heading)
        main_title.setObjectName("pageTitle")
        main_title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        outer.addWidget(main_title)

        intro_lbl = QLabel(intro)
        intro_lbl.setObjectName("helpLabel")
        intro_lbl.setWordWrap(True)
        outer.addWidget(intro_lbl)

        modify_cb = QCheckBox("Modifica unica per tutte le tabelle")
        modify_cb.setToolTip(
            "Se attivo, consente modifiche in tabella."
        )
        save_btn = QPushButton("Salva")
        save_btn.setObjectName("primaryActionButton")
        refresh_btn = QPushButton("Aggiorna")
        refresh_btn.setObjectName("secondaryActionButton")
        delete_btn = QPushButton("Elimina")
        delete_btn.setObjectName("dangerActionButton")

        bar = QHBoxLayout()
        bar.setSpacing(12)
        bar.addWidget(modify_cb)
        bar.addStretch(1)
        bar.addWidget(save_btn)
        bar.addWidget(refresh_btn)
        bar.addWidget(delete_btn)

        pages = [table_page]
        modify_cb.toggled.connect(
            lambda c: self._catalog_edit_pages_toggled(c, pages)
        )
        save_btn.clicked.connect(
            lambda: self._catalog_save_pages_generic(
                modify_cb, pages, dialog_title
            )
        )
        refresh_btn.clicked.connect(
            lambda: self._catalog_refresh_pages(pages)
        )
        delete_btn.clicked.connect(
            lambda: self._catalog_delete_single_table(
                modify_cb, table_page, dialog_title
            )
        )

        outer.addLayout(bar)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        inner = QWidget()
        inner_l = QVBoxLayout(inner)
        inner_l.setContentsMargins(0, 0, 4, 0)
        inner_l.setSpacing(20)
        inner_l.addWidget(table_page)
        scroll.setWidget(inner)
        outer.addWidget(scroll, 1)

        return page, modify_cb

    def eventFilter(self, watched, event) -> bool:  # type: ignore[override]
        if event.type() == QEvent.Type.FocusIn:
            for pg in (
                getattr(self, "_catalog_env_page", None),
                getattr(self, "_catalog_rel_page", None),
                getattr(self, "_catalog_pt_page", None),
                getattr(self, "_catalog_prod_page", None),
            ):
                if pg is not None and watched is pg.table:
                    self._catalog_focused_page = pg
                    break
            for pg in (
                getattr(self, "_rr_comp_page", None),
                getattr(self, "_rr_roles_page", None),
                getattr(self, "_rr_res_page", None),
            ):
                if pg is not None and watched is pg.table:
                    self._rr_focused_page = pg
                    break
        return super().eventFilter(watched, event)

    def _build_product_catalog_page(self) -> QWidget:
        """Unica pagina: toolbar comune; Ambienti e Release affiancati; Tipi; Prodotti."""
        page = QWidget()
        outer = QVBoxLayout(page)
        outer.setContentsMargins(14, 14, 14, 14)
        outer.setSpacing(12)

        main_title = QLabel("Catalogo prodotti")
        main_title.setObjectName("pageTitle")
        main_title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        outer.addWidget(main_title)

        intro = QLabel(
            "In questa pagina puoi creare gli ambienti, le release, i tipi prodotto e confezionare "
            "i prodotti da associare al cliente. Puoi incollare direttamente righe con Ctrl+V."
        )
        intro.setObjectName("helpLabel")
        intro.setWordWrap(True)
        outer.addWidget(intro)

        self._catalog_env_page = self._build_environments_page(catalog_section=True)
        self._catalog_rel_page = self._build_releases_page(catalog_section=True)
        self._catalog_pt_page = self._build_product_types_page(catalog_section=True)
        self._catalog_prod_page = self._build_products_page(catalog_section=True)
        self._catalog_focused_page = self._catalog_env_page

        self._catalog_modify_cb = QCheckBox("Modifica unica per tutte le tabelle")
        self._catalog_modify_cb.setToolTip(
            "Se attivo, consente modifiche in tutte le tabelle del catalogo."
        )
        self._catalog_save_btn = QPushButton("Salva")
        self._catalog_save_btn.setObjectName("primaryActionButton")
        self._catalog_refresh_btn = QPushButton("Aggiorna")
        self._catalog_refresh_btn.setObjectName("secondaryActionButton")
        self._catalog_delete_btn = QPushButton("Elimina")
        self._catalog_delete_btn.setObjectName("dangerActionButton")

        bar = QHBoxLayout()
        bar.setSpacing(12)
        bar.addWidget(self._catalog_modify_cb)
        bar.addStretch(1)
        bar.addWidget(self._catalog_save_btn)
        bar.addWidget(self._catalog_refresh_btn)
        bar.addWidget(self._catalog_delete_btn)

        self._catalog_modify_cb.toggled.connect(self._on_catalog_edit_toggled)
        self._catalog_save_btn.clicked.connect(self._on_catalog_save_all)
        self._catalog_refresh_btn.clicked.connect(self._on_catalog_refresh_all)
        self._catalog_delete_btn.clicked.connect(self._on_catalog_delete_focused)

        for p in self._catalog_table_pages():
            p.table.installEventFilter(self)

        outer.addLayout(bar)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        inner = QWidget()
        inner_l = QVBoxLayout(inner)
        inner_l.setContentsMargins(0, 0, 4, 0)
        inner_l.setSpacing(20)

        row_h = QHBoxLayout()
        row_h.setSpacing(16)
        self._catalog_env_page.setMinimumWidth(360)
        self._catalog_rel_page.setMinimumWidth(360)
        row_h.addWidget(self._catalog_env_page, 1)
        row_h.addWidget(self._catalog_rel_page, 1)
        inner_l.addLayout(row_h)

        sep1 = QFrame()
        sep1.setFrameShape(QFrame.Shape.HLine)
        sep1.setFixedHeight(1)
        sep1.setObjectName("settingsCatalogSep")
        inner_l.addWidget(sep1)

        inner_l.addWidget(self._catalog_pt_page)

        sep2 = QFrame()
        sep2.setFrameShape(QFrame.Shape.HLine)
        sep2.setFixedHeight(1)
        sep2.setObjectName("settingsCatalogSep")
        inner_l.addWidget(sep2)

        inner_l.addWidget(self._catalog_prod_page)

        scroll.setWidget(inner)
        outer.addWidget(scroll, 1)

        self._catalog_page_widget = page
        return page

    def _build_roles_resources_page(self) -> QWidget:
        """Unica pagina: toolbar comune; Competenze, Ruoli e Risorse in verticale (stile catalogo)."""
        page = QWidget()
        outer = QVBoxLayout(page)
        outer.setContentsMargins(14, 14, 14, 14)
        outer.setSpacing(12)

        main_title = QLabel("Ruoli, risorse e competenze")
        main_title.setObjectName("pageTitle")
        main_title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        outer.addWidget(main_title)

        intro = QLabel(
            "Definisci prima le competenze, poi i ruoli e le persone (risorse). "
            "Nelle risorse puoi indicare una competenza (facoltativa). "
            "Puoi incollare righe con Ctrl+V; tasto destro sulla tabella per «Nuovo»."
        )
        intro.setObjectName("helpLabel")
        intro.setWordWrap(True)
        outer.addWidget(intro)

        self._rr_comp_page = self._build_competences_table(catalog_section=True)
        self._rr_roles_page = self._build_roles_page(catalog_section=True)
        self._rr_res_page = self._build_resources_page(catalog_section=True)
        self._rr_focused_page = self._rr_comp_page

        self._rr_modify_cb = QCheckBox("Modifica unica per tutte le tabelle")
        self._rr_modify_cb.setToolTip(
            "Se attivo, consente modifiche nelle tre tabelle (Competenze, Ruoli, Risorse)."
        )
        self._rr_save_btn = QPushButton("Salva")
        self._rr_save_btn.setObjectName("primaryActionButton")
        self._rr_refresh_btn = QPushButton("Aggiorna")
        self._rr_refresh_btn.setObjectName("secondaryActionButton")
        self._rr_delete_btn = QPushButton("Elimina")
        self._rr_delete_btn.setObjectName("dangerActionButton")

        bar = QHBoxLayout()
        bar.setSpacing(12)
        bar.addWidget(self._rr_modify_cb)
        bar.addStretch(1)
        bar.addWidget(self._rr_save_btn)
        bar.addWidget(self._rr_refresh_btn)
        bar.addWidget(self._rr_delete_btn)

        self._rr_modify_cb.toggled.connect(self._on_rr_edit_toggled)
        self._rr_save_btn.clicked.connect(self._on_rr_save_all)
        self._rr_refresh_btn.clicked.connect(self._on_rr_refresh_all)
        self._rr_delete_btn.clicked.connect(self._on_rr_delete_focused)

        for p in self._rr_table_pages():
            p.table.installEventFilter(self)

        outer.addLayout(bar)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        inner = QWidget()
        inner_l = QVBoxLayout(inner)
        inner_l.setContentsMargins(0, 0, 4, 0)
        inner_l.setSpacing(20)

        inner_l.addWidget(self._rr_comp_page)

        sep0 = QFrame()
        sep0.setFrameShape(QFrame.Shape.HLine)
        sep0.setFixedHeight(1)
        sep0.setObjectName("settingsCatalogSep")
        inner_l.addWidget(sep0)

        inner_l.addWidget(self._rr_roles_page)

        sep1 = QFrame()
        sep1.setFrameShape(QFrame.Shape.HLine)
        sep1.setFixedHeight(1)
        sep1.setObjectName("settingsCatalogSep")
        inner_l.addWidget(sep1)

        inner_l.addWidget(self._rr_res_page)

        scroll.setWidget(inner)
        outer.addWidget(scroll, 1)

        self._rr_page_widget = page
        return page

    @staticmethod
    def _tools_apply_standard_button_size(btn: QPushButton) -> None:
        """Pulsanti Setup Strumenti: stessa larghezza e altezza per tutti."""
        btn.setFixedSize(260, 32)

    def _build_tools_setup_page(self) -> QWidget:
        page = QWidget()
        page.setObjectName("toolsSetupPage")
        outer = QVBoxLayout(page)
        outer.setContentsMargins(16, 16, 16, 12)
        outer.setSpacing(0)

        title_lbl = QLabel("Setup Strumenti")
        title_lbl.setObjectName("pageTitle")
        title_lbl.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        outer.addWidget(title_lbl)

        intro_lbl = QLabel(
            "Configura salvataggi, widget in alto, calendario lavorativo e festività per l’agenda."
        )
        intro_lbl.setObjectName("subText")
        intro_lbl.setWordWrap(True)
        outer.addWidget(intro_lbl)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        inner = QWidget()
        inner_l = QVBoxLayout(inner)
        inner_l.setContentsMargins(0, 8, 4, 0)
        inner_l.setSpacing(18)

        dir_group = QGroupBox("Directory salvataggi")
        dir_group.setObjectName("toolsSettingsSection")
        dir_form = QFormLayout(dir_group)
        dir_form.setSpacing(10)
        dir_form.setContentsMargins(12, 14, 12, 12)
        dir_row = QHBoxLayout()
        dir_row.setSpacing(10)
        self._tools_dir_edit = QLineEdit()
        self._tools_dir_edit.setPlaceholderText(
            "Vuoto = cartella predefinita (%LOCALAPPDATA%\\HDManagerDesktop\\calculator_exports)"
        )
        self._tools_dir_edit.setMinimumHeight(32)
        btn_browse = QPushButton("Sfoglia…")
        btn_browse.setObjectName("secondaryActionButton")
        self._tools_apply_standard_button_size(btn_browse)
        dir_row.addWidget(self._tools_dir_edit, 1)
        dir_row.addWidget(btn_browse)
        dir_form.addRow("Cartella predefinita:", dir_row)
        inner_l.addWidget(dir_group)

        widgets_group = QGroupBox("Barra in alto (logo e strumenti)")
        widgets_group.setObjectName("toolsSettingsSection")
        wg_l = QVBoxLayout(widgets_group)
        wg_l.setSpacing(10)
        wg_l.setContentsMargins(12, 14, 12, 12)
        self._tools_agenda_header_cb = QCheckBox("Abilita widget promemoria appuntamenti")
        self._tools_agenda_header_cb.setToolTip(
            "Mostra la barra in alto con i promemoria per oggi e domani (accanto al logo)."
        )
        self._tools_notes_widget_cb = QCheckBox("Abilita widget Note")
        self._tools_notes_widget_cb.setToolTip(
            "Mostra in alto (sopra agli appuntamenti) le note con scadenza; gestiscile anche dalla finestra Note (pulsante in alto a destra)."
        )
        wg_l.addWidget(self._tools_agenda_header_cb)
        wg_l.addWidget(self._tools_notes_widget_cb)
        inner_l.addWidget(widgets_group)

        cal_row = QHBoxLayout()
        cal_row.setSpacing(16)
        sched_outer = QGroupBox("Orario lavorativo")
        sched_outer.setObjectName("toolsSettingsSection")
        sched_outer.setMinimumWidth(340)
        sched_outer.setToolTip(
            "Riepilogo dei giorni lavorativi. Per modificare orari e pausa usa «Configura giorni e turni»."
        )
        sched_outer_l = QVBoxLayout(sched_outer)
        sched_outer_l.setSpacing(10)
        sched_outer_l.setContentsMargins(12, 14, 12, 12)

        sched_table_wrap = QFrame()
        sched_table_wrap.setObjectName("clientDashboardTableWrap")
        sched_table_wrap.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        stw_l = QVBoxLayout(sched_table_wrap)
        stw_l.setContentsMargins(0, 0, 0, 0)
        self._tools_schedule_table = QTableWidget(7, 4)
        self._tools_schedule_table.setObjectName("settingsCatalogDataTable")
        self._tools_schedule_table.setShowGrid(False)
        self._tools_schedule_table.setHorizontalHeaderLabels(
            ["Giorno", "Lavorativo", "Turno", "Pausa"]
        )
        self._tools_schedule_table.verticalHeader().setVisible(False)
        self._tools_schedule_table.horizontalHeader().setStretchLastSection(True)
        self._tools_schedule_table.setEditTriggers(
            QAbstractItemView.EditTrigger.NoEditTriggers
        )
        self._tools_schedule_table.setSelectionMode(
            QAbstractItemView.SelectionMode.NoSelection
        )
        self._tools_schedule_table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self._tools_schedule_table.setMinimumHeight(200)
        stw_l.addWidget(self._tools_schedule_table)
        sched_outer_l.addWidget(sched_table_wrap, 1)

        btn_sched = QPushButton("Configura giorni e turni…")
        btn_sched.setObjectName("secondaryActionButton")
        self._tools_apply_standard_button_size(btn_sched)
        btn_sched.setToolTip(
            "Apre una finestra per impostare, per ogni giorno della settimana, "
            "se è lavorativo e gli orari di turno e pausa."
        )
        btn_sched.clicked.connect(self._tools_open_work_schedule_dialog)
        sched_outer_l.addWidget(btn_sched, alignment=Qt.AlignmentFlag.AlignLeft)

        hol_group = QGroupBox("Festività (agenda)")
        hol_group.setObjectName("toolsSettingsSection")
        hol_group.setMinimumWidth(340)
        hol_l = QVBoxLayout(hol_group)
        hol_l.setSpacing(10)
        hol_l.setContentsMargins(12, 14, 12, 12)
        hol_group.setToolTip(
            "Giorno e mese ricorrenti ogni anno; in agenda sfondo rosso intenso. "
            "Pasqua e Pasquetta sono calcolate automaticamente (non in tabella)."
        )
        btn_it = QPushButton("Aggiungi festività fisse italiane")
        btn_it.setObjectName("secondaryActionButton")
        self._tools_apply_standard_button_size(btn_it)
        btn_it.clicked.connect(self._tools_add_italian_holidays)
        hol_l.addWidget(btn_it, alignment=Qt.AlignmentFlag.AlignLeft)

        table_wrap = QFrame()
        table_wrap.setObjectName("clientDashboardTableWrap")
        table_wrap.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        tw_l = QVBoxLayout(table_wrap)
        tw_l.setContentsMargins(0, 0, 0, 0)
        self._tools_holidays_table = QTableWidget(0, 3)
        self._tools_holidays_table.setObjectName("settingsCatalogDataTable")
        self._tools_holidays_table.setShowGrid(False)
        self._tools_holidays_table.setHorizontalHeaderLabels(["Giorno", "Mese", "Etichetta"])
        self._tools_holidays_table.horizontalHeader().setStretchLastSection(True)
        self._tools_holidays_table.setMinimumHeight(200)
        tw_l.addWidget(self._tools_holidays_table)
        hol_l.addWidget(table_wrap, 1)

        hol_row_btns = QVBoxLayout()
        hol_row_btns.setSpacing(8)
        btn_h_add = QPushButton("Aggiungi riga")
        btn_h_add.setObjectName("secondaryActionButton")
        btn_h_rem = QPushButton("Rimuovi selezionata")
        btn_h_rem.setObjectName("secondaryActionButton")
        self._tools_apply_standard_button_size(btn_h_add)
        self._tools_apply_standard_button_size(btn_h_rem)
        btn_h_add.clicked.connect(self._tools_holiday_add_row)
        btn_h_rem.clicked.connect(self._tools_holiday_remove_row)
        hol_row_btns.addWidget(btn_h_add, alignment=Qt.AlignmentFlag.AlignLeft)
        hol_row_btns.addWidget(btn_h_rem, alignment=Qt.AlignmentFlag.AlignLeft)
        hol_l.addLayout(hol_row_btns)

        cal_row.addWidget(sched_outer, 1)
        cal_row.addWidget(hol_group, 1)
        inner_l.addLayout(cal_row)

        help_lbl = QLabel(
            "<b>Directory</b>: destinazione per «Salva vista» (calcolatrice), "
            "file .txt delle query SQL e altri salvataggi Strumenti (puoi comunque scegliere cartelle diverse dove l’app lo consente). "
            "<b>Orario</b>: modifica solo dalla finestra «Configura giorni e turni». "
            "<b>Festività</b>: giorno/mese e etichetta; Pasqua e Pasquetta sono gestite in automatico. "
            "Le impostazioni sono salvate nel database."
        )
        help_lbl.setObjectName("toolsHelpBlock")
        help_lbl.setWordWrap(True)
        help_lbl.setTextFormat(Qt.TextFormat.RichText)
        help_lbl.setOpenExternalLinks(False)
        inner_l.addWidget(help_lbl)

        scroll.setWidget(inner)
        outer.addWidget(scroll, 1)

        footer = QHBoxLayout()
        footer.setContentsMargins(0, 12, 0, 0)
        footer.addStretch(1)
        save_btn = QPushButton("Salva impostazioni")
        save_btn.setObjectName("primaryActionButton")
        self._tools_apply_standard_button_size(save_btn)
        footer.addWidget(save_btn)
        outer.addLayout(footer)

        btn_browse.clicked.connect(self._tools_browse_calculator_directory)
        save_btn.clicked.connect(self._tools_setup_save)

        self._load_tools_setup_page()
        return page

    def _tools_populate_holidays_table(self, rows: list[dict[str, Any]]) -> None:
        self._tools_holidays_table.setRowCount(0)
        for r in rows:
            i = self._tools_holidays_table.rowCount()
            self._tools_holidays_table.insertRow(i)
            day: int | None = None
            month: int | None = None
            if "month" in r and "day" in r:
                try:
                    month = int(r["month"])
                    day = int(r["day"])
                except (TypeError, ValueError):
                    day = month = None
            if day is None and r.get("date"):
                qd = QDate.fromString(str(r["date"]).strip()[:10], Qt.DateFormat.ISODate)
                if qd.isValid():
                    month = qd.month()
                    day = qd.day()
            if day is None or month is None:
                continue
            self._tools_holidays_table.setItem(i, 0, QTableWidgetItem(str(day)))
            self._tools_holidays_table.setItem(i, 1, QTableWidgetItem(str(month)))
            self._tools_holidays_table.setItem(i, 2, QTableWidgetItem(str(r.get("label", ""))))

    def _tools_holiday_add_row(self) -> None:
        i = self._tools_holidays_table.rowCount()
        self._tools_holidays_table.insertRow(i)
        cd = QDate.currentDate()
        self._tools_holidays_table.setItem(i, 0, QTableWidgetItem(str(cd.day())))
        self._tools_holidays_table.setItem(i, 1, QTableWidgetItem(str(cd.month())))
        self._tools_holidays_table.setItem(i, 2, QTableWidgetItem(""))

    def _tools_holiday_remove_row(self) -> None:
        r = self._tools_holidays_table.currentRow()
        if r >= 0:
            self._tools_holidays_table.removeRow(r)

    def _tools_add_italian_holidays(self) -> None:
        existing: set[tuple[int, int]] = set()
        for rr in range(self._tools_holidays_table.rowCount()):
            it0 = self._tools_holidays_table.item(rr, 0)
            it1 = self._tools_holidays_table.item(rr, 1)
            if it0 and it1:
                try:
                    da = int(it0.text().strip())
                    mo = int(it1.text().strip())
                    existing.add((mo, da))
                except ValueError:
                    pass
        for mo, da, lab in fixed_recurring_italian_holidays():
            if (mo, da) in existing:
                continue
            i = self._tools_holidays_table.rowCount()
            self._tools_holidays_table.insertRow(i)
            self._tools_holidays_table.setItem(i, 0, QTableWidgetItem(str(da)))
            self._tools_holidays_table.setItem(i, 1, QTableWidgetItem(str(mo)))
            self._tools_holidays_table.setItem(i, 2, QTableWidgetItem(lab))
            existing.add((mo, da))

    def _tools_populate_schedule_table(self) -> None:
        """Aggiorna la tabella riepilogo orari (sola lettura, come la tabella festività)."""
        if not hasattr(self, "_tools_schedule_table"):
            return
        short = ["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"]
        sched = self._tools_work_schedule
        self._tools_schedule_table.setRowCount(7)
        for d in range(7):
            row = sched.get(d)
            if row is None:
                row = sched.get(str(d))
            if not isinstance(row, dict):
                row = {}
            if not row.get("lavorativo"):
                turn = "—"
                pause = "—"
                lab = "No"
            else:
                lab = "Sì"
                t0 = str(row.get("inizio", ""))
                t1 = str(row.get("fine", ""))
                turn = f"{t0}–{t1}" if (t0 or t1) else "—"
                if row.get("pausa_abilitata"):
                    p0 = str(row.get("pausa_inizio", ""))
                    p1 = str(row.get("pausa_fine", ""))
                    pause = f"{p0}–{p1}"
                else:
                    pause = "Senza pausa"
            for col, text in enumerate((short[d], lab, turn, pause)):
                it = QTableWidgetItem(text)
                it.setFlags(Qt.ItemFlag.ItemIsEnabled)
                self._tools_schedule_table.setItem(d, col, it)
        self._tools_schedule_table.resizeColumnsToContents()

    def _tools_refresh_schedule_summary(self) -> None:
        self._tools_populate_schedule_table()

    def _tools_open_work_schedule_dialog(self) -> None:
        dlg = ToolsWorkScheduleDialog(self, self._tools_work_schedule)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        self._tools_work_schedule = dlg.result_schedule()
        self._tools_refresh_schedule_summary()

    def _load_tools_setup_page(self) -> None:
        svc = self._settings_service()
        if svc is None:
            self._tools_agenda_header_cb.setChecked(True)
            self._tools_notes_widget_cb.setChecked(False)
            self._tools_work_schedule = {d: {} for d in range(7)}
            self._tools_refresh_schedule_summary()
            return
        self._tools_agenda_header_cb.setChecked(svc.get_agenda_header_widget_enabled())
        self._tools_notes_widget_cb.setChecked(svc.get_notes_widget_enabled())
        self._tools_dir_edit.setText(svc.get_calculator_save_directory())
        hol = svc.get_public_holidays()
        if not hol:
            hol = [
                {"month": mo, "day": da, "label": lab}
                for mo, da, lab in fixed_recurring_italian_holidays()
            ]
        self._tools_populate_holidays_table(hol)
        sch = svc.get_work_schedule()
        self._tools_work_schedule = {}
        for d in range(7):
            row = sch.get(d)
            if row is None:
                row = sch.get(str(d))
            self._tools_work_schedule[d] = dict(row) if isinstance(row, dict) else {}
        self._tools_refresh_schedule_summary()

    def _tools_browse_calculator_directory(self) -> None:
        current = (self._tools_dir_edit.text() or "").strip()
        start = current if current else str(Path.home())
        path = QFileDialog.getExistingDirectory(self, "Directory Salvataggi", start)
        if path:
            self._tools_dir_edit.setText(path)

    def _tools_setup_save(self) -> None:
        svc = self._settings_service()
        if svc is None:
            QMessageBox.warning(self, "Setup Strumenti", "Servizio impostazioni non disponibile.")
            return
        raw_dir = (self._tools_dir_edit.text() or "").strip()
        if raw_dir:
            p = Path(raw_dir).expanduser()
            if not p.exists():
                try:
                    p.mkdir(parents=True, exist_ok=True)
                except OSError as exc:
                    QMessageBox.warning(
                        self,
                        "Setup Strumenti",
                        f"Impossibile creare la cartella:\n{exc}",
                    )
                    return
            if not p.is_dir():
                QMessageBox.warning(self, "Setup Strumenti", "Il percorso non è una cartella valida.")
                return
            svc.set_calculator_save_directory(str(p.resolve()))
        else:
            svc.set_calculator_save_directory("")

        hol_rows: list[dict[str, Any]] = []
        for r in range(self._tools_holidays_table.rowCount()):
            it0 = self._tools_holidays_table.item(r, 0)
            it1 = self._tools_holidays_table.item(r, 1)
            it2 = self._tools_holidays_table.item(r, 2)
            try:
                day = int((it0.text() if it0 else "").strip())
                month = int((it1.text() if it1 else "").strip())
            except ValueError:
                continue
            lab = (it2.text() if it2 else "").strip() or "Festività"
            hol_rows.append({"month": month, "day": day, "label": lab})
        svc.set_public_holidays(hol_rows)

        schedule = deepcopy(self._tools_work_schedule)
        err = _validate_work_schedule_dict(schedule)
        if err:
            QMessageBox.warning(self, "Setup Strumenti", err)
            return
        svc.set_work_schedule(schedule)

        svc.set_agenda_header_widget_enabled(self._tools_agenda_header_cb.isChecked())
        svc.set_notes_widget_enabled(self._tools_notes_widget_cb.isChecked())

        self._notify_data_changed()
        QMessageBox.information(self, "Setup Strumenti", "Impostazioni salvate.")

    # Pages
    def _build_competences_table(self, *, catalog_section: bool = False) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        emb = catalog_section
        return EditableTablePage(
            title="Competenze",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("name", "Nome", 360, required=True),
            ],
            load_rows=lambda: [
                {
                    "id": row["id"],
                    "code": row["code"],
                    "name": row["name"],
                }
                for row in repo.list_competences()
            ],
            save_row=lambda row_id, data: repo.upsert_competence(row_id, data["name"]),
            delete_row=repo.delete_competence,
            get_help_text=lambda: (
                "Elenco delle competenze aziendali (usate anche nei ruoli e, in modo facoltativo, sulle risorse)."
            ),
            on_change=self._notify_data_changed,
            embedded=emb,
            embedded_toolbar=not catalog_section,
            show_help=not catalog_section,
            catalog_context_menu=catalog_section,
            catalog_enable_all_edit=self._rr_enable_all_edit_mode if catalog_section else None,
            use_catalog_styling=catalog_section,
        )

    def _build_product_types_page(self, *, embedded: bool = False, catalog_section: bool = False) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        emb = embedded or catalog_section
        page_title = "Tipi prodotto" if emb else "Setup Tipi Prodotti"
        return EditableTablePage(
            title=page_title,
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("name", "Nome", 260, required=True),
                TableColumn("flag_ip", "Flag IP", 100, editor="bool"),
                TableColumn("flag_host", "Flag HOST", 110, editor="bool"),
                TableColumn(
                    "flag_preconfigured",
                    "Flag PreConfigurata",
                    170,
                    editor="bool",
                ),
                TableColumn("flag_url", "Flag URL", 110, editor="bool"),
                TableColumn("flag_port", "Flag Porta", 110, editor="bool"),
            ],
            load_rows=lambda: [
                {
                    **row,
                    "flag_ip": str(row.get("flag_ip", 0)),
                    "flag_host": str(row.get("flag_host", 0)),
                    "flag_preconfigured": str(row.get("flag_preconfigured", 0)),
                    "flag_url": str(row.get("flag_url", 0)),
                    "flag_port": str(row.get("flag_port", 0)),
                }
                for row in repo.list_product_types()
            ],
            save_row=lambda row_id, data: repo.upsert_product_type(row_id, **data),
            delete_row=repo.delete_product_type,
            get_help_text=lambda: (
                "Inserisci il tipo di prodotto con i flag a seconda della tipologia e l'utilizzo. "
                "Ogni tipo determina le gestione degli accessi."
            ),
            on_change=self._notify_data_changed,
            embedded=emb,
            embedded_toolbar=not catalog_section,
            show_help=not catalog_section,
            catalog_context_menu=catalog_section,
            catalog_enable_all_edit=self._catalog_enable_all_edit_mode if catalog_section else None,
            use_catalog_styling=catalog_section,
        )

    def _build_tags_table(self, *, catalog_section: bool = False) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        emb = catalog_section
        return EditableTablePage(
            title="Tag",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("name", "Nome", 220, required=True),
                TableColumn("color", "Colore", 140, required=True, editor="color"),
                TableColumn(
                    "client_name",
                    "Cliente",
                    220,
                    editor="combo",
                    options_loader=lambda: [
                        row["label"] for row in repo.list_clients_lookup()
                    ],
                ),
            ],
            load_rows=lambda: [
                {
                    "id": row["id"],
                    "code": row["code"],
                    "name": row["name"],
                    "color": row.get("color") or "#0f766e",
                    "client_name": row.get("client_name") or "",
                }
                for row in repo.list_tags()
            ],
            save_row=lambda row_id, data: repo.upsert_tag(row_id, **data),
            delete_row=repo.delete_tag,
            get_help_text=lambda: (
                "Crea i tag da utilizzare e associare per le varie entita."
            ),
            on_change=self._notify_data_changed,
            embedded=emb,
            embedded_toolbar=not catalog_section,
            show_help=not catalog_section,
            catalog_context_menu=catalog_section,
            catalog_enable_all_edit=self._tag_enable_all_edit_mode if catalog_section else None,
            use_catalog_styling=catalog_section,
        )

    def _build_tags_page(self) -> QWidget:
        self._tag_table_page = self._build_tags_table(catalog_section=True)
        page, cb = self._assemble_catalog_toolbar_page(
            dialog_title="Tag",
            main_heading="Tag",
            intro=(
                "Crea i tag da utilizzare e associare per le varie entità. "
                "Tasto destro sulla tabella per «Nuovo»; puoi incollare righe con Ctrl+V."
            ),
            table_page=self._tag_table_page,
        )
        self._tag_modify_cb = cb
        self._tag_page_widget = page
        return page

    def _build_products_page(self, *, embedded: bool = False, catalog_section: bool = False) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        emb = embedded or catalog_section
        return EditableTablePage(
            title="Prodotti",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("name", "Nome", 220, required=True),
                TableColumn(
                    "product_type",
                    "Tipo Prodotto",
                    220,
                    required=True,
                    editor="combo",
                    options_loader=lambda: [row["label"] for row in repo.list_product_types_lookup()],
                ),
                TableColumn(
                    "clients",
                    "Clienti",
                    280,
                    editor="multi",
                    options_loader=lambda: [row["label"] for row in repo.list_clients_lookup()],
                ),
                TableColumn(
                    "environments",
                    "Ambienti",
                    280,
                    editor="multi",
                    options_loader=lambda: [
                        row["label"] for row in repo.list_environments_lookup()
                    ],
                ),
            ],
            load_rows=lambda: [
                {
                    "id": row["id"],
                    "code": row["code"],
                    "name": row["name"],
                    "product_type": row["product_type"],
                    "clients": row.get("clients") or "",
                    "environments": row.get("environments") or "",
                }
                for row in repo.list_products()
            ],
            save_row=lambda row_id, data: repo.upsert_product(row_id, **data),
            delete_row=repo.delete_product,
            get_help_text=lambda: (
                "Gestione dei tuoi prodotti."
            ),
            on_change=self._notify_data_changed,
            embedded=emb,
            embedded_toolbar=not catalog_section,
            show_help=not catalog_section,
            catalog_context_menu=catalog_section,
            catalog_enable_all_edit=self._catalog_enable_all_edit_mode if catalog_section else None,
            use_catalog_styling=catalog_section,
        )

    def _build_environments_page(self, *, embedded: bool = False, catalog_section: bool = False) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        emb = embedded or catalog_section
        return EditableTablePage(
            title="Ambienti",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("name", "Nome", 320, required=True),
            ],
            load_rows=lambda: [
                {"id": row["id"], "code": row["code"], "name": row["name"]}
                for row in repo.list_environments()
            ],
            save_row=lambda row_id, data: repo.upsert_environment(row_id, **data),
            delete_row=repo.delete_environment,
            get_help_text=lambda: (
                "Crea gli ambienti da associare ai prodotti. "
                "Puoi incollare righe da Excel/CSV con Ctrl+V."
            ),
            on_change=self._notify_data_changed,
            embedded=emb,
            embedded_toolbar=not catalog_section,
            show_help=not catalog_section,
            catalog_context_menu=catalog_section,
            catalog_enable_all_edit=self._catalog_enable_all_edit_mode if catalog_section else None,
            use_catalog_styling=catalog_section,
        )

    def _build_releases_page(self, *, embedded: bool = False, catalog_section: bool = False) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        emb = embedded or catalog_section
        return EditableTablePage(
            title="Release",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("name", "Nome", 320, required=True),
            ],
            load_rows=lambda: [
                {"id": row["id"], "code": row["code"], "name": row["name"]}
                for row in repo.list_releases()
            ],
            save_row=lambda row_id, data: repo.upsert_release(row_id, **data),
            delete_row=repo.delete_release,
            get_help_text=lambda: (
                "Crea la lista delle versioni da associare agli ambienti. "
                "Puoi incollare righe da Excel/CSV con Ctrl+V."
            ),
            on_change=self._notify_data_changed,
            embedded=emb,
            embedded_toolbar=not catalog_section,
            show_help=not catalog_section,
            catalog_context_menu=catalog_section,
            catalog_enable_all_edit=self._catalog_enable_all_edit_mode if catalog_section else None,
            use_catalog_styling=catalog_section,
        )

    def _build_clients_table(self, *, catalog_section: bool = False) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        emb = catalog_section
        return EditableTablePage(
            title="Clienti",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("name", "Nome", 220, required=True),
                TableColumn("location", "Localita", 180, required=True),
                TableColumn("link", "Link", 300),
                TableColumn(
                    "resources",
                    "Risorse",
                    340,
                    editor="multi",
                    options_loader=lambda: [
                        row["label"] for row in repo.list_resources_lookup()
                    ],
                ),
            ],
            load_rows=lambda: [
                {
                    "id": row["id"],
                    "code": row["code"],
                    "name": row["name"],
                    "location": row["location"],
                    "link": row.get("link") or "",
                    "resources": row.get("resources") or "",
                }
                for row in repo.list_clients()
            ],
            save_row=lambda row_id, data: repo.upsert_client(row_id, **data),
            delete_row=repo.delete_client,
            get_help_text=lambda: (
                "Creazione e Gestione dei Clienti. "
                "Puoi incollare righe da Excel/CSV con Ctrl+V."
            ),
            on_change=self._notify_data_changed,
            embedded=emb,
            embedded_toolbar=not catalog_section,
            show_help=not catalog_section,
            catalog_context_menu=catalog_section,
            catalog_enable_all_edit=self._clients_enable_all_edit_mode if catalog_section else None,
            use_catalog_styling=catalog_section,
        )

    def _build_clients_page(self) -> QWidget:
        self._clients_table_page = self._build_clients_table(catalog_section=True)
        page, cb = self._assemble_catalog_toolbar_page(
            dialog_title="Clienti",
            main_heading="Clienti",
            intro=(
                "Creazione e gestione dei clienti e delle risorse associate. "
                "Tasto destro sulla tabella per «Nuovo»; puoi incollare righe con Ctrl+V."
            ),
            table_page=self._clients_table_page,
        )
        self._clients_modify_cb = cb
        self._clients_page_widget = page
        return page

    def _build_resources_page(self, *, catalog_section: bool = False) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        emb = catalog_section

        def _competence_options() -> list[str]:
            names = [row["label"] for row in repo.list_competences_lookup()]
            return [""] + names

        return EditableTablePage(
            title="Risorse",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("name", "Nome", 150, required=True),
                TableColumn("surname", "Cognome", 170, required=True),
                TableColumn(
                    "role_name",
                    "Ruolo",
                    200,
                    required=True,
                    editor="combo",
                    options_loader=lambda: [row["label"] for row in repo.list_roles_lookup()],
                ),
                TableColumn(
                    "competence_name",
                    "Competenza",
                    200,
                    required=False,
                    editor="combo",
                    options_loader=_competence_options,
                ),
                TableColumn("phone", "Telefono", 150),
                TableColumn("email", "Email", 240),
                TableColumn("linkedin", "LinkedIn", 220),
                TableColumn("photo_link", "Link foto", 240),
                TableColumn("note", "Note", 200),
            ],
            load_rows=lambda: [
                {
                    "id": row["id"],
                    "code": row["code"],
                    "name": row["name"],
                    "surname": row["surname"],
                    "role_name": row.get("role_name") or "",
                    "competence_name": row.get("competence_name") or "",
                    "phone": row.get("phone") or "",
                    "email": row.get("email") or "",
                    "linkedin": row.get("linkedin") or "",
                    "photo_link": row.get("photo_link") or "",
                    "note": row.get("note") or "",
                }
                for row in repo.list_resources()
            ],
            save_row=lambda row_id, data: repo.upsert_resource(row_id, **data),
            delete_row=repo.delete_resource,
            get_help_text=lambda: "Creazione e Gestione delle risorse della tua azienda.",
            on_change=self._notify_data_changed,
            embedded=emb,
            embedded_toolbar=not catalog_section,
            show_help=not catalog_section,
            catalog_context_menu=catalog_section,
            catalog_enable_all_edit=self._rr_enable_all_edit_mode if catalog_section else None,
            use_catalog_styling=catalog_section,
        )

    def _build_roles_page(self, *, catalog_section: bool = False) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        emb = catalog_section
        return EditableTablePage(
            title="Ruoli",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("name", "Nome", 220, required=True),
                TableColumn(
                    "display_order",
                    "Ordine visualizzazione",
                    190,
                    required=False,
                ),
                TableColumn(
                    "multi_clients",
                    "Più di uno sullo stesso cliente",
                    140,
                    editor="bool",
                ),
            ],
            load_rows=lambda: [
                {
                    "id": row["id"],
                    "code": row["code"],
                    "name": row["name"],
                    "display_order": str(row.get("display_order", "") or ""),
                    "multi_clients": str(row.get("multi_clients", 0)),
                }
                for row in repo.list_roles()
            ],
            save_row=lambda row_id, data: repo.upsert_role(row_id, **data),
            delete_row=repo.delete_role,
            get_help_text=lambda: (
                "Creazioni dei Ruoli da associare alle tuo risorse es PM, Consulenti... "
                "Puoi incollare righe da Excel/CSV con Ctrl+V."
            ),
            on_change=self._notify_data_changed,
            embedded=emb,
            embedded_toolbar=not catalog_section,
            show_help=not catalog_section,
            catalog_context_menu=catalog_section,
            catalog_enable_all_edit=self._rr_enable_all_edit_mode if catalog_section else None,
            use_catalog_styling=catalog_section,
        )

    def _build_vpns_table(self, *, catalog_section: bool = False) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        emb = catalog_section
        return EditableTablePage(
            title="VPN",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("connection_name", "Nome Connessione", 220, required=True),
                TableColumn("server_address", "Indirizzo Server", 220, required=True),
                TableColumn(
                    "vpn_type",
                    "Tipo VPN",
                    150,
                    required=True,
                    editor="combo",
                    options_loader=lambda: ["Vpn Proprietario", "VPN Windows"],
                ),
                TableColumn(
                    "vpn_windows_name",
                    "VPN Windows",
                    200,
                    editor="combo",
                    options_loader=lambda: list(self._windows_vpn_names),
                ),
                TableColumn(
                    "access_info_type",
                    "Tipo Info Accesso",
                    170,
                    required=True,
                    editor="combo",
                    options_loader=lambda: ["Utente/Password", "File Configurato"],
                ),
                TableColumn("username", "Nome Utente", 160, required=True),
                TableColumn("password", "Password", 150, required=True),
                TableColumn("vpn_path", "Percorso VPN", 220),
                TableColumn(
                    "clients",
                    "Clienti",
                    260,
                    editor="multi",
                    options_loader=lambda: [
                        row["label"] for row in repo.list_clients_lookup()
                    ],
                ),
            ],
            load_rows=lambda: [
                {
                    "id": row["id"],
                    "code": row["code"],
                    "connection_name": row["connection_name"],
                    "server_address": row["server_address"],
                    "vpn_type": row["vpn_type"],
                    "vpn_windows_name": row["connection_name"]
                    if row.get("vpn_type") == "VPN Windows"
                    else "",
                    "access_info_type": row.get("access_info_type") or "",
                    "username": row["username"],
                    "password": row["password"],
                    "vpn_path": row.get("vpn_path") or "",
                    "clients": row.get("clients") or "",
                }
                for row in repo.list_vpns()
            ],
            save_row=lambda row_id, data: self._save_vpn_with_clients(repo, row_id, data),
            delete_row=repo.delete_vpn,
            get_help_text=lambda: "Gestione delle VPN da associare al tuo cliente.",
            on_change=self._notify_data_changed,
            embedded=emb,
            embedded_toolbar=not catalog_section,
            show_help=not catalog_section,
            catalog_context_menu=catalog_section,
            catalog_enable_all_edit=self._vpn_enable_all_edit_mode if catalog_section else None,
            use_catalog_styling=catalog_section,
        )

    def _build_vpns_page(self) -> QWidget:
        self._vpn_table_page = self._build_vpns_table(catalog_section=True)
        page, cb = self._assemble_catalog_toolbar_page(
            dialog_title="VPN",
            main_heading="VPN",
            intro=(
                "Gestione delle VPN da associare ai clienti. "
                "Tasto destro sulla tabella per «Nuovo»; puoi incollare righe con Ctrl+V."
            ),
            table_page=self._vpn_table_page,
        )
        self._vpn_modify_cb = cb
        self._vpn_page_widget = page
        return page

    def _build_product_info_page(self) -> QWidget:
        from app.version import (
            PRODUCT_AUTHOR,
            PRODUCT_CONTACT_EMAIL,
            PRODUCT_KIND,
            __version__,
            release_date_display_it,
        )

        widget = QWidget()
        widget.setObjectName("toolsSetupPage")
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(16)

        title = QLabel("Informazione Prodotto")
        title.setObjectName("pageTitle")
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        layout.addWidget(title)

        intro = QLabel(
            "Informazioni sulla versione installata, sull’autore e su come contattarlo per segnalazioni o richieste."
        )
        intro.setWordWrap(True)
        intro.setObjectName("toolsHelpBlock")
        layout.addWidget(intro)

        g_app = QGroupBox("Applicazione")
        f_app = QFormLayout(g_app)
        f_app.setSpacing(10)
        f_app.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        f_app.addRow("Versione:", QLabel(__version__))
        f_app.addRow("Data ultimo rilascio:", QLabel(release_date_display_it()))
        layout.addWidget(g_app)

        g_author = QGroupBox("Autore")
        f_auth = QFormLayout(g_author)
        f_auth.setSpacing(10)
        f_auth.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        f_auth.addRow("Creatore:", QLabel(PRODUCT_AUTHOR))
        f_auth.addRow("Tipo:", QLabel(PRODUCT_KIND))
        layout.addWidget(g_author)

        g_contact = QGroupBox("Contatti")
        v_c = QVBoxLayout(g_contact)
        v_c.setSpacing(8)
        hint = QLabel(
            "Per feedback, supporto o segnalazioni puoi scrivere all’indirizzo qui sotto."
        )
        hint.setWordWrap(True)
        hint.setObjectName("toolsHelpBlock")
        v_c.addWidget(hint)
        mail_lbl = QLabel(
            f'<a href="mailto:{PRODUCT_CONTACT_EMAIL}">{PRODUCT_CONTACT_EMAIL}</a>'
        )
        mail_lbl.setOpenExternalLinks(True)
        mail_lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        v_c.addWidget(mail_lbl)
        layout.addWidget(g_contact)

        layout.addStretch(1)
        return widget

    def _build_packages_page(self) -> QWidget:
        widget = QWidget()
        widget.setObjectName("toolsSetupPage")
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(16)

        title = QLabel("Pacchetti")
        title.setObjectName("pageTitle")
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        layout.addWidget(title)

        intro = QLabel(
            "<b>Esporta</b>: scegli le tabelle e il formato (Excel: un foglio per tabella; "
            "XML; JSON; CSV).<br><br>"
            "<b>Importa</b>: file JSON (pacchetto Core/Risorse/VPN oppure backup completo esportato da qui), "
            "file XML prodotto dall’export, oppure un CSV il cui nome file coincide con il "
            "<i>nome tecnico</i> della tabella nel database (es. <code>clients.csv</code>)."
        )
        intro.setWordWrap(True)
        intro.setObjectName("toolsHelpBlock")
        intro.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(intro)

        base_repo = (
            self.repository.repository
            if hasattr(self.repository, "repository")
            else self.repository
        )
        packaging = PackagingService(base_repo)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)
        btn_export = QPushButton("Esporta")
        btn_export.setObjectName("primaryActionButton")
        btn_import = QPushButton("Importa")
        btn_import.setObjectName("secondaryActionButton")
        self._tools_apply_standard_button_size(btn_export)
        self._tools_apply_standard_button_size(btn_import)
        btn_row.addWidget(btn_export)
        btn_row.addWidget(btn_import)
        btn_row.addStretch(1)
        layout.addLayout(btn_row)

        def on_export() -> None:
            conn = base_repo.connection_factory()
            try:
                tables = list_exportable_tables(conn)
                if not tables:
                    QMessageBox.warning(self, "Esporta", "Nessuna tabella trovata nel database.")
                    return
                dlg = PackageExportDialog(tables, self)
                if dlg.exec() != QDialog.DialogCode.Accepted:
                    return
                sel = dlg.selected_tables()
                fmt = dlg.selected_format()
                if not sel:
                    QMessageBox.warning(self, "Esporta", "Seleziona almeno una tabella.")
                    return
                if fmt == EXPORT_FORMAT_CSV and len(sel) > 1:
                    folder = QFileDialog.getExistingDirectory(
                        self,
                        "Seleziona la cartella per i file CSV",
                    )
                    if not folder:
                        return
                    target = Path(folder)
                else:
                    ext_map = {
                        EXPORT_FORMAT_XLSX: (".xlsx", "Excel (*.xlsx)"),
                        EXPORT_FORMAT_XML: (".xml", "XML (*.xml)"),
                        EXPORT_FORMAT_JSON: (".json", "JSON (*.json)"),
                        EXPORT_FORMAT_CSV: (".csv", "CSV (*.csv)"),
                    }
                    ext, filt = ext_map[fmt]
                    path, _ = QFileDialog.getSaveFileName(
                        self,
                        "Salva file",
                        f"export{ext}",
                        filt,
                    )
                    if not path:
                        return
                    target = Path(path)
                    if target.suffix.lower() != ext.lower():
                        target = target.with_suffix(ext)
                try:
                    run_export(conn=conn, tables=sel, fmt=fmt, target_path=target)
                except Exception as exc:  # pragma: no cover - UI only
                    QMessageBox.critical(self, "Errore esportazione", str(exc))
                    return
                QMessageBox.information(self, "Esporta", f"Salvato in:\n{target}")
            finally:
                conn.close()

        def on_import() -> None:
            path, _ = QFileDialog.getOpenFileName(
                self,
                "Importa dati",
                "",
                "JSON (*.json);;XML (*.xml);;CSV (*.csv);;Tutti i file (*.*)",
            )
            if not path:
                return
            p = Path(path)
            suffix = p.suffix.lower()
            try:
                if suffix == ".json":
                    raw = json.loads(p.read_text(encoding="utf-8"))
                    if not isinstance(raw, dict):
                        raise ValueError("File JSON non valido.")
                    ptype = str(raw.get("package_type") or "").strip()
                    if ptype in ("core", "resources", "vpns"):
                        packaging.import_package(p)
                    elif raw.get("format") == EXPORT_JSON_MAGIC and isinstance(
                        raw.get("tables"), dict
                    ):
                        td = {
                            k: v
                            for k, v in raw["tables"].items()
                            if isinstance(v, list)
                        }
                        conn = base_repo.connection_factory()
                        try:
                            import_tables_from_bundle(conn, td)
                        finally:
                            conn.close()
                    else:
                        QMessageBox.warning(
                            self,
                            "Importa",
                            "Questo JSON non è riconosciuto. Usa un pacchetto Core, Risorse o VPN, "
                            "oppure un backup completo esportato dalla stessa funzione «Esporta».",
                        )
                        return
                elif suffix == ".xml":
                    conn = base_repo.connection_factory()
                    try:
                        import_tables_from_xml(conn, p)
                    finally:
                        conn.close()
                elif suffix == ".csv":
                    conn = base_repo.connection_factory()
                    try:
                        import_tables_from_csv_file(conn, p)
                    finally:
                        conn.close()
                else:
                    QMessageBox.warning(
                        self,
                        "Importa",
                        "Seleziona un file .json, .xml o .csv.",
                    )
                    return
            except Exception as exc:  # pragma: no cover - UI only
                QMessageBox.critical(self, "Errore durante l’importazione", str(exc))
                return
            self._notify_data_changed()
            QMessageBox.information(self, "Importa", "Importazione completata.")

        btn_export.clicked.connect(on_export)
        btn_import.clicked.connect(on_import)

        layout.addStretch(1)
        return widget

    def _refresh_windows_vpn_async(self) -> None:
        if self._vpn_worker is not None and self._vpn_worker.isRunning():
            return
        self._vpn_worker = VpnListWorker(self.repository)
        self._vpn_worker.finished.connect(self._on_windows_vpn_loaded)
        self._vpn_worker.start()

    def _on_windows_vpn_loaded(self, names: list[str]) -> None:
        self._windows_vpn_names = names
        if self._vpn_table_page is not None:
            self._vpn_table_page.refresh_page()

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QWidget {
                font-family: "Segoe UI";
                font-size: 12px;
                color: #1e293b;
                background: #f3f7fa;
            }
            #sideMenu {
                background: #0f172a;
                color: #e2e8f0;
                border: none;
                padding: 8px;
            }
            #sideMenu::item {
                padding: 10px 12px;
                margin: 4px 6px;
                border-radius: 8px;
            }
            #sideMenu::item:selected {
                background: #2563eb;
                color: #ffffff;
            }
            #contentStack {
                background: #ffffff;
                border-left: 1px solid #dbe5ee;
            }
            #pageTitle {
                color: #0f172a;
            }
            #subSectionTitle {
                color: #0f172a;
                font-size: 13px;
                font-weight: 700;
            }
            #settingsCatalogSep {
                background: #e2e8f0;
                border: none;
                max-height: 1px;
            }
            #clientDashboardTableWrap {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 10px;
            }
            #settingsCatalogDataTable {
                background: #ffffff;
                border: none;
                border-radius: 10px;
                gridline-color: transparent;
                outline: none;
                selection-background-color: #ebf3ff;
                selection-color: #001a41;
            }
            #settingsCatalogDataTable::item {
                padding: 8px 10px;
                border-bottom: 1px solid #f1f5f9;
            }
            #settingsCatalogDataTable::item:selected {
                background: #ebf3ff;
                color: #001a41;
            }
            #settingsCatalogDataTable QHeaderView::section {
                background: #f1f5f9;
                color: #475569;
                font-size: 11px;
                font-weight: 700;
                border: none;
                border-bottom: 1px solid #e2e8f0;
                padding: 10px 10px;
            }
            #settingsCatalogDataTable QHeaderView::section:first {
                border-top-left-radius: 9px;
            }
            #settingsCatalogDataTable QHeaderView::section:last {
                border-top-right-radius: 9px;
            }
            #helpLabel {
                color: #0c4a6e;
                background: #e0f2fe;
                border: 1px solid #bae6fd;
                border-radius: 8px;
                padding: 8px;
            }
            QWidget#toolsSetupPage {
                background: #ffffff;
            }
            QGroupBox#toolsSettingsSection {
                border: 1px solid #e2e8f0;
                border-radius: 10px;
                margin-top: 16px;
                padding: 14px 12px 12px 12px;
                background: #f8fafc;
                font-weight: 600;
            }
            QGroupBox#toolsSettingsSection::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 12px;
                padding: 0 6px;
                color: #0f172a;
                font-weight: 700;
                font-size: 13px;
            }
            #toolsHelpBlock {
                color: #0c4a6e;
                background: #e0f2fe;
                border: 1px solid #bae6fd;
                border-radius: 10px;
                padding: 12px 14px;
            }
            QTableWidget {
                background: #ffffff;
                border: 1px solid #cbd5e1;
                gridline-color: #e7edf3;
                selection-background-color: #2563eb;
                selection-color: #ffffff;
            }
            QHeaderView::section {
                background: #0f172a;
                border: 1px solid #0f172a;
                padding: 6px;
                color: #ffffff;
                font-weight: 700;
            }
            QPushButton {
                background: #ffffff;
                border: 1px solid #94a3b8;
                border-radius: 8px;
                padding: 6px 12px;
                font-weight: 600;
                color: #0f172a;
            }
            QPushButton:hover {
                background: #e0f2fe;
                border-color: #0284c7;
            }
            QPushButton:checked {
                background: #2563eb;
                color: #ffffff;
                border: 1px solid #1d4ed8;
            }
            #primaryActionButton {
                background: #2563eb;
                color: #ffffff;
                border: 1px solid #1d4ed8;
                font-weight: 700;
            }
            #primaryActionButton:hover {
                background: #1d4ed8;
                border-color: #1e40af;
            }
            #dangerActionButton {
                background: #dc2626;
                color: #ffffff;
                border: 1px solid #b91c1c;
                font-weight: 700;
            }
            #dangerActionButton:hover {
                background: #b91c1c;
                border-color: #991b1b;
            }
            #secondaryActionButton {
                background: #ffffff;
                color: #0f172a;
                border: 1px solid #94a3b8;
                font-weight: 600;
            }
            #secondaryActionButton:hover {
                background: #e0f2fe;
                border-color: #0284c7;
            }
            #statusLabel {
                color: #334155;
            }
            """
        )

    def _notify_data_changed(self) -> None:
        self.data_changed.emit()

    def _save_vpn_with_clients(self, repo, row_id: int | None, data: dict) -> int:
        """Salva VPN assicurando che il campo Clienti sia sempre passato correttamente."""
        clients_raw = data.get("clients")
        if clients_raw is None:
            clients_val = ""
        elif isinstance(clients_raw, list):
            clients_val = ", ".join(str(x).strip() for x in clients_raw if str(x).strip())
        else:
            clients_val = str(clients_raw).strip() if clients_raw else ""
        payload = dict(data)
        payload["clients"] = clients_val
        return repo.upsert_vpn(row_id, **payload)


class VpnListWorker(QThread):
    finished = pyqtSignal(list)

    def __init__(self, repository, parent=None) -> None:
        super().__init__(parent)
        self.repository = repository

    def run(self) -> None:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        names = repo.list_windows_vpn_connections()
        self.finished.emit(names)
