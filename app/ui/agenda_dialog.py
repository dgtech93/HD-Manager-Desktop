"""Dialog moderno per creazione/modifica impegni agenda."""

from __future__ import annotations

import uuid
from typing import Any

from PyQt6.QtCore import QDate, QDateTime, QSignalBlocker, Qt, QTime
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTimeEdit,
    QVBoxLayout,
    QWidget,
)

from app.agenda_expand import fmt_dt, normalize_item, parse_item_datetime


class AgendaItemDialog(QDialog):
    def __init__(
        self,
        parent: QWidget | None,
        initial: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Nuovo impegno")
        self.setMinimumWidth(480)
        self.resize(500, 560)
        self.setModal(True)

        self.setStyleSheet(
            """
            QDialog { background: #f8fafc; }
            QGroupBox {
                font-weight: 700;
                font-size: 12px;
                color: #0f172a;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
                margin-top: 14px;
                padding: 16px 14px 14px 14px;
                background: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px;
            }
            QLabel#agendaDlgHint {
                color: #64748b;
                font-size: 11px;
            }
            QLineEdit, QComboBox, QDateEdit, QTimeEdit {
                padding: 8px 10px;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                background: #ffffff;
                min-height: 20px;
            }
            QLineEdit:focus, QComboBox:focus, QDateEdit:focus, QTimeEdit:focus {
                border-color: #2563eb;
            }
            """
        )

        self.title_edit = QLineEdit()
        self.title_edit.setPlaceholderText("Titolo dell’impegno")

        self.kind_combo = QComboBox()
        self.kind_combo.addItem("Appuntamento", "appointment")
        self.kind_combo.addItem("Attività da finire", "task")
        self.kind_combo.addItem("Ferie", "vacation")
        self.kind_combo.addItem("Permessi", "leave")

        self.date_from = QDateEdit()
        self.date_to = QDateEdit()
        for de in (self.date_from, self.date_to):
            de.setCalendarPopup(True)
            de.setDisplayFormat("dd/MM/yyyy")
        self.range_cb = QCheckBox("Intervallo su più giorni lavorativi (esclusi giorni festivi impostati)")
        self.range_cb.setChecked(False)

        self.time_start = QTimeEdit()
        self.time_end = QTimeEdit()
        self.time_start.setDisplayFormat("HH:mm")
        self.time_end.setDisplayFormat("HH:mm")

        self.done_cb = QCheckBox("Completata")

        hint = QLabel(
            "In un intervallo, le ore Da–A si applicano a ogni giornata lavorativa «bianca» "
            "(turno in Setup Strumenti, esclusi festivi). Ferie: giornata intera su ogni giorno."
        )
        hint.setObjectName("agendaDlgHint")
        hint.setWordWrap(True)

        root = QVBoxLayout(self)
        root.setSpacing(12)
        root.setContentsMargins(18, 18, 18, 18)

        header = QLabel("Impegno")
        header.setStyleSheet("font-size: 18px; font-weight: 800; color: #0f172a;")
        root.addWidget(header)

        g0 = QGroupBox("Titolo e tipo")
        f0 = QFormLayout(g0)
        f0.setSpacing(10)
        f0.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        f0.addRow("Titolo:", self.title_edit)
        f0.addRow("Tipo:", self.kind_combo)
        root.addWidget(g0)

        g1 = QGroupBox("Periodo")
        f1 = QFormLayout(g1)
        f1.setSpacing(10)
        f1.addRow("Da:", self.date_from)
        f1.addRow("A:", self.date_to)
        f1.addRow("", self.range_cb)
        self._row_time = QWidget()
        row_t = QHBoxLayout(self._row_time)
        row_t.setContentsMargins(0, 0, 0, 0)
        row_t.addWidget(QLabel("Orario"))
        row_t.addWidget(self.time_start)
        row_t.addWidget(QLabel("→"))
        row_t.addWidget(self.time_end)
        f1.addRow("Orario:", self._row_time)
        root.addWidget(g1)

        g2 = QGroupBox("Stato")
        f2 = QVBoxLayout(g2)
        f2.addWidget(self.done_cb)
        root.addWidget(g2)

        root.addWidget(hint)
        root.addStretch(1)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Save).setText("Salva")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Annulla")
        buttons.accepted.connect(self._accept_validate)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

        today = QDate.currentDate()
        if initial:
            it = normalize_item(initial)
            self.setWindowTitle("Modifica impegno")
            self.title_edit.setText(it["title"])
            k = self.kind_combo.findData(it["kind"])
            if k >= 0:
                self.kind_combo.setCurrentIndex(k)
            st = parse_item_datetime(it["start"])
            en = parse_item_datetime(it["end"])
            self.date_from.setDate(st.date())
            md = bool(it.get("multi_day")) or st.date() != en.date()
            self.range_cb.setChecked(md)
            if md:
                self.date_to.setDate(en.date())
            else:
                self.date_to.setDate(st.date())
            self.time_start.setTime(st.time())
            self.time_end.setTime(en.time())
            self.done_cb.setChecked(it["done"])
        else:
            self.date_from.setDate(today)
            self.date_to.setDate(today)
            now_t = QDateTime.currentDateTime().time()
            self.time_start.setTime(now_t)
            self.time_end.setTime(now_t.addSecs(3600))

        self.kind_combo.currentIndexChanged.connect(self._on_kind_changed)
        self.range_cb.toggled.connect(self._on_range_toggled)
        self.date_from.dateChanged.connect(self._on_date_from_changed)
        self.time_start.timeChanged.connect(self._ensure_end_after_start)
        self.date_to.dateChanged.connect(self._ensure_end_after_start)
        self.time_end.timeChanged.connect(self._ensure_end_after_start)

        self._on_kind_changed()
        self._on_range_toggled()
        self._ensure_end_after_start()

    def _on_date_from_changed(self) -> None:
        self._sync_range_from_dates()
        self._ensure_end_after_start()

    def _sync_range_from_dates(self) -> None:
        if not self.range_cb.isChecked():
            self.date_to.setDate(self.date_from.date())
        elif self.date_to.date() < self.date_from.date():
            self.date_to.setDate(self.date_from.date())

    def _ensure_end_after_start(self) -> None:
        """Fine (data+ora) sempre strettamente successiva all’inizio; aggiorna automaticamente se serve."""
        kind = self.kind_combo.currentData()
        if kind in ("vacation", "leave"):
            return
        df = self.date_from.date()
        dt = self.date_to.date()
        ts = QDateTime(df, self.time_start.time())
        te = QDateTime(dt, self.time_end.time())
        if te > ts:
            return
        adj = ts.addSecs(3600)
        with QSignalBlocker(self.date_to), QSignalBlocker(self.time_end):
            self.date_to.setDate(adj.date())
            self.time_end.setTime(adj.time())

    def _on_range_toggled(self) -> None:
        on = self.range_cb.isChecked()
        self.date_to.setEnabled(on)
        if not on:
            self.date_to.setDate(self.date_from.date())
        self._ensure_end_after_start()

    def _on_kind_changed(self) -> None:
        kind = self.kind_combo.currentData()
        is_vac = kind == "vacation"
        is_leave = kind == "leave"
        self.done_cb.setVisible(not is_vac and not is_leave)
        self._row_time.setVisible(not is_vac)
        self._ensure_end_after_start()

    def _accept_validate(self) -> None:
        df = self.date_from.date()
        dt = self.date_to.date()
        if self.range_cb.isChecked() and dt < df:
            QMessageBox.warning(self, "Periodo", "La data «A» non può precedere «Da».")
            return
        self.accept()

    def to_item(self, existing_id: str | None = None) -> dict[str, Any]:
        kind = str(self.kind_combo.currentData())
        df = self.date_from.date()
        dt = self.date_to.date()
        if not self.range_cb.isChecked():
            dt = df
        multi = self.range_cb.isChecked() and df < dt

        if kind == "vacation":
            ts = QDateTime(df, QTime(0, 0))
            te = QDateTime(dt, QTime(23, 59, 59))
        else:
            ts = QDateTime(df, self.time_start.time())
            te = QDateTime(dt, self.time_end.time())
            if te <= ts and not multi:
                te = ts.addSecs(3600)
            if multi and te <= ts:
                te = QDateTime(dt, self.time_end.time())
                if te <= ts:
                    te = ts.addSecs(3600)

        return normalize_item(
            {
                "id": existing_id or str(uuid.uuid4()),
                "title": self.title_edit.text(),
                "kind": kind,
                "start": fmt_dt(ts),
                "end": fmt_dt(te),
                "done": self.done_cb.isChecked() if kind not in ("vacation", "leave") else False,
                "multi_day": multi,
            }
        )
