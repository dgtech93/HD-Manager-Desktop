"""Agenda (finestra dedicata) + widget promemoria nell’header."""

from __future__ import annotations

import uuid
from typing import Any

from PyQt6.QtCore import QDate, QDateTime, Qt, QTime, QTimer
from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QFrame,
    QGraphicsDropShadowEffect,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.agenda_expand import (
    expand_agenda_items,
    fmt_dt as _fmt_dt,
    holiday_month_day_set_from_rows,
    normalize_item as _normalize_item,
    parse_item_datetime as _parse_item_datetime,
)
from app.agenda_schedule import is_public_holiday, is_working_day
from app.ui.agenda_dialog import AgendaItemDialog
from app.ui.agenda_month_cell import MonthDayCellWidget
from app.ui.agenda_time_grid import (
    build_day_timeline,
    build_week_timeline,
    schedule_scroll_to_first_work,
    wire_column_signals,
)


def _events_for_day(items: list[dict[str, Any]], qd: QDate) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for it in items:
        st = _parse_item_datetime(it["start"])
        if st.isValid() and st.date() == qd:
            out.append(it)
    out.sort(key=lambda x: _parse_item_datetime(str(x["start"])))
    return out


def _short_title(s: str, n: int = 18) -> str:
    t = (s or "").strip()
    return t if len(t) <= n else t[: n - 1] + "…"


class AgendaUpcomingHeaderWidget(QWidget):
    """Promemoria oggi/domani: scheda + elenco (tipo, titolo, data/ora, stato)."""

    def __init__(self, main_window: QWidget, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._mw = main_window
        self._root = QHBoxLayout(self)
        self._root.setContentsMargins(4, 2, 8, 4)
        self._root.setSpacing(0)
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        self.setMinimumSize(220, 142)
        self.setMaximumHeight(142)
        self.setMaximumWidth(320)
        self.setObjectName("agendaUpcomingHeader")
        # Tick periodico: countdown, «In corso» / scomparsa a fine slot senza dipendere dalla finestra Agenda.
        self._status_timer = QTimer(self)
        self._status_timer.setInterval(15_000)
        self._status_timer.timeout.connect(self.refresh)
        self._status_timer.start()
        self._deadline_timer = QTimer(self)
        self._deadline_timer.setSingleShot(True)
        self._deadline_timer.timeout.connect(self.refresh)
        self.refresh()

    @staticmethod
    def _kind_caption(kind: str) -> str:
        k = str(kind or "appointment")
        if k == "task":
            return "ATTIVITÀ"
        if k == "vacation":
            return "FERIE"
        if k == "leave":
            return "PERMESSO"
        return "APPUNT."

    @staticmethod
    def _day_caption(st: QDateTime, today: QDate, tomorrow: QDate) -> str:
        d = st.date()
        if d == today:
            return "Oggi"
        if d == tomorrow:
            return "Domani"
        return st.toString("dd/MM")

    @staticmethod
    def _status_caption(st: QDateTime, it: dict[str, Any], now: QDateTime) -> str:
        en = _parse_item_datetime(str(it.get("end", "")))
        if now < st:
            secs = now.secsTo(st)
            h, r = divmod(secs, 3600)
            m, s = divmod(r, 60)
            if h > 0:
                return f"Tra {h}h {m}m"
            if m > 0:
                return f"Tra {m} min"
            return f"Tra {s}s"
        if en.isValid() and now < en:
            return "In corso"
        return "Terminato"

    @staticmethod
    def _status_kind(st: QDateTime, it: dict[str, Any], now: QDateTime) -> str:
        """Per colore stato: upcoming | ongoing | ended."""
        en = _parse_item_datetime(str(it.get("end", "")))
        if now < st:
            return "upcoming"
        if en.isValid() and now < en:
            return "ongoing"
        return "ended"

    def _settings(self):
        repo = getattr(self._mw, "repository", None)
        if repo is not None and hasattr(repo, "settings"):
            return repo.settings
        return None

    def _arm_next_deadline(self, items: list[dict[str, Any]], now: QDateTime) -> None:
        self._deadline_timer.stop()
        today = QDate.currentDate()
        t1 = today.addDays(1)
        deltas: list[int] = []
        for it in items:
            if it.get("done"):
                continue
            st = _parse_item_datetime(it["start"])
            en = _parse_item_datetime(it["end"])
            if not st.isValid():
                continue
            if st.date() not in (today, t1):
                continue
            if en.isValid() and now >= en:
                continue
            if now < st:
                if st.isValid():
                    s = now.secsTo(st)
                    if s > 0:
                        deltas.append(s)
            elif en.isValid() and now < en:
                s = now.secsTo(en)
                if s > 0:
                    deltas.append(s)
        if not deltas:
            return
        wait_ms = min(deltas) * 1000
        # Nessun limite artificiale a 1 h: altrimenti si ritarda l’aggiornamento a inizio/fine impegno.
        wait_ms = max(250, min(wait_ms, 86_400_000))
        self._deadline_timer.start(int(wait_ms))

    def refresh(self) -> None:
        while self._root.count():
            item = self._root.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        card = QFrame()
        card.setObjectName("agendaHeaderCard")
        card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        card.setMaximumHeight(132)
        shadow = QGraphicsDropShadowEffect(card)
        shadow.setBlurRadius(12)
        shadow.setOffset(0, 2)
        shadow.setColor(QColor(15, 23, 42, 45))
        card.setGraphicsEffect(shadow)

        outer = QVBoxLayout(card)
        outer.setContentsMargins(8, 5, 8, 6)
        outer.setSpacing(0)

        svc = self._settings()
        if svc is None:
            lbl = QLabel("—")
            lbl.setObjectName("agendaHeaderCardMuted")
            outer.addWidget(lbl)
            self._root.addWidget(card, 1)
            return

        sch = svc.get_work_schedule()
        hset = holiday_month_day_set_from_rows(svc.get_public_holidays())
        raw_items = expand_agenda_items(
            [_normalize_item(x) for x in svc.get_agenda_items()],
            sch,
            hset,
        )
        now = QDateTime.currentDateTime()
        today = QDate.currentDate()
        t1 = today.addDays(1)
        upcoming: list[tuple[QDateTime, dict[str, Any]]] = []
        for it in raw_items:
            if it.get("done"):
                continue
            st = _parse_item_datetime(it["start"])
            en = _parse_item_datetime(it["end"])
            if not st.isValid():
                continue
            if en.isValid() and now >= en:
                continue
            d = st.date()
            if d != today and d != t1:
                continue
            upcoming.append((st, it))
        upcoming.sort(key=lambda x: x[0])
        self._arm_next_deadline(raw_items, now)

        head = QHBoxLayout()
        head.setSpacing(6)
        ht = QLabel("Oggi e domani")
        ht.setObjectName("agendaHeaderCardTitle")
        head.addWidget(ht)
        head.addStretch(1)
        btn_new = QPushButton("Nuovo Appuntamento")
        btn_new.setObjectName("agendaHeaderNewButton")
        btn_new.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_new.clicked.connect(self._open_new_appointment)
        head.addWidget(btn_new, 0, Qt.AlignmentFlag.AlignRight)
        outer.addLayout(head)

        sep = QFrame()
        sep.setObjectName("agendaHeaderCardSeparator")
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFixedHeight(1)
        outer.addWidget(sep)
        outer.addSpacing(4)

        scroll = QScrollArea()
        scroll.setObjectName("agendaHeaderScroll")
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setMinimumHeight(36)
        scroll.setMaximumHeight(82)
        scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        inner = QWidget()
        inner.setObjectName("agendaHeaderScrollInner")
        col = QVBoxLayout(inner)
        col.setContentsMargins(0, 0, 2, 0)
        col.setSpacing(4)

        if not upcoming:
            empty = QLabel("Nessun impegno per oggi e domani")
            empty.setObjectName("agendaHeaderCardMuted")
            empty.setWordWrap(True)
            col.addWidget(empty)
        else:
            for st, it in upcoming[:12]:
                col.addWidget(self._agenda_item_row(st, it, now, today, t1))
        scroll.setWidget(inner)
        outer.addWidget(scroll)
        self._root.addWidget(card, 1)

    def _open_new_appointment(self) -> None:
        """Stesso flusso della vista Agenda → «Nuovo…»."""
        dlg = AgendaItemDialog(self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        item = dlg.to_item()
        svc = self._settings()
        if svc is None:
            return
        items = [_normalize_item(x) for x in svc.get_agenda_items()]
        items.append(item)
        svc.set_agenda_items(items)
        if hasattr(self._mw, "refresh_views"):
            self._mw.refresh_views()

    def _agenda_item_row(
        self,
        st: QDateTime,
        it: dict[str, Any],
        now: QDateTime,
        today: QDate,
        tomorrow: QDate,
    ) -> QFrame:
        fr = QFrame()
        fr.setObjectName("agendaHeaderItemCard")
        k = str(it.get("kind") or "appointment")
        if k not in ("appointment", "task", "vacation", "leave"):
            k = "appointment"
        fr.setProperty("kind", k)
        fr.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        fr.setMaximumHeight(44)
        lay = QVBoxLayout(fr)
        lay.setContentsMargins(6, 4, 6, 4)
        lay.setSpacing(2)

        row1 = QHBoxLayout()
        row1.setSpacing(6)
        kind = QLabel(self._kind_caption(k))
        kind.setObjectName("agendaHeaderItemKind")
        kind.setProperty("kind", k)
        title = QLabel(_short_title(str(it.get("title", "")), 42))
        title.setObjectName("agendaHeaderItemTitle")
        title.setProperty("kind", k)
        title.setWordWrap(False)
        row1.addWidget(kind, 0)
        row1.addWidget(title, 1)
        lay.addLayout(row1)

        day_part = self._day_caption(st, today, tomorrow)
        meta = QLabel(f'{st.toString("dd/MM  HH:mm")} · {day_part}')
        meta.setObjectName("agendaHeaderItemMeta")
        meta.setProperty("kind", k)

        row2 = QHBoxLayout()
        row2.setSpacing(4)
        row2.addWidget(meta, 1)
        sk = self._status_kind(st, it, now)
        status = QLabel(self._status_caption(st, it, now))
        status.setObjectName("agendaHeaderItemStatus")
        status.setProperty("status", sk)
        status.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        row2.addWidget(status, 0, Qt.AlignmentFlag.AlignRight)
        lay.addLayout(row2)
        return fr


class AgendaWorkspaceWidget(QWidget):
    """Vista Agenda: giorno / settimana / mese con orari da Setup Strumenti."""

    def __init__(self, main_window: QWidget, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._mw = main_window
        self._ctl = main_window.repository
        self._anchor = QDate.currentDate()
        self._view = "month"
        self._selected_event_id: str | None = None
        self._cached_week_start: QDate | None = None
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self._build_ui()
        self._apply_agenda_table_styles()
        self._refresh_all()

    def _settings(self):
        if hasattr(self._ctl, "settings"):
            return self._ctl.settings
        return None

    def _schedule(self) -> dict[int, dict[str, Any]]:
        s = self._settings()
        if s is None:
            return {}
        return s.get_work_schedule()

    def _holiday_month_day_set(self) -> set[tuple[int, int]]:
        s = self._settings()
        if s is None:
            return set()
        return holiday_month_day_set_from_rows(s.get_public_holidays())

    def _stored_items(self) -> list[dict[str, Any]]:
        s = self._settings()
        if s is None:
            return []
        return [_normalize_item(x) for x in s.get_agenda_items()]

    def _display_items(self) -> list[dict[str, Any]]:
        return expand_agenda_items(self._stored_items(), self._schedule(), self._holiday_month_day_set())

    def _save_items(self, items: list[dict[str, Any]]) -> None:
        s = self._settings()
        if s is None:
            return
        s.set_agenda_items(items)
        if hasattr(self._mw, "refresh_views"):
            self._mw.refresh_views()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        bar = QHBoxLayout()
        bar.addWidget(QLabel("Vista:"))
        self._mode = QComboBox()
        self._mode.addItems(["Mese", "Settimana", "Giorno"])
        self._mode.currentIndexChanged.connect(self._on_mode_change)
        self._prev = QPushButton("◀")
        self._next = QPushButton("▶")
        self._today = QPushButton("Oggi")
        self._prev.setObjectName("secondaryActionButton")
        self._next.setObjectName("secondaryActionButton")
        self._today.setObjectName("secondaryActionButton")
        self._period_label = QLabel("")
        self._period_label.setObjectName("sectionTitle")
        self._year_spin = QSpinBox()
        self._year_spin.setRange(2000, 2100)
        self._year_spin.setPrefix("Anno ")
        self._year_spin.valueChanged.connect(self._on_year_month_changed)
        bar.addWidget(self._mode)
        bar.addWidget(self._prev)
        bar.addWidget(self._today)
        bar.addWidget(self._next)
        bar.addWidget(self._period_label)
        bar.addStretch(1)
        bar.addWidget(self._year_spin)
        root.addLayout(bar)

        self._stack = QStackedWidget()
        self._month_table = QTableWidget(6, 7)
        self._month_table.setObjectName("agendaMonthTable")
        self._month_table.setHorizontalHeaderLabels(["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"])
        self._month_table.verticalHeader().setVisible(False)
        self._month_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._month_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._month_table.cellClicked.connect(self._on_month_cell_clicked)
        self._month_table.cellDoubleClicked.connect(self._on_month_cell_double_clicked)
        self._week_container = QWidget()
        self._week_container.setObjectName("agendaWeekContainer")
        self._week_layout = QVBoxLayout(self._week_container)
        self._week_layout.setContentsMargins(0, 0, 0, 0)
        self._day_container = QWidget()
        self._day_container.setObjectName("agendaDayContainer")
        self._day_layout = QVBoxLayout(self._day_container)
        self._day_layout.setContentsMargins(0, 0, 0, 0)

        self._stack.addWidget(self._month_table)
        self._stack.addWidget(self._week_container)
        self._stack.addWidget(self._day_container)
        root.addWidget(self._stack, 1)

        hint = QLabel(
            "Settimana e giorno: ogni impegno è un blocco colorato con altezza proporzionale alla durata; "
            "se si sovrappongono, compaiono affiancati (clic sul blocco per selezionare quello giusto). "
            "Lo sfondo rosa indica minuti fuori turno (orario in Setup Strumenti); grigio chiaro = in fascia lavorativa. "
            "All’apertura la vista scorre alla prima ora lavorativa. "
            "Doppio clic sull’impegno per modificarlo; doppio clic su una zona vuota per crearne uno. "
            "«Nuovo» aggiunge un impegno."
        )
        hint.setObjectName("subText")
        hint.setWordWrap(True)
        root.addWidget(hint)

        btn_row = QHBoxLayout()
        self._btn_new = QPushButton("Nuovo…")
        self._btn_edit = QPushButton("Modifica…")
        self._btn_del = QPushButton("Elimina")
        self._btn_new.setObjectName("primaryActionButton")
        self._btn_edit.setObjectName("secondaryActionButton")
        self._btn_del.setObjectName("dangerActionButton")
        btn_row.addWidget(self._btn_new)
        btn_row.addWidget(self._btn_edit)
        btn_row.addWidget(self._btn_del)
        btn_row.addStretch(1)
        root.addLayout(btn_row)

        self._prev.clicked.connect(self._nav_prev)
        self._next.clicked.connect(self._nav_next)
        self._today.clicked.connect(self._nav_today)
        self._btn_new.clicked.connect(self._new_item)
        self._btn_edit.clicked.connect(self._edit_selected)
        self._btn_del.clicked.connect(self._delete_selected)

        self._on_mode_change()

    def _apply_agenda_table_styles(self) -> None:
        self.setStyleSheet(
            """
            QTableWidget#agendaMonthTable {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
                gridline-color: #e2e8f0;
            }
            QTableWidget#agendaMonthTable QHeaderView::section {
                background: #f8fafc;
                color: #0f172a;
                padding: 10px 6px;
                border: none;
                border-bottom: 2px solid #e2e8f0;
                font-weight: 700;
                font-size: 12px;
            }
            QWidget#agendaWeekContainer, QWidget#agendaDayContainer {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
            }
            QScrollArea#agendaWeekScroll, QScrollArea#agendaDayScroll {
                background: #ffffff;
                border: none;
            }
            """
        )

    def _on_mode_change(self) -> None:
        self._view = ["month", "week", "day"][self._mode.currentIndex()]
        self._stack.setCurrentIndex(self._mode.currentIndex())
        self._year_spin.setVisible(self._view == "month")
        self._refresh_all()

    def _on_year_month_changed(self) -> None:
        if self._view == "month":
            y = self._year_spin.value()
            d = self._anchor
            self._anchor = QDate(y, d.month(), min(d.day(), QDate(y, d.month(), 1).daysInMonth()))
            self._refresh_all()

    def _nav_today(self) -> None:
        self._anchor = QDate.currentDate()
        self._refresh_all()

    def _nav_prev(self) -> None:
        if self._view == "month":
            self._anchor = self._anchor.addMonths(-1)
        elif self._view == "week":
            self._anchor = self._anchor.addDays(-7)
        else:
            self._anchor = self._anchor.addDays(-1)
        self._refresh_all()

    def _nav_next(self) -> None:
        if self._view == "month":
            self._anchor = self._anchor.addMonths(1)
        elif self._view == "week":
            self._anchor = self._anchor.addDays(7)
        else:
            self._anchor = self._anchor.addDays(1)
        self._refresh_all()

    def _week_start(self, d: QDate) -> QDate:
        return d.addDays(1 - d.dayOfWeek())

    def refresh(self) -> None:
        """Aggiorna griglie e lista (es. dopo modifica Setup Strumenti)."""
        self._refresh_all()

    def _refresh_all(self) -> None:
        sch = self._schedule()
        self._year_spin.blockSignals(True)
        self._year_spin.setValue(self._anchor.year())
        self._year_spin.blockSignals(False)
        if self._view == "month":
            self._period_label.setText(self._anchor.toString("MMMM yyyy"))
            self._paint_month(sch)
        elif self._view == "week":
            ws = self._week_start(self._anchor)
            we = ws.addDays(6)
            self._period_label.setText(f"{ws.toString('dd/MM')} – {we.toString('dd/MM/yyyy')}")
            self._paint_week(sch, ws)
        else:
            self._period_label.setText(self._anchor.toString("dddd dd/MM/yyyy"))
            self._paint_day(sch)

    def _paint_month(self, sch: dict[int, dict[str, Any]]) -> None:
        items = self._display_items()
        hset = self._holiday_month_day_set()
        y, m = self._anchor.year(), self._anchor.month()
        first = QDate(y, m, 1)
        days_in_m = first.daysInMonth()
        start_col = first.dayOfWeek() - 1
        for r in range(6):
            for c in range(7):
                self._month_table.removeCellWidget(r, c)
                self._month_table.setItem(r, c, None)
            self._month_table.setRowHeight(r, 124)
        d = 1
        r, c = 0, start_col
        while d <= days_in_m:
            qd = QDate(y, m, d)
            day_ev = _events_for_day(items, qd)
            is_h = is_public_holiday(qd, hset)
            non_work = not is_working_day(sch, qd)
            selected_day = self._anchor == qd and self._anchor.year() == y and self._anchor.month() == m
            cell = MonthDayCellWidget(
                qd, day_ev, is_h, non_work, selected_day, self._month_table
            )
            cell.daySelected.connect(self._on_month_day_selected)
            cell.pillDoubleClicked.connect(self._edit_by_id)
            cell.emptyAreaDoubleClicked.connect(self._new_item_for_date)
            self._month_table.setCellWidget(r, c, cell)
            c += 1
            if c > 6:
                c = 0
                r += 1
            d += 1
        for col in range(7):
            self._month_table.horizontalHeader().setSectionResizeMode(
                col, QHeaderView.ResizeMode.Stretch
            )

    def _sync_month_day_selection(self) -> None:
        """Aggiorna il bordo giorno selezionato senza ricostruire le celle (evita di rompere il doppio clic)."""
        for r in range(6):
            for c in range(7):
                w = self._month_table.cellWidget(r, c)
                if isinstance(w, MonthDayCellWidget):
                    w.set_selected(self._anchor == w._qd)

    def _on_month_day_selected(self, qd: QDate, eid: object) -> None:
        self._anchor = qd
        self._selected_event_id = str(eid) if eid else None
        self._sync_month_day_selection()

    def _on_month_cell_clicked(self, row: int, col: int) -> None:
        if self._month_table.cellWidget(row, col) is not None:
            return
        self._selected_event_id = None

    def _on_month_cell_double_clicked(self, row: int, col: int) -> None:
        if self._month_table.cellWidget(row, col) is not None:
            return

    def _clear_layout(self, lay: QVBoxLayout) -> None:
        while lay.count():
            item = lay.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()

    def _paint_week(self, sch: dict[int, dict[str, Any]], ws: QDate) -> None:
        self._cached_week_start = ws
        items = self._display_items()
        hd = self._holiday_month_day_set()
        self._clear_layout(self._week_layout)
        timeline, cols, scroll = build_week_timeline(ws, sch, items, holiday_month_days=hd)
        for col in cols:
            wire_column_signals(
                col,
                self._on_timeline_event_click,
                self._on_timeline_event_dbl,
                self._on_timeline_empty_click,
                self._on_timeline_empty_dbl,
            )
        self._week_layout.addWidget(timeline, 1)
        week_dates = [ws.addDays(i) for i in range(7)]
        schedule_scroll_to_first_work(scroll, sch, week_dates)

    def _paint_day(self, sch: dict[int, dict[str, Any]]) -> None:
        items = self._display_items()
        hd = self._holiday_month_day_set()
        qd = self._anchor
        self._clear_layout(self._day_layout)
        timeline, cols, scroll = build_day_timeline(qd, sch, items, holiday_month_days=hd)
        for col in cols:
            wire_column_signals(
                col,
                self._on_timeline_event_click,
                self._on_timeline_event_dbl,
                self._on_timeline_empty_click,
                self._on_timeline_empty_dbl,
            )
        self._day_layout.addWidget(timeline, 1)
        schedule_scroll_to_first_work(scroll, sch, [qd])

    def _on_timeline_event_click(self, eid: str) -> None:
        self._selected_event_id = str(eid) if eid else None

    def _on_timeline_event_dbl(self, eid: str) -> None:
        if eid:
            self._edit_by_id(str(eid))

    def _on_timeline_empty_click(self, _dt: object) -> None:
        self._selected_event_id = None

    def _on_timeline_empty_dbl(self, dt: object) -> None:
        if isinstance(dt, QDateTime) and dt.isValid():
            now = QDateTime.currentDateTime()
            if dt < now:
                QMessageBox.warning(
                    self,
                    "Agenda",
                    "Hai selezionato un orario già trascorso rispetto all’ora attuale.\n\n"
                    "Puoi comunque procedere con l’inserimento e correggere data e ora nel modulo.",
                )
            self._new_item_for_datetime(dt)

    def _new_item_for_datetime(self, dt: QDateTime) -> None:
        if not dt.isValid():
            return
        en = dt.addSecs(3600)
        seed = _normalize_item(
            {
                "id": str(uuid.uuid4()),
                "title": "",
                "start": _fmt_dt(dt),
                "end": _fmt_dt(en),
                "kind": "appointment",
                "done": False,
            }
        )
        dlg = AgendaItemDialog(self, seed)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        item = dlg.to_item()
        items = self._stored_items()
        items.append(item)
        self._save_items(items)
        self._refresh_all()

    def _edit_by_id(self, eid: str) -> None:
        items = self._stored_items()
        cur = next((x for x in items if str(x.get("id")) == eid), None)
        if not cur:
            return
        dlg = AgendaItemDialog(self, cur)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        new = dlg.to_item(existing_id=str(cur["id"]))
        items = [x for x in items if str(x.get("id")) != str(cur["id"])]
        items.append(new)
        self._save_items(items)
        self._refresh_all()

    def _new_item_for_date(self, qd: QDate, hour: int | None = None) -> None:
        if hour is None:
            # Ora di inizio = ora corrente al momento dell’apertura (sul giorno cliccato), non 9:00 fisse.
            now = QDateTime.currentDateTime()
            st = QDateTime(qd, now.time())
            en = st.addSecs(3600)
        else:
            st = QDateTime(qd, QTime(hour, 0))
            en = st.addSecs(3600)
        seed = _normalize_item(
            {
                "id": str(uuid.uuid4()),
                "title": "",
                "start": _fmt_dt(st),
                "end": _fmt_dt(en),
                "kind": "appointment",
                "done": False,
            }
        )
        dlg = AgendaItemDialog(self, seed)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        item = dlg.to_item()
        items = self._stored_items()
        items.append(item)
        self._save_items(items)
        self._refresh_all()

    def _new_item(self) -> None:
        dlg = AgendaItemDialog(self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        item = dlg.to_item()
        items = self._stored_items()
        items.append(item)
        self._save_items(items)
        self._refresh_all()

    def _edit_selected(self) -> None:
        if not self._selected_event_id:
            QMessageBox.information(
                self,
                "Agenda",
                "Seleziona un appuntamento nel calendario (clic su una cella con un impegno).",
            )
            return
        self._edit_by_id(self._selected_event_id)

    def _delete_selected(self) -> None:
        if not self._selected_event_id:
            QMessageBox.information(self, "Agenda", "Seleziona prima un appuntamento nel calendario.")
            return
        eid = self._selected_event_id
        ans = QMessageBox.question(
            self,
            "Agenda",
            "Eliminare questo impegno?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if ans != QMessageBox.StandardButton.Yes:
            return
        items = [x for x in self._stored_items() if str(x.get("id")) != str(eid)]
        self._save_items(items)
        self._selected_event_id = None
        self._refresh_all()

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self._refresh_all()
