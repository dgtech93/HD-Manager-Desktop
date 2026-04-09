"""Celle giorno per la vista calendario mensile (pill impilate)."""

from __future__ import annotations

from typing import Any

from PyQt6.QtCore import QDate, Qt, pyqtSignal
from PyQt6.QtGui import QFontMetrics
from PyQt6.QtWidgets import QFrame, QHBoxLayout, QLabel, QSizePolicy, QVBoxLayout

from app.ui.agenda_time_grid import event_palette, parse_agenda_datetime

_MAX_PILLS = 5


class _MonthPill(QFrame):
    """Singola pill: striscia colorata a sinistra + ora + titolo (eliso)."""

    clicked = pyqtSignal(str)
    doubleClicked = pyqtSignal(str)

    def __init__(self, ev: dict[str, Any], parent: QFrame | None = None) -> None:
        super().__init__(parent)
        self._eid = str(ev.get("id", ""))
        self.setObjectName("agendaMonthPill")
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.setMinimumHeight(22)
        self.setMaximumHeight(26)

        bg, accent = event_palette(ev)
        self.setStyleSheet(
            f"QFrame#agendaMonthPill {{ background: {bg.name()}; border: none; "
            f"border-radius: 6px; border-left: 4px solid {accent.name()}; }}"
        )

        lay = QHBoxLayout(self)
        lay.setContentsMargins(6, 2, 6, 2)
        lay.setSpacing(6)

        st = parse_agenda_datetime(str(ev.get("start", "")))
        time_lbl = QLabel(st.toString("HH:mm") if st.isValid() else "—")
        time_lbl.setStyleSheet("font-weight: 700; font-size: 10px; color: #0f172a; background: transparent;")
        time_lbl.setFixedWidth(38)
        time_lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

        title = str(ev.get("title", "")).strip() or "(Senza titolo)"
        self._title_full = title
        self._title_lbl = QLabel(title)
        self._title_lbl.setStyleSheet("font-size: 10px; color: #334155; background: transparent;")
        self._title_lbl.setWordWrap(False)
        self._title_lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

        lay.addWidget(time_lbl, 0)
        lay.addWidget(self._title_lbl, 1)

    def resizeEvent(self, e) -> None:
        super().resizeEvent(e)
        self._apply_elide()

    def _apply_elide(self) -> None:
        w = max(20, self.width() - 50)
        fm = QFontMetrics(self._title_lbl.font())
        self._title_lbl.setText(fm.elidedText(self._title_full, Qt.TextElideMode.ElideRight, w))

    def showEvent(self, e) -> None:
        super().showEvent(e)
        self._apply_elide()

    def mousePressEvent(self, e) -> None:
        if e.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit(self._eid)
        super().mousePressEvent(e)

    def mouseDoubleClickEvent(self, e) -> None:
        if e.button() == Qt.MouseButton.LeftButton:
            self.doubleClicked.emit(self._eid)
        super().mouseDoubleClickEvent(e)


class MonthDayCellWidget(QFrame):
    """Contenuto cella mese: numero giorno + pill per ogni impegno."""

    daySelected = pyqtSignal(object, object)  # QDate, str | None (event id)
    pillDoubleClicked = pyqtSignal(str)
    emptyAreaDoubleClicked = pyqtSignal(object)  # QDate

    def __init__(
        self,
        qd: QDate,
        events: list[dict[str, Any]],
        is_holiday: bool,
        non_work: bool,
        selected_day: bool,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._qd = qd
        self._holiday = is_holiday
        self._non_work = non_work
        self._selected = selected_day
        self.setObjectName("agendaMonthDayCell")
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self._apply_frame_style()

        root = QVBoxLayout(self)
        root.setContentsMargins(6, 6, 6, 6)
        root.setSpacing(4)

        head = QHBoxLayout()
        head.setSpacing(0)
        self._day_lbl = QLabel(str(qd.day()))
        day_lbl = self._day_lbl
        self._apply_day_number_style()
        day_lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        head.addWidget(day_lbl, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        head.addStretch(1)
        root.addLayout(head)

        shown = events[:_MAX_PILLS]
        rest = max(0, len(events) - _MAX_PILLS)

        for ev in shown:
            pill = _MonthPill(ev, self)
            pill.clicked.connect(lambda eid, _qd=self._qd: self.daySelected.emit(_qd, eid))
            pill.doubleClicked.connect(self.pillDoubleClicked.emit)
            root.addWidget(pill)

        if rest > 0:
            more = QLabel(f"+{rest} altri")
            more.setStyleSheet("font-size: 10px; color: #64748b; font-weight: 600;")
            more.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
            root.addWidget(more)

        root.addStretch(1)

    def _apply_frame_style(self) -> None:
        if self._selected:
            border = "2px solid #2563eb"
            bg = "#eff6ff"
        elif self._holiday:
            border = "1px solid #b91c1c"
            bg = "#fecaca"
        elif self._non_work:
            border = "1px solid #fecaca"
            bg = "#fff1f2"
        else:
            border = "1px solid #e2e8f0"
            bg = "#ffffff"
        self.setStyleSheet(
            f"QFrame#agendaMonthDayCell {{ background: {bg}; border: {border}; border-radius: 10px; }}"
        )

    def _apply_day_number_style(self) -> None:
        self._day_lbl.setStyleSheet(
            "font-weight: 800; font-size: 13px; background: transparent;"
            + (" color: #1d4ed8;" if self._selected else " color: #0f172a;")
        )

    def set_selected(self, selected: bool) -> None:
        self._selected = selected
        self._apply_frame_style()
        self._apply_day_number_style()

    def mousePressEvent(self, e) -> None:
        if e.button() == Qt.MouseButton.LeftButton:
            self.daySelected.emit(self._qd, None)
        super().mousePressEvent(e)

    def mouseDoubleClickEvent(self, e) -> None:
        if e.button() == Qt.MouseButton.LeftButton:
            child = self.childAt(e.position().toPoint())
            p = child
            while p is not None and p != self:
                if isinstance(p, _MonthPill):
                    super().mouseDoubleClickEvent(e)
                    return
                p = p.parentWidget()
            self.emptyAreaDoubleClicked.emit(self._qd)
        super().mouseDoubleClickEvent(e)
