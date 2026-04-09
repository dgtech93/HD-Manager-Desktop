"""Griglia oraria proporzionale (stile calendario) per vista settimana / giorno."""

from __future__ import annotations

import hashlib
from typing import Any, Callable

from PyQt6.QtCore import QEvent, QObject, QDate, QRect, Qt, QTime, QTimer, pyqtSignal
from PyQt6.QtGui import QColor, QFont, QPainter, QPen
from PyQt6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from app.agenda_schedule import (
    first_work_start_minute_for_dates,
    is_public_holiday,
    py_weekday_from_qdate,
    _work_segments_minutes,
)

# Altezza giornata: 24h × 60 min × pixel per minuto
DEFAULT_PPM = 1.35


class _SyncInnerWidthToViewport(QObject):
    """Mantiene la larghezza della griglia uguale alla viewport dello scroll."""

    def __init__(self, scroll: QScrollArea, inner: QWidget) -> None:
        super().__init__(scroll)
        self._scroll = scroll
        self._inner = inner

    def eventFilter(self, obj: QObject | None, ev: QEvent | None) -> bool:  # type: ignore[override]
        if ev is not None and ev.type() == QEvent.Type.Resize:
            w = self._scroll.viewport().width()
            if w > 0:
                self._inner.setMinimumWidth(w)
        return False


def _parse_dt(raw: str):
    from PyQt6.QtCore import QDateTime

    if not raw or not str(raw).strip():
        return QDateTime()
    s = str(raw).strip()
    if " " in s and "T" not in s:
        s = s.replace(" ", "T", 1)
    dt = QDateTime.fromString(s, Qt.DateFormat.ISODate)
    if dt.isValid():
        return dt
    for fmt in (
        "yyyy-MM-ddTHH:mm:ss",
        "yyyy-MM-dd HH:mm:ss",
        "yyyy-MM-ddTHH:mm",
        "yyyy-MM-dd HH:mm",
    ):
        dt = QDateTime.fromString(s, fmt)
        if dt.isValid():
            return dt
    return QDateTime()


def _minutes_from_midnight(dt) -> float:
    if not dt.isValid():
        return 0.0
    t = dt.time()
    return t.hour() * 60 + t.minute() + t.second() / 60.0


def _overlapping_group(ev: dict[str, Any], day_events: list[dict[str, Any]]) -> list[dict[str, Any]]:
    st = _parse_dt(str(ev.get("start", "")))
    en = _parse_dt(str(ev.get("end", "")))
    if not st.isValid():
        return [ev]
    if not en.isValid() or en <= st:
        en = st.addSecs(3600)
    group: list[dict[str, Any]] = []
    for e in day_events:
        est = _parse_dt(str(e.get("start", "")))
        een = _parse_dt(str(e.get("end", "")))
        if not est.isValid():
            continue
        if not een.isValid() or een <= est:
            een = est.addSecs(3600)
        if est < en and een > st:
            group.append(e)
    group.sort(
        key=lambda x: (
            _parse_dt(str(x.get("start", ""))),
            str(x.get("id", "")),
        )
    )
    return group


def _lane_index(ev: dict[str, Any], group: list[dict[str, Any]]) -> tuple[int, int]:
    eid = str(ev.get("id", ""))
    for i, e in enumerate(group):
        if str(e.get("id", "")) == eid:
            return i, max(1, len(group))
    return 0, 1


def _event_colors(ev: dict[str, Any]) -> tuple[QColor, QColor]:
    """(sfondo, bordo) — variazione stabile per id + tipo."""
    kind = str(ev.get("kind", "appointment"))
    eid = str(ev.get("id", ""))
    h = int(hashlib.md5(eid.encode("utf-8"), usedforsecurity=False).hexdigest()[:8], 16)
    if kind == "vacation":
        hue = 145 + (h % 25)
        return QColor.fromHsv(hue, 42, 255), QColor.fromHsv(hue, 110, 175)
    if kind == "leave":
        hue = 285 + (h % 30)
        return QColor.fromHsv(hue % 360, 38, 255), QColor.fromHsv(hue % 360, 95, 195)
    hue = (h % 50) + (30 if kind == "task" else 210)  # ambra vs blu
    hue = hue % 360
    if kind == "task":
        return QColor.fromHsv(hue, 55, 255), QColor.fromHsv(hue, 120, 180)
    return QColor.fromHsv(hue, 40, 255), QColor.fromHsv(hue, 100, 220)


def event_palette(ev: dict[str, Any]) -> tuple[QColor, QColor]:
    """Stessi colori dei blocchi settimana/giorno (es. vista mese a pill)."""
    return _event_colors(ev)


def parse_agenda_datetime(raw: str):
    """Parse ISO-like agenda datetime string."""
    return _parse_dt(raw)


class _EventBlock(QFrame):
    clicked = pyqtSignal(str)
    doubleClicked = pyqtSignal(str)

    def __init__(self, ev: dict[str, Any], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._ev = ev
        self._eid = str(ev.get("id", ""))
        self.setObjectName("agendaEventBlock")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(4, 3, 4, 3)
        lay.setSpacing(0)
        title = QLabel(str(ev.get("title", "")).strip() or "(Senza titolo)")
        title.setWordWrap(True)
        title.setStyleSheet("font-weight: 600; font-size: 11px; color: #0f172a; background: transparent;")
        tm = _parse_dt(str(ev.get("start", "")))
        if tm.isValid():
            sub = QLabel(tm.toString("HH:mm"))
            sub.setStyleSheet("font-size: 10px; color: #475569; background: transparent;")
            lay.addWidget(sub)
        lay.addWidget(title)
        for lab in self.findChildren(QLabel):
            lab.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

    def mousePressEvent(self, e) -> None:
        if e.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit(self._eid)
        super().mousePressEvent(e)

    def mouseDoubleClickEvent(self, e) -> None:
        if e.button() == Qt.MouseButton.LeftButton:
            self.doubleClicked.emit(self._eid)
        super().mouseDoubleClickEvent(e)


class DayTimelineColumn(QWidget):
    """Un giorno: sfondo a ore + blocchi evento posizionati in pixel."""

    eventClicked = pyqtSignal(str)
    eventDoubleClicked = pyqtSignal(str)
    emptyClicked = pyqtSignal(object)  # QDateTime
    emptyDoubleClicked = pyqtSignal(object)

    def __init__(
        self,
        qd,
        schedule: dict[int, dict[str, Any]],
        items: list[dict[str, Any]],
        ppm: float = DEFAULT_PPM,
        holiday_month_days: set[tuple[int, int]] | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._qd = qd
        self._sch = schedule
        self._items = items
        self._ppm = ppm
        self._holiday_month_days = holiday_month_days or set()
        self._day_events = [e for e in items if _parse_dt(str(e.get("start", ""))).date() == qd]
        self._blocks: list[_EventBlock] = []
        h_px = int(24 * 60 * ppm)
        self.setFixedHeight(h_px)
        self.setMinimumWidth(80)
        self.setMouseTracking(True)

    def resizeEvent(self, e) -> None:
        super().resizeEvent(e)
        self._layout_blocks()

    def _layout_blocks(self) -> None:
        for b in self._blocks:
            b.deleteLater()
        self._blocks.clear()
        w = self.width()
        ppm = self._ppm
        for ev in self._day_events:
            st = _parse_dt(str(ev.get("start", "")))
            en = _parse_dt(str(ev.get("end", "")))
            if not st.isValid():
                continue
            if not en.isValid() or en <= st:
                en = st.addSecs(3600)
            group = _overlapping_group(ev, self._day_events)
            lane, nlanes = _lane_index(ev, group)
            top = int(_minutes_from_midnight(st) * ppm)
            dur_m = max(1.0, (en.toMSecsSinceEpoch() - st.toMSecsSinceEpoch()) / 60000.0)
            h = max(18, int(dur_m * ppm))
            col_w = w / max(1, nlanes)
            left = int(lane * col_w)
            bw = max(24, int(col_w) - 2)
            bg, bd = _event_colors(ev)
            block = _EventBlock(ev, self)
            block.setGeometry(left + 1, top, bw, h)
            block.setStyleSheet(
                f"QFrame#agendaEventBlock {{ background: {bg.name()}; border-left: 4px solid {bd.name()}; "
                f"border-radius: 6px; }}"
            )
            block.clicked.connect(self.eventClicked.emit)
            block.doubleClicked.connect(self.eventDoubleClicked.emit)
            block.show()
            self._blocks.append(block)

    def _color_for_minute(self, m: int) -> QColor:
        """Sfondo: festività (rosso intenso) > fuori turno (rosa) > lavoro (grigio chiaro)."""
        if is_public_holiday(self._qd, self._holiday_month_days):
            return QColor(248, 113, 113)
        row = self._sch.get(py_weekday_from_qdate(self._qd))
        work = QColor(248, 250, 252)
        non_work = QColor(255, 228, 230)
        if not row or not row.get("lavorativo"):
            return non_work
        for s0, s1 in _work_segments_minutes(row):
            if s0 <= m < s1:
                return work
        return non_work

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        ppm = self._ppm
        w = self.width()
        p = QPainter(self)
        m = 0
        total_m = 24 * 60
        while m < total_m:
            c = self._color_for_minute(m)
            m2 = m + 1
            while m2 < total_m and self._color_for_minute(m2) == c:
                m2 += 1
            y0 = int(m * ppm)
            y1 = int(m2 * ppm)
            p.fillRect(QRect(0, y0, w, max(1, y1 - y0)), c)
            m = m2

        p.setPen(QPen(QColor(226, 232, 240), 1))
        for hour in range(24):
            y = int(hour * 60 * ppm)
            p.drawLine(0, y, w, y)
        y_half = int(30 * ppm)
        pen_half = QPen(QColor(241, 245, 249), 1, Qt.PenStyle.DotLine)
        p.setPen(pen_half)
        for hour in range(24):
            y = int(hour * 60 * ppm) + y_half
            p.drawLine(0, y, w, y)

    def _time_at_y(self, y: float):
        from PyQt6.QtCore import QDateTime, QTime

        m = y / self._ppm
        m = max(0.0, min(24 * 60 - 1, m))
        mi = int(m)
        sec = int((m - mi) * 60)
        h, mm = divmod(mi, 60)
        return QDateTime(self._qd, QTime(h, mm, min(59, sec)))

    def _hit_event_block(self, pos) -> _EventBlock | None:
        w = self.childAt(pos)
        p = w
        while p is not None and p != self:
            if isinstance(p, _EventBlock):
                return p
            p = p.parentWidget()
        return None

    def mousePressEvent(self, e) -> None:
        if self._hit_event_block(e.pos()) is not None:
            super().mousePressEvent(e)
            return
        self.emptyClicked.emit(self._time_at_y(e.position().y()))
        super().mousePressEvent(e)

    def mouseDoubleClickEvent(self, e) -> None:
        if self._hit_event_block(e.pos()) is not None:
            super().mouseDoubleClickEvent(e)
            return
        self.emptyDoubleClicked.emit(self._time_at_y(e.position().y()))
        super().mouseDoubleClickEvent(e)


class TimeRulerColumn(QWidget):
    def __init__(self, ppm: float = DEFAULT_PPM, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._ppm = ppm
        self.setFixedWidth(52)
        self.setFixedHeight(int(24 * 60 * ppm))

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        p = QPainter(self)
        p.setPen(QColor(100, 116, 139))
        p.setFont(QFont("Segoe UI", 9))
        ppm = self._ppm
        for hour in range(24):
            y = int(hour * 60 * ppm) + 2
            p.drawText(4, y + 12, f"{hour:02d}:00")


def build_week_timeline(
    ws,
    schedule: dict[int, dict[str, Any]],
    items: list[dict[str, Any]],
    ppm: float = DEFAULT_PPM,
    holiday_month_days: set[tuple[int, int]] | None = None,
) -> tuple[QWidget, list[DayTimelineColumn], QScrollArea]:
    """ws = lunedì della settimana. Ritorna (widget con intestazione + scroll, colonne giorno, area scroll)."""
    outer = QWidget()
    outer.setObjectName("agendaWeekTimeline")
    vl = QVBoxLayout(outer)
    vl.setContentsMargins(0, 0, 0, 0)
    vl.setSpacing(0)

    header = QWidget()
    header.setObjectName("agendaWeekHeader")
    hlay = QHBoxLayout(header)
    hlay.setContentsMargins(0, 0, 0, 0)
    hlay.setSpacing(0)
    sp = QWidget()
    sp.setFixedWidth(52)
    hlay.addWidget(sp)
    for i in range(7):
        dd = ws.addDays(i)
        lbl = QLabel(f"{dd.toString('ddd')}\n{dd.toString('dd/MM')}")
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl.setWordWrap(True)
        lbl.setStyleSheet(
            "font-weight: 700; font-size: 12px; color: #0f172a; padding: 8px 4px; "
            "background: #f8fafc; border-bottom: 2px solid #e2e8f0;"
        )
        hlay.addWidget(lbl, 1)

    inner = QWidget()
    row = QHBoxLayout(inner)
    row.setContentsMargins(0, 0, 0, 0)
    row.setSpacing(0)
    row.addWidget(TimeRulerColumn(ppm))
    cols: list[DayTimelineColumn] = []
    for i in range(7):
        col = DayTimelineColumn(ws.addDays(i), schedule, items, ppm, holiday_month_days)
        col.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        row.addWidget(col, 1)
        cols.append(col)

    scroll = QScrollArea()
    scroll.setObjectName("agendaWeekScroll")
    scroll.setWidgetResizable(False)
    scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
    scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
    scroll.setFrameShape(QScrollArea.Shape.NoFrame)
    sync = _SyncInnerWidthToViewport(scroll, inner)
    scroll.viewport().installEventFilter(sync)
    inner.setMinimumWidth(scroll.viewport().width())
    scroll.setWidget(inner)

    vl.addWidget(header)
    vl.addWidget(scroll, 1)
    return outer, cols, scroll


def build_day_timeline(
    qd,
    schedule: dict[int, dict[str, Any]],
    items: list[dict[str, Any]],
    ppm: float = DEFAULT_PPM,
    holiday_month_days: set[tuple[int, int]] | None = None,
) -> tuple[QWidget, list[DayTimelineColumn], QScrollArea]:
    outer = QWidget()
    outer.setObjectName("agendaDayTimeline")
    vl = QVBoxLayout(outer)
    vl.setContentsMargins(0, 0, 0, 0)
    vl.setSpacing(0)

    head = QLabel(qd.toString("dddd d MMMM yyyy"))
    head.setAlignment(Qt.AlignmentFlag.AlignCenter)
    head.setStyleSheet(
        "font-weight: 700; font-size: 13px; color: #0f172a; padding: 10px; "
        "background: #f8fafc; border-bottom: 2px solid #e2e8f0;"
    )
    vl.addWidget(head)

    inner = QWidget()
    row = QHBoxLayout(inner)
    row.setContentsMargins(0, 0, 0, 0)
    row.setSpacing(0)
    row.addWidget(TimeRulerColumn(ppm))
    col = DayTimelineColumn(qd, schedule, items, ppm, holiday_month_days)
    col.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    row.addWidget(col, 1)
    cols = [col]

    scroll = QScrollArea()
    scroll.setObjectName("agendaDayScroll")
    scroll.setWidgetResizable(False)
    scroll.setFrameShape(QScrollArea.Shape.NoFrame)
    sync = _SyncInnerWidthToViewport(scroll, inner)
    scroll.viewport().installEventFilter(sync)
    inner.setMinimumWidth(scroll.viewport().width())
    scroll.setWidget(inner)

    vl.addWidget(scroll, 1)
    return outer, cols, scroll


def wire_column_signals(
    col: DayTimelineColumn,
    on_event_click: Callable[[str], None],
    on_event_dbl: Callable[[str], None],
    on_empty_click: Callable[[object], None],
    on_empty_dbl: Callable[[object], None],
) -> None:
    col.eventClicked.connect(on_event_click)
    col.eventDoubleClicked.connect(on_event_dbl)
    col.emptyClicked.connect(on_empty_click)
    col.emptyDoubleClicked.connect(on_empty_dbl)


def schedule_scroll_to_first_work(
    scroll: QScrollArea,
    schedule: dict[int, dict[str, Any]],
    dates: list[QDate],
    ppm: float = DEFAULT_PPM,
    margin_px: int = 16,
) -> None:
    """Porta lo scroll verticale alla prima ora lavorativa (Setup Strumenti) dopo il layout."""
    first = first_work_start_minute_for_dates(schedule, dates)
    y = max(0, int(first * ppm) - margin_px)

    def apply() -> None:
        sb = scroll.verticalScrollBar()
        mx = sb.maximum()
        if mx <= 0:
            QTimer.singleShot(50, apply)
            return
        sb.setValue(min(y, mx))

    QTimer.singleShot(0, apply)
