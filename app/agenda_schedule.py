"""Regole orario lavorativo (Setup Strumenti) per colorazione agenda."""

from __future__ import annotations

from typing import Any

from PyQt6.QtCore import QDate, Qt

from app.italian_holidays import is_movable_italian_holiday


def py_weekday_from_qdate(qd: QDate) -> int:
    """Lunedì=0 … Domenica=6 (come datetime.weekday())."""
    return qd.dayOfWeek() - 1


def _minutes(hhmm: str) -> int:
    parts = (hhmm or "00:00").replace(".", ":").split(":")
    h = int(parts[0]) if parts else 0
    m = int(parts[1]) if len(parts) > 1 else 0
    return max(0, min(24 * 60 - 1, h * 60 + m))


def _work_segments_minutes(row: dict[str, Any]) -> list[tuple[int, int]]:
    """Fasce [start,end) in minuti da mezzanotte, esclusa pausa."""
    start = _minutes(str(row.get("inizio", "09:00")))
    end = _minutes(str(row.get("fine", "18:00")))
    if end <= start:
        return []
    if not row.get("pausa_abilitata"):
        return [(start, end)]
    p1 = _minutes(str(row.get("pausa_inizio", "13:00")))
    p2 = _minutes(str(row.get("pausa_fine", "14:00")))
    if p1 >= p2 or p1 <= start or p2 >= end:
        return [(start, end)]
    out: list[tuple[int, int]] = []
    if p1 > start:
        out.append((start, min(p1, end)))
    if p2 < end:
        out.append((max(p2, start), end))
    return out


def is_working_day(schedule: dict[int, dict[str, Any]], qd: QDate) -> bool:
    wd = py_weekday_from_qdate(qd)
    row = schedule.get(wd)
    if not row:
        return False
    return bool(row.get("lavorativo", False))


def is_public_holiday(qd: QDate, user_month_days: set[tuple[int, int]]) -> bool:
    """Festività: giorno/mese da Setup + Pasqua/Pasquetta (calcolate)."""
    if (qd.month(), qd.day()) in user_month_days:
        return True
    if is_movable_italian_holiday(qd):
        return True
    return False


def is_white_agenda_day(
    schedule: dict[int, dict[str, Any]], qd: QDate, holiday_month_days: set[tuple[int, int]]
) -> bool:
    """Giorno bianco in agenda: turno lavorativo e non festivo."""
    if is_public_holiday(qd, holiday_month_days):
        return False
    return is_working_day(schedule, qd)


def hour_is_non_work(schedule: dict[int, dict[str, Any]], qd: QDate, hour: int) -> bool:
    """True se l’intera ora [hour, hour+1) è fuori dal lavoro (sfondo rosso)."""
    wd = py_weekday_from_qdate(qd)
    row = schedule.get(wd)
    if not row or not row.get("lavorativo"):
        return True
    h0 = hour * 60
    h1 = h0 + 60
    for s0, s1 in _work_segments_minutes(row):
        if s0 < h1 and s1 > h0:
            return False
    return True


def first_work_start_minute_for_dates(
    schedule: dict[int, dict[str, Any]], dates: list[QDate],
) -> int:
    """Primo minuto (da mezzanotte) in cui inizia una fascia lavorativa tra i giorni dati."""
    best: int | None = None
    for qd in dates:
        row = schedule.get(py_weekday_from_qdate(qd))
        if not row or not row.get("lavorativo"):
            continue
        for s0, s1 in _work_segments_minutes(row):
            if best is None or s0 < best:
                best = s0
    if best is not None:
        return best
    return 8 * 60


def overlaps_work_segment(
    schedule: dict[int, dict[str, Any]], qd: QDate, start_min: int, end_min: int
) -> bool:
    """True se l’intervallo [start_min, end_min) interseca almeno un minuto lavorativo."""
    wd = py_weekday_from_qdate(qd)
    row = schedule.get(wd)
    if not row or not row.get("lavorativo"):
        return False
    for s0, s1 in _work_segments_minutes(row):
        if s0 < end_min and s1 > start_min:
            return True
    return False
