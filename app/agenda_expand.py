"""Normalizzazione voci agenda ed espansione intervalli su giorni lavorativi «bianchi»."""

from __future__ import annotations

import uuid
from typing import Any

from PyQt6.QtCore import QDate, QDateTime, QTime, Qt

from app.agenda_schedule import is_white_agenda_day


def parse_item_datetime(raw: str) -> QDateTime:
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


def fmt_dt(dt: QDateTime) -> str:
    return dt.toString("yyyy-MM-ddTHH:mm:ss")


_VALID_KINDS = frozenset({"appointment", "task", "vacation", "leave"})


def normalize_item(row: dict[str, Any]) -> dict[str, Any]:
    out = dict(row)
    out["id"] = str(out.get("id") or "").strip() or str(uuid.uuid4())
    out["title"] = str(out.get("title") or "").strip() or "(Senza titolo)"
    out["kind"] = str(out.get("kind") or "appointment").strip()
    if out["kind"] not in _VALID_KINDS:
        out["kind"] = "appointment"
    out["done"] = bool(out.get("done", False))
    st = parse_item_datetime(str(out.get("start", "")))
    en = parse_item_datetime(str(out.get("end", "")))
    if not en.isValid() or en <= st:
        en = st.addSecs(3600)
    out["start"] = fmt_dt(st)
    out["end"] = fmt_dt(en)
    if st.date() != en.date():
        out["multi_day"] = True
    else:
        out["multi_day"] = False
    return out


def expand_agenda_items(
    items: list[dict[str, Any]],
    schedule: dict[int, dict[str, Any]],
    holiday_month_days: set[tuple[int, int]],
) -> list[dict[str, Any]]:
    """Espande le voci multi-giorno in segmenti per ogni giorno lavorativo bianco (non festivo)."""
    out: list[dict[str, Any]] = []
    for it in items:
        it = normalize_item(dict(it))
        if not it.get("multi_day"):
            out.append(it)
            continue
        st = parse_item_datetime(it["start"])
        en = parse_item_datetime(it["end"])
        if not st.isValid():
            out.append(it)
            continue
        if not en.isValid() or en <= st:
            en = st.addSecs(3600)
        d0 = st.date()
        d1 = en.date()
        if d1 < d0:
            out.append(it)
            continue
        if d0 == d1:
            out.append({**it, "multi_day": False})
            continue
        segs: list[dict[str, Any]] = []
        d = d0
        while d <= d1:
            if not is_white_agenda_day(schedule, d, holiday_month_days):
                d = d.addDays(1)
                continue
            if d == d0:
                ts = st
                te = QDateTime(d, QTime(23, 59, 59)) if d < d1 else en
            elif d == d1:
                ts = QDateTime(d, QTime(0, 0))
                te = en
            else:
                ts = QDateTime(d, QTime(0, 0))
                te = QDateTime(d, QTime(23, 59, 59))
            if te <= ts:
                te = ts.addSecs(60)
            segs.append({**it, "start": fmt_dt(ts), "end": fmt_dt(te), "multi_day": False})
            d = d.addDays(1)
        if not segs:
            out.append(it)
        else:
            out.extend(segs)
    return out


def holiday_month_day_set_from_rows(rows: list[dict[str, Any]]) -> set[tuple[int, int]]:
    """Da righe {month, day, label} (o legacy {date}) a set (mese, giorno)."""
    s: set[tuple[int, int]] = set()
    for r in rows:
        if not isinstance(r, dict):
            continue
        m: int | None = None
        d: int | None = None
        if "month" in r and "day" in r:
            try:
                m = int(r["month"])
                d = int(r["day"])
            except (TypeError, ValueError):
                m = d = None
        if m is None and r.get("date"):
            qd = QDate.fromString(str(r["date"]).strip()[:10], Qt.DateFormat.ISODate)
            if qd.isValid():
                m, d = qd.month(), qd.day()
        if m is None or d is None:
            continue
        if 1 <= m <= 12 and 1 <= d <= 31:
            s.add((m, d))
    return s
