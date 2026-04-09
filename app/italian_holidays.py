"""Festività nazionali italiane (date mobili + fisse)."""

from __future__ import annotations

from datetime import date, timedelta

from PyQt6.QtCore import QDate, Qt


def _easter_sunday(y: int) -> date:
    """Algoritmo gregoriano (Anonymous)."""
    a = y % 19
    b = y // 100
    c = y % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(y, month, day)


def _to_qdate(d: date) -> QDate:
    return QDate(d.year, d.month, d.day)


def fixed_recurring_italian_holidays() -> list[tuple[int, int, str]]:
    """Solo giorno/mese (ricorrenti ogni anno), senza Pasqua/Pasquetta (calcolate a parte)."""
    return [
        (1, 1, "Capodanno"),
        (1, 6, "Epifania"),
        (4, 25, "Festa della Liberazione"),
        (5, 1, "Festa del Lavoro"),
        (6, 2, "Festa della Repubblica"),
        (8, 15, "Ferragosto"),
        (11, 1, "Ognissanti"),
        (12, 8, "Immacolata Concezione"),
        (12, 25, "Natale"),
        (12, 26, "Santo Stefano"),
    ]


def is_movable_italian_holiday(qd: QDate) -> bool:
    """Pasqua e Pasquetta per l’anno di qd (non memorizzate come gg/mm fissi)."""
    y = qd.year()
    e = _easter_sunday(y)
    pasqua = _to_qdate(e)
    pasquetta = _to_qdate(e + timedelta(days=1))
    return qd == pasqua or qd == pasquetta


def italian_public_holidays_for_year(year: int) -> list[tuple[QDate, str]]:
    """Coppie (data, etichetta) festività civili italiane per l'anno indicato."""
    easter = _easter_sunday(year)
    pasqua = _to_qdate(easter)
    pasquetta = _to_qdate(easter + timedelta(days=1))

    fixed: list[tuple[tuple[int, int], str]] = [
        ((1, 1), "Capodanno"),
        ((1, 6), "Epifania"),
        ((4, 25), "Festa della Liberazione"),
        ((5, 1), "Festa del Lavoro"),
        ((6, 2), "Festa della Repubblica"),
        ((8, 15), "Ferragosto"),
        ((11, 1), "Ognissanti"),
        ((12, 8), "Immacolata Concezione"),
        ((12, 25), "Natale"),
        ((12, 26), "Santo Stefano"),
    ]
    out: list[tuple[QDate, str]] = [
        (pasqua, "Pasqua"),
        (pasquetta, "Lunedì dell'Angelo (Pasquetta)"),
    ]
    for (mo, da), label in fixed:
        out.append((QDate(year, mo, da), label))
    out.sort(key=lambda x: x[0].toJulianDay())
    return out
