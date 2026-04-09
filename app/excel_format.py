"""Tipo e formato di visualizzazione per colonne Excel (import, confronto, concatenazione)."""

from __future__ import annotations

import math
from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import ROUND_HALF_UP, Decimal
from typing import Any


@dataclass
class ColumnFormatSpec:
    """Come mostrare (e confrontare come testo) i valori di una colonna."""

    kind: str = "auto"
    decimal_places: int = 2
    date_pattern: str = "%d/%m/%Y"


KIND_AUTO = "auto"
KIND_STRING = "string"
KIND_INTEGER = "integer"
KIND_DECIMAL = "decimal"
KIND_DATE = "date"
KIND_DATETIME = "datetime"


def formats_or_auto(n: int, fmts: list[ColumnFormatSpec] | None) -> list[ColumnFormatSpec]:
    """Allinea la lista di formati a ``n`` colonne: padding con Automatico o troncamento.

    Evita di perdere i formati scelti dall'utente quando mancano o avanzano voci
    rispetto alle colonne effettive (import/confronto/concatenazione).
    """
    if n <= 0:
        return []
    if fmts is None:
        return [ColumnFormatSpec() for _ in range(n)]
    if len(fmts) == n:
        return fmts
    if len(fmts) > n:
        return fmts[:n]
    out = list(fmts)
    out.extend(ColumnFormatSpec() for _ in range(n - len(fmts)))
    return out


def _norm_plain(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _coerce_datetime(val: Any) -> datetime | date | time | None:
    """Prova a ottenere date/datetime/time da valori Excel / Python.

    openpyxl ``from_excel`` può restituire ``datetime.time`` per seriali 0–1 (solo ora).
    """
    if val is None:
        return None
    if isinstance(val, datetime):
        return val
    if isinstance(val, time):
        return val
    if isinstance(val, date) and not isinstance(val, datetime):
        return datetime.combine(val, time.min)
    if isinstance(val, (int, float)):
        try:
            from openpyxl.utils.datetime import from_excel

            r = from_excel(float(val))
            if r is not None:
                return r
        except Exception:
            pass
        try:
            import xlrd
            from xlrd import xldate_as_datetime

            return xldate_as_datetime(float(val), 0)
        except Exception:
            pass
    return None


def _format_auto_datetimeish(dtv: datetime | date | time) -> str:
    """Formato anteprima in modalità auto (date, datetime o solo ora)."""
    if isinstance(dtv, time):
        return dtv.strftime("%H:%M:%S")
    if isinstance(dtv, datetime):
        return dtv.strftime(
            "%d/%m/%Y %H:%M:%S" if dtv.time() != time.min else "%d/%m/%Y"
        )
    return dtv.strftime("%d/%m/%Y")


def format_cell_as_string(raw: Any, spec: ColumnFormatSpec) -> str:
    """Valore mostrato come stringa secondo il tipo scelto."""
    kind = spec.kind or KIND_AUTO
    if raw is None:
        return ""

    if kind == KIND_AUTO:
        dt = _coerce_datetime(raw)
        if dt is not None:
            return _format_auto_datetimeish(dt)
        if isinstance(raw, (int, float)) and not isinstance(raw, bool):
            if float(raw).is_integer():
                return str(int(raw))
            return str(raw)
        return _norm_plain(raw)

    if kind == KIND_STRING:
        return _norm_plain(raw)

    if kind == KIND_INTEGER:
        try:
            if isinstance(raw, (int, float)) and not isinstance(raw, bool):
                return str(int(round(float(raw))))
            s = _norm_plain(raw)
            if not s:
                return ""
            return str(int(round(float(s.replace(",", ".")))))
        except (ValueError, TypeError):
            return _norm_plain(raw)

    if kind == KIND_DECIMAL:
        places = max(0, min(12, int(spec.decimal_places)))
        try:
            if isinstance(raw, (int, float)) and not isinstance(raw, bool):
                d = Decimal(str(float(raw)))
            else:
                s = _norm_plain(raw).replace(",", ".")
                if not s:
                    return ""
                d = Decimal(s)
        except Exception:
            return _norm_plain(raw)
        q = Decimal(10) ** -places
        return str(d.quantize(q, rounding=ROUND_HALF_UP))

    if kind in (KIND_DATE, KIND_DATETIME):
        dt = _coerce_datetime(raw)
        if dt is None:
            return _norm_plain(raw)
        pat = spec.date_pattern or "%d/%m/%Y"
        try:
            if kind == KIND_DATE:
                if isinstance(dt, time):
                    return dt.strftime("%H:%M:%S")
                if isinstance(dt, date) and not isinstance(dt, datetime):
                    return dt.strftime(pat)
                return dt.date().strftime(pat)
            if isinstance(dt, time):
                return datetime.combine(date.today(), dt).strftime(pat)
            if isinstance(dt, date) and not isinstance(dt, datetime):
                return datetime.combine(dt, time.min).strftime(pat)
            return dt.strftime(pat)
        except Exception:
            return _format_auto_datetimeish(dt) if isinstance(dt, (datetime, date, time)) else _norm_plain(raw)

    return _norm_plain(raw)


def _ordinal_for_chronological_sort(raw: Any) -> float | None:
    """Ordine temporale per ordinare valori distinti (None = non trattabile come data/ora)."""
    dtv = _coerce_datetime(raw)
    if dtv is None:
        return None
    if isinstance(dtv, datetime):
        return dtv.timestamp()
    if isinstance(dtv, date) and not isinstance(dtv, datetime):
        return datetime.combine(dtv, time.min).timestamp()
    if isinstance(dtv, time):
        return dtv.hour * 3600 + dtv.minute * 60 + dtv.second + dtv.microsecond * 1e-6
    return None


def sort_unique_display_values_for_spec(
    pairs: list[tuple[Any, str]],
    spec: ColumnFormatSpec,
) -> list[str]:
    """Ordina stringhe già formattate per elenchi (filtri, combobox).

    Colonne **Data** / **Data e ora**: ordine cronologico sul valore grezzo.
    **Automatico**: cronologico solo se tutti i valori sono interpretabili come data/ora;
    se la colonna è mista (testo + date), ordine alfabetico come prima.
    Altre colonne: ordine alfabetico sul testo mostrato.
    """
    if not pairs:
        return []
    kind = spec.kind or KIND_AUTO

    def _alpha() -> list[str]:
        return [p[1] for p in sorted(pairs, key=lambda p: (p[1].lower(), p[1]))]

    if kind not in (KIND_AUTO, KIND_DATE, KIND_DATETIME):
        return _alpha()

    if kind == KIND_AUTO:
        ords = [_ordinal_for_chronological_sort(p[0]) for p in pairs]
        if any(o is not None for o in ords) and any(o is None for o in ords):
            return _alpha()

    def sort_key(p: tuple[Any, str]) -> tuple:
        raw, disp = p
        o = _ordinal_for_chronological_sort(raw)
        if o is not None:
            return (0, o, disp.lower())
        return (1, disp.lower(), disp)

    return [p[1] for p in sorted(pairs, key=sort_key)]


def format_row_strings(row: list[Any], specs: list[ColumnFormatSpec] | None) -> list[str]:
    n = len(row)
    fmts = formats_or_auto(n, specs)
    return [format_cell_as_string(row[i] if i < len(row) else None, fmts[i]) for i in range(n)]


def format_matrix_strings(
    rows: list[list[Any]],
    headers_len: int,
    specs: list[ColumnFormatSpec] | None,
) -> list[list[str]]:
    fmts = formats_or_auto(headers_len, specs)
    out: list[list[str]] = []
    for row in rows:
        padded = list(row) + [None] * max(0, headers_len - len(row))
        out.append(
            [
                format_cell_as_string(padded[i] if i < len(padded) else None, fmts[i])
                for i in range(headers_len)
            ]
        )
    return out


def _sample_non_empty_cells(values: list[Any], limit: int) -> list[Any]:
    out: list[Any] = []
    for v in values:
        if v is None:
            continue
        if isinstance(v, str) and not str(v).strip():
            continue
        out.append(v)
        if len(out) >= limit:
            break
    return out


def _try_parse_floatish(v: Any) -> float | None:
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def _is_whole_number(f: float) -> bool:
    if math.isnan(f) or math.isinf(f):
        return False
    return abs(f - round(f)) < 1e-9


def _decimal_places_hint(f: float) -> int:
    try:
        d = Decimal(str(f))
        t = d.as_tuple()
        if t.exponent >= 0:
            return 0
        return min(6, -int(t.exponent))
    except Exception:
        return 2


def _parse_string_datetime(s: str) -> tuple[str, str] | None:
    """Se la stringa è una data/ora nota, restituisce (KIND_DATE|KIND_DATETIME, pattern strftime)."""
    s = s.strip()
    if not s:
        return None
    dt_patterns = (
        ("%d/%m/%Y %H:%M:%S", KIND_DATETIME),
        ("%d/%m/%Y %H:%M", KIND_DATETIME),
        ("%Y-%m-%d %H:%M:%S", KIND_DATETIME),
        ("%Y-%m-%dT%H:%M:%S", KIND_DATETIME),
        ("%Y-%m-%d %H:%M", KIND_DATETIME),
    )
    for pat, kind in dt_patterns:
        try:
            datetime.strptime(s, pat)
            return (kind, pat)
        except ValueError:
            continue
    d_patterns = (
        ("%d/%m/%Y", KIND_DATE),
        ("%d/%m/%y", KIND_DATE),
        ("%Y-%m-%d", KIND_DATE),
        ("%d-%m-%Y", KIND_DATE),
        ("%Y/%m/%d", KIND_DATE),
    )
    for pat, kind in d_patterns:
        try:
            datetime.strptime(s, pat)
            return (kind, pat)
        except ValueError:
            continue
    return None


def infer_column_format_spec(values: list[Any], *, sample_limit: int = 500) -> ColumnFormatSpec:
    """
    Suggerisce un ``ColumnFormatSpec`` da un campione di celle (stessa colonna).
    Euristica: date/ora (Excel, Python o stringhe comuni), interi, decimali, altrimenti testo.
    """
    sample = _sample_non_empty_cells(values, sample_limit)
    if not sample:
        return ColumnFormatSpec(kind=KIND_AUTO)

    kind_votes: dict[str, int] = {}
    pattern_votes: dict[str, int] = {}
    max_dec = 0

    for v in sample:
        if isinstance(v, bool):
            kind_votes[KIND_STRING] = kind_votes.get(KIND_STRING, 0) + 1
            continue

        dt = _coerce_datetime(v)
        if dt is not None:
            if isinstance(dt, time):
                kind_votes[KIND_DATETIME] = kind_votes.get(KIND_DATETIME, 0) + 1
                pattern_votes["%d/%m/%Y %H:%M:%S"] = pattern_votes.get("%d/%m/%Y %H:%M:%S", 0) + 1
                continue
            if isinstance(dt, datetime):
                t = dt.time()
                if t.hour == 0 and t.minute == 0 and t.second == 0 and t.microsecond == 0:
                    kind_votes[KIND_DATE] = kind_votes.get(KIND_DATE, 0) + 1
                    pattern_votes["%d/%m/%Y"] = pattern_votes.get("%d/%m/%Y", 0) + 1
                else:
                    kind_votes[KIND_DATETIME] = kind_votes.get(KIND_DATETIME, 0) + 1
                    pattern_votes["%d/%m/%Y %H:%M:%S"] = pattern_votes.get("%d/%m/%Y %H:%M:%S", 0) + 1
                continue
            if isinstance(dt, date) and not isinstance(dt, datetime):
                kind_votes[KIND_DATE] = kind_votes.get(KIND_DATE, 0) + 1
                pattern_votes["%d/%m/%Y"] = pattern_votes.get("%d/%m/%Y", 0) + 1
                continue

        if isinstance(v, str):
            ps = _parse_string_datetime(v)
            if ps:
                k, pat = ps
                kind_votes[k] = kind_votes.get(k, 0) + 1
                pattern_votes[pat] = pattern_votes.get(pat, 0) + 1
                continue

        f = _try_parse_floatish(v)
        if f is not None:
            if _is_whole_number(f):
                kind_votes[KIND_INTEGER] = kind_votes.get(KIND_INTEGER, 0) + 1
            else:
                kind_votes[KIND_DECIMAL] = kind_votes.get(KIND_DECIMAL, 0) + 1
                max_dec = max(max_dec, _decimal_places_hint(f))
            continue

        kind_votes[KIND_STRING] = kind_votes.get(KIND_STRING, 0) + 1

    total = len(sample)
    thr = max(1, int(total * 0.55))

    n_dt = kind_votes.get(KIND_DATETIME, 0)
    n_d = kind_votes.get(KIND_DATE, 0)
    if n_dt + n_d >= thr:
        if n_dt >= n_d:
            cand_dt = [p for p in pattern_votes if ("%H" in p or "T" in p)]
            if not cand_dt:
                pat = "%d/%m/%Y %H:%M"
            else:
                pat = max(cand_dt, key=lambda p: pattern_votes.get(p, 0))
            return ColumnFormatSpec(kind=KIND_DATETIME, date_pattern=pat)
        cand_d = [p for p in pattern_votes if "%H" not in p and "T" not in p]
        if not cand_d:
            pat = "%d/%m/%Y"
        else:
            pat = max(cand_d, key=lambda p: pattern_votes.get(p, 0))
        return ColumnFormatSpec(kind=KIND_DATE, date_pattern=pat)

    ni = kind_votes.get(KIND_INTEGER, 0)
    ndec = kind_votes.get(KIND_DECIMAL, 0)
    if ni + ndec >= thr:
        if ndec >= ni and ndec > 0:
            places = max(2, min(6, max_dec if max_dec > 0 else 2))
            return ColumnFormatSpec(kind=KIND_DECIMAL, decimal_places=places)
        return ColumnFormatSpec(kind=KIND_INTEGER)

    if kind_votes.get(KIND_STRING, 0) >= thr:
        return ColumnFormatSpec(kind=KIND_STRING)

    # misto: preferisci testo se non c’è una maggioranza netta
    if kind_votes.get(KIND_STRING, 0) > 0:
        return ColumnFormatSpec(kind=KIND_STRING)

    return ColumnFormatSpec(kind=KIND_AUTO)
