"""Logica di confronto / estrazione su dataset tabellari (import Excel)."""

from __future__ import annotations

from collections import Counter, defaultdict
from typing import Any

from app.excel_format import (
    ColumnFormatSpec,
    format_cell_as_string,
    formats_or_auto,
    sort_unique_display_values_for_spec,
)
from app.ui.excel_import_dialog import ExcelImportOutcome


def _norm_cell(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def row_key_full(row: list[Any], formats: list[ColumnFormatSpec] | None = None) -> tuple[str, ...]:
    if formats is None:
        return tuple(_norm_cell(x) for x in row)
    return tuple(
        format_cell_as_string(row[i] if i < len(row) else None, formats[i]) for i in range(len(formats))
    )


def row_key_at(row: list[Any], col_idx: int, formats: list[ColumnFormatSpec] | None = None) -> str:
    if col_idx < 0 or col_idx >= len(row):
        return ""
    v = row[col_idx]
    if formats is None or col_idx >= len(formats):
        return _norm_cell(v)
    return format_cell_as_string(v, formats[col_idx])


def join_key_tuple(
    row: list[Any],
    col_indices: list[int],
    formats: list[ColumnFormatSpec] | None = None,
) -> tuple[str, ...]:
    return tuple(row_key_at(row, j, formats) for j in col_indices)


def wide_headers_for_tables(
    table_labels: list[str], outcomes: list[ExcelImportOutcome]
) -> list[str]:
    """Una colonna per ogni campo di ogni foglio, prefissata con l'etichetta dell'import."""
    out: list[str] = []
    for lab, o in zip(table_labels, outcomes, strict=True):
        for h in o.headers:
            out.append(f"{lab} — {h}")
    return out


def wide_column_formats_for_tables(outcomes: list[ExcelImportOutcome]) -> list[ColumnFormatSpec]:
    """Allinea i formati colonna degli import all'ordine delle colonne in ``wide_headers_for_tables``."""
    out: list[ColumnFormatSpec] = []
    for o in outcomes:
        n = len(o.headers)
        out.extend(formats_or_auto(n, o.column_formats))
    return out


def first_row_by_join_key(
    rows: list[list[Any]],
    key_cols: list[int],
    formats: list[ColumnFormatSpec] | None = None,
) -> dict[tuple[str, ...], list[Any]]:
    """Prima riga per ogni chiave composita (chiavi duplicate nel foglio: si usa la prima)."""
    m: dict[tuple[str, ...], list[Any]] = {}
    for r in rows:
        k = join_key_tuple(r, key_cols, formats)
        if k not in m:
            m[k] = list(r)
    return m


def row_passes_optional_filter(
    row: list[Any],
    filter_col: int | None,
    filter_val: str | None,
    formats: list[ColumnFormatSpec] | None = None,
) -> bool:
    if filter_val is None or not str(filter_val).strip():
        return True
    if filter_col is None:
        return True
    return row_key_at(row, filter_col, formats).lower() == str(filter_val).strip().lower()


def multi_inner_join_wide(
    table_labels: list[str],
    outcomes: list[ExcelImportOutcome],
    key_cols_per_table: list[list[int]],
    optional_filters: list[list[tuple[int, str]]] | None = None,
) -> tuple[list[str], list[list[Any]]]:
    """
    Inner join su chiave composita: stesso numero di componenti per ogni tabella.
    Risultato: tutte le colonne di tutti i fogli affiancate (ordine: tab1 tutte, tab2 tutte, …).

    optional_filters: per tabella una lista di (colonna, valore) in AND; lista vuota = nessun filtro.
    """
    if not outcomes:
        return [], []
    n = len(outcomes)
    if len(table_labels) != n or len(key_cols_per_table) != n:
        raise ValueError("Etichette, outcome e chiavi devono avere la stessa lunghezza")
    klen = len(key_cols_per_table[0])
    if klen == 0:
        raise ValueError("Definisci almeno un campo nella chiave")
    for kc in key_cols_per_table:
        if len(kc) != klen:
            raise ValueError("Ogni foglio deve avere lo stesso numero di campi nella chiave")

    if optional_filters is None:
        optional_filters = [[] for _ in range(n)]
    elif len(optional_filters) != n:
        raise ValueError("Filtri opzionali: una lista per tabella")

    aligned_fmts = [formats_or_auto(len(o.headers), o.column_formats) for o in outcomes]

    filtered_rows: list[list[list[Any]]] = []
    for o, flist, fmts in zip(outcomes, optional_filters, aligned_fmts, strict=True):
        part = [list(r) for r in o.rows if row_matches_all_field_filters(r, flist, fmts)]
        filtered_rows.append(part)

    maps = [
        first_row_by_join_key(part, key_cols_per_table[i], aligned_fmts[i])
        for i, part in enumerate(filtered_rows)
    ]
    key_sets = [set(m.keys()) for m in maps]
    common = set.intersection(*key_sets) if key_sets else set()

    headers = wide_headers_for_tables(table_labels, outcomes)
    out_rows: list[list[Any]] = []
    for k in sorted(common):
        merged: list[Any] = []
        for i, m in enumerate(maps):
            merged.extend(m[k])
        out_rows.append(merged)

    return headers, out_rows


def multi_anti_join_wide(
    label_a: str,
    outcome_a: ExcelImportOutcome,
    key_cols_a: list[int],
    label_b: str,
    outcome_b: ExcelImportOutcome,
    key_cols_b: list[int],
    optional_filter_a: list[tuple[int, str]] | None = None,
    optional_filter_b: list[tuple[int, str]] | None = None,
) -> tuple[list[str], list[list[Any]]]:
    """
    Righe del foglio A la cui chiave non compare nel foglio B (stessa logica di chiave composita).
    Output wide: colonne A + colonne B riempite con None.
    """
    fa = optional_filter_a or []
    fb = optional_filter_b or []
    fa_fmt = formats_or_auto(len(outcome_a.headers), outcome_a.column_formats)
    fb_fmt = formats_or_auto(len(outcome_b.headers), outcome_b.column_formats)
    rows_a = [list(r) for r in outcome_a.rows if row_matches_all_field_filters(r, fa, fa_fmt)]
    rows_b = [list(r) for r in outcome_b.rows if row_matches_all_field_filters(r, fb, fb_fmt)]

    keys_b = {join_key_tuple(r, key_cols_b, fb_fmt) for r in rows_b}
    headers = wide_headers_for_tables([label_a, label_b], [outcome_a, outcome_b])
    nb = len(outcome_b.headers)
    out: list[list[Any]] = []
    for r in rows_a:
        if join_key_tuple(r, key_cols_a, fa_fmt) not in keys_b:
            out.append(list(r) + [None] * nb)
    return headers, out


def _cell_display_str(outcome: ExcelImportOutcome, row: list[Any], col_idx: int) -> str:
    v = row[col_idx] if col_idx < len(row) else None
    n = len(outcome.headers)
    fmts = formats_or_auto(n, outcome.column_formats)
    if col_idx < 0 or col_idx >= len(fmts):
        return "" if v is None else str(v).strip()
    return format_cell_as_string(v, fmts[col_idx])


def unique_sorted_values_for_column(
    outcome: ExcelImportOutcome,
    col_idx: int,
    *,
    max_items: int = 500,
) -> list[str]:
    """Valori distinti della colonna, ordinati (per popolare combobox)."""
    seen: set[str] = set()
    for r in outcome.rows:
        if col_idx < 0 or col_idx >= len(r):
            continue
        seen.add(_cell_display_str(outcome, r, col_idx))
    out = sorted(seen, key=lambda x: (x.lower(), x))
    return out[:max_items]


def row_matches_all_field_filters(
    row: list[Any],
    filters: list[tuple[int, str]],
    formats: list[ColumnFormatSpec] | None = None,
) -> bool:
    """AND su colonne (confronto case-insensitive come single_filter_fields)."""
    for col_idx, wanted in filters:
        got = row_key_at(row, col_idx, formats)
        if wanted.strip().lower() != got.lower():
            return False
    return True


def unique_sorted_values_for_column_after_filters(
    outcome: ExcelImportOutcome,
    col_idx: int,
    prior_filters: list[tuple[int, str]],
    *,
    max_items: int = 500,
) -> list[str]:
    """Valori distinti dopo aver applicato prior_filters (AND) alle righe."""
    fmts = formats_or_auto(len(outcome.headers), outcome.column_formats)
    by_disp: dict[str, Any] = {}
    for r in outcome.rows:
        if col_idx < 0 or col_idx >= len(r):
            continue
        if not row_matches_all_field_filters(r, prior_filters, fmts):
            continue
        disp = _cell_display_str(outcome, r, col_idx)
        if disp not in by_disp:
            by_disp[disp] = r[col_idx] if col_idx < len(r) else None
    spec = fmts[col_idx] if 0 <= col_idx < len(fmts) else ColumnFormatSpec()
    pairs = [(by_disp[d], d) for d in by_disp]
    out = sort_unique_display_values_for_spec(pairs, spec)
    return out[:max_items]


def unique_sorted_values_for_column_joined(
    outcomes: list[ExcelImportOutcome],
    key_cols_per_table: list[list[int]],
    optional_filters_per_table: list[list[tuple[int, str]]],
    target_table: int,
    col_idx: int,
    *,
    target_skip_row: int | None = None,
    max_items: int = 500,
) -> list[str]:
    """
    Valori distinti nella colonna `col_idx` del foglio `target_table`, limitati alle righe
    la cui chiave di join compare nell'intersezione dopo aver applicato i filtri (AND) per tabella.
    `target_skip_row`: indice della riga filtro del foglio target da escludere (combobox in compilazione).
    """
    n = len(outcomes)
    if n == 0:
        return []
    if len(key_cols_per_table) != n or len(optional_filters_per_table) != n:
        raise ValueError("Outcome, chiavi e filtri devono avere la stessa lunghezza")
    aligned_fmts = [formats_or_auto(len(o.headers), o.column_formats) for o in outcomes]

    rows_per_table: list[list[list[Any]]] = []
    for i in range(n):
        fl = list(optional_filters_per_table[i])
        if i == target_table and target_skip_row is not None and 0 <= target_skip_row < len(fl):
            fl = fl[:target_skip_row] + fl[target_skip_row + 1:]
        fmts = aligned_fmts[i]
        part: list[list[Any]] = []
        for r in outcomes[i].rows:
            if row_matches_all_field_filters(r, fl, fmts):
                part.append(list(r))
        rows_per_table.append(part)

    key_sets: list[set[tuple[str, ...]]] = []
    for i, rows in enumerate(rows_per_table):
        key_sets.append(
            {join_key_tuple(r, key_cols_per_table[i], aligned_fmts[i]) for r in rows}
        )
    if not key_sets:
        return []
    common = set.intersection(*key_sets)
    if not common:
        return []

    by_disp: dict[str, Any] = {}
    kcols = key_cols_per_table[target_table]
    o_tgt = outcomes[target_table]
    tgt_fmt = aligned_fmts[target_table]
    for r in rows_per_table[target_table]:
        if join_key_tuple(r, kcols, tgt_fmt) not in common:
            continue
        if col_idx < 0 or col_idx >= len(r):
            continue
        disp = _cell_display_str(o_tgt, r, col_idx)
        if disp not in by_disp:
            by_disp[disp] = r[col_idx] if col_idx < len(r) else None
    spec = tgt_fmt[col_idx] if 0 <= col_idx < len(tgt_fmt) else ColumnFormatSpec()
    pairs = [(by_disp[d], d) for d in by_disp]
    out = sort_unique_display_values_for_spec(pairs, spec)
    return out[:max_items]


def composite_field_label(outcome: ExcelImportOutcome, source_label: str, col_idx: int) -> str:
    h = outcome.headers[col_idx] if col_idx < len(outcome.headers) else f"Col {col_idx + 1}"
    return f"{source_label} — {h}"


# --- Operazioni su singolo foglio ---


def single_duplicates_exact(
    headers: list[str],
    rows: list[list[Any]],
    column_formats: list[ColumnFormatSpec] | None = None,
) -> tuple[list[str], list[list[Any]]]:
    groups: dict[tuple[str, ...], list[list[Any]]] = defaultdict(list)
    for r in rows:
        groups[row_key_full(r, column_formats)].append(list(r))
    out: list[list[Any]] = []
    for _key, group in groups.items():
        if len(group) > 1:
            out.extend(group)
    return headers, out


def single_duplicates_by_column(
    headers: list[str],
    rows: list[list[Any]],
    col_idx: int,
    column_formats: list[ColumnFormatSpec] | None = None,
) -> tuple[list[str], list[list[Any]]]:
    groups: dict[str, list[list[Any]]] = defaultdict(list)
    for r in rows:
        groups[row_key_at(r, col_idx, column_formats)].append(list(r))
    out: list[list[Any]] = []
    for _k, group in groups.items():
        if len(group) > 1:
            out.extend(group)
    return headers, out


def single_unique_exact(
    headers: list[str],
    rows: list[list[Any]],
    column_formats: list[ColumnFormatSpec] | None = None,
) -> tuple[list[str], list[list[Any]]]:
    c = Counter(row_key_full(r, column_formats) for r in rows)
    out = [list(r) for r in rows if c[row_key_full(r, column_formats)] == 1]
    return headers, out


def single_unique_by_column(
    headers: list[str],
    rows: list[list[Any]],
    col_idx: int,
    column_formats: list[ColumnFormatSpec] | None = None,
) -> tuple[list[str], list[list[Any]]]:
    c = Counter(row_key_at(r, col_idx, column_formats) for r in rows)
    out = [list(r) for r in rows if c[row_key_at(r, col_idx, column_formats)] == 1]
    return headers, out


def single_filter_fields(
    headers: list[str],
    rows: list[list[Any]],
    filters: list[tuple[int, str]],
    column_formats: list[ColumnFormatSpec] | None = None,
) -> tuple[list[str], list[list[Any]]]:
    if not filters:
        return headers, [list(r) for r in rows]
    out: list[list[Any]] = []
    for r in rows:
        ok = True
        for col_idx, wanted in filters:
            got = row_key_at(r, col_idx, column_formats)
            if wanted.strip().lower() != got.lower():
                ok = False
                break
        if ok:
            out.append(list(r))
    return headers, out
