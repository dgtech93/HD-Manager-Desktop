"""Concatenazione testi da colonne/righe (incluso formato IN (...) per SQL)."""

from __future__ import annotations

from typing import Any, Iterable

from app.excel_format import ColumnFormatSpec, format_cell_as_string, formats_or_auto


def sql_escape_single_quotes(s: str) -> str:
    return s.replace("'", "''")


def column_cell_strings(
    rows: list[list[Any]],
    col_idx: int,
    column_formats: list[ColumnFormatSpec] | None,
    headers_len: int,
) -> list[str]:
    fmts = formats_or_auto(headers_len, column_formats)
    spec = fmts[col_idx]
    out: list[str] = []
    for row in rows:
        v = row[col_idx] if col_idx < len(row) else None
        out.append(format_cell_as_string(v, spec))
    return out


def _row_matches_filter(cell_str: str, op: str, needle: str) -> bool:
    """Confronto sul valore già formattato come stringa (come in anteprima)."""
    needle = needle.strip()
    if op == "eq":
        return cell_str == needle
    if op == "ne":
        return cell_str != needle
    if op == "contains":
        return needle in cell_str
    if op == "not_contains":
        return needle not in cell_str
    if op == "starts":
        return cell_str.startswith(needle)
    if op == "ends":
        return cell_str.endswith(needle)
    if op == "empty":
        return cell_str == ""
    if op == "not_empty":
        return cell_str != ""
    return True


def filter_data_rows(
    rows: list[list[Any]],
    conditions: list[tuple[int, str, str]],
    *,
    column_formats: list[ColumnFormatSpec] | None = None,
    headers_len: int,
) -> list[list[Any]]:
    """Restituisce solo le righe per cui tutte le condizioni (AND) sono vere.

    ``conditions``: ``(indice_colonna, op, valore_testo)`` — ``op`` come in
    :func:`_row_matches_filter`. Valore cella = :func:`format_cell_as_string`.
    """
    if not conditions:
        return rows
    fmts = formats_or_auto(headers_len, column_formats)
    out: list[list[Any]] = []
    for row in rows:
        ok = True
        for col_idx, op, needle in conditions:
            if col_idx < 0 or col_idx >= headers_len:
                ok = False
                break
            v = row[col_idx] if col_idx < len(row) else None
            s = format_cell_as_string(v, fmts[col_idx])
            if not _row_matches_filter(s, op, needle):
                ok = False
                break
        if ok:
            out.append(row)
    return out


def concat_column_with_separator(
    rows: list[list[Any]],
    col_idx: int,
    separator: str,
    *,
    column_formats: list[ColumnFormatSpec] | None = None,
    headers_len: int,
) -> str:
    parts = [s for s in column_cell_strings(rows, col_idx, column_formats, headers_len) if s]
    return separator.join(parts)


def build_sql_in_clause(values: Iterable[str], vertical: bool, use_quotes: bool) -> str:
    vals = list(values)
    if not vals:
        return "IN ()"
    parts: list[str] = []
    for v in vals:
        if use_quotes:
            parts.append(f"'{sql_escape_single_quotes(v)}'")
        else:
            parts.append(v)
    if not vertical:
        return "IN (" + ",".join(parts) + ")"
    if len(parts) == 1:
        return f"IN ({parts[0]})"
    lines = [f"IN ({parts[0]},", *[f"{p}," for p in parts[1:-1]]]
    lines.append(parts[-1])
    lines.append(")")
    return "\n".join(lines)


def concat_column_sql_in(
    rows: list[list[Any]],
    col_idx: int,
    vertical: bool,
    use_quotes: bool,
    *,
    column_formats: list[ColumnFormatSpec] | None = None,
    headers_len: int,
) -> str:
    vals = [s for s in column_cell_strings(rows, col_idx, column_formats, headers_len) if s]
    return build_sql_in_clause(vals, vertical, use_quotes)


def concat_rows_merge_columns(
    rows: list[list[Any]],
    col_indices: list[int],
    separator: str,
    *,
    column_formats: list[ColumnFormatSpec] | None = None,
    headers_len: int,
) -> list[str]:
    """Unisce le colonne nell'ordine indicato; valori null/vuoti sono omessi (nessun separatore extra)."""
    fmts = formats_or_auto(headers_len, column_formats)
    out: list[str] = []
    for row in rows:
        parts: list[str] = []
        for j in col_indices:
            v = row[j] if j < len(row) else None
            s = format_cell_as_string(v, fmts[j])
            if s:
                parts.append(s)
        out.append(separator.join(parts))
    return out
