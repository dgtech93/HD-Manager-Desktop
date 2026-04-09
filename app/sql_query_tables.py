"""Estrazione e sostituzione nomi di tabella da testo SQL (euristiche SQLite-like)."""

from __future__ import annotations

import re
from typing import Dict, List, Tuple


# Parole chiave da non considerare come nomi tabella
_SQL_KW = frozenset(
    {
        "SELECT",
        "WHERE",
        "FROM",
        "JOIN",
        "LEFT",
        "RIGHT",
        "INNER",
        "OUTER",
        "CROSS",
        "FULL",
        "NATURAL",
        "ON",
        "AND",
        "OR",
        "NOT",
        "IN",
        "IS",
        "AS",
        "BY",
        "GROUP",
        "ORDER",
        "HAVING",
        "LIMIT",
        "OFFSET",
        "UNION",
        "ALL",
        "DISTINCT",
        "CASE",
        "WHEN",
        "THEN",
        "ELSE",
        "END",
        "NULL",
        "TRUE",
        "FALSE",
        "WITH",
        "RECURSIVE",
        "EXISTS",
        "BETWEEN",
        "LIKE",
        "GLOB",
        "ESCAPE",
        "CAST",
        "COLLATE",
        "INSERT",
        "INTO",
        "VALUES",
        "UPDATE",
        "SET",
        "DELETE",
        "CREATE",
        "DROP",
        "ALTER",
        "TABLE",
        "INDEX",
        "VIEW",
        "TRIGGER",
        "BEGIN",
        "COMMIT",
        "ROLLBACK",
        "PRAGMA",
        "WITHOUT",
        "ROWID",
        "PRIMARY",
        "KEY",
        "FOREIGN",
        "REFERENCES",
        "DEFAULT",
        "CONSTRAINT",
        "CHECK",
        "UNIQUE",
        "MERGE",
    }
)

_ID = r"(?:`[^`]+`|\"[^\"]+\"|\[[^\]]+\]|[\w.]+)"


def _strip_sql_comments(sql: str) -> str:
    out: list[str] = []
    i = 0
    n = len(sql)
    while i < n:
        if i + 1 < n and sql[i : i + 2] == "--":
            while i < n and sql[i] not in "\r\n":
                i += 1
            continue
        if i + 1 < n and sql[i : i + 2] == "/*":
            j = sql.find("*/", i + 2)
            if j == -1:
                break
            i = j + 2
            continue
        out.append(sql[i])
        i += 1
    return "".join(out)


def _normalize_identifier(raw: str) -> str:
    raw = raw.strip()
    if len(raw) >= 2 and raw[0] == raw[-1] == "`":
        return raw[1:-1]
    if len(raw) >= 2 and raw[0] == raw[-1] == '"':
        return raw[1:-1]
    if len(raw) >= 2 and raw[0] == "[" and raw[-1] == "]":
        return raw[1:-1]
    return raw


def _is_keyword(name: str) -> bool:
    return name.strip().upper() in _SQL_KW


def _split_top_level_commas(s: str) -> list[str]:
    parts: list[str] = []
    buf: list[str] = []
    depth = 0
    in_single = False
    in_double = False
    for ch in s:
        if ch == "'" and not in_double:
            in_single = not in_single
        elif ch == '"' and not in_single:
            in_double = not in_double
        if not in_single and not in_double:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            if ch == "," and depth == 0:
                parts.append("".join(buf))
                buf = []
                continue
        buf.append(ch)
    if buf:
        parts.append("".join(buf))
    return [p.strip() for p in parts if p.strip()]


def _first_table_token(segment: str) -> str | None:
    seg = segment.strip()
    if not seg or seg.startswith("("):
        return None
    m = re.match(rf"^({_ID})", seg, re.IGNORECASE)
    if not m:
        return None
    name = _normalize_identifier(m.group(1))
    if _is_keyword(name):
        return None
    return name


def _extract_from_clause_body(sql: str) -> str | None:
    m = re.search(r"\bFROM\b", sql, re.IGNORECASE)
    if not m:
        return None
    start = m.end()
    i = start
    depth = 0
    in_single = False
    in_double = False
    buf: list[str] = []
    n = len(sql)
    kw_stop = re.compile(
        r"^\s*(WHERE|GROUP\s+BY|ORDER\s+BY|HAVING|LIMIT|UNION|EXCEPT|INTERSECT)\b",
        re.IGNORECASE | re.DOTALL,
    )
    while i < n:
        c = sql[i]
        if c == "'" and not in_double:
            in_single = not in_single
        elif c == '"' and not in_single:
            in_double = not in_double
        if not in_single and not in_double:
            if c == "(":
                depth += 1
            elif c == ")":
                depth -= 1
            if depth == 0:
                rest = sql[i:]
                if kw_stop.match(rest):
                    break
        buf.append(c)
        i += 1
    return "".join(buf).strip() or None


# JOIN con prefisso opzionale; include anche "JOIN" senza INNER/LEFT (inner implicito)
_JOIN_TABLE_RE = re.compile(
    rf"\b(?:(?:INNER|CROSS|NATURAL|(?:LEFT|RIGHT|FULL)(?:\s+OUTER)?)\s+)?JOIN\s+({_ID})",
    re.IGNORECASE | re.DOTALL,
)

# Ogni «FROM identificatore» (anche dentro SELECT figlie / EXISTS / APPLY). Non cattura «FROM (».
_FROM_TABLE_RE = re.compile(rf"\bFROM\s+({_ID})", re.IGNORECASE | re.DOTALL)

# Primo segmento senza «.» finale (altrimenti [\w.]+ assorbe «GDM.» e non matcha «GDM.[a]»).
_QUAL_NO_DOT = r"(?:`[^`]+`|\"[^\"]+\"|\[[^\]]+\]|[A-Za-z_][\w]*)"

# alias.tabella o [Tab].[Campo] (secondo segmento: anche schema.tabella come nome colonna se quotato)
# Nessun «\b» finale: dopo «.[a]» segue spesso «,» o «)» e «\b» tra «]» e «,» non vale.
_QUALIFIED_COL_RE = re.compile(
    rf"\b({_QUAL_NO_DOT})\s*\.\s*({_ID})",
    re.IGNORECASE | re.DOTALL,
)

_JOIN_ON_ALIAS_RE = re.compile(
    rf"\bJOIN\s+{_ID}\s+(?:AS\s+)?({_ID})(?=\s+ON\b)",
    re.IGNORECASE | re.DOTALL,
)


def _tables_after_joins_in_fragment(fragment: str) -> list[str]:
    """Tabelle citate dopo JOIN / INNER JOIN / LEFT JOIN / … (anche JOIN senza prefisso)."""
    out: list[str] = []
    for m in _JOIN_TABLE_RE.finditer(fragment):
        name = _normalize_identifier(m.group(1))
        if name and not _is_keyword(name):
            out.append(name)
    return out


def _tables_from_from_clause_body(body: str) -> list[str]:
    """
    Per ogni segmento separato da virgola (join implicito): prima tabella + tutte dopo JOIN.
    """
    found: list[str] = []
    for part in _split_top_level_commas(body):
        t = _first_table_token(part)
        if t:
            found.append(t)
        found.extend(_tables_after_joins_in_fragment(part))
    return found


def extract_sql_table_names(sql: str) -> list[str]:
    """
    Elenco ordinato di nomi di tabella probabili (FROM, JOIN, UPDATE, INSERT INTO, DELETE FROM).
    Include anche FROM nelle sottoquery; deduplica per nome.
    """
    s = _strip_sql_comments(sql or "")
    if not s.strip():
        return []

    found: list[str] = []

    def add(name: str | None) -> None:
        if not name:
            return
        n = name.strip()
        if not n or _is_keyword(n):
            return
        found.append(n)

    # Clauses FROM … (prima tabella per segmento + tutte le tabelle dopo JOIN nello stesso frammento)
    body = _extract_from_clause_body(s)
    if body:
        for t in _tables_from_from_clause_body(body):
            add(t)

    # JOIN ovunque nel testo (backup e query senza FROM “pulito”)
    for m in _JOIN_TABLE_RE.finditer(s):
        add(_normalize_identifier(m.group(1)))

    # FROM <tabella> in query annidate (prima tabella di ogni clausola FROM)
    for m in _FROM_TABLE_RE.finditer(s):
        add(_normalize_identifier(m.group(1)))

    # UPDATE …
    for m in re.finditer(rf"\bUPDATE\s+({_ID})", s, re.IGNORECASE):
        add(_normalize_identifier(m.group(1)))

    # INSERT INTO …
    for m in re.finditer(rf"\bINSERT\s+INTO\s+({_ID})", s, re.IGNORECASE):
        add(_normalize_identifier(m.group(1)))

    # DELETE FROM …
    for m in re.finditer(rf"\bDELETE\s+FROM\s+({_ID})", s, re.IGNORECASE):
        add(_normalize_identifier(m.group(1)))

    # Dedup preservando ordine
    seen: set[str] = set()
    out: list[str] = []
    for x in found:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


def apply_table_replacements(sql: str, mapping: Dict[str, str]) -> str:
    """
    Sostituisce identificatori interi (non sottostringhe).
    Chiavi con valore vuoto o uguale alla chiave vengono ignorate.
    """
    pairs = [(k, v) for k, v in mapping.items() if v is not None and str(v).strip() and str(v).strip() != k]
    if not pairs:
        return sql
    pairs.sort(key=lambda kv: len(kv[0]), reverse=True)
    out = sql
    for old, new in pairs:
        old_s = str(old).strip()
        new_s = str(new).strip()
        if not old_s:
            continue
        escaped = re.escape(old_s)
        pat = re.compile(r"(?<![\w.])" + escaped + r"(?![\w.])")
        out = pat.sub(new_s, out)
    return out


def apply_alias_replacements(sql: str, mapping: Dict[str, str]) -> str:
    """
    Sostituisce alias di tabella ovunque compaiano, incluso in «Alias.campo».
    A differenza di apply_table_replacements, consente il punto subito dopo l’alias (GDM.[x]).
    """
    pairs = [(k, v) for k, v in mapping.items() if v is not None and str(v).strip() and str(v).strip() != k]
    if not pairs:
        return sql
    pairs.sort(key=lambda kv: len(kv[0]), reverse=True)
    out = sql
    for old, new in pairs:
        old_s = str(old).strip()
        new_s = str(new).strip()
        if not old_s:
            continue
        escaped = re.escape(old_s)
        # (?![\w]): dopo l’alias può esserci «.» (campo qualificato); non sostituire parti di identificatori più lunghi.
        pat = re.compile(r"(?<![\w.])" + escaped + r"(?!\w)")
        out = pat.sub(new_s, out)
    return out


def _norm_ident_key(s: str) -> str:
    return _normalize_identifier(s).strip().lower()


def column_ref_storage_key(qualifier: str, column: str) -> str:
    return f"{_norm_ident_key(qualifier)}\x1f{_norm_ident_key(column)}"


def _looks_like_numeric_literal_pair(qualifier: str, column: str) -> bool:
    q = _normalize_identifier(qualifier).strip()
    c = _normalize_identifier(column).strip()
    if not q or not c:
        return False
    if q.isdigit() and c.isdigit():
        return True
    return False


def _clause_body_end(sql: str, body_start: int) -> int:
    """Fine corpo clausola WHERE/HAVING (profondità parentesi + stringhe)."""
    n = len(sql)
    i = body_start
    depth = 0
    in_single = False
    in_double = False
    while i < n:
        c = sql[i]
        if in_single:
            if c == "'" and i + 1 < n and sql[i + 1] == "'":
                i += 2
                continue
            if c == "'":
                in_single = False
            i += 1
            continue
        if in_double:
            if c == '"':
                in_double = False
            i += 1
            continue
        if c == "'":
            in_single = True
            i += 1
            continue
        if c == '"':
            in_double = True
            i += 1
            continue
        if c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
        if depth == 0:
            rest = sql[i:]
            if re.match(
                r"^\s*(ORDER\s+BY|GROUP\s+BY|HAVING|LIMIT|UNION|EXCEPT|INTERSECT|OFFSET|FETCH|FOR\b)\b",
                rest,
                re.IGNORECASE,
            ):
                return i
        i += 1
    return n


def extract_sql_where_having_spans(sql: str) -> List[Tuple[int, int]]:
    """Intervalli caratteri (start, end) dei corpi WHERE e HAVING (commenti già ignorati dal chiamante)."""
    s = _strip_sql_comments(sql or "")
    if not s.strip():
        return []
    spans: List[Tuple[int, int]] = []
    for kw in ("WHERE", "HAVING"):
        for m in re.finditer(rf"\b{kw}\b", s, re.IGNORECASE):
            body_start = m.end()
            body_end = _clause_body_end(s, body_start)
            if body_end > body_start:
                spans.append((body_start, body_end))
    return spans


def _offset_in_spans(pos: int, spans: Iterable[Tuple[int, int]]) -> bool:
    for a, b in spans:
        if a <= pos < b:
            return True
    return False


def extract_sql_qualified_column_refs(sql: str) -> List[dict]:
    """Riferimenti alias.tabella o [Tab].[Campo] con flag se compaiono in WHERE/HAVING."""
    s = _strip_sql_comments(sql or "")
    if not s.strip():
        return []
    spans = extract_sql_where_having_spans(s)
    seen: set[str] = set()
    out: List[dict] = []
    for m in _QUALIFIED_COL_RE.finditer(s):
        if _looks_like_numeric_literal_pair(m.group(1), m.group(2)):
            continue
        qual = _normalize_identifier(m.group(1))
        col = _normalize_identifier(m.group(2))
        key = column_ref_storage_key(qual, col)
        if key in seen:
            continue
        seen.add(key)
        in_wh = _offset_in_spans(m.start(), spans)
        out.append(
            {
                "key": key,
                "qualifier": qual,
                "column": col,
                "column_raw": m.group(2),
                "sample": m.group(0).strip(),
                "in_where_having": in_wh,
            }
        )
    return out


def extract_sql_alias_suggestions(sql: str) -> List[str]:
    """Possibili alias di tabella (qualificatori in a.b, JOIN … ON, FROM …)."""
    s = _strip_sql_comments(sql or "")
    if not s.strip():
        return []
    found: List[str] = []
    seen: set[str] = set()

    def add(raw: str) -> None:
        name = _normalize_identifier(raw).strip()
        if not name or _is_keyword(name):
            return
        low = name.lower()
        if low in seen:
            return
        seen.add(low)
        found.append(name)

    for m in _QUALIFIED_COL_RE.finditer(s):
        if _looks_like_numeric_literal_pair(m.group(1), m.group(2)):
            continue
        add(m.group(1))

    for m in _JOIN_ON_ALIAS_RE.finditer(s):
        add(m.group(1))

    for m in re.finditer(
        rf"\bFROM\s+{_ID}\s+(?:AS\s+)?({_ID})(?=\s*(?:WHERE|JOIN|INNER|LEFT|RIGHT|CROSS|OUTER|GROUP|ORDER|HAVING|LIMIT|UNION|,|\)|$))",
        s,
        re.IGNORECASE | re.DOTALL,
    ):
        add(m.group(1))

    return found


def apply_qualified_column_name_replacements(sql: str, mapping: Dict[str, str]) -> str:
    """
    Sostituisce solo il nome campo (secondo segmento) in «qualificatore.campo».
    mapping: chiave column_ref_storage_key(qual, col) → nuovo identificatore di colonna (es. [Nuovo nome]).
    Il qualificatore (alias) va modificato nella sezione Alias.
    """
    pairs = [(k, v) for k, v in mapping.items() if v is not None and str(v).strip()]
    if not pairs:
        return sql
    norm_map = {str(k).strip(): str(v).strip() for k, v in pairs}
    s = sql
    parts: List[str] = []
    last = 0
    for m in _QUALIFIED_COL_RE.finditer(s):
        if _looks_like_numeric_literal_pair(m.group(1), m.group(2)):
            parts.append(s[last : m.end()])
            last = m.end()
            continue
        qual = _normalize_identifier(m.group(1))
        col = _normalize_identifier(m.group(2))
        key = column_ref_storage_key(qual, col)
        if key not in norm_map:
            parts.append(s[last : m.end()])
            last = m.end()
            continue
        parts.append(s[last : m.start(2)])
        parts.append(norm_map[key])
        last = m.end()
    parts.append(s[last:])
    return "".join(parts)


def apply_gestione_replacements(
    sql: str,
    *,
    column_map: Dict[str, str] | None = None,
    table_map: Dict[str, str] | None = None,
    alias_map: Dict[str, str] | None = None,
) -> str:
    """
    Ordine: solo nome campo (dopo il punto) → tabelle → alias.
    """
    out = sql
    out = apply_qualified_column_name_replacements(out, column_map or {})
    out = apply_table_replacements(out, table_map or {})
    out = apply_alias_replacements(out, alias_map or {})
    return out
