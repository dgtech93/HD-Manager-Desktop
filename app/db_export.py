"""
Export / import dati tabellari dal database SQLite (backup selettivo, più formati).
"""

from __future__ import annotations

import csv
import json
import re
import sqlite3
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from xml.dom import minidom

EXPORT_FORMAT_JSON = "json"
EXPORT_FORMAT_XML = "xml"
EXPORT_FORMAT_XLSX = "xlsx"
EXPORT_FORMAT_CSV = "csv"

EXPORT_JSON_MAGIC = "HDManagerDbExport"
EXPORT_JSON_VERSION = 1

# Nomi tecnici (SQLite) → etichette per l’interfaccia in italiano
TABLE_LABELS_IT: dict[str, str] = {
    "app_settings": "Impostazioni applicazione",
    "archive_favorites": "Preferiti archivio",
    "archive_files": "File in archivio",
    "archive_folders": "Cartelle archivio",
    "archive_links": "Link in archivio",
    "client_contacts": "Contatti clienti",
    "client_notes": "Note clienti",
    "client_pending_relations": "Relazioni clienti in attesa",
    "client_resources": "Associazioni cliente–risorsa",
    "clients": "Clienti",
    "competences": "Competenze",
    "environments": "Ambienti",
    "product_clients": "Associazioni prodotto–cliente",
    "product_credential_environments": "Credenziali per ambiente",
    "product_credentials": "Credenziali prodotto",
    "product_environments": "Associazioni prodotto–ambiente",
    "product_types": "Tipi prodotto",
    "products": "Prodotti",
    "releases": "Release",
    "resources": "Risorse",
    "roles": "Ruoli",
    "tags": "Tag",
    "vpns": "VPN",
}


def table_label_it(internal_name: str) -> str:
    """Etichetta italiana per il nome tabella SQLite; se non mappata, restituisce il nome tecnico."""
    return TABLE_LABELS_IT.get(internal_name, internal_name)


# Inserimento rispettando dipendenze FK (genitori prima dei figli)
_TABLE_INSERT_ORDER: list[str] = [
    "competences",
    "releases",
    "product_types",
    "environments",
    "roles",
    "vpns",
    "resources",
    "clients",
    "client_resources",
    "products",
    "product_clients",
    "product_environments",
    "product_credentials",
    "product_credential_environments",
    "tags",
    "archive_folders",
    "archive_files",
    "archive_links",
    "archive_favorites",
    "client_contacts",
    "client_notes",
    "client_pending_relations",
    "app_settings",
]


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _safe_table_name(name: str) -> bool:
    return bool(name) and re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name) is not None


def list_exportable_tables(conn: sqlite3.Connection) -> list[str]:
    rows = conn.execute(
        """
        SELECT name FROM sqlite_master
        WHERE type='table' AND name NOT LIKE 'sqlite_%'
        ORDER BY name COLLATE NOCASE
        """
    ).fetchall()
    return [str(r[0]) for r in rows if _safe_table_name(str(r[0]))]


def fetch_table_rows(
    conn: sqlite3.Connection, table: str
) -> tuple[list[str], list[dict[str, Any]]]:
    if not _safe_table_name(table):
        raise ValueError(f"Nome tabella non valido: {table}")
    cur = conn.execute(f'SELECT * FROM "{table}"')
    rows = cur.fetchall()
    if not rows:
        return [], []
    cols = list(rows[0].keys())
    out = [dict(zip(cols, tuple(r))) for r in rows]
    return cols, out


def _xlsx_sheet_title(name: str) -> str:
    t = name[:31]
    for bad in ("\\", "/", "?", "*", "[", "]", ":", "'"):
        t = t.replace(bad, "_")
    return t or "Foglio"


def export_tables_xlsx(path: Path, conn: sqlite3.Connection, tables: list[str]) -> None:
    from openpyxl import Workbook

    if not tables:
        raise ValueError("Seleziona almeno una tabella.")
    path = Path(path)
    wb = Workbook()
    first = True
    for table in tables:
        cols, data_rows = fetch_table_rows(conn, table)
        sheet_title = _xlsx_sheet_title(table_label_it(table))
        if first:
            ws = wb.active
            ws.title = sheet_title
            first = False
        else:
            ws = wb.create_sheet(title=sheet_title)
        if cols:
            ws.append(list(cols))
        for r in data_rows:
            ws.append([r.get(c) for c in cols])
    wb.save(str(path))


def export_tables_json(path: Path, conn: sqlite3.Connection, tables: list[str]) -> None:
    if not tables:
        raise ValueError("Seleziona almeno una tabella.")
    payload: dict[str, Any] = {
        "format": EXPORT_JSON_MAGIC,
        "version": EXPORT_JSON_VERSION,
        "generated_at": _now_iso(),
        "tables": {},
    }
    for t in tables:
        _cols, rows = fetch_table_rows(conn, t)
        payload["tables"][t] = rows
    path = Path(path)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def export_tables_xml(path: Path, conn: sqlite3.Connection, tables: list[str]) -> None:
    if not tables:
        raise ValueError("Seleziona almeno una tabella.")
    root = ET.Element(
        "database_export",
        version="1",
        generated=_now_iso(),
    )
    for table in tables:
        _cols, rows = fetch_table_rows(conn, table)
        tel = ET.SubElement(root, "table", name=table)
        for row in rows:
            rel = ET.SubElement(tel, "row")
            for k, v in row.items():
                col_el = ET.SubElement(rel, "col", name=str(k))
                col_el.text = "" if v is None else str(v)
    rough = ET.tostring(root, encoding="unicode")
    parsed = minidom.parseString(rough)
    pretty = parsed.toprettyxml(indent="  ", encoding="utf-8")
    Path(path).write_bytes(pretty)


def export_tables_csv(path: Path, conn: sqlite3.Connection, tables: list[str]) -> None:
    if not tables:
        raise ValueError("Seleziona almeno una tabella.")
    path = Path(path)
    if len(tables) == 1:
        cols, data_rows = fetch_table_rows(conn, tables[0])
        with path.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=";")
            if cols:
                w.writerow(cols)
            for r in data_rows:
                w.writerow([r.get(c) for c in cols])
        return

    path.mkdir(parents=True, exist_ok=True)
    for table in tables:
        cols, data_rows = fetch_table_rows(conn, table)
        fp = path / f"{table}.csv"
        with fp.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=";")
            if cols:
                w.writerow(cols)
            for r in data_rows:
                w.writerow([r.get(c) for c in cols])


def run_export(
    *,
    conn: sqlite3.Connection,
    tables: list[str],
    fmt: str,
    target_path: str | Path,
) -> None:
    """Esegue l’export nel formato scelto."""
    target_path = Path(target_path)
    tset = [t for t in tables if _safe_table_name(t)]
    if not tset:
        raise ValueError("Nessuna tabella valida selezionata.")
    if fmt == EXPORT_FORMAT_XLSX:
        export_tables_xlsx(target_path, conn, tset)
    elif fmt == EXPORT_FORMAT_JSON:
        export_tables_json(target_path, conn, tset)
    elif fmt == EXPORT_FORMAT_XML:
        export_tables_xml(target_path, conn, tset)
    elif fmt == EXPORT_FORMAT_CSV:
        export_tables_csv(target_path, conn, tset)
    else:
        raise ValueError(f"Formato non supportato: {fmt}")


def _insert_order_for_tables(tables: set[str]) -> list[str]:
    ordered = [t for t in _TABLE_INSERT_ORDER if t in tables]
    rest = sorted(tables - set(ordered), key=str.lower)
    return ordered + rest


def _table_column_names(conn: sqlite3.Connection, table: str) -> set[str]:
    if not _safe_table_name(table):
        return set()
    rows = conn.execute(f'PRAGMA table_info("{table}")').fetchall()
    return {str(r[1]) for r in rows}


def import_tables_from_bundle(conn: sqlite3.Connection, tables_data: dict[str, list[dict[str, Any]]]) -> None:
    """Importa righe da un bundle JSON (chiave = nome tabella). FK disattivate durante l’operazione."""
    if not tables_data:
        raise ValueError("Nessuna tabella nel file.")
    names = {t for t in tables_data if _safe_table_name(t)}
    insert_order = _insert_order_for_tables(names)
    conn.execute("PRAGMA foreign_keys=OFF")
    try:
        for t in names:
            conn.execute(f'DELETE FROM "{t}"')
        for t in insert_order:
            rows = tables_data.get(t) or []
            if not rows:
                continue
            valid_cols = _table_column_names(conn, t)
            if not valid_cols:
                continue
            for row in rows:
                if not isinstance(row, dict):
                    continue
                cols = [c for c in row.keys() if str(c) in valid_cols]
                if not cols:
                    continue
                vals = [row.get(c) for c in cols]
                ph = ",".join(["?"] * len(cols))
                qcols = ",".join(f'"{c}"' for c in cols)
                conn.execute(f'INSERT INTO "{t}" ({qcols}) VALUES ({ph})', vals)
        conn.commit()
    finally:
        conn.execute("PRAGMA foreign_keys=ON")


def import_tables_from_xml(conn: sqlite3.Connection, path: Path) -> None:
    tree = ET.parse(path)
    root = tree.getroot()
    if root.tag != "database_export":
        raise ValueError("XML non valido: la radice deve essere «database_export».")
    tables_data: dict[str, list[dict[str, Any]]] = {}
    for tel in root.findall("table"):
        tname = tel.get("name")
        if not tname or not _safe_table_name(tname):
            continue
        rows_out: list[dict[str, Any]] = []
        for rel in tel.findall("row"):
            row: dict[str, Any] = {}
            for col_el in rel.findall("col"):
                cname = col_el.get("name")
                if cname:
                    row[str(cname)] = col_el.text
            rows_out.append(row)
        tables_data[tname] = rows_out
    import_tables_from_bundle(conn, tables_data)


def import_tables_from_json_file(conn: sqlite3.Connection, path: Path) -> None:
    raw = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(raw, dict) and raw.get("format") == EXPORT_JSON_MAGIC and "tables" in raw:
        td = raw["tables"]
        if not isinstance(td, dict):
            raise ValueError("Struttura JSON non valida.")
        tables_data = {k: v for k, v in td.items() if isinstance(v, list)}
        import_tables_from_bundle(conn, tables_data)
        return
    raise ValueError(
        "Il JSON non è un backup esportato dall’app (struttura «format» / «tables» non trovata)."
    )


def import_tables_from_csv_file(conn: sqlite3.Connection, path: Path) -> None:
    """Importa un singolo CSV in una tabella: nome tabella = stem del file (es. clients.csv → clients)."""
    stem = path.stem
    if not _safe_table_name(stem):
        raise ValueError(
            f"Il nome file deve coincidere con il nome tecnico della tabella (es. clients.csv). "
            f"Nome ricevuto: {stem}"
        )
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        sample = f.read(4096)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=";,")
        except csv.Error:
            dialect = csv.excel
            dialect.delimiter = ";"
        reader = csv.reader(f, dialect)
        rows_list = list(reader)
    if not rows_list:
        raise ValueError("CSV vuoto.")
    headers = [h.strip() for h in rows_list[0]]
    data_rows: list[dict[str, Any]] = []
    for parts in rows_list[1:]:
        if len(parts) < len(headers):
            parts = parts + [""] * (len(headers) - len(parts))
        data_rows.append({headers[i]: parts[i] if i < len(parts) else "" for i in range(len(headers))})
    import_tables_from_bundle(conn, {stem: data_rows})
