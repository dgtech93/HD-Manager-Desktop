"""Esportazione tabella in file .xlsx."""

from __future__ import annotations

from pathlib import Path
from typing import Any


def export_table_xlsx(path: str | Path, headers: list[str], rows: list[list[Any]]) -> None:
    from openpyxl import Workbook

    path = Path(path)
    wb = Workbook()
    ws = wb.active
    ws.append(list(headers))
    for r in rows:
        out = []
        for x in r:
            if x is None:
                out.append(None)
            elif isinstance(x, (int, float, bool)):
                out.append(x)
            else:
                out.append(str(x))
        ws.append(out)
    wb.save(str(path))
