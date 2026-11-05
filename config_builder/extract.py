"""Extract compact, referenced cell content without sending whole XLSX files."""
from datetime import date, datetime, time
from pathlib import Path
import warnings

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell


def extract(path: Path, sheets: list[str] | None = None) -> dict:
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        wb = load_workbook(path, data_only=False)
    try:
        unknown = set(sheets or []) - set(wb.sheetnames)
        if unknown:
            raise ValueError(f"Unknown sheets: {sorted(unknown)}")
        result = {"filename": path.name, "sheets": [], "warnings": [str(w.message) for w in caught]}
        for index, ws in enumerate(wb.worksheets):
            if sheets and ws.title not in sheets:
                continue
            cells = []
            # Iterate the loaded cell collection: avoids million-row styled ranges.
            for cell in sorted(ws._cells.values(), key=lambda c: (c.row, c.column)):
                if isinstance(cell, MergedCell) or cell.value is None:
                    continue
                value = cell.value
                if isinstance(value, str) and not value.strip():
                    continue
                cells.append({
                    "id": f"s{index}:{cell.coordinate}",
                    "address": cell.coordinate,
                    "value": value.isoformat() if isinstance(value, (date, datetime, time)) else value,
                    "type": cell.data_type,
                    "format": cell.number_format,
                    "bold": bool(cell.font.bold),
                })
            result["sheets"].append({
                "name": ws.title,
                "index": index,
                "merged_ranges": [str(r) for r in ws.merged_cells.ranges],
                "cells": cells,
                "images": len(ws._images),
                "charts": len(ws._charts),
            })
        return result
    finally:
        wb.close()



