#!/usr/bin/env python3
"""Generate the third-party fixtures in fixtures/ with openpyxl.

The point of these files is independence: excel-lite's tests read workbooks
that were produced by a completely different implementation, so a shared bug
in our own writer cannot hide a reader bug.

    python3 scripts/make-fixtures.py

Requires: pip install openpyxl
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / "fixtures"
FIXTURES.mkdir(exist_ok=True)


def mixed() -> None:
    """Values of every kind, plus a second sheet and a blank row."""
    workbook = Workbook()
    data = workbook.active
    data.title = "Data"
    data.append(["name", "qty", "price", "when", "active", "note"])
    data.append(["Ali", 3, 19.99, dt.date(2023, 7, 4), True, "plain text"])
    data.append(["Ünïcode ✓", 0, 0.5, dt.datetime(2022, 11, 30, 13, 45, 30), False, 'comma, "quote"'])
    data.append([])  # blank row: the reader must skip it
    data.append(["Nina", 1299, 1234567.891, dt.date(2024, 2, 29), True, "😀 emoji"])
    data.column_dimensions["A"].width = 18
    data.freeze_panes = "A2"

    notes = workbook.create_sheet("Notes")
    notes.append(["key", "value"])
    notes.append(["owner", "finance"])
    notes.append(["updated", dt.date(2024, 1, 15)])
    notes["A1"].font = Font(bold=True, color="FFFFFFFF")
    notes["A1"].fill = PatternFill("solid", fgColor="FF1F2937")
    notes["B2"].alignment = Alignment(horizontal="right")

    workbook.save(FIXTURES / "openpyxl-mixed.xlsx")


def formats() -> None:
    """Custom number formats and sparse cells."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Formats"
    rows = [
        ["label", "money", "pct", "date", "datetime", "sparse"],
        ["a", 1234.5, 0.125, dt.date(2020, 1, 1), dt.datetime(2020, 1, 1, 6, 30), "x"],
        ["b", -20, 1.0, dt.date(1999, 12, 31), dt.datetime(1999, 12, 31, 23, 59, 59), None],
        ["c", 0, 0.0, dt.date(2024, 2, 29), None, "z"],
    ]
    for row in rows:
        sheet.append(row)
    for row in sheet.iter_rows(min_row=2, min_col=2, max_col=2):
        for cell in row:
            cell.number_format = "#,##0.00"
    for row in sheet.iter_rows(min_row=2, min_col=3, max_col=3):
        for cell in row:
            cell.number_format = "0.00%"
    for row in sheet.iter_rows(min_row=2, min_col=4, max_col=4):
        for cell in row:
            cell.number_format = "yyyy-mm-dd"
    for row in sheet.iter_rows(min_row=2, min_col=5, max_col=5):
        for cell in row:
            cell.number_format = "yyyy-mm-dd hh:mm:ss"
    sheet.column_dimensions[get_column_letter(1)].width = 12
    workbook.save(FIXTURES / "openpyxl-formats.xlsx")


if __name__ == "__main__":
    mixed()
    formats()
    for path in sorted(FIXTURES.glob("*.xlsx")):
        print(f"{path.name}: {path.stat().st_size} bytes")
