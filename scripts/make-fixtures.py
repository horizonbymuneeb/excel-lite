#!/usr/bin/env python3
"""Generate the third-party fixtures in fixtures/ with openpyxl + hand-written OOXML.

The point of these files is independence: excel-lite's tests read workbooks
that were produced by a completely different implementation, so a shared bug
in our own writer cannot hide a reader bug.

Most fixtures come from openpyxl. `handmade-cells.xlsx` is assembled directly
from OOXML part XML because the cell types it covers (`t="str"` cached formula
results, `t="d"` ISO dates, out-of-range shared-string indexes, `<row>` elements
with no cells, `<rowBreaks>`) are things no high-level writer will emit for us.

    python3 scripts/make-fixtures.py

Requires: pip install openpyxl
"""

from __future__ import annotations

import datetime as dt
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import CALENDAR_MAC_1904

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / "fixtures"
FIXTURES.mkdir(exist_ok=True)

# Fixed timestamps so the checked-in fixtures are byte-for-byte reproducible; CI fails
# when `python3 scripts/make-fixtures.py` changes anything under fixtures/.
FIXED_TIME = (2024, 1, 1, 0, 0, 0)
FIXED_STAMP = dt.datetime(2024, 1, 1, 0, 0, 0)


def _fix_properties(workbook: Workbook) -> None:
    workbook.properties.creator = "excel-lite fixtures"
    workbook.properties.lastModifiedBy = "excel-lite fixtures"
    workbook.properties.created = FIXED_STAMP
    workbook.properties.modified = FIXED_STAMP


def _normalize_zip(path: Path) -> None:
    """Rewrite a workbook with sorted entries and fixed timestamps."""
    with zipfile.ZipFile(path) as source:
        entries = [(info.filename, source.read(info.filename)) for info in source.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target:
        for name, payload in sorted(entries):
            info = zipfile.ZipInfo(name, FIXED_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            target.writestr(info, payload)


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

    _fix_properties(workbook)
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
    _fix_properties(workbook)
    workbook.save(FIXTURES / "openpyxl-formats.xlsx")


def mac_dates() -> None:
    """A 1904 date-system workbook (Excel for Mac's old epoch)."""
    workbook = Workbook()
    workbook.epoch = CALENDAR_MAC_1904
    sheet = workbook.active
    sheet.title = "Mac"
    sheet.append(["label", "date", "datetime", "time"])
    sheet.append(["aligned", dt.date(2024, 1, 2), dt.datetime(2024, 1, 2, 3, 4, 5), dt.time(6, 30)])
    sheet.append(["epoch", dt.date(1904, 1, 1), dt.datetime(1904, 1, 1), dt.time(0, 0)])
    sheet.column_dimensions["B"].width = 14
    _fix_properties(workbook)
    workbook.save(FIXTURES / "openpyxl-1904.xlsx")


def shapes() -> None:
    """Rich-text runs, merged cells, sparse cells, an error cell and a boolean."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Shapes"
    sheet["A1"] = CellRichText(
        [TextBlock(InlineFont(b=True), "Bold "), "plain ", TextBlock(InlineFont(i=True), "italic")]
    )
    sheet["B1"] = "plain"
    sheet["C1"] = "#DIV/0!"  # openpyxl writes this as t="e"
    sheet["D1"] = True
    sheet["E1"] = 12.5
    sheet.append(["x", None, "z"])  # B2 is genuinely absent -> sparse row
    sheet.merge_cells("A3:C3")
    sheet["A3"] = "merged block"  # only the top-left cell holds a value
    sheet.append(["only"])
    _fix_properties(workbook)
    workbook.save(FIXTURES / "openpyxl-shapes.xlsx")


def non_dates() -> None:
    """Custom number formats that contain letters but are NOT dates, plus time-only cells."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Formats"
    sheet.append(["label", "money", "pct", "text", "general", "quoted", "red", "literal", "time", "elapsed"])
    sheet.append(["a", 1234.5, 0.125, 12345, 42, 42, 42, 42, dt.time(6, 30), dt.time(18, 45)])
    sheet.append(["b", -20, 1.0, 7, -3.25, 99, -1, 3, dt.time(0, 0), dt.time(1, 30)])
    # openpyxl maps these to builtin ids (4, 10, 49, 0, 46, ...) or custom 164+.
    formats = {
        "B": "#,##0.00",
        "C": "0.00%",
        "D": "@",
        "E": "General",
        "F": '"kg"#,##0',
        "G": "[Red]0.00",
        "H": '"days"0',
        "I": "hh:mm:ss",
        "J": "[h]:mm:ss",
    }
    for column, code in formats.items():
        for row in range(2, 4):
            sheet[f"{column}{row}"].number_format = code
    _fix_properties(workbook)
    workbook.save(FIXTURES / "openpyxl-nondates.xlsx")


def serial_dates() -> None:
    """Raw serial numbers styled as dates, across the 1900 leap-year bug."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Serials"
    sheet.append(["serial", "as_date"])
    for serial in (0, 1, 2, 59, 60, 61, 36525):
        sheet.append([serial, serial])
        sheet.cell(row=sheet.max_row, column=2).number_format = "yyyy-mm-dd"
    _fix_properties(workbook)
    workbook.save(FIXTURES / "openpyxl-serial-dates.xlsx")


def header_edges() -> None:
    """Duplicate, blank and numeric headers, plus rows that are longer/shorter than the header."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Headers"
    sheet.append(["name", "name", 123, None, "name"])
    sheet.append(["a", "b", "c", "d", "e", "f"])  # one column past the header
    sheet.append(["g"])  # shorter than the header
    _fix_properties(workbook)
    workbook.save(FIXTURES / "openpyxl-headers.xlsx")


def wide() -> None:
    """Two hundred columns in both the header and the data rows."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Wide"
    sheet.append([f"c{index}" for index in range(200)])
    sheet.append([index for index in range(200)])
    sheet.append([index * 1.5 - 10 for index in range(200)])
    _fix_properties(workbook)
    workbook.save(FIXTURES / "openpyxl-wide.xlsx")


# ---------------------------------------------------------------------------
# hand-written OOXML: cell types no high-level writer produces on request
# ---------------------------------------------------------------------------

RAW_SHEET = (
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<dimension ref="A1:N7"/><sheetData>'
    '<row r="1">'
    '<c r="A1" t="s"><v>0</v></c>'
    '<c r="B1" t="s"><v>1</v></c>'
    '<c r="C1" t="inlineStr"><is><t>inline</t></is></c>'
    '<c r="D1" t="str"><f>CONCAT("ca","ched")</f><v>cached</v></c>'
    '<c r="E1" t="e"><v>#N/A</v></c>'
    '<c r="F1" t="d"><v>2024-05-17T12:34:56Z</v></c>'
    '<c r="G1" t="d"><v>2024-05-17T18:04:56+05:30</v></c>'
    '<c r="H1" t="b"><v>1</v></c>'
    '<c r="I1" t="b"><v>0</v></c>'
    '<c r="J1" t="b"><v>true</v></c>'
    '<c r="K1" t="s"><v>2</v></c>'
    '<c r="L1" t="s"><v>99</v></c>'
    '<c r="M1" t="n"><v></v></c>'
    '<c r="N1"><v>42</v></c>'
    "</row>"
    '<row r="2"><c r="A2"><v>1</v></c><c r="C2"><v>3</v></c></row>'
    '<row r="3"><c r="A3" t="s"><v>0</v></c><c r="B3"/><c r="C3"/></row>'
    '<row r="4"/>'
    '<row r="5"><c r="A5" t="n"><v>5</v></c><c r="B5"/><c r="C5" t="n"><v></v></c></row>'
    '<row r="6"><c r="A6" t="inlineStr"><is><r><t>a</t></r><r><t>b</t></r></is></c></row>'
    '<row r="7"><c r="A7" t="s"><v>0</v></c><c r="B7"/></row>'
    "</sheetData>"
    '<rowBreaks count="1" manualBreakCount="1"><brk id="2" max="16383" man="1"/></rowBreaks>'
    "</worksheet>"
)

ROWS_SHEET = (
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<dimension ref="A1:A4"/><sheetData>'
    '<row r="1"><c r="A1"><v>1</v></c></row>'
    '<row r="2"/>'
    '<row r="4"><c r="A4"><v>4</v></c></row>'
    "</sheetData></worksheet>"
)

SHARED_STRINGS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">'
    "<si><t>shared</t></si>"
    '<si><r><rPr><b/></rPr><t xml:space="preserve">rich </t></r><r><t>text</t></r></si>'
    "<si><t>a &amp; b &lt; c &gt; d</t></si>"
    "</sst>"
)

STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<numFmts count="0"/>'
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    "</styleSheet>"
)

WORKBOOK = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<bookViews><workbookView activeTab="0"/></bookViews>'
    '<sheets><sheet name="Raw" sheetId="1" r:id="rId1"/>'
    '<sheet name="Rows" sheetId="2" r:id="rId2"/></sheets></workbook>'
)

WORKBOOK_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
    "</Relationships>"
)

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    "</Types>"
)


def handmade() -> None:
    """Cell types that only appear in hand-written OOXML (t="str", t="d", empty <row>…)."""
    files = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>"
        ),
        "xl/workbook.xml": WORKBOOK,
        "xl/_rels/workbook.xml.rels": WORKBOOK_RELS,
        "xl/sharedStrings.xml": SHARED_STRINGS,
        "xl/styles.xml": STYLES,
        "xl/worksheets/sheet1.xml": RAW_SHEET,
        "xl/worksheets/sheet2.xml": ROWS_SHEET,
    }
    target = FIXTURES / "handmade-cells.xlsx"
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, body in sorted(files.items()):
            info = zipfile.ZipInfo(name, FIXED_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            archive.writestr(info, body.encode("utf-8"))


if __name__ == "__main__":
    mixed()
    formats()
    mac_dates()
    shapes()
    non_dates()
    serial_dates()
    header_edges()
    wide()
    handmade()
    for path in sorted(FIXTURES.glob("*.xlsx")):
        _normalize_zip(path)
    for path in sorted(FIXTURES.glob("*.xlsx")):
        print(f"{path.name}: {path.stat().st_size} bytes")
