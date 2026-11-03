#!/usr/bin/env python3
"""Read an excel-lite workbook with openpyxl and assert the values.

Independent check that what we write is a real .xlsx, not just something our
own reader happens to accept:

    node -e "const {writeSheet}=await import('./dist/index.js'); await writeSheet('/tmp/sample.xlsx',[{name:'Muneeb',amount:1234.5,when:new Date(Date.UTC(2024,4,17)),active:true}],{columns:[{key:'name'},{key:'amount',format:'#,##0.00'},{key:'when',format:'yyyy-mm-dd'},{key:'active'}]})" --input-type=module
    python3 scripts/verify-with-openpyxl.py /tmp/sample.xlsx
"""

from __future__ import annotations

import datetime as dt
import sys

from openpyxl import load_workbook


def main(path: str) -> int:
    workbook = load_workbook(path)
    print("sheets:", workbook.sheetnames)
    sheet = workbook[workbook.sheetnames[0]]
    print("dimensions:", sheet.dimensions, "max_row:", sheet.max_row, "max_col:", sheet.max_column)

    problems: list[str] = []

    def expect(ref: str, expected: object) -> None:
        actual = sheet[ref].value
        if actual != expected:
            problems.append(f"{ref}: expected {expected!r}, got {actual!r}")
        else:
            print(f"  {ref} = {actual!r}")

    expect("A1", "name")
    expect("B1", "amount")
    expect("C1", "when")
    expect("A2", "Muneeb")
    expect("B2", 1234.5)
    when = sheet["C2"].value
    if isinstance(when, dt.datetime) and when.date() == dt.date(2024, 5, 17):
        print(f"  C2 = {when!r} (date ok)")
    else:
        problems.append(f"C2: expected 2024-05-17, got {when!r}")

    if sheet["B2"].number_format != "#,##0.00":
        problems.append(f"B2 number_format: {sheet['B2'].number_format!r}")
    else:
        print("  B2 format = #,##0.00")

    if problems:
        print("\nFAILED")
        for problem in problems:
            print(" -", problem)
        return 1
    print("\nOK: openpyxl reads every value excel-lite wrote")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else "/tmp/sample.xlsx"))
