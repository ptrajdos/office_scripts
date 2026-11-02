#!/usr/bin/env python3
"""Python port of column_with_spaces.vba (RemoveFirstWord).

Removes everything up to and including the first space in each cell of a column.
"""
import argparse
from openpyxl import load_workbook


def remove_first_word(workbook_path, sheet_name, column, output_path=None):
    wb = load_workbook(workbook_path)
    ws = wb[sheet_name]

    for row in range(1, ws.max_row + 1):
        cell = ws[f"{column}{row}"]
        value = cell.value
        if isinstance(value, str) and " " in value:
            cell.value = value[value.index(" ") + 1:]

    wb.save(output_path or workbook_path)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Remove the first word from each cell of a column in an XLSX file.")
    parser.add_argument("workbook", help="XLSX file to process")
    parser.add_argument("--sheet", default="Sheet1", help="Sheet name (default: Sheet1)")
    parser.add_argument("--column", default="A", help="Column letter (default: A)")
    parser.add_argument("--output", help="Output XLSX path (default: overwrite input)")
    args = parser.parse_args()
    remove_first_word(args.workbook, args.sheet, args.column.upper(), args.output)
