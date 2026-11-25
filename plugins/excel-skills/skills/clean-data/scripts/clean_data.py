#!/usr/bin/env python3
"""Clean a messy sheet into a tidy table, and log every change.

    python clean_data.py INPUT [--sheet NAME] [--out FILE] [--month-first] [--keep-duplicates]
                         [--text-columns "Col A,Col B"] [--overwrite]

Reads .xlsx/.xlsm/.csv. Writes a NEW workbook (never changes the input) with two sheets:
  Cleaned  the tidy table: real numbers and dates, trimmed text, PK mobiles as 0300-1234567,
           exact duplicate rows and empty rows removed, title rows above the header dropped
  Changes  one line per change: original row, column, old value, new value, reason
Prints a JSON summary for the assistant to explain.
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import Workbook  # noqa: E402

from xl_io import (DATE_FORMAT, emit, find_column, guess_day_first, is_blank,  # noqa: E402
                   normalize_pk_phone, number_format_for, output_path, parse_date,
                   parse_number, read_table, save, write_table)

ID_HINT = re.compile(r"cnic|nic\b|ntn|strn|iban|account|a/c|acc\b|code|\bid\b|sku|barcode|zip|postal|pin\b|cheque|invoice\s*(no|#|number)|ref", re.I)
PHONE_HINT = re.compile(r"phone|mobile|cell|whats\s*app|contact|tel\b", re.I)
MAX_LOG = 5000
THRESHOLD = 0.8  # share of filled cells that must parse before a column is converted


def classify(name, values, day_first):
    filled = [v for v in values if not is_blank(v)]
    if not filled:
        return "empty"
    if PHONE_HINT.search(name):
        if sum(normalize_pk_phone(v) is not None for v in filled) >= len(filled) * THRESHOLD:
            return "phone"
        return "text"
    if not ID_HINT.search(name):
        dates = sum(parse_date(v, day_first) is not None and not isinstance(v, (int, float)) for v in filled)
        if dates >= len(filled) * THRESHOLD:
            return "date"
        if sum(parse_number(v) is not None for v in filled) >= len(filled) * THRESHOLD:
            return "number"
    return "text"


def tidy_text(v):
    if isinstance(v, str):
        t = " ".join(v.replace(" ", " ").split())
        return t if t else None
    return v


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input")
    ap.add_argument("--sheet")
    ap.add_argument("--out")
    ap.add_argument("--month-first", action="store_true", help="dates are MM/DD (default: detect, else DD/MM)")
    ap.add_argument("--keep-duplicates", action="store_true")
    ap.add_argument("--text-columns", default="", help="comma-separated columns to leave as text")
    ap.add_argument("--overwrite", action="store_true")
    a = ap.parse_args()

    t = read_table(a.input, a.sheet)
    header, rows, rownums = t["header"], t["rows"], t["row_numbers"]
    keep_text = {find_column(header, c.strip()) for c in a.text_columns.split(",") if c.strip()}
    log = []
    changes = 0

    def note(row_i, col_i, old, new, reason):
        nonlocal changes
        changes += 1
        if len(log) < MAX_LOG:
            log.append([rownums[row_i], header[col_i], old, new, reason])

    kinds = {}
    unconverted = []
    for c, name in enumerate(header):
        column = [r[c] for r in rows]
        day_first = False if a.month_first else guess_day_first(column)
        kind = "text" if c in keep_text else classify(name, column, day_first)
        kinds[name] = kind
        for i, r in enumerate(rows):
            old = r[c]
            if is_blank(old):
                if old is not None:
                    r[c] = None
                continue
            new, reason = tidy_text(old), "trimmed spaces"
            parsed = None
            if kind == "date":
                parsed = parse_date(old, day_first)
                reason = "text to date"
            elif kind == "number":
                parsed = parse_number(old)
                reason = "text to number"
            elif kind == "phone":
                parsed = normalize_pk_phone(old)
                reason = "mobile number format"
            if kind in ("date", "number", "phone"):
                if parsed is None:
                    reason = "trimmed spaces"
                    unconverted.append({"row": rownums[i], "column": name, "value": old})
                else:
                    new = parsed
            if new != old or type(new) is not type(old):
                note(i, c, old, new, reason)
            r[c] = new

    removed_dupes = 0
    if not a.keep_duplicates:
        seen = set()
        kept, kept_nums = [], []
        for i, r in enumerate(rows):
            key = tuple(str(v) for v in r)
            if key in seen:
                removed_dupes += 1
                changes += 1
                if len(log) < MAX_LOG:
                    log.append([rownums[i], "(whole row)", "duplicate", None, "exact duplicate row removed"])
                continue
            seen.add(key)
            kept.append(r)
            kept_nums.append(rownums[i])
        rows, rownums = kept, kept_nums

    formats = {}
    for c, name in enumerate(header):
        if kinds[name] == "date":
            formats[c] = DATE_FORMAT
        elif kinds[name] == "number":
            formats[c] = number_format_for([r[c] for r in rows])

    wb = Workbook()
    ws = wb.active
    ws.title = "Cleaned"
    write_table(ws, header, rows, formats)
    log_ws = wb.create_sheet("Changes")
    write_table(log_ws, ["Original row", "Column", "Old value", "New value", "What changed"], log)
    if changes > MAX_LOG:
        log_ws.append([None, f"... and {changes - MAX_LOG} more changes not listed"])
    out = save(wb, output_path(a.input, "cleaned", a.out), a.overwrite)

    emit({
        "output": out,
        "sheet_read": t["sheet"],
        "header_row": t["header_row"],
        "rows_in": len(t["row_numbers"]),
        "rows_out": len(rows),
        "duplicates_removed": removed_dupes,
        "empty_rows_removed": t["empty_rows_removed"],
        "column_types": kinds,
        "changes": changes,
        "values_not_converted": unconverted[:20],
        "values_not_converted_count": len(unconverted),
    })


if __name__ == "__main__":
    main()
