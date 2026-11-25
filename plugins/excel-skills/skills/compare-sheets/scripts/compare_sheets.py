#!/usr/bin/env python3
"""Compare two versions of a sheet by a key column: what was added, removed and changed.

    python compare_sheets.py OLD NEW [--key "Invoice No"] [--old-sheet NAME] [--new-sheet NAME]
                             [--out FILE] [--ignore-case] [--overwrite]

OLD and NEW can be different files, or the same file with --old-sheet/--new-sheet.
--key may list several columns ("Customer,Month"). Without --key, the first column whose values
are unique and filled in both tables is used.
Numbers are compared as numbers ("1,50,000" equals 150000) and dates as dates.
Writes a NEW workbook with Summary, Changed, Added and Removed sheets, and prints a JSON summary.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import PatternFill  # noqa: E402

from xl_io import (TOTAL_FONT, autosize, emit, fail, find_column, is_blank,  # noqa: E402
                   output_path, parse_date, parse_number, read_table, save, style_header,
                   write_table)

CHANGED_FILL = PatternFill("solid", fgColor="FFF4CC")


def normalize(v, ignore_case):
    if is_blank(v):
        return ""
    n = parse_number(v)
    if n is not None:
        n = round(float(n), 2)
        return int(n) if n.is_integer() else n
    d = parse_date(v)
    if d is not None and not isinstance(v, (int, float)):
        return d.date().isoformat()
    s = " ".join(str(v).split())
    return s.lower() if ignore_case else s


def pick_key(old, new, ignore_case):
    for name in old["header"]:
        if name not in new["header"]:
            continue
        ok = True
        for t in (old, new):
            c = t["header"].index(name)
            vals = [normalize(r[c], ignore_case) for r in t["rows"]]
            if any(v == "" for v in vals) or len(set(vals)) != len(vals):
                ok = False
                break
        if ok:
            return [name]
    fail("Could not find a column with unique values in both tables. "
         "Pass --key with the column that identifies each row (for example an invoice number or CNIC). "
         f"Columns: {', '.join(old['header'])}")


def index(table, key_cols, ignore_case):
    idx = [find_column(table["header"], k) for k in key_cols]
    out, dupes = {}, []
    for i, r in enumerate(table["rows"]):
        key = tuple(normalize(r[c], ignore_case) for c in idx)
        if all(k == "" for k in key):
            continue
        if key in out:
            dupes.append({"key": " | ".join(str(k) for k in key), "row": table["row_numbers"][i]})
            continue
        out[key] = i
    return out, dupes


def label(key):
    return " | ".join(str(k) for k in key)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("old")
    ap.add_argument("new")
    ap.add_argument("--key")
    ap.add_argument("--old-sheet")
    ap.add_argument("--new-sheet")
    ap.add_argument("--out")
    ap.add_argument("--ignore-case", action="store_true")
    ap.add_argument("--overwrite", action="store_true")
    a = ap.parse_args()

    old = read_table(a.old, a.old_sheet)
    new = read_table(a.new, a.new_sheet)
    key_cols = [k.strip() for k in a.key.split(",")] if a.key else pick_key(old, new, a.ignore_case)
    key_cols = [old["header"][find_column(old["header"], k)] for k in key_cols]
    for k in key_cols:
        find_column(new["header"], k)

    old_idx, old_dupes = index(old, key_cols, a.ignore_case)
    new_idx, new_dupes = index(new, key_cols, a.ignore_case)
    common = [h for h in old["header"] if find_column(new["header"], h, required=False) is not None
              and h not in key_cols]
    only_old = [h for h in old["header"] if find_column(new["header"], h, required=False) is None]
    only_new = [h for h in new["header"] if find_column(old["header"], h, required=False) is None]

    added = [k for k in new_idx if k not in old_idx]
    removed = [k for k in old_idx if k not in new_idx]
    changed = []  # (key, column, old value, new value)
    changed_rows = set()
    for k, oi in old_idx.items():
        ni = new_idx.get(k)
        if ni is None:
            continue
        for col in common:
            ov = old["rows"][oi][find_column(old["header"], col)]
            nv = new["rows"][ni][find_column(new["header"], col)]
            if normalize(ov, a.ignore_case) != normalize(nv, a.ignore_case):
                changed.append((k, col, ov, nv))
                changed_rows.add(k)

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    summary = [
        ["Compared", f"{os.path.basename(a.old)} [{old['sheet']}]", f"{os.path.basename(a.new)} [{new['sheet']}]"],
        ["Matched by", ", ".join(key_cols), None],
        ["Rows", len(old["rows"]), len(new["rows"])],
        ["Rows changed", len(changed_rows), None],
        ["Cells changed", len(changed), None],
        ["Rows added (only in new)", len(added), None],
        ["Rows removed (only in old)", len(removed), None],
        ["Columns only in old", ", ".join(only_old) or "-", None],
        ["Columns only in new", ", ".join(only_new) or "-", None],
        ["Duplicate keys skipped", len(old_dupes) + len(new_dupes), None],
    ]
    ws.append(["", "Old", "New"])
    for r in summary:
        ws.append(r)
    style_header(ws)
    for row in ws.iter_rows(min_row=2, max_col=1):
        row[0].font = TOTAL_FONT
    autosize(ws, max_width=60)

    ws = wb.create_sheet("Changed")
    write_table(ws, key_cols + ["Column", "Old value", "New value"],
                [list(k) + [col, ov, nv] for k, col, ov, nv in changed])
    for row in ws.iter_rows(min_row=2, min_col=len(key_cols) + 3, max_col=len(key_cols) + 3):
        row[0].fill = CHANGED_FILL

    ws = wb.create_sheet("Added")
    write_table(ws, ["New file row"] + new["header"],
                [[new["row_numbers"][new_idx[k]]] + new["rows"][new_idx[k]] for k in added])
    ws = wb.create_sheet("Removed")
    write_table(ws, ["Old file row"] + old["header"],
                [[old["row_numbers"][old_idx[k]]] + old["rows"][old_idx[k]] for k in removed])
    out = save(wb, output_path(a.new, "comparison", a.out), a.overwrite)

    emit({
        "output": out,
        "matched_by": key_cols,
        "rows_old": len(old["rows"]),
        "rows_new": len(new["rows"]),
        "rows_changed": len(changed_rows),
        "cells_changed": len(changed),
        "added": len(added),
        "removed": len(removed),
        "columns_only_in_old": only_old,
        "columns_only_in_new": only_new,
        "duplicate_keys": (old_dupes + new_dupes)[:20],
        "examples": {
            "changed": [{"key": label(k), "column": c, "old": o, "new": n} for k, c, o, n in changed[:15]],
            "added": [label(k) for k in added[:15]],
            "removed": [label(k) for k in removed[:15]],
        },
    })


if __name__ == "__main__":
    main()
