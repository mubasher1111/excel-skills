---
name: compare-sheets
description: Compare two versions of an Excel or CSV sheet (or two sheets in one file) by a key column such as invoice number, CNIC, SKU or employee ID, and report rows added, removed and changed, cell by cell, in a new workbook. Use when the user asks what changed between two files, to compare price lists, stock counts, payroll months, customer lists or ledgers, or to find differences or mismatches between sheets.
---

# Compare two sheets

## Steps

1. Get the two files (old first, then new). If both tables are in one workbook, pass the same
   file twice with `--old-sheet` and `--new-sheet`.
2. Decide the key: the column that identifies a row (invoice no, CNIC, SKU, employee code).
   If the user didn't say and it isn't obvious, the script picks the first column with unique,
   filled values in both tables. Several columns can be combined: `--key "Customer,Month"`.
3. Run from this skill's folder:

   ```bash
   python scripts/compare_sheets.py "old.xlsx" "new.xlsx" --key "Invoice No"
   ```

   Options: `--old-sheet`, `--new-sheet`, `--out "path.xlsx"`, `--ignore-case`, `--overwrite`.
   If openpyxl is missing, run `pip install openpyxl`.
4. Explain the JSON result: how many rows changed, were added and removed, columns that exist
   in only one file, duplicate keys that were skipped, and a few examples of each. Point them to
   the output workbook (Summary, Changed, Added, Removed sheets).

## How values are compared

- Numbers as numbers: `1,50,000` equals `150000`; differences under 0.005 are ignored.
- Dates as dates: `05/11/2026` equals a real date of 5 Nov 2026.
- Text after trimming spaces; case matters unless `--ignore-case`.

If the script reports duplicate keys, tell the user: those rows can't be matched reliably, and
they should check them or choose a more specific key.
