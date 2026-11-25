---
name: clean-data
description: Clean a messy Excel or CSV sheet into a tidy table - text numbers like "Rs. 1,50,000/-" to real numbers, DD/MM text dates to real dates, Pakistani mobiles to 0300-1234567, extra spaces, empty rows, duplicate rows and title rows above the header - and log every change on a separate sheet. Use when the user asks to clean, fix, tidy or prepare a spreadsheet, bank or ERP export, customer or stock list, or when SUM/VLOOKUP/pivots fail because values are stored as text.
---

# Clean a messy sheet

The script does the work and never changes the user's file: it writes a new workbook with a
**Cleaned** sheet and a **Changes** sheet (original row, column, old value, new value, reason).

## Steps

1. Find the file. If the user named a sheet, pass `--sheet`. Only `.xlsx`, `.xlsm` and `.csv`
   are supported; for `.xls`, ask them to save it as `.xlsx` first.
2. Run the script from this skill's folder:

   ```bash
   python scripts/clean_data.py "path/to/file.xlsx"
   ```

   Options:
   - `--out "path/cleaned.xlsx"` - output file (default: `<name> - cleaned.xlsx` next to the input)
   - `--month-first` - dates are MM/DD (default: detected from values like 25/11, else DD/MM)
   - `--keep-duplicates` - keep exact duplicate rows
   - `--text-columns "Code,Account"` - columns to leave exactly as text
   - `--overwrite` - replace an existing output file

   If it says openpyxl is missing, run `pip install openpyxl` and try again.
3. Read the JSON it prints and tell the user, in plain words:
   - where the cleaned file is
   - rows before and after, duplicates and empty rows removed
   - what each column became (`column_types`: date, number, phone, text)
   - any `values_not_converted` (for example "n/a" in a date column), so they can fix those by hand
   - that every change is listed on the Changes sheet

## Rules the script follows (explain these if asked)

- A column is converted only when at least 80% of its filled cells parse, so a mostly-text
  column is never damaged.
- Columns whose names suggest identifiers (CNIC, NTN, IBAN, account, code, ID, invoice no, ref)
  are never turned into numbers, so leading zeros and dashes survive. Values like "0300" stay text.
- Phone columns (name contains phone, mobile, cell, WhatsApp, contact) are formatted only when
  they hold Pakistani mobiles.
- Brackets mean negative: `(2,000)` becomes -2000. `18%` becomes 0.18.
- Exact duplicate rows are removed after conversion; near-duplicates are left alone. If the user
  wants near-duplicates found (same name, different spelling), do that as a separate step and
  show them the candidates before deleting anything.

## Don't

- Don't edit the original file, and don't delete rows that aren't exact duplicates without asking.
- Don't change the case of names or text unless the user asks.
