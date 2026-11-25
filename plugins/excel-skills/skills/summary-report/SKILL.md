---
name: summary-report
description: Turn a data sheet (sales, expenses, invoices, stock movements) into a live Excel summary report - totals and counts by customer, product, city or any column, share of total, month-by-month totals and a bar chart - built with SUMIFS/COUNTIFS formulas so it updates when the data changes. Use when the user asks for a summary, report, pivot, totals by category, sales by customer or month, or a dashboard-style sheet from their data.
---

# Build a summary report

## Steps

1. Look at the sheet's columns (read the first rows) and agree what to total and by what:
   - `--group`: the column to group by (Customer, Product, City, Category)
   - `--value`: the amount column(s) to total; repeat for several (`--value Amount --value Qty`)
   - optional `--group2`: a second level (for example Customer then City)
   - optional `--date`: a date column, to add a **By Month** sheet
2. Run from this skill's folder:

   ```bash
   python scripts/summary_report.py "sales.xlsx" --group "Customer" --value "Amount" --date "Invoice Date"
   ```

   Options: `--sheet`, `--out "path.xlsx"`, `--overwrite`. If openpyxl is missing, run
   `pip install openpyxl`.
3. Report the numbers from the JSON: grand total, top groups, month totals, and any values that
   weren't numbers or dates (`non_numeric_values_ignored`, `unreadable_dates`) - those rows are
   left out of the totals, so the user should know.

## What the report contains

- **Summary**: one row per group, largest first, with Count, Total per value column, % of total,
  a grand total and a bar chart of the top 15.
- **By Month** (with `--date`): the first value column per group per month (up to 36 months).
- **Data**: the source table with real numbers and dates.

Every figure is a formula on the Data sheet, so the user can add rows or fix values and the
report recalculates. Excel calculates the formulas when the file is opened; some previewers
show blanks until then.

Grouping ignores upper/lower case ("Ali Traders" and "ali traders" are one group), the same way
SUMIFS matches text.
