---
name: excel-formulas
description: Write, explain and fix Excel and Google Sheets formulas the way an experienced accountant would - lookups (XLOOKUP, INDEX/MATCH, VLOOKUP), SUMIFS/COUNTIFS by date range and month, text and date cleanup, ageing buckets, running balances, lakh/crore number formats, amounts in words, and errors like #N/A, #VALUE!, #SPILL! and #REF!. Use when the user asks for a formula, asks why a formula gives an error or a wrong total, or describes a calculation they want in a spreadsheet.
---

# Excel formulas

## How to answer

1. **Pin down the layout first.** A formula is only right for a specific layout. If the user
   hasn't said, state your assumptions in one line (sheet names, which column holds what, the
   header row) or ask for a screenshot or the first rows.
2. **Ask which Excel, only when it matters.** XLOOKUP, FILTER, UNIQUE, LET, TEXTSPLIT and
   dynamic arrays need Excel 365/2021 or Google Sheets. If the user may be on Excel 2016/2019,
   give the classic version too (INDEX/MATCH, SUMPRODUCT). [reference.md](reference.md) lists both.
3. **Give the formula ready to paste**, for the first data row, and say where to put it and
   whether to fill it down. Use `$` correctly so it survives being copied.
4. **Explain it in one or two plain sentences**, then one worked example with real numbers
   from their data when you have it.
5. **For a wrong result, find the cause before rewriting.** The usual suspects, in order:
   numbers or dates stored as text (green triangle, left-aligned), extra spaces, a range that
   doesn't cover new rows, a relative reference that shifted, approximate match by accident
   (VLOOKUP without FALSE), and manual calculation mode. If the data itself is messy, suggest
   the clean-data skill.

## House rules

- Prefer SUMIFS/COUNTIFS over SUMPRODUCT for readability, and whole-column or Table references
  (`Sales[Amount]`) over fixed ranges that stop at row 500.
- Never hard-code numbers that live in cells (tax rates, thresholds): reference the cell.
- Wrap lookups in IFNA (not IFERROR) so real errors still show.
- Dates in criteria: `">="&DATE(2026,11,1)`, never a text date like `">=01/11/2026"`.
- Pakistani formatting: lakh/crore needs a custom number format (see reference.md); amounts
  in words need a formula or script, because Excel has no built-in function for it.
- In Google Sheets, mention differences when they matter (for example `ARRAYFORMULA`, and
  `;` as separator in some locales).

For long or repeated jobs (building a whole report, cleaning a sheet, comparing files), suggest
the summary-report, clean-data or compare-sheets skills instead of a wall of formulas.
