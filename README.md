# 📊 excel-skills

**Excel skills for Claude.** Five ready-made skills that teach Claude how to do real spreadsheet
work: clean a messy export, write and fix formulas, compare two files, build a live summary
report, and make a payroll sheet with Pakistani income tax.

> Ask Claude to "clean this bank export" and, without skills, it improvises. With excel-skills
> it runs tested scripts, keeps your original file untouched, logs every change, and builds
> reports with real Excel formulas you can keep using.

[![Tests](https://github.com/mubasher1111/excel-skills/actions/workflows/test.yml/badge.svg)](https://github.com/mubasher1111/excel-skills/actions/workflows/test.yml)
![License](https://img.shields.io/badge/license-MIT-blue)
![Claude Code plugin](https://img.shields.io/badge/Claude%20Code-plugin-D97757)

📖 **[Full guide](https://zexcart.site/excel-skills/)**

---

## The skills

| Skill | Ask Claude something like | What you get |
|---|---|---|
| `clean-data` | "Clean up customers.xlsx" | Real numbers (`Rs. 1,50,000/-` → 150000) and dates (`05/11/2026` → 5 Nov 2026), trimmed text, mobiles as `0300-1234567`, duplicate and empty rows removed, title rows dropped. A **Changes** sheet lists every edit. |
| `excel-formulas` | "Formula to total sales for each customer by month?" / "Why does my VLOOKUP give #N/A?" | A formula ready to paste, explained, with a version for older Excel. Covers lookups, SUMIFS by date, ageing, running balances, lakh/crore formats and an error guide. |
| `compare-sheets` | "What changed between last month's and this month's price list?" | Rows added, removed and changed, cell by cell, matched by a key such as invoice no, SKU or CNIC. Numbers and dates compared by value. |
| `summary-report` | "Make a sales summary by customer and month" | Totals, counts, % of total, a month-by-month sheet and a chart, all as SUMIFS/COUNTIFS formulas that update with the data. |
| `salary-sheet` | "Make November's payroll from staff.xlsx" | Gross, monthly income tax (FBR salaried slabs FY 2026-27), deductions and net pay as formulas, with lakh/crore formatting and a print-ready layout. |

Your files stay on your computer. The scripts read only the files you name, always write a new
file, and never connect to the internet.

## Install

### Claude Code

```text
/plugin marketplace add mubasher1111/excel-skills
/plugin install excel-skills@zexcart
```

Then just ask, or call a skill directly, for example `/excel-skills:clean-data customers.xlsx`.

### Claude.ai and Claude Desktop

Download a skill's `.zip` from the [latest release](https://github.com/mubasher1111/excel-skills/releases/latest),
then go to **Settings → Capabilities → Skills → Upload skill**. Upload one zip per skill.
(Skills need code execution to be turned on in Settings → Capabilities.)

### Requirements

The four skills with scripts need Python 3.10+ and `openpyxl`. Claude.ai already has both. On your
own computer, if Claude reports that openpyxl is missing:

```bash
pip install openpyxl
```

`excel-formulas` needs nothing.

## Examples

**Clean a bank or ERP export**

> "Clean up `Downloads/ledger-export.xlsx`. The dates are day/month."

Claude runs `clean_data.py`, then tells you: 1,240 rows in, 1,236 out, 4 exact duplicates removed,
"Amount" converted to numbers, "Date" to dates, 3 values it couldn't convert (with their rows).

**Payroll with tax**

> "Make the November 2026 salary sheet for Crescent Garments from `staff.xlsx`."

An employee on Rs 2,00,000 a month pays Rs 13,000 tax a month (Rs 24 lakh a year: 1,16,000 + 20%
of the amount over 22 lakh). Change any salary in the sheet and the tax and net pay update.

## Limits

- `.xlsx`, `.xlsm` and `.csv` only. Save old `.xls` files as `.xlsx` first.
- The scripts read cell values, not formulas or formatting, from your file.
- The payroll tax is an estimate for regular salaried employees, not tax advice.

## Development

```bash
git clone https://github.com/mubasher1111/excel-skills.git
cd excel-skills
pip install openpyxl
python -m unittest discover -s tests -v
```

Shared helpers live in `shared/xl_io.py`. After editing, run `python tools/sync.py` to copy them
into each skill (each skill folder must work on its own when uploaded to Claude.ai), and
`python tools/sync.py --zip` to build the upload zips in `dist/`.

## License

[MIT](LICENSE). Made by [ZexCart](https://zexcart.site). More free tools:
[zexcart.site/dev-tools](https://zexcart.site/dev-tools/).

If it saves you an evening in Excel, please ⭐ **star the repo**.
