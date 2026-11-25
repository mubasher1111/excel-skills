#!/usr/bin/env python3
"""Make a monthly payroll sheet for a Pakistani business, with FBR salary tax for FY 2026-27.

    python salary_sheet.py --out FILE [--employees FILE] [--month "November 2026"]
                           [--company "Name"] [--overwrite]

--employees: optional .xlsx/.csv with columns such as Name, CNIC, Designation, Basic,
             Allowances (taxable), Exempt Allowances, Deductions, Advance. Headers are matched loosely.
             Without it, the sheet has 5 sample rows to overwrite.

Writes a NEW workbook:
  Payroll    one row per employee; Gross, annual taxable income, monthly tax and net pay are
             formulas, so editing any salary updates everything
  Tax Slabs  the FY 2026-27 salaried slabs the tax formula reads (edit here when rates change)
  Notes      what the sheet assumes
Prints a JSON summary with the same figures computed in Python.
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Alignment, Border, Side  # noqa: E402
from openpyxl.workbook.properties import CalcProperties  # noqa: E402

from xl_io import (LAKH_FORMAT, TOTAL_FONT, autosize, emit, fail, is_blank,  # noqa: E402
                   parse_number, read_table, save, style_header)

FISCAL_YEAR = "2026-27"
# FY 2026-27 (Tax Year 2027), Finance Act 2026, salaried individuals: (over, up_to, fixed, rate)
SLABS = [
    (0, 600_000, 0, 0.00),
    (600_000, 1_200_000, 0, 0.01),
    (1_200_000, 2_200_000, 6_000, 0.11),
    (2_200_000, 3_200_000, 116_000, 0.20),
    (3_200_000, 4_100_000, 316_000, 0.25),
    (4_100_000, 5_600_000, 541_000, 0.29),
    (5_600_000, 7_000_000, 976_000, 0.32),
    (7_000_000, None, 1_424_000, 0.35),
]

COLUMNS = ["#", "Employee name", "CNIC", "Designation", "Basic salary", "Taxable allowances",
           "Exempt allowances", "Gross salary", "Annual taxable income", "Income tax (monthly)",
           "Other deductions", "Advance recovery", "Net pay"]
FIELDS = {  # employees-file column -> payroll column index (0-based in COLUMNS)
    1: r"name|employee",
    2: r"cnic|nic|id\s*card",
    3: r"designation|title|position|post|role",
    4: r"basic",
    5: r"^allowances?$|taxable\s*allow|house\s*rent|conveyance|other\s*allow",
    6: r"exempt|medical",
    10: r"deduction|eobi|loan",
    11: r"advance",
}
SAMPLE = [
    ["Ali Raza", "35202-1234567-1", "Accountant", 60000, 15000, 5000, 0, 0],
    ["Sana Tariq", "35201-7654321-2", "HR Officer", 85000, 20000, 8500, 0, 5000],
    ["Usman Khalid", "35202-1111111-3", "Production Manager", 150000, 40000, 15000, 2000, 0],
    ["Hira Ahmed", "35202-2222222-4", "Designer", 45000, 10000, 4500, 0, 0],
    ["Bilal Hussain", "35202-3333333-5", "Driver", 32000, 5000, 0, 0, 2000],
]


def annual_tax(income):
    for over, up_to, fixed, rate in SLABS:
        if up_to is None or income <= up_to:
            return fixed + (income - over) * rate
    return 0  # pragma: no cover


def monthly_tax(basic, taxable_allowances):
    return round(annual_tax((basic + taxable_allowances) * 12) / 12)


def load_employees(path):
    t = read_table(path)
    header = t["header"]
    mapping = {}
    for target, pattern in FIELDS.items():
        for i, h in enumerate(header):
            if i not in mapping.values() and re.search(pattern, h, re.I):
                mapping[target] = i
                break
    if 1 not in mapping or 4 not in mapping:
        fail("The employees file needs at least a name column and a basic salary column. "
             f"Columns found: {', '.join(header)}")
    people = []
    for r in t["rows"]:
        if is_blank(r[mapping[1]]):
            continue
        row = []
        for target in (1, 2, 3, 4, 5, 6, 10, 11):
            v = r[mapping[target]] if target in mapping else None
            if target >= 4:
                v = parse_number(v) or 0
            row.append(v)
        people.append(row)
    if not people:
        fail("No employees found in the file.")
    return people, {COLUMNS[k]: header[v] for k, v in mapping.items()}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", required=True)
    ap.add_argument("--employees")
    ap.add_argument("--month", default="")
    ap.add_argument("--company", default="")
    ap.add_argument("--overwrite", action="store_true")
    a = ap.parse_args()

    if a.employees:
        people, mapping = load_employees(a.employees)
        sample = False
    else:
        people, mapping, sample = SAMPLE, {}, True

    wb = Workbook()
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    ws = wb.active
    ws.title = "Payroll"
    title = " - ".join(x for x in [a.company, f"Payroll {a.month}".strip()] if x)
    ws["A1"] = title
    ws["A1"].font = TOTAL_FONT
    ws["A2"] = (f"Income tax: FBR salaried slabs FY {FISCAL_YEAR} (see Tax Slabs sheet). "
                "Edit the blue columns; the rest are formulas.")
    ws.append([])
    ws.append(COLUMNS)
    style_header(ws, 4)
    first = 5
    slab_ref = "'Tax Slabs'!$A$5:$D$12"
    for i, p in enumerate(people):
        r = first + i
        name, cnic, desig, basic, taxable, exempt, deductions, advance = p
        ws.append([
            i + 1, name, cnic, desig, basic, taxable, exempt,
            f"=E{r}+F{r}+G{r}",
            f"=(E{r}+F{r})*12",
            f"=ROUND((VLOOKUP(I{r},{slab_ref},3,TRUE)+(I{r}-VLOOKUP(I{r},{slab_ref},1,TRUE))"
            f"*VLOOKUP(I{r},{slab_ref},4,TRUE))/12,0)",
            deductions, advance,
            f"=H{r}-J{r}-K{r}-L{r}",
        ])
    last = first + len(people) - 1
    tr = last + 1
    ws.append(["", "Total"] + [None] * 2 + [f"=SUM({c}{first}:{c}{last})" for c in "EFGHIJKLM"])
    thin = Side(style="thin", color="999999")
    for cell in ws[tr]:
        cell.font = TOTAL_FONT
        cell.border = Border(top=thin, bottom=Side(style="double", color="999999"))
    for row in ws.iter_rows(min_row=first, max_row=tr, min_col=5, max_col=13):
        for cell in row:
            cell.number_format = LAKH_FORMAT
    from openpyxl.styles import Font
    for row in ws.iter_rows(min_row=first, max_row=last):
        for idx in (1, 2, 3, 4, 5, 6, 10, 11):
            row[idx].font = Font(color="1F4E9E")
    ws.freeze_panes = "C5"
    autosize(ws, max_width=28)
    ws.column_dimensions["A"].width = 5
    ws.row_dimensions[4].height = 32
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "4:4"

    sl = wb.create_sheet("Tax Slabs")
    sl["A1"] = f"Salaried income tax slabs, FY {FISCAL_YEAR} (Tax Year 2027)"
    sl["A1"].font = TOTAL_FONT
    sl["A2"] = "Tax = fixed amount + rate x (annual income - 'Over'). Update this table if the law changes."
    sl.append([])
    sl.append(["Over (Rs)", "Up to (Rs)", "Fixed tax (Rs)", "Rate on the excess"])
    style_header(sl, 4)
    for over, up_to, fixed, rate in SLABS:
        sl.append([over, up_to if up_to is not None else "and above", fixed, rate])
    for row in sl.iter_rows(min_row=5, max_row=12):
        for cell in row[:3]:
            cell.number_format = LAKH_FORMAT
        row[3].number_format = "0%"
    autosize(sl, max_width=24)
    sl.column_dimensions["A"].width = 16

    notes = wb.create_sheet("Notes")
    lines = [
        "How this sheet works",
        "Gross salary = basic + taxable allowances + exempt allowances.",
        "Annual taxable income = (basic + taxable allowances) x 12. Exempt allowances are not taxed.",
        "Income tax (monthly) = annual tax from the Tax Slabs sheet / 12, rounded to the nearest rupee.",
        "Net pay = gross salary - income tax - other deductions - advance recovery.",
        "",
        "Assumptions",
        "This is an estimate for regular salaried employees. It does not include tax credits, "
        "arrears, bonuses taxed separately, or other income. Check final figures with your tax advisor.",
        "Fill the blue columns. Keep the Tax Slabs table sorted from the lowest slab.",
    ]
    for line in lines:
        notes.append([line])
    notes["A1"].font = TOTAL_FONT
    notes["A7"].font = TOTAL_FONT
    notes.column_dimensions["A"].width = 110
    for row in notes.iter_rows():
        row[0].alignment = Alignment(wrap_text=True)

    out = os.path.expanduser(a.out)
    if not out.lower().endswith(".xlsx"):
        out += ".xlsx"
    out = save(wb, out, a.overwrite)

    figures = []
    for name, _cnic, _d, basic, taxable, exempt, deductions, advance in people:
        tax = monthly_tax(basic, taxable)
        gross = basic + taxable + exempt
        figures.append({"name": name, "gross": gross, "monthly_tax": tax,
                        "net_pay": gross - tax - deductions - advance})
    emit({
        "output": out,
        "fiscal_year": FISCAL_YEAR,
        "employees": len(people),
        "sample_data": sample,
        "columns_used": mapping,
        "total_gross": sum(f["gross"] for f in figures),
        "total_tax": sum(f["monthly_tax"] for f in figures),
        "total_net_pay": sum(f["net_pay"] for f in figures),
        "per_employee": figures[:50],
    })


if __name__ == "__main__":
    main()
