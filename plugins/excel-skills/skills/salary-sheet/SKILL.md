---
name: salary-sheet
description: Create a monthly payroll / salary sheet in Excel for a Pakistani business with FBR salaried income tax for FY 2026-27 calculated by formulas - gross salary, taxable and exempt allowances, monthly income tax, deductions, advance recovery and net pay, with lakh/crore number formatting. Use when the user asks for a salary sheet, payroll, pay register, or income tax on salaries in Pakistan, or wants to work out tax and take-home pay for employees.
---

# Make a payroll sheet

## Steps

1. Ask for the month and company name if they're not given. If the user has an employee list,
   use it (`--employees`); it needs at least a name and a basic salary column. Other columns are
   matched by name: CNIC, Designation, Allowances / House Rent / Conveyance (taxable),
   Exempt / Medical (not taxed), Deductions / EOBI / Loan, Advance.
2. Run from this skill's folder:

   ```bash
   python scripts/salary_sheet.py --out "Payroll November 2026.xlsx" --month "November 2026" --company "Crescent Garments" --employees "staff.xlsx"
   ```

   Without `--employees`, the sheet has 5 sample rows for the user to overwrite - say so clearly.
   `--overwrite` replaces an existing file. If openpyxl is missing, run `pip install openpyxl`.
3. Check `columns_used` in the JSON: tell the user which of their columns went where, especially
   which allowances were treated as taxable and which as exempt, and ask them to confirm.
4. Give the totals (gross, tax, net pay) and a few employees' figures from `per_employee`.

## How tax is worked out

Annual taxable income = (basic + taxable allowances) x 12. Annual tax comes from the FY 2026-27
salaried slabs on the **Tax Slabs** sheet, then divided by 12 and rounded to the rupee:

| Annual income (Rs) | Tax |
|---|---|
| up to 6,00,000 | 0 |
| 6,00,001 - 12,00,000 | 1% of the amount over 6 lakh |
| 12,00,001 - 22,00,000 | 6,000 + 11% over 12 lakh |
| 22,00,001 - 32,00,000 | 1,16,000 + 20% over 22 lakh |
| 32,00,001 - 41,00,000 | 3,16,000 + 25% over 32 lakh |
| 41,00,001 - 56,00,000 | 5,41,000 + 29% over 41 lakh |
| 56,00,001 - 70,00,000 | 9,76,000 + 32% over 56 lakh |
| above 70,00,000 | 14,24,000 + 35% over 70 lakh |

Example: Rs 2,00,000 a month is Rs 24 lakh a year: 1,16,000 + 20% of 2,00,000 = 1,56,000 a year,
so Rs 13,000 a month.

## Be clear about limits

This is an estimate for regular salaried employees. It doesn't include tax credits, arrears,
bonuses or other income, and it isn't tax advice: suggest the user confirms final figures with
their tax advisor. If the user says the rates have changed, edit the Tax Slabs sheet (keep it
sorted from the lowest slab) rather than the formulas.
