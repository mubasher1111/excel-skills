"""End-to-end tests: run each script like Claude would, then open the workbook it wrote."""
import filecmp
import json
import os
import subprocess
import sys
import tempfile
import unittest
from datetime import datetime

from openpyxl import Workbook, load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKILLS = os.path.join(ROOT, "plugins", "excel-skills", "skills")
sys.path.insert(0, os.path.join(SKILLS, "salary-sheet", "scripts"))
import salary_sheet  # noqa: E402
import xl_io  # noqa: E402


def script(skill, name):
    return os.path.join(SKILLS, skill, "scripts", name)


def run(path, *args, ok=True):
    p = subprocess.run([sys.executable, path, *args], capture_output=True, text=True, encoding="utf-8")
    out = json.loads(p.stdout)
    if ok:
        assert p.returncode == 0, p.stdout + p.stderr
    return out


def xlsx(path, rows, title="Sheet1"):
    wb = Workbook()
    ws = wb.active
    ws.title = title
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


class Parsing(unittest.TestCase):
    def test_numbers(self):
        p = xl_io.parse_number
        self.assertEqual(p("Rs. 1,50,000/-"), 150000)
        self.assertEqual(p("(1,200.50)"), -1200.5)
        self.assertEqual(p("PKR 2,500"), 2500)
        self.assertEqual(p("-45"), -45)
        self.assertAlmostEqual(p("18%"), 0.18)
        self.assertIsNone(p("0300"))
        self.assertIsNone(p("35202-1234567-1"))
        self.assertIsNone(p("abc"))
        self.assertIsNone(p(True))

    def test_dates(self):
        d = xl_io.parse_date
        self.assertEqual(d("05/11/2026"), datetime(2026, 11, 5))
        self.assertEqual(d("05/11/2026", day_first=False), datetime(2026, 5, 11))
        self.assertEqual(d("5-Nov-26"), datetime(2026, 11, 5))
        self.assertEqual(d("Nov 5, 2026"), datetime(2026, 11, 5))
        self.assertEqual(d("2026-11-05"), datetime(2026, 11, 5))
        self.assertIsNone(d("31/02/2026"))
        self.assertIsNone(d("hello"))
        self.assertTrue(xl_io.guess_day_first(["25/11/2026", "01/12/2026"]))
        self.assertFalse(xl_io.guess_day_first(["11/25/2026"]))

    def test_phones(self):
        f = xl_io.normalize_pk_phone
        self.assertEqual(f("+92 300 1234567"), "0300-1234567")
        self.assertEqual(f("923001234567"), "0300-1234567")
        self.assertEqual(f("3001234567"), "0300-1234567")
        self.assertEqual(f("0300-1234567"), "0300-1234567")
        self.assertIsNone(f("042-35761234"))

    def test_shared_copies_identical(self):
        shared = os.path.join(ROOT, "shared", "xl_io.py")
        for name in os.listdir(SKILLS):
            copy = os.path.join(SKILLS, name, "scripts", "xl_io.py")
            if os.path.exists(copy):
                self.assertTrue(filecmp.cmp(shared, copy, shallow=False), f"{name}: run python tools/sync.py")


class Scripts(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def p(self, name):
        return os.path.join(self.tmp, name)

    def test_clean_data(self):
        src = xlsx(self.p("customers.xlsx"), [
            ["Customer list - exported 25 Nov 2026"],
            [None],
            ["Name", "Mobile", "Joined", "Balance", "CNIC"],
            ["  Ali   Raza ", "+92 300 1234567", "05/11/2026", "Rs. 1,50,000", "35202-1234567-1"],
            ["Sana", "03211234567", "25/10/2026", "(2,000)", "35201-7654321-2"],
            [None, None, None, None, None],
            ["Sana", "03211234567", "25/10/2026", "(2,000)", "35201-7654321-2"],
            ["Usman", "n/a", "1-Dec-26", "500", "35202-1111111-3"],
            ["Hira", "0333 7654321", "10/11/2026", "PKR 750", "35202-2222222-4"],
        ])
        r = run(script("clean-data", "clean_data.py"), src)
        self.assertEqual(r["header_row"], 3)
        self.assertEqual(r["rows_in"], 5)
        self.assertEqual(r["rows_out"], 4)
        self.assertEqual(r["duplicates_removed"], 1)
        self.assertEqual(r["empty_rows_removed"], 1)
        self.assertEqual(r["column_types"], {"Name": "text", "Mobile": "phone", "Joined": "date",
                                             "Balance": "number", "CNIC": "text"})
        self.assertEqual(r["values_not_converted_count"], 1)  # "n/a" mobile
        wb = load_workbook(r["output"])
        ws = wb["Cleaned"]
        self.assertEqual([c.value for c in ws[2]],
                         ["Ali Raza", "0300-1234567", datetime(2026, 11, 5), 150000, "35202-1234567-1"])
        self.assertEqual(ws["D3"].value, -2000)
        self.assertEqual(ws["C4"].value, datetime(2026, 12, 1))
        self.assertEqual(ws["C2"].number_format, "dd-mmm-yyyy")
        self.assertGreater(wb["Changes"].max_row, 5)
        # never overwrites
        again = run(script("clean-data", "clean_data.py"), src, ok=False)
        self.assertIn("already exists", again["error"])

    def test_clean_csv_month_first(self):
        src = self.p("us.csv")
        with open(src, "w", encoding="utf-8") as f:
            f.write("Date,Amount\n11/25/2026,\"1,200\"\n12/01/2026,300\n")
        r = run(script("clean-data", "clean_data.py"), src, "--out", self.p("out.xlsx"))
        ws = load_workbook(r["output"])["Cleaned"]
        self.assertEqual(ws["A2"].value, datetime(2026, 11, 25))
        self.assertEqual(ws["A3"].value, datetime(2026, 12, 1))
        self.assertEqual(ws["B2"].value, 1200)

    def test_errors(self):
        r = run(script("clean-data", "clean_data.py"), self.p("missing.xlsx"), ok=False)
        self.assertIn("File not found", r["error"])
        open(self.p("old.xls"), "w").close()
        r = run(script("clean-data", "clean_data.py"), self.p("old.xls"), ok=False)
        self.assertIn("Save As", r["error"])

    def test_compare(self):
        old = xlsx(self.p("old.xlsx"), [
            ["Invoice No", "Customer", "Amount", "Status"],
            ["INV-1", "Ali", 1000, "Paid"],
            ["INV-2", "Sana", 2500, "Due"],
            ["INV-3", "Usman", 700, "Due"],
        ])
        new = xlsx(self.p("new.xlsx"), [
            ["Invoice No", "Customer", "Amount", "Status", "Notes"],
            ["INV-1", "Ali", "1,000", "Paid", None],
            ["INV-2", "Sana", 2700, "Paid", "revised"],
            ["INV-4", "Hira", 900, "Due", None],
        ])
        r = run(script("compare-sheets", "compare_sheets.py"), old, new)
        self.assertEqual(r["matched_by"], ["Invoice No"])
        self.assertEqual((r["rows_changed"], r["cells_changed"], r["added"], r["removed"]), (1, 2, 1, 1))
        self.assertEqual(r["columns_only_in_new"], ["Notes"])
        self.assertEqual(r["examples"]["added"], ["INV-4"])
        changed = {(c["column"], c["old"], c["new"]) for c in r["examples"]["changed"]}
        self.assertEqual(changed, {("Amount", 2500, 2700), ("Status", "Due", "Paid")})
        wb = load_workbook(r["output"])
        self.assertEqual(wb.sheetnames, ["Summary", "Changed", "Added", "Removed"])
        self.assertEqual(wb["Removed"]["C2"].value, "Usman")

    def test_compare_needs_key(self):
        old = xlsx(self.p("a.xlsx"), [["Name", "Qty"], ["x", 1], ["x", 1]])
        new = xlsx(self.p("b.xlsx"), [["Name", "Qty"], ["x", 1]])
        r = run(script("compare-sheets", "compare_sheets.py"), old, new, ok=False)
        self.assertIn("--key", r["error"])

    def test_summary_report(self):
        src = xlsx(self.p("sales.xlsx"), [
            ["Date", "Customer", "City", "Amount"],
            ["05/10/2026", "Ali Traders", "Lahore", "10,000"],
            ["15/10/2026", "Sana & Co", "Karachi", 5000],
            ["03/11/2026", "ali traders", "Lahore", 2500],
            ["20/11/2026", "Usman*Store", "Lahore", 1500],
            ["21/11/2026", None, "Multan", 100],
        ])
        r = run(script("summary-report", "summary_report.py"), src, "--group", "customer",
                "--value", "Amount", "--date", "Date")
        self.assertEqual(r["rows"], 5)
        self.assertEqual(r["groups"], 4)  # "Ali Traders" and "ali traders" are one group, like SUMIFS
        self.assertEqual(r["grand_total"], {"Amount": 19100})
        self.assertEqual(r["top_groups"][0], {"Customer": "Ali Traders", "count": 2, "Amount": 12500})
        self.assertEqual(r["by_month"], {"2026-10": 15000, "2026-11": 4100})
        wb = load_workbook(r["output"])
        self.assertEqual(wb.sheetnames, ["Summary", "By Month", "Data"])
        s = wb["Summary"]
        self.assertEqual([c.value for c in s[4]], ["Customer", "Count", "Total Amount", "% of Amount"])
        self.assertEqual(s["A5"].value, "Ali Traders")
        self.assertEqual(s["C5"].value, "=SUMIFS(Data!$D:$D,Data!$B:$B,$A5)")
        formulas = [s.cell(row=r, column=3).value for r in range(5, 9)]
        self.assertIn('=SUMIFS(Data!$D:$D,Data!$B:$B,"=Usman~*Store")', formulas)
        self.assertIn('=SUMIFS(Data!$D:$D,Data!$B:$B,"=")', formulas)
        self.assertEqual(s["C9"].value, "=SUM(C5:C8)")
        self.assertEqual(wb["Data"]["D2"].value, 10000)
        self.assertEqual(wb["Data"]["A2"].value, datetime(2026, 10, 5))
        self.assertIn('DATE(2026,11,1)', wb["By Month"]["B5"].value)
        self.assertEqual(len(s._charts), 1)

    def test_salary_tax_matches_fbr_examples(self):
        # Rs 2,00,000 a month -> Rs 24 lakh a year -> 116,000 + 20% of 200,000 = 156,000 -> 13,000 a month
        self.assertEqual(salary_sheet.monthly_tax(200000, 0), 13000)
        self.assertEqual(salary_sheet.monthly_tax(50000, 0), 0)
        self.assertEqual(salary_sheet.monthly_tax(100000, 0), 500)     # 12 lakh: 1% of 6 lakh / 12
        self.assertEqual(salary_sheet.monthly_tax(1000000, 0), 264500)  # 1.2 crore: 1,424,000 + 35% of 50 lakh

    def test_salary_sheet(self):
        r = run(script("salary-sheet", "salary_sheet.py"), "--out", self.p("payroll"), "--month", "November 2026")
        self.assertTrue(r["sample_data"])
        self.assertEqual(r["employees"], 5)
        wb = load_workbook(r["output"])
        ws = wb["Payroll"]
        self.assertEqual(ws["B4"].value, "Employee name")
        self.assertEqual(ws["H5"].value, "=E5+F5+G5")
        self.assertTrue(ws["J5"].value.startswith("=ROUND((VLOOKUP(I5,'Tax Slabs'!$A$5:$D$12,3,TRUE)"))
        self.assertEqual(ws["M10"].value, "=SUM(M5:M9)")
        # The VLOOKUP table holds the same slabs as the Python calculation, sorted for approximate match.
        slabs = [[c.value for c in row] for row in wb["Tax Slabs"].iter_rows(min_row=5, max_row=12)]
        overs = [s[0] for s in slabs]
        self.assertEqual(overs, sorted(overs))

        def excel_like(income):  # what VLOOKUP(...,TRUE) does with that table
            row = [s for s in slabs if s[0] <= income][-1]
            return round((row[2] + (income - row[0]) * row[3]) / 12)

        for monthly in (30000, 50000, 75000, 100000, 150000, 200000, 300000, 500000, 700000, 1000000):
            self.assertEqual(excel_like(monthly * 12), salary_sheet.monthly_tax(monthly, 0), monthly)
        usman = next(e for e in r["per_employee"] if e["name"] == "Usman Khalid")
        self.assertEqual(usman["gross"], 205000)
        self.assertEqual(usman["monthly_tax"], salary_sheet.monthly_tax(150000, 40000))

    def test_salary_from_employees_file(self):
        emp = self.p("staff.csv")
        with open(emp, "w", encoding="utf-8") as f:
            f.write("Employee Name,Designation,Basic Pay,House Rent,Medical,Advance\n"
                    "Ali,Clerk,\"40,000\",10000,4000,1000\nSana,Manager,200000,0,0,0\n")
        r = run(script("salary-sheet", "salary_sheet.py"), "--out", self.p("p.xlsx"), "--employees", emp)
        self.assertFalse(r["sample_data"])
        self.assertEqual(r["employees"], 2)
        self.assertEqual(r["columns_used"]["Basic salary"], "Basic Pay")
        self.assertEqual(r["columns_used"]["Taxable allowances"], "House Rent")
        self.assertEqual(r["columns_used"]["Exempt allowances"], "Medical")
        sana = r["per_employee"][1]
        self.assertEqual((sana["monthly_tax"], sana["net_pay"]), (13000, 187000))
        ali = r["per_employee"][0]
        self.assertEqual(ali["net_pay"], 54000 - ali["monthly_tax"] - 1000)


if __name__ == "__main__":
    unittest.main()
