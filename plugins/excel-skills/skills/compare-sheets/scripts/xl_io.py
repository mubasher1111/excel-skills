"""Shared helpers for the excel-skills scripts: read sheets, parse messy values, write tidy workbooks.

Each skill keeps its own copy of this file (scripts/xl_io.py) so a skill folder works on its own.
Edit shared/xl_io.py and run `python tools/sync.py`; a test checks the copies are identical.
"""
import csv
import json
import os
import re
import sys
from datetime import date, datetime

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    print(json.dumps({"error": "The openpyxl package is missing. Install it with: pip install openpyxl"}))
    sys.exit(1)

HEADER_FILL = PatternFill("solid", fgColor="0E6E6A")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FONT = Font(bold=True)
# Pakistani/Indian grouping: 1,50,000 and 1,25,00,000 (positive values; negatives use normal grouping).
LAKH_FORMAT = r"[>=10000000]##\,##\,##\,##0;[>=100000]##\,##\,##0;##,##0"
DATE_FORMAT = "dd-mmm-yyyy"


def fail(message):
    """Print a JSON error for the assistant to relay, and stop."""
    print(json.dumps({"error": message}, ensure_ascii=False))
    sys.exit(1)


def emit(result):
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))


def is_blank(value):
    return value is None or (isinstance(value, str) and not value.strip())


# ---------------------------------------------------------------- reading

def read_rows(path, sheet=None):
    """Return (rows, sheet_name). Rows are lists of raw cell values."""
    path = os.path.expanduser(path)
    if not os.path.exists(path):
        fail(f"File not found: {path}")
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        try:
            wb = load_workbook(path, data_only=True, read_only=True)
        except PermissionError:
            fail(f"Can't open {path}. If it is open in Excel, close it and try again.")
        if sheet and sheet not in wb.sheetnames:
            fail(f'Sheet "{sheet}" not found. Sheets in this file: {", ".join(wb.sheetnames)}')
        ws = wb[sheet] if sheet else wb.worksheets[0]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        name = ws.title
        wb.close()
        return rows, name
    if ext in (".csv", ".txt"):
        for encoding in ("utf-8-sig", "cp1252"):
            try:
                with open(path, newline="", encoding=encoding) as f:
                    sample = f.read(4096)
                    f.seek(0)
                    try:
                        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
                    except csv.Error:
                        dialect = csv.excel
                    rows = [list(r) for r in csv.reader(f, dialect)]
                return rows, os.path.splitext(os.path.basename(path))[0][:31]
            except UnicodeDecodeError:
                continue
        fail("Could not read the CSV file's text encoding. Save it again as 'CSV UTF-8' from Excel.")
    if ext == ".xls":
        fail("Old .xls files are not supported. Open the file in Excel and use File > Save As > Excel Workbook (.xlsx).")
    fail(f"Unsupported file type '{ext}'. Use .xlsx, .xlsm or .csv.")


def find_header(rows, scan=15):
    """Index of the header row: skips title rows above the table (bank and ERP exports often have them)."""
    counts = [sum(not is_blank(v) for v in r) for r in rows[:scan]]
    widest = max(counts, default=0)
    for i, r in enumerate(rows[:scan]):
        filled = [v for v in r if not is_blank(v)]
        if (len(filled) >= max(2, widest * 0.6)
                and sum(isinstance(v, str) for v in filled) >= len(filled) * 0.6):
            return i
    return 0


def read_table(path, sheet=None):
    """Read a table. Returns dict(header, rows, row_numbers, sheet, empty_rows_removed).

    row_numbers holds the original sheet row (1-based) of each data row, for reporting.
    """
    raw, name = read_rows(path, sheet)
    numbered = [(i + 1, r) for i, r in enumerate(raw) if any(not is_blank(v) for v in r)]
    if not numbered:
        fail("The sheet is empty.")
    h = find_header([r for _, r in numbered])
    header_row = numbered[h][1]
    body = numbered[h + 1:]
    width = max(len(r) for _, r in numbered[h:])
    header = []
    seen = {}
    for i in range(width):
        v = header_row[i] if i < len(header_row) else None
        label = " ".join(str(v).split()) if not is_blank(v) else f"Column {i + 1}"
        if label in seen:
            seen[label] += 1
            label = f"{label} ({seen[label]})"
        else:
            seen[label] = 1
        header.append(label)
    rows = [list(r) + [None] * (width - len(r)) for _, r in body]
    keep = [i for i in range(width)
            if not header[i].startswith("Column ") or any(not is_blank(r[i]) for r in rows)]
    first_data_row = numbered[h][0] + 1
    last_row = numbered[-1][0]
    empty_inside = (last_row - first_data_row + 1) - len(body) if body else 0
    return {
        "header": [header[i] for i in keep],
        "rows": [[r[i] for i in keep] for r in rows],
        "row_numbers": [n for n, _ in body],
        "sheet": name,
        "header_row": numbered[h][0],
        "empty_rows_removed": max(0, empty_inside),
    }


def find_column(header, wanted, required=True):
    """Find a column by name, ignoring case, spaces and punctuation. Also accepts a letter (A, B, ...)."""
    if wanted is None:
        return None
    norm = lambda s: re.sub(r"[^a-z0-9]", "", str(s).lower())
    for i, h in enumerate(header):
        if norm(h) == norm(wanted):
            return i
    if re.fullmatch(r"[A-Za-z]{1,2}", wanted.strip()):
        idx = 0
        for ch in wanted.strip().upper():
            idx = idx * 26 + (ord(ch) - 64)
        if idx - 1 < len(header):
            return idx - 1
    if required:
        fail(f'Column "{wanted}" not found. Columns: {", ".join(header)}')
    return None


# ---------------------------------------------------------------- parsing

_CURRENCY_PREFIX = re.compile(r"^(?:rs\.?|pkr|inr|usd|aed|sar|\$|£|€)", re.I)
_CURRENCY_SUFFIX = re.compile(r"(?:/-|rs\.?|pkr|inr|usd|aed|sar)$", re.I)
_NUMBER = re.compile(r"\d{1,3}(?:,\d{2,3})+(?:\.\d+)?|\d+(?:\.\d+)?|\.\d+")


def parse_number(value):
    """'Rs. 1,50,000/-' -> 150000, '(1,200.50)' -> -1200.5, '12%' -> 0.12. Returns None if not a number.

    Leading-zero codes such as '0300' or '007' are left as text.
    """
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    if not isinstance(value, str):
        return None
    s = value.strip().replace(" ", "").replace(" ", "")
    if not s:
        return None
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative, s = True, s[1:-1]
    if s.startswith("-"):
        negative, s = (not negative), s[1:]
    s = _CURRENCY_SUFFIX.sub("", _CURRENCY_PREFIX.sub("", s))
    if s.startswith("-"):
        negative, s = (not negative), s[1:]
    percent = s.endswith("%")
    if percent:
        s = s[:-1]
    if not _NUMBER.fullmatch(s) or re.fullmatch(r"0\d+", s):
        return None
    n = float(s.replace(",", ""))
    if percent:
        n = n / 100
    elif n.is_integer() and "." not in s:
        n = int(n)
    return -n if negative else n


_MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}


def parse_date(value, day_first=True):
    """Parse '05/11/2026', '5-Nov-26', 'Nov 5, 2026', '2026-11-05'. Returns datetime or None."""
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    if not isinstance(value, str):
        return None
    s = value.strip()
    m = re.fullmatch(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?", s)
    if m:
        y, mo, d = map(int, m.groups())
    else:
        m = re.fullmatch(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})(?: \d{1,2}:\d{2}(?::\d{2})?)?", s)
        if m:
            a, b, y = map(int, m.groups())
            d, mo = (a, b) if day_first else (b, a)
        else:
            m = re.fullmatch(r"(\d{1,2})[-\s]([A-Za-z]{3,9})[-\s,]+(\d{2}|\d{4})", s)
            if m:
                d, mo, y = int(m[1]), _MONTHS.get(m[2][:3].lower()), int(m[3])
            else:
                m = re.fullmatch(r"([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{4})", s)
                if not m:
                    return None
                mo, d, y = _MONTHS.get(m[1][:3].lower()), int(m[2]), int(m[3])
            if mo is None:
                return None
        if y < 100:
            y += 2000
    try:
        return datetime(y, mo, d)
    except ValueError:
        return None


def guess_day_first(values):
    """True for DD/MM, False for MM/DD, judged from values like 25/11 or 11/25. Defaults to DD/MM."""
    day_first = month_first = 0
    for v in values:
        if isinstance(v, str):
            m = re.match(r"\s*(\d{1,2})[-/.](\d{1,2})[-/.]\d{2,4}", v)
            if m:
                a, b = int(m[1]), int(m[2])
                day_first += a > 12
                month_first += b > 12
    return not (month_first > day_first)


def normalize_pk_phone(value):
    """'+92 300 1234567' / '923001234567' / '3001234567' -> '0300-1234567'. None if not a PK mobile."""
    if value is None:
        return None
    digits = re.sub(r"\D", "", str(value))
    if digits.startswith("0092"):
        digits = digits[2:]
    if digits.startswith("92") and len(digits) == 12:
        digits = "0" + digits[2:]
    elif len(digits) == 10 and digits.startswith("3"):
        digits = "0" + digits
    if len(digits) == 11 and digits.startswith("03"):
        return f"{digits[:4]}-{digits[4:]}"
    return None


# ---------------------------------------------------------------- writing

def _display_len(v):
    if v is None:
        return 0
    if isinstance(v, (datetime, date)):
        return 11
    if isinstance(v, float):
        return len(f"{v:,.2f}")
    if isinstance(v, int):
        return len(f"{v:,}")
    return max((len(line) for line in str(v).splitlines()), default=0)


def style_header(ws, row=1):
    for cell in ws[row]:
        if cell.value is not None:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(vertical="center", wrap_text=True)


def autosize(ws, min_width=8, max_width=50):
    widths = {}
    for row in ws.iter_rows(max_row=min(ws.max_row, 500)):
        for cell in row:
            v = cell.value
            n = 12 if isinstance(v, str) and v.startswith("=") else _display_len(v)
            widths[cell.column_letter] = max(widths.get(cell.column_letter, 0), n)
    for letter, n in widths.items():
        ws.column_dimensions[letter].width = min(max(min_width, n + 2), max_width)


def write_table(ws, header, rows, formats=None):
    """Write a header + rows with a styled header, filter, frozen top row and sensible widths.

    formats: {column_index: number_format}
    """
    ws.append(header)
    for r in rows:
        ws.append(r)
    style_header(ws)
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(header))}{len(rows) + 1}"
    for idx, fmt in (formats or {}).items():
        for (cell,) in ws.iter_rows(min_row=2, max_row=len(rows) + 1, min_col=idx + 1, max_col=idx + 1):
            if not isinstance(cell.value, str):
                cell.number_format = fmt
    autosize(ws)


def number_format_for(values):
    has_decimals = any(isinstance(v, float) and not float(v).is_integer() for v in values)
    return "#,##0.00" if has_decimals else "#,##0"


def output_path(input_path, suffix, out=None):
    if out:
        out = os.path.expanduser(out)
        if not out.lower().endswith((".xlsx", ".xlsm")):
            out += ".xlsx"
        return out
    base, _ = os.path.splitext(os.path.expanduser(input_path))
    return f"{base} - {suffix}.xlsx"


def save(wb, path, overwrite=False):
    if os.path.exists(path) and not overwrite:
        fail(f"{path} already exists. Choose another name, or pass --overwrite to replace it.")
    folder = os.path.dirname(os.path.abspath(path))
    os.makedirs(folder, exist_ok=True)
    try:
        wb.save(path)
    except PermissionError:
        fail(f"Can't save {path}. If it is open in Excel, close it and try again.")
    return os.path.abspath(path)
