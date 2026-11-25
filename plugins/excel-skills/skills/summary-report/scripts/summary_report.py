#!/usr/bin/env python3
"""Build a live summary report (totals by group, and by month) from a data sheet.

    python summary_report.py INPUT --group "Customer" --value "Amount" [--value "Qty"]
                             [--group2 "City"] [--date "Invoice Date"] [--sheet NAME]
                             [--out FILE] [--overwrite]

Writes a NEW workbook:
  Summary   one row per group: count, sum of each value column, share of the total, grand total
            and a bar chart. Every number is a COUNTIFS/SUMIFS formula on the Data sheet, so the
            report updates when the data changes.
  By Month  (with --date) the first value column per group per month, also as formulas
  Data      the source table with real numbers and dates
Prints a JSON summary with the same totals, computed in Python, for the assistant to explain.
"""
import argparse
import os
import sys
from collections import OrderedDict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import Workbook  # noqa: E402
from openpyxl.chart import BarChart, Reference  # noqa: E402
from openpyxl.utils import get_column_letter as L  # noqa: E402
from openpyxl.workbook.properties import CalcProperties  # noqa: E402

from xl_io import (DATE_FORMAT, TOTAL_FONT, autosize, emit, fail, find_column,  # noqa: E402
                   is_blank, output_path, parse_date, parse_number, read_table, save,
                   style_header, write_table)

BLANK = "(blank)"
MONEY = "#,##0;-#,##0"
MONEY2 = "#,##0.00;-#,##0.00"


def criteria(value, cell_ref):
    """SUMIFS criteria that matches `value` exactly (no wildcards, blanks handled)."""
    if value == BLANK:
        return '"="'
    if isinstance(value, str) and any(ch in value for ch in '*?~<>="'):
        esc = value.replace("~", "~~").replace("*", "~*").replace("?", "~?").replace('"', '""')
        return f'"={esc}"'
    return cell_ref


def group_key(v):
    if is_blank(v):
        return BLANK.casefold(), BLANK
    if isinstance(v, str):
        t = " ".join(v.split())
        return t.casefold(), t
    return v, v


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input")
    ap.add_argument("--group", required=True)
    ap.add_argument("--group2")
    ap.add_argument("--value", action="append", required=True, help="repeat for several columns")
    ap.add_argument("--date")
    ap.add_argument("--sheet")
    ap.add_argument("--out")
    ap.add_argument("--overwrite", action="store_true")
    a = ap.parse_args()

    t = read_table(a.input, a.sheet)
    header, rows = t["header"], t["rows"]
    g1 = find_column(header, a.group)
    g2 = find_column(header, a.group2) if a.group2 else None
    vals = []
    for spec in a.value:
        vals += [find_column(header, v.strip()) for v in spec.split(",") if v.strip()]
    dc = find_column(header, a.date) if a.date else None

    # Real numbers and dates in the Data sheet, so SUMIFS works.
    skipped = {header[v]: 0 for v in vals}
    bad_dates = 0
    for r in rows:
        for v in vals:
            if is_blank(r[v]):
                r[v] = None
                continue
            n = parse_number(r[v])
            if n is None:
                skipped[header[v]] += 1
            else:
                r[v] = n
        if dc is not None and not is_blank(r[dc]):
            d = parse_date(r[dc])
            if d is None:
                bad_dates += 1
            else:
                r[dc] = d
        for g in (g1, g2):
            if g is not None and isinstance(r[g], str):
                r[g] = " ".join(r[g].split()) or None

    # Python totals (for the JSON summary and for row order).
    groups = OrderedDict()
    for r in rows:
        k1, d1 = group_key(r[g1])
        k2, d2 = group_key(r[g2]) if g2 is not None else (None, None)
        e = groups.setdefault((k1, k2), {"labels": (d1, d2), "count": 0, "sums": [0.0] * len(vals)})
        e["count"] += 1
        for j, v in enumerate(vals):
            if isinstance(r[v], (int, float)):
                e["sums"][j] += r[v]
    if not groups:
        fail("No data rows found under the header.")
    order = sorted(groups.values(), key=lambda e: -abs(e["sums"][0]))
    totals = [sum(e["sums"][j] for e in order) for j in range(len(vals))]

    wb = Workbook()
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    summary = wb.active
    summary.title = "Summary"
    data = wb.create_sheet("Data")
    formats = {v: MONEY2 for v in vals}
    if dc is not None:
        formats[dc] = DATE_FORMAT
    write_table(data, header, rows, formats)
    col = lambda i: f"Data!${L(i + 1)}:${L(i + 1)}"

    # ---- Summary sheet
    labels = [header[g1]] + ([header[g2]] if g2 is not None else [])
    nlab = len(labels)
    summary["A1"] = f"Summary by {' and '.join(labels)}"
    summary["A1"].font = TOTAL_FONT
    summary["A2"] = f"Source: {os.path.basename(a.input)} [{t['sheet']}], {len(rows)} rows. Formulas update when the Data sheet changes."
    head = labels + ["Count"] + [f"Total {header[v]}" for v in vals] + [f"% of {header[vals[0]]}"]
    summary.append([])
    summary.append(head)
    style_header(summary, 4)
    first = 5
    for i, e in enumerate(order):
        r = first + i
        crit = [col(g1), criteria(e["labels"][0], f"$A{r}")]
        if g2 is not None:
            crit += [col(g2), criteria(e["labels"][1], f"$B{r}")]
        crit_s = ",".join(crit)
        line = list(e["labels"][:nlab]) + [f"=COUNTIFS({crit_s})"]
        line += [f"=SUMIFS({col(v)},{crit_s})" for v in vals]
        sum_col = L(nlab + 2)
        line.append(f"=IFERROR({sum_col}{r}/{sum_col}${first + len(order)},0)")
        summary.append(line)
    tr = first + len(order)
    total = ["Grand total"] + [None] * (nlab - 1)
    for j in range(len(vals) + 1):
        c = L(nlab + 1 + j)
        total.append(f"=SUM({c}{first}:{c}{tr - 1})")
    total.append(f"=IFERROR({L(nlab + 2)}{tr}/{L(nlab + 2)}{tr},0)")
    summary.append(total)
    for cell in summary[tr]:
        cell.font = TOTAL_FONT
    for row in summary.iter_rows(min_row=first, max_row=tr):
        row[nlab].number_format = "#,##0"
        for j in range(len(vals)):
            row[nlab + 1 + j].number_format = MONEY
        row[nlab + 1 + len(vals)].number_format = "0.0%"
    summary.freeze_panes = f"A{first}"
    autosize(summary, max_width=40)
    summary.column_dimensions["A"].width = max(summary.column_dimensions["A"].width, 18)

    if len(order) > 1:
        chart = BarChart()
        chart.type = "bar"
        chart.title = f"Total {header[vals[0]]} by {header[g1]}"
        chart.legend = None
        n = min(len(order), 15)
        chart.add_data(Reference(summary, min_col=nlab + 2, min_row=first, max_row=first + n - 1))
        chart.set_categories(Reference(summary, min_col=1, min_row=first, max_row=first + n - 1))
        chart.height, chart.width = 7 + n * 0.45, 16
        summary.add_chart(chart, f"{L(nlab + len(vals) + 4)}4")

    # ---- By Month sheet
    months = []
    if dc is not None:
        months = sorted({(r[dc].year, r[dc].month) for r in rows if hasattr(r[dc], "year")})
        if len(months) > 36:
            months = months[-36:]
    if months:
        bm = wb.create_sheet("By Month", 1)
        v0 = vals[0]
        bm["A1"] = f"Total {header[v0]} by month"
        bm["A1"].font = TOTAL_FONT
        bm.append([])
        bm.append([])
        names = [f"{y}-{m:02d}" for y, m in months]
        bm.append(labels + names + ["Total"])
        style_header(bm, 4)
        for i, e in enumerate(order):
            r = first + i
            crit = [col(g1), criteria(e["labels"][0], f"$A{r}")]
            if g2 is not None:
                crit += [col(g2), criteria(e["labels"][1], f"$B{r}")]
            line = list(e["labels"][:nlab])
            for y, m in months:
                ny, nm = (y + 1, 1) if m == 12 else (y, m + 1)
                line.append(f"=SUMIFS({col(v0)},{','.join(crit)},{col(dc)},\">=\"&DATE({y},{m},1),"
                            f"{col(dc)},\"<\"&DATE({ny},{nm},1))")
            line.append(f"=SUM({L(nlab + 1)}{r}:{L(nlab + len(months))}{r})")
            bm.append(line)
        tot = ["Grand total"] + [None] * (nlab - 1)
        for j in range(len(months) + 1):
            c = L(nlab + 1 + j)
            tot.append(f"=SUM({c}{first}:{c}{first + len(order) - 1})")
        bm.append(tot)
        for cell in bm[first + len(order)]:
            cell.font = TOTAL_FONT
        for row in bm.iter_rows(min_row=first, min_col=nlab + 1):
            for cell in row:
                cell.number_format = MONEY
        bm.freeze_panes = bm.cell(row=first, column=nlab + 1)
        autosize(bm, max_width=40)

    out = save(wb, output_path(a.input, "summary", a.out), a.overwrite)

    month_totals = {}
    if months:
        for r in rows:
            if hasattr(r[dc], "year") and isinstance(r[vals[0]], (int, float)):
                k = f"{r[dc].year}-{r[dc].month:02d}"
                month_totals[k] = month_totals.get(k, 0) + r[vals[0]]
    emit({
        "output": out,
        "rows": len(rows),
        "groups": len(order),
        "grand_total": {header[v]: round(totals[j], 2) for j, v in enumerate(vals)},
        "top_groups": [
            {**{labels[i]: e["labels"][i] for i in range(nlab)}, "count": e["count"],
             **{header[v]: round(e["sums"][j], 2) for j, v in enumerate(vals)}}
            for e in order[:15]],
        "by_month": {k: round(v, 2) for k, v in sorted(month_totals.items())},
        "non_numeric_values_ignored": {k: v for k, v in skipped.items() if v},
        "unreadable_dates": bad_dates,
    })


if __name__ == "__main__":
    main()
