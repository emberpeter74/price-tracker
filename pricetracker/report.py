"""The report command: merge all run files into the comparison workbook."""
from __future__ import annotations

from collections import Counter, defaultdict
from decimal import Decimal
from pathlib import Path

import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .excel_io import InputError, _autosize, _style_header, save_workbook
from .models import CONFIRMED, Label, PriceResult
from .runfile import RUN_HEADERS, as_row, list_runs, read_run

REPORT_NAME = "Price comparison.xlsx"
RED, GREEN = "FFC7CE", "C6EFCE"
MONEY, PERCENT = "#,##0.00", "0.0%"
CHARTS_PER_ROW, CHART_ROWS, CHART_COLS = 2, 17, 10


def _num(value):
    return float(value) if isinstance(value, Decimal) else value


def _append(ws, row) -> None:
    ws.append([_num(v) for v in row])


def _finish(ws) -> None:
    _style_header(ws)
    _autosize(ws)


def _newest_per_pair(results: list[PriceResult]) -> dict[tuple[str, str], PriceResult]:
    newest: dict[tuple[str, str], PriceResult] = {}
    for r in results:
        key = (r.item_code, r.company)
        if key not in newest or r.run_at >= newest[key].run_at:
            newest[key] = r
    return newest


def _confirmed(r: PriceResult) -> bool:
    return r.label in CONFIRMED and r.price_ex is not None


def _latest_sheet(ws, newest: dict, companies: list[str]) -> None:
    headers = (["Item code", "Name", "Client price ex VAT"] + companies +
               ["Cheapest competitor", "Cheapest price", "Difference £", "Difference %"])
    ws.append(headers)
    by_item: dict[str, dict[str, PriceResult]] = defaultdict(dict)
    for (code, company), r in newest.items():
        by_item[code][company] = r
    for code in sorted(by_item):
        per = by_item[code]
        last = max(per.values(), key=lambda r: r.run_at)
        client = last.client_price_ex
        prices = {c: per[c].price_ex for c in companies if c in per and _confirmed(per[c])}
        cheapest = min(prices, key=lambda c: (prices[c], c)) if prices else None
        diff = prices[cheapest] - client if cheapest and client is not None else None
        pct = float(diff / client) if diff is not None and client else None
        _append(ws, [code, last.name, client] + [prices.get(c) for c in companies] +
                [cheapest, prices.get(cheapest) if cheapest else None, diff, pct])
        row = ws.max_row
        if diff is not None:
            fill = PatternFill("solid", start_color=RED if diff < 0 else GREEN)
            for col in (3, len(headers) - 1, len(headers)):
                ws.cell(row, col).fill = fill
        for col in range(3, len(headers)):
            ws.cell(row, col).number_format = MONEY
        ws.cell(row, len(headers)).number_format = PERCENT
    _finish(ws)
    ws.freeze_panes = "D2"


def _unverified_sheet(ws, newest: dict) -> None:
    ws.append(["Item code", "Name", "Company", "Client price ex VAT", "Price ex VAT", "Price inc VAT",
               "Product page link", "Remark"])
    for key in sorted(newest):
        r = newest[key]
        if r.label == Label.UNVERIFIED:
            _append(ws, [r.item_code, r.name, r.company, r.client_price_ex, r.price_ex, r.price_inc,
                         r.url or None, r.remark or None])
    _finish(ws)


def _history_sheet(ws, results: list[PriceResult]) -> None:
    ws.append(RUN_HEADERS)
    for r in sorted(results, key=lambda r: (r.run_at, r.item_code, r.company)):
        _append(ws, as_row(r))
    _finish(ws)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(RUN_HEADERS))}{ws.max_row}"


def _graphs(ws_graphs, ws_data, results: list[PriceResult]) -> None:
    ws_graphs["A1"] = ("Price history per item, ex VAT. A line needs at least two runs on different "
                       "days; one run shows a single point.")
    ws_graphs["A1"].font = Font(italic=True)
    runs = sorted({r.run_at for r in results})
    by_item: dict[str, list[PriceResult]] = defaultdict(list)
    for r in results:
        by_item[r.item_code].append(r)
    n = 0
    for code in sorted(by_item):
        rows = by_item[code]
        companies = sorted({r.company for r in rows if _confirmed(r)})
        if not companies:
            continue
        name = max(rows, key=lambda r: r.run_at).name
        title = f"{code} – {name}"
        top = ws_data.max_row + (2 if ws_data.max_row > 1 else 0)
        top = 1 if ws_data.max_row == 1 and ws_data["A1"].value is None else top
        ws_data.cell(top, 1, title).font = Font(bold=True)
        for i, head in enumerate(["Run", "Client price"] + companies, 1):
            ws_data.cell(top + 1, i, head)
        for j, run_at in enumerate(runs):
            at = [r for r in rows if r.run_at == run_at]
            client = next((r.client_price_ex for r in at if r.client_price_ex is not None), None)
            price = {r.company: r.price_ex for r in at if _confirmed(r)}
            line = [f"{run_at:%d %b %Y %H:%M}", client] + [price.get(c) for c in companies]
            for i, v in enumerate(line, 1):
                ws_data.cell(top + 2 + j, i, _num(v))
        last = top + 1 + len(runs)

        chart = LineChart()
        chart.title, chart.height, chart.width = title[:80], 8, 16
        chart.y_axis.title, chart.y_axis.number_format = "£ ex VAT", "#,##0"
        chart.add_data(Reference(ws_data, min_col=2, max_col=2 + len(companies), min_row=top + 1,
                                 max_row=last), titles_from_data=True)
        chart.set_categories(Reference(ws_data, min_col=1, min_row=top + 2, max_row=last))
        for series in chart.series:
            series.marker.symbol = "circle"
            series.marker.size = 6
        col = 1 + (n % CHARTS_PER_ROW) * CHART_COLS
        ws_graphs.add_chart(chart, f"{get_column_letter(col)}{3 + (n // CHARTS_PER_ROW) * CHART_ROWS}")
        n += 1
    ws_data.column_dimensions["A"].width = 20


def _coverage_sheet(ws, results: list[PriceResult]) -> None:
    labels = [Label.CLIENT_LINK, Label.VERIFIED, Label.MANUAL, Label.UNVERIFIED, Label.NOT_FOUND,
              Label.BLOCKED, Label.ERROR]
    ws.append(["Company"] + [l.value for l in labels] + ["Total"])
    latest = max(r.run_at for r in results)
    counts: dict[str, Counter] = defaultdict(Counter)
    for r in results:
        if r.run_at == latest:
            counts[r.company][r.label] += 1
    for company in sorted(counts):
        c = counts[company]
        ws.append([company] + [c[l] for l in labels] + [sum(c.values())])
    _finish(ws)


def build_report(output_dir: Path) -> Path:
    output_dir = Path(output_dir)
    runs = list_runs(output_dir)
    results = [r for path in runs for r in read_run(path)]
    if not results:
        raise InputError(f"No run files with prices were found in {output_dir}. Run the scan first.")
    newest = _newest_per_pair(results)
    companies = sorted({r.company for r in results})

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Latest comparison"
    _latest_sheet(ws, newest, companies)
    _unverified_sheet(wb.create_sheet("Unverified matches"), newest)
    _history_sheet(wb.create_sheet("History"), results)
    graphs, data = wb.create_sheet("Graphs"), wb.create_sheet("Graph data")
    _graphs(graphs, data, results)
    _coverage_sheet(wb.create_sheet("Coverage"), results)
    return save_workbook(wb, output_dir / REPORT_NAME)
