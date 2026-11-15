from datetime import datetime
from decimal import Decimal as D

import openpyxl
import pytest

from pricetracker.excel_io import InputError
from pricetracker.models import Label
from pricetracker.report import GREEN, RED, build_report
from pricetracker.runfile import RUN_HEADERS, run_file_name, write_run
from test_runfile import result

D1, D2, D3 = datetime(2026, 11, 18, 9, 0), datetime(2026, 11, 19, 9, 0), datetime(2026, 11, 20, 9, 0)


def r(run_at, company, ex, label=Label.VERIFIED, code="A1", client="100", **kw):
    ex = D(ex) if ex is not None else None
    return result(run_at=run_at, item_code=code, name=f"Item {code}", company=company,
                  client_price_ex=D(client) if client is not None else None, price_ex=ex,
                  price_inc=(ex * D("1.2")).quantize(D("0.01")) if ex is not None else None,
                  calculated="Inc VAT" if ex is not None else "", label=label, **kw)


def build(tmp_path, *runs):
    for rows in runs:
        write_run(tmp_path / run_file_name(rows[0].run_at), rows)
    return openpyxl.load_workbook(build_report(tmp_path))


def table(ws):
    rows = list(ws.iter_rows(values_only=True))
    return [dict(zip(rows[0], row)) for row in rows[1:]]


def test_sheets_in_order(tmp_path):
    wb = build(tmp_path, [r(D1, "A", "90")])
    assert wb.sheetnames == ["Latest comparison", "Unverified matches", "History", "Graphs",
                             "Graph data", "Coverage"]


def test_latest_uses_newest_run_per_item_and_company(tmp_path):
    wb = build(tmp_path, [r(D1, "A", "90"), r(D1, "B", "110")], [r(D2, "A", "95", client="105")])
    row = table(wb["Latest comparison"])[0]
    assert row["A"] == 95 and row["B"] == 110 and row["Client price ex VAT"] == 105


def test_unverified_blocked_and_not_found_never_appear_as_prices_in_latest(tmp_path):
    wb = build(tmp_path, [r(D1, "A", "90"), r(D1, "B", "50", Label.UNVERIFIED), r(D1, "C", None, Label.BLOCKED)],
               [r(D2, "A", None, Label.NOT_FOUND)])
    row = table(wb["Latest comparison"])[0]
    assert row["A"] is None and row["B"] is None and row["C"] is None
    assert row["Cheapest competitor"] is None and row["Difference £"] is None


def test_cheapest_and_difference(tmp_path):
    wb = build(tmp_path, [r(D1, "A", "90"), r(D1, "B", "110"), r(D1, "M", "95", Label.MANUAL),
                          r(D1, "L", "99", Label.CLIENT_LINK)])
    row = table(wb["Latest comparison"])[0]
    assert (row["Cheapest competitor"], row["Cheapest price"]) == ("A", 90)
    assert row["Difference £"] == -10 and row["Difference %"] == pytest.approx(-0.10)
    assert row["M"] == 95 and row["L"] == 99


def fill_of(ws, header, row=2):
    col = [c.value for c in ws[1]].index(header) + 1
    return ws.cell(row, col).fill.start_color.rgb[-6:]


def test_client_more_expensive_than_cheapest_is_filled_red_and_cheaper_green(tmp_path):
    wb = build(tmp_path, [r(D1, "A", "90", code="A1"), r(D1, "A", "120", code="B2"),
                          r(D1, "A", "100", code="C3")])
    ws = wb["Latest comparison"]
    assert [ws.cell(i, 1).value for i in (2, 3, 4)] == ["A1", "B2", "C3"]
    assert fill_of(ws, "Difference £", 2) == RED and fill_of(ws, "Client price ex VAT", 2) == RED
    assert fill_of(ws, "Difference £", 3) == GREEN and fill_of(ws, "Difference £", 4) == GREEN


def test_item_without_client_price_has_empty_difference(tmp_path):
    wb = build(tmp_path, [r(D1, "A", "90", client=None)])
    row = table(wb["Latest comparison"])[0]
    assert row["Cheapest price"] == 90 and row["Difference £"] is None and row["Difference %"] is None


def test_history_has_one_row_per_result_across_runs(tmp_path):
    wb = build(tmp_path, [r(D1, "A", "90"), r(D1, "B", "110")], [r(D2, "A", "95")])
    ws = wb["History"]
    assert [c.value for c in ws[1]] == RUN_HEADERS and ws.max_row == 4
    assert ws.auto_filter.ref == "A1:L4"


def test_one_chart_per_item_with_confirmed_prices(tmp_path):
    wb = build(tmp_path, [r(D1, "A", "90", code="A1"), r(D1, "A", "50", Label.UNVERIFIED, code="B2"),
                          r(D1, "A", "70", code="C3")], [r(D2, "A", "91", code="A1")])
    assert len(wb["Graphs"]._charts) == 2


def test_chart_series_include_client_price_and_each_company(tmp_path):
    wb = build(tmp_path, [r(D1, "A", "90"), r(D1, "B", "110"), r(D1, "C", None, Label.BLOCKED)],
               [r(D2, "A", "95"), r(D2, "B", "108")], [r(D3, "A", "96")])
    chart = wb["Graphs"]._charts[0]
    assert len(chart.series) == 3          # client, A, B  (C never had a price)
    data = list(wb["Graph data"].iter_rows(values_only=True))
    assert data[0][0] == "A1 – Item A1" and data[1][:4] == ("Run", "Client price", "A", "B")
    assert data[2][1:4] == (100, 90, 110) and data[4][1:4] == (100, 96, None)
    assert data[2][0] == "18 Nov 2026 09:00"


def test_single_run_still_builds_charts(tmp_path):
    wb = build(tmp_path, [r(D1, "A", "90")])
    assert len(wb["Graphs"]._charts) == 1


def test_coverage_counts_by_label_for_latest_run(tmp_path):
    wb = build(tmp_path, [r(D1, "A", "90", code="A1"), r(D1, "A", "90", code="B2")],
               [r(D2, "A", "90", code="A1"), r(D2, "A", None, Label.BLOCKED, code="B2"),
                r(D2, "B", "80", Label.UNVERIFIED, code="A1")])
    rows = {row["Company"]: row for row in table(wb["Coverage"])}
    assert (rows["A"]["Verified"], rows["A"]["Blocked"], rows["A"]["Total"]) == (1, 1, 2)
    assert (rows["B"]["Unverified"], rows["B"]["Verified"], rows["B"]["Total"]) == (1, 0, 1)


def test_unverified_matches_sheet_lists_them(tmp_path):
    wb = build(tmp_path, [r(D1, "A", "90"), r(D1, "B", "50", Label.UNVERIFIED, url="https://b/p",
                                              remark="Matched by name only")])
    rows = table(wb["Unverified matches"])
    assert len(rows) == 1 and rows[0]["Company"] == "B" and rows[0]["Price ex VAT"] == 50
    assert rows[0]["Product page link"] == "https://b/p" and rows[0]["Remark"] == "Matched by name only"


def test_no_run_files_raises_input_error(tmp_path):
    with pytest.raises(InputError, match="scan"):
        build_report(tmp_path)


def test_report_file_open_in_excel_falls_back_to_new_name(tmp_path, monkeypatch):
    write_run(tmp_path / run_file_name(D1), [r(D1, "A", "90")])
    real_save = openpyxl.Workbook.save

    def fake_save(self, filename):
        if str(filename).endswith("Price comparison.xlsx"):
            raise PermissionError("open")
        return real_save(self, filename)
    monkeypatch.setattr(openpyxl.Workbook, "save", fake_save)
    assert build_report(tmp_path).name == "Price comparison (new).xlsx"


def test_report_is_not_read_back_as_a_run_file(tmp_path):
    build(tmp_path, [r(D1, "A", "90")])
    wb = openpyxl.load_workbook(build_report(tmp_path))      # second build, report already present
    assert wb["History"].max_row == 2
