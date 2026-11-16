from datetime import date, datetime
from decimal import Decimal

import openpyxl
import pytest

from conftest import COMP_HEADERS, LINK_HEADERS
from pricetracker import excel_io
from pricetracker.excel_io import (InputError, join_note, read_competitors, read_items, read_links,
                                   save_workbook, split_note, write_competitors, write_links,
                                   write_table)
from pricetracker.issues import ISSUE_HEADERS, IssueLog
from pricetracker.models import Access, Connection, ProductLink, VatBasis


def test_reads_items_with_mangled_cells(tmp_path, make_xlsx):
    p = make_xlsx(tmp_path, "Items.xlsx",
                  ["Item code (MPN) ", "name", "Brand", "Competitor group",
                   "Client price ex VAT", "EAN", "Search text"],
                  [[305000, "PIUSI E80 Pump,\xa0 230v ", "PIUSI", "Piusi Equipment", 304,
                    8050031000670, None],
                   [None, None, None, None, None, None, None],
                   ["ZVA25", "Elaflex ZVA25", "Elaflex", "Piusi Equipment", None, "#N/A", ""]])
    a, b = read_items(p)
    assert a.code == "305000" and a.name == "PIUSI E80 Pump, 230v"
    assert a.client_price_ex == Decimal("304")
    assert a.ean == "8050031000670" and a.row == 2
    assert b.row == 4 and b.client_price_ex is None and b.ean == ""


def test_float_code_and_price_text(tmp_path, make_xlsx):
    p = make_xlsx(tmp_path, "Items.xlsx", ["Item code (MPN)", "Name", "Competitor group",
                                           "Client price ex VAT"],
                  [[70062.0, "Filter", "G", "£1,902.50"], ["A1", "x", "G", "n/a"]])
    a, b = read_items(p)
    assert a.code == "70062" and a.client_price_ex == Decimal("1902.50")
    assert a.brand == "" and a.ean == "" and a.search_text == ""
    assert b.client_price_ex is None


def test_missing_required_header_names_the_file_and_column(tmp_path, make_xlsx):
    p = make_xlsx(tmp_path, "Items.xlsx", ["Name"], [["x"]])
    with pytest.raises(InputError, match=r"Items\.xlsx.*Item code \(MPN\)"):
        read_items(p)


def test_missing_file_raises_input_error_but_links_are_optional(tmp_path):
    with pytest.raises(InputError, match=r"Items\.xlsx"):
        read_items(tmp_path / "Items.xlsx")
    assert read_links(tmp_path / "Product links.xlsx") == []


def test_reads_competitors_with_empty_and_filled_statuses(tmp_path, make_xlsx):
    p = make_xlsx(tmp_path, "Competitors.xlsx", COMP_HEADERS, [
        ["Tanks", "Tanks R Us", "https://www.tanksrus.co.uk/", None, None, None, None, None],
        ["Tanks", "Fuel Tank Shop", "https://www.fueltankshop.co.uk/", "refusing", "Reachable",
         "inc vat", datetime(2026, 11, 1, 8, 0), "https://x/search?q={q}"],
    ])
    a, b = read_competitors(p)
    assert a.search_status is None and a.page_status is None and a.vat_basis == VatBasis.UNKNOWN
    assert b.search_status == Access.REFUSING and b.page_status == Access.REACHABLE
    assert b.vat_basis == VatBasis.INC and b.search_address == "https://x/search?q={q}"
    assert b.last_checked == datetime(2026, 11, 1, 8, 0)


def test_reads_links_with_manual_price(tmp_path, make_xlsx):
    p = make_xlsx(tmp_path, "Product links.xlsx", LINK_HEADERS, [
        ["A1", "Landowner", None, None, "yes", None, None, 555, "Ex VAT", date(2026, 11, 10)],
        ["A2", "Agriemach", "https://www.agriemach.com/p", "Entered by client", None, None,
         None, None, None, None],
    ])
    a, b = read_links(p)
    assert a.locked is True and a.manual_price == Decimal("555")
    assert a.manual_vat == VatBasis.EX and a.manual_date == date(2026, 11, 10) and a.url == ""
    assert b.locked is False and b.source == "Entered by client" and b.manual_price is None


def test_write_links_preserves_client_columns_and_appends(tmp_path, make_xlsx):
    p = make_xlsx(tmp_path, "Product links.xlsx", LINK_HEADERS + ["My notes"], [
        ["A1", "Agriemach", "https://old/p", "Found by tool", "Yes", None, None, 12.5, "Ex VAT",
         date(2026, 11, 10), "keep me"],
    ])
    rows = read_links(p)
    rows[0].url = "https://new/p"
    rows[0].last_result = "Verified"
    rows[0].last_checked = datetime(2026, 11, 20, 9, 30)
    rows.append(ProductLink("A2", "Hytek", "https://hytekgb.com/p", "Found by tool", False,
                            "Unverified", datetime(2026, 11, 20, 9, 30), None, None, None, 0))
    out = write_links(p, rows)
    ws = openpyxl.load_workbook(out).active
    assert ws.max_row == 3
    assert ws["C2"].value == "https://new/p" and ws["F2"].value == "Verified"
    assert ws["E2"].value == "Yes" and ws["H2"].value == 12.5 and ws["K2"].value == "keep me"
    assert [ws.cell(3, c).value for c in (1, 2, 3, 4, 6)] == [
        "A2", "Hytek", "https://hytekgb.com/p", "Found by tool", "Unverified"]
    again = read_links(out)
    assert again[0].locked is True and again[1].row == 3


def test_write_links_never_writes_a_read_only_url(tmp_path, make_xlsx):
    p = make_xlsx(tmp_path, "Product links.xlsx", LINK_HEADERS, [
        ["A1", "Chemfast", "https://www.chemfastuk.com/collections/fuelbox", "Entered by client",
         None, None, None, None, None, None]])
    rows = read_links(p)
    rows[0].url = ""
    rows[0].url_readonly = True
    rows[0].last_result = "Not found"
    ws = openpyxl.load_workbook(write_links(p, rows)).active
    assert ws["C2"].value == "https://www.chemfastuk.com/collections/fuelbox"
    assert ws["F2"].value == "Not found"


def test_write_links_creates_missing_file(tmp_path):
    p = tmp_path / "Product links.xlsx"
    write_links(p, [ProductLink("A1", "Hytek", "https://hytekgb.com/p", "Found by tool", False,
                                "Verified", None, None, None, None, 0)])
    assert read_links(p)[0].url == "https://hytekgb.com/p"


def test_write_competitors_updates_tool_columns_and_adds_missing_ones(tmp_path, make_xlsx):
    p = make_xlsx(tmp_path, "Competitors.xlsx", ["Competitor group", "Company", "Website", "Note"],
                  [["Tanks", "Tanks R Us", "https://www.tanksrus.co.uk/", "mine"]])
    rows = read_competitors(p)
    rows[0].search_status = Access.NO_SEARCH
    rows[0].page_status = Access.REACHABLE
    rows[0].last_checked = datetime(2026, 11, 20, 9, 30)
    rows[0].search_address = "https://www.tanksrus.co.uk/?s={q}"
    out = write_competitors(p, rows)
    ws = openpyxl.load_workbook(out).active
    assert ws["D2"].value == "mine" and ws["A2"].value == "Tanks"
    again = read_competitors(out)[0]
    assert again.search_status == Access.NO_SEARCH and again.page_status == Access.REACHABLE
    assert again.search_address == "https://www.tanksrus.co.uk/?s={q}"
    assert again.vat_basis == VatBasis.UNKNOWN


def test_note_round_trip_and_junk_is_ignored():
    assert join_note("forbidden by the site's rules (robots.txt)", "") == \
        "Search: forbidden by the site's rules (robots.txt)"
    both = join_note("a", "b")
    assert both == "Search: a\nProduct pages: b" and split_note(both) == ("a", "b")
    assert split_note("hello\nSearch: x") == ("x", "") and split_note("") == ("", "")
    assert join_note("", "") == ""


def test_competitors_new_columns_round_trip(tmp_path, make_xlsx):
    p = make_xlsx(tmp_path, "Competitors.xlsx", COMP_HEADERS,
                  [["Tanks", "T4E", "https://www.t4e.co.uk/", None, None, None, None, None]])
    rows = read_competitors(p)
    assert rows[0].connection is None and rows[0].site_checked is None and rows[0].search_note == ""
    rows[0].connection = Connection.HTTP1
    rows[0].search_note, rows[0].page_note = "x", "y"
    rows[0].site_checked = datetime(2026, 11, 20, 9, 30)
    out = write_competitors(p, rows)
    headers = [c.value for c in openpyxl.load_workbook(out).active[1]]
    assert headers[-3:] == ["Connection", "Status note", "Site checked"]
    again = read_competitors(out)[0]
    assert (again.connection, again.search_note, again.page_note, again.site_checked) == (
        Connection.HTTP1, "x", "y", datetime(2026, 11, 20, 9, 30))


def test_junk_connection_reads_as_untested(tmp_path, make_xlsx):
    p = make_xlsx(tmp_path, "Competitors.xlsx", COMP_HEADERS + ["Connection"],
                  [["Tanks", "T4E", "https://www.t4e.co.uk/", None, None, None, None, None, "fast"]])
    assert read_competitors(p)[0].connection is None


def test_save_falls_back_when_file_is_open(tmp_path, monkeypatch):
    wb = openpyxl.Workbook()
    target = tmp_path / "Product links.xlsx"
    real_save = openpyxl.Workbook.save

    def fake_save(self, filename):
        if str(filename) == str(target):
            raise PermissionError("open in Excel")
        return real_save(self, filename)

    monkeypatch.setattr(openpyxl.Workbook, "save", fake_save)
    out = save_workbook(wb, target)
    assert out.name == "Product links (new).xlsx" and out.exists()


def test_write_table_round_trip(tmp_path):
    p = write_table(tmp_path / "t.xlsx", "Data", ["A", "B"], [[1, "x"], [2, None]])
    ws = openpyxl.load_workbook(p)["Data"]
    assert [c.value for c in ws[1]] == ["A", "B"] and ws["B2"].value == "x" and ws.max_row == 3


def test_empty_issue_log_writes_no_issues_found(tmp_path):
    p = IssueLog().write(tmp_path / "Data issues.xlsx")
    ws = openpyxl.load_workbook(p).active
    assert [c.value for c in ws[1]] == ISSUE_HEADERS and ws["A2"].value == "No issues found"


def test_issue_log_writes_errors_before_questions(tmp_path):
    log = IssueLog()
    log.add("Question", "Items.xlsx", 5, "A1", "No client price", fix="Fill in the price")
    log.add("Error", "Items.xlsx", 3, "A2", "Duplicate item code", "A2", "", "Remove one row")
    assert len(log) == 2
    ws = openpyxl.load_workbook(log.write(tmp_path / "Data issues.xlsx")).active
    assert [ws.cell(2, c).value for c in range(1, 9)] == [
        "Error", "Items.xlsx", 3, "A2", "Duplicate item code", "A2", None, "Remove one row"]
    assert ws["A3"].value == "Question"


def test_issue_log_drops_exact_duplicates():
    log = IssueLog()
    for _ in range(2):
        log.add("Question", "Items.xlsx", 5, "A1", "No client price")
    assert len(log) == 1
