from datetime import datetime
from decimal import Decimal as D

import openpyxl

from pricetracker.models import Label, PriceResult
from pricetracker.runfile import RUN_HEADERS, list_runs, read_run, run_file_name, write_run

RUN_AT = datetime(2026, 11, 20, 9, 30)


def result(**kw):
    base = dict(run_at=RUN_AT, item_code="F00373000", name="PIUSI EX50", group="Piusi Equipment",
                client_price_ex=D("343"), company="Agriemach", price_ex=D("121.78"),
                price_inc=D("146.14"), calculated="Ex VAT", label=Label.VERIFIED,
                url="https://www.agriemach.com/p", remark="")
    base.update(kw)
    return PriceResult(**base)


def test_run_file_name():
    assert run_file_name(RUN_AT) == "Prices 2026-11-20 0930.xlsx"


def test_write_then_read_round_trips_all_fields(tmp_path):
    rows = [result(),
            result(item_code="305000", company="Hytek", client_price_ex=None, price_ex=None,
                   price_inc=None, calculated="", label=Label.BLOCKED, url="", remark="Site blocks automated reading")]
    path = write_run(tmp_path / run_file_name(RUN_AT), rows)
    ws = openpyxl.load_workbook(path)["Prices"]
    assert [c.value for c in ws[1]] == RUN_HEADERS and ws["G2"].value == 121.78
    assert read_run(path) == rows


def test_numeric_item_code_survives_as_text(tmp_path):
    path = write_run(tmp_path / "Prices 2026-11-20 0930.xlsx", [result(item_code="305000")])
    assert read_run(path)[0].item_code == "305000"


def test_list_runs_is_sorted_and_ignores_other_files(tmp_path):
    for name in ["Prices 2026-11-21 0800.xlsx", "Prices 2026-11-20 0930.xlsx",
                 "~$Prices 2026-11-21 0800.xlsx", "Price comparison.xlsx", "Data issues.xlsx"]:
        (tmp_path / name).write_bytes(b"")
    assert [p.name for p in list_runs(tmp_path)] == ["Prices 2026-11-20 0930.xlsx",
                                                     "Prices 2026-11-21 0800.xlsx"]
    assert list_runs(tmp_path / "missing") == []
