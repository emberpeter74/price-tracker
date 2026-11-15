"""The run file: every price found in one scan."""
from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from pathlib import Path

import openpyxl

from .excel_io import decimal, text, write_table
from .models import Label, PriceResult

RUN_HEADERS = ["Run date and time", "Item code", "Name", "Competitor group", "Client price ex VAT",
               "Company", "Price ex VAT", "Price inc VAT", "Calculated", "Match label",
               "Product page link", "Remark"]
SHEET = "Prices"


def run_file_name(run_at: datetime) -> str:
    return f"Prices {run_at:%Y-%m-%d %H%M}.xlsx"


def as_row(r: PriceResult) -> list:
    return [r.run_at, r.item_code, r.name, r.group, r.client_price_ex, r.company, r.price_ex,
            r.price_inc, r.calculated or None, r.label.value, r.url or None, r.remark or None]


def write_run(path: Path, results: list[PriceResult]) -> Path:
    return write_table(Path(path), SHEET, RUN_HEADERS, [as_row(r) for r in results])


def _money(value) -> Decimal | None:
    d = decimal(value)
    return d.quantize(Decimal("0.01")) if d is not None else None


def read_run(path: Path) -> list[PriceResult]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[SHEET] if SHEET in wb.sheetnames else wb.worksheets[0]
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        row = list(row) + [None] * (len(RUN_HEADERS) - len(row))
        if not isinstance(row[0], datetime) or not text(row[1]):
            continue
        try:
            label = Label(text(row[9]))
        except ValueError:
            continue
        price_ex, price_inc = decimal(row[6]), decimal(row[7])
        out.append(PriceResult(row[0], text(row[1]), text(row[2]), text(row[3]), decimal(row[4]),
                               text(row[5]), _money(price_ex), _money(price_inc), text(row[8]),
                               label, text(row[10]), text(row[11])))
    wb.close()
    return out


def list_runs(output_dir: Path) -> list[Path]:
    """Run files in the folder, oldest first. Excel's '~$' lock files are skipped."""
    folder = Path(output_dir)
    if not folder.is_dir():
        return []
    return sorted(p for p in folder.glob("Prices *.xlsx") if not p.name.startswith("~$"))
