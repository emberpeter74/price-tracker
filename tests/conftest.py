from datetime import datetime
from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest

from pricetracker.models import Competitor, Item, ProductLink, VatBasis

ITEM_HEADERS = ["Item code (MPN)", "Name", "Brand", "Competitor group",
                "Client price ex VAT", "EAN", "Search text"]
COMP_HEADERS = ["Competitor group", "Company", "Website", "Search status",
                "Product page status", "VAT basis", "Last checked", "Search address"]
LINK_HEADERS = ["Item code", "Company", "Product page link", "Source", "Locked",
                "Last result", "Last checked", "Manual price", "Manual price VAT basis",
                "Manual price date"]

FIXTURES = Path(__file__).parent / "fixtures"


def _make_xlsx(folder, name, headers, rows, sheet="Sheet1"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(headers)
    for r in rows:
        ws.append(list(r))
    path = Path(folder) / name
    wb.save(path)
    return path


@pytest.fixture
def make_xlsx():
    return _make_xlsx


@pytest.fixture
def load_page():
    def load(name):
        return (FIXTURES / "pages" / name).read_text(encoding="utf-8", errors="replace")
    return load


def item(code="F00373000", name="PIUSI EX50 Fuel Transfer Pump, 230v", brand="PIUSI",
         group="Piusi Equipment", price="343", ean="", search_text="", row=2):
    return Item(code, name, brand, group, Decimal(price) if price is not None else None,
                ean, search_text, row)


def comp(company="Agriemach", website="https://www.agriemach.com/", group="Piusi Equipment",
         search_status=None, page_status=None, vat=VatBasis.INC, search_address="", row=2):
    return Competitor(group, company, website, search_status, page_status, vat, None,
                      search_address, row)


def link(item_code="F00373000", company="Agriemach", url="", source="Found by tool",
         locked=False, manual_price=None, manual_vat=None, manual_date=None, row=2):
    return ProductLink(item_code, company, url, source, locked, "", None,
                       Decimal(manual_price) if manual_price is not None else None,
                       manual_vat, manual_date, row)


NOW = datetime(2026, 11, 20, 9, 30)
