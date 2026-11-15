"""Data model shared by every module."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from enum import Enum


class Access(str, Enum):
    REACHABLE = "Reachable"
    REFUSING = "Refusing"
    UNREACHABLE = "Unreachable"
    NO_SEARCH = "Search not worked out"


class VatBasis(str, Enum):
    INC = "Inc VAT"
    EX = "Ex VAT"
    BOTH = "Both shown"
    UNKNOWN = "Unknown"


class Label(str, Enum):
    CLIENT_LINK = "Client link"
    VERIFIED = "Verified"
    UNVERIFIED = "Unverified"
    MANUAL = "Manual price"
    NOT_FOUND = "Not found"
    BLOCKED = "Blocked"
    ERROR = "Error"


CONFIRMED = {Label.CLIENT_LINK, Label.VERIFIED, Label.MANUAL}
SOURCE_TOOL = "Found by tool"
SOURCE_CLIENT = "Entered by client"
ERROR = "Error"
QUESTION = "Question"


@dataclass
class Item:
    code: str
    name: str
    brand: str
    group: str
    client_price_ex: Decimal | None
    ean: str
    search_text: str
    row: int


@dataclass
class Competitor:
    group: str
    company: str
    website: str
    search_status: Access | None
    page_status: Access | None
    vat_basis: VatBasis
    last_checked: datetime | None
    search_address: str  # tool column: URL template containing "{q}", or ""
    row: int


@dataclass
class ProductLink:
    item_code: str
    company: str
    url: str
    source: str
    locked: bool
    last_result: str
    last_checked: datetime | None
    manual_price: Decimal | None
    manual_vat: VatBasis | None
    manual_date: date | None
    row: int


@dataclass
class PriceResult:
    run_at: datetime
    item_code: str
    name: str
    group: str
    client_price_ex: Decimal | None
    company: str
    price_ex: Decimal | None
    price_inc: Decimal | None
    calculated: str  # "", "Ex VAT" or "Inc VAT": which figure was computed
    label: Label
    url: str
    remark: str


@dataclass
class Issue:
    kind: str  # "Error" or "Question"
    file: str
    row: int | None
    subject: str
    problem: str
    client_value: str = ""
    seen_value: str = ""
    fix: str = ""
