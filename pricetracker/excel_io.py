"""Reading and writing the Excel input files."""
from __future__ import annotations

import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from .models import Access, Competitor, Connection, Item, ProductLink, VatBasis

ITEM_COLS = {"code": "Item code (MPN)", "name": "Name", "brand": "Brand",
             "group": "Competitor group", "price": "Client price ex VAT", "ean": "EAN",
             "search": "Search text"}
COMP_COLS = {"group": "Competitor group", "company": "Company", "website": "Website",
             "search_status": "Search status", "page_status": "Product page status",
             "vat": "VAT basis", "checked": "Last checked", "address": "Search address",
             "connection": "Connection", "note": "Status note", "site_checked": "Site checked"}
_SEARCH_PREFIX, _PAGE_PREFIX = "Search: ", "Product pages: "
LINK_COLS = {"code": "Item code", "company": "Company", "url": "Product page link",
             "source": "Source", "locked": "Locked", "result": "Last result",
             "checked": "Last checked", "mprice": "Manual price",
             "mvat": "Manual price VAT basis", "mdate": "Manual price date"}
_EMPTY_MARKERS = {"#n/a", "null", "#value!", "#ref!"}


class InputError(Exception):
    """An input file or a required column is missing. The message is shown to the user."""


# ---------- cell cleaning ----------

def text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "Yes" if value else "No"
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    s = re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()
    return "" if s.lower() in _EMPTY_MARKERS else s


def decimal(value) -> Decimal | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    s = text(value).replace("£", "").replace(",", "").strip()
    try:
        return Decimal(s) if s else None
    except InvalidOperation:
        return None


def _enum(cls, value):
    s = text(value).casefold()
    for member in cls:
        if member.value.casefold() == s:
            return member
    return None


def _datetime(value) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    s = text(value)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def _date(value) -> date | None:
    d = _datetime(value)
    return d.date() if d else None


def _yes(value) -> bool:
    return text(value).casefold() in {"yes", "y", "true", "1", "x"}


def join_note(search_note: str, page_note: str) -> str:
    """The Status note cell: one line per part that has something to say."""
    parts = [(_SEARCH_PREFIX, search_note), (_PAGE_PREFIX, page_note)]
    return "\n".join(prefix + note for prefix, note in parts if note)


def split_note(cell) -> tuple[str, str]:
    """(search note, product page note) from a Status note cell; other lines are ignored."""
    search = page = ""
    for line in ("" if cell is None else str(cell)).splitlines():
        line = line.strip()
        if line.startswith(_SEARCH_PREFIX):
            search = line[len(_SEARCH_PREFIX):].strip()
        elif line.startswith(_PAGE_PREFIX):
            page = line[len(_PAGE_PREFIX):].strip()
    return search, page


# ---------- generic sheet access ----------

def _norm(header) -> str:
    return text(header).casefold()


def _open(path: Path):
    path = Path(path)
    if not path.exists():
        raise InputError(f"The input file {path.name} was not found in {path.parent}.")
    try:
        wb = openpyxl.load_workbook(path)
    except PermissionError:
        raise InputError(f"{path.name} could not be opened. Close it in Excel and try again.")
    except Exception as exc:  # corrupt or not an xlsx
        raise InputError(f"{path.name} could not be read as an Excel file ({exc}).")
    return wb, wb.worksheets[0]


def _header_map(ws) -> dict[str, int]:
    return {_norm(c.value): c.column for c in ws[1] if text(c.value)}


def _columns(ws, path: Path, cols: dict[str, str], required: list[str]) -> dict[str, int | None]:
    headers = _header_map(ws)
    found = {key: headers.get(_norm(title)) for key, title in cols.items()}
    missing = [cols[k] for k in required if found[k] is None]
    if missing:
        raise InputError(f"{Path(path).name} has no column named {', '.join(missing)}. "
                         f"Add it to the first row.")
    return found


def _rows(ws, cols: dict[str, int | None]):
    """Yield (row number, {key: raw value}) for every row that is not completely blank."""
    for r in range(2, ws.max_row + 1):
        values = {k: (ws.cell(r, c).value if c else None) for k, c in cols.items()}
        if any(text(v) for v in values.values()):
            yield r, values


# ---------- readers ----------

def read_items(path: Path) -> list[Item]:
    _, ws = _open(path)
    cols = _columns(ws, path, ITEM_COLS, ["code", "name", "group"])
    return [Item(text(v["code"]), text(v["name"]), text(v["brand"]), text(v["group"]),
                 decimal(v["price"]), text(v["ean"]), text(v["search"]), r)
            for r, v in _rows(ws, cols)]


def read_competitors(path: Path) -> list[Competitor]:
    _, ws = _open(path)
    cols = _columns(ws, path, COMP_COLS, ["group", "company", "website"])
    return [Competitor(text(v["group"]), text(v["company"]), text(v["website"]),
                       _enum(Access, v["search_status"]), _enum(Access, v["page_status"]),
                       _enum(VatBasis, v["vat"]) or VatBasis.UNKNOWN, _datetime(v["checked"]),
                       text(v["address"]), r, _enum(Connection, v["connection"]),
                       *split_note(v["note"]), _datetime(v["site_checked"]))
            for r, v in _rows(ws, cols)]


def read_links(path: Path) -> list[ProductLink]:
    if not Path(path).exists():
        return []
    _, ws = _open(path)
    cols = _columns(ws, path, LINK_COLS, ["code", "company"])
    return [ProductLink(text(v["code"]), text(v["company"]), text(v["url"]), text(v["source"]),
                        _yes(v["locked"]), text(v["result"]), _datetime(v["checked"]),
                        decimal(v["mprice"]), _enum(VatBasis, v["mvat"]), _date(v["mdate"]), r)
            for r, v in _rows(ws, cols)]


# ---------- writers ----------

def save_workbook(wb, path: Path) -> Path:
    """Save; if the file is open in Excel, save beside it as '<name> (new).xlsx'."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(path)
        return path
    except PermissionError:
        alt = path.with_name(f"{path.stem} (new){path.suffix}")
        wb.save(alt)
        return alt


def _style_header(ws) -> None:
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"


def _autosize(ws, limit: int = 60) -> None:
    widths: dict[int, int] = {}
    for row in ws.iter_rows(values_only=True):
        for i, v in enumerate(row, 1):
            if v is not None:
                widths[i] = max(widths.get(i, 0), len(str(v)))
    for i, w in widths.items():
        ws.column_dimensions[get_column_letter(i)].width = min(max(w + 2, 8), limit)


def write_table(path: Path, sheet: str, headers: list[str], rows: list[list]) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(headers)
    for row in rows:
        ws.append([float(v) if isinstance(v, Decimal) else v for v in row])
    _style_header(ws)
    _autosize(ws)
    return save_workbook(wb, path)


def _ensure_columns(ws, cols: dict[str, str]) -> dict[str, int]:
    """Column number for every title, adding any missing title to the end of the header row."""
    headers = _header_map(ws)
    out = {}
    for key, title in cols.items():
        col = headers.get(_norm(title))
        if col is None:
            col = max(headers.values(), default=0) + 1
            ws.cell(1, col).value = title
            ws.cell(1, col).font = Font(bold=True)
            headers[_norm(title)] = col
        out[key] = col
    return out


def _load_or_new(path: Path, sheet: str):
    path = Path(path)
    if path.exists():
        wb = openpyxl.load_workbook(path)
        return wb, wb.worksheets[0]
    wb = openpyxl.Workbook()
    wb.active.title = sheet
    return wb, wb.active


def _val(member):
    return member.value if member is not None else None


def write_competitors(path: Path, rows: list[Competitor]) -> Path:
    """Update the tool's columns in place; client columns and formatting are left alone."""
    wb, ws = _load_or_new(path, "Competitors")
    c = _ensure_columns(ws, COMP_COLS)
    for comp in rows:
        r = comp.row
        ws.cell(r, c["search_status"]).value = _val(comp.search_status)
        ws.cell(r, c["page_status"]).value = _val(comp.page_status)
        if not text(ws.cell(r, c["vat"]).value):
            ws.cell(r, c["vat"]).value = comp.vat_basis.value
        ws.cell(r, c["checked"]).value = comp.last_checked
        ws.cell(r, c["address"]).value = comp.search_address or None
        ws.cell(r, c["connection"]).value = _val(comp.connection)
        ws.cell(r, c["note"]).value = join_note(comp.search_note, comp.page_note) or None
        ws.cell(r, c["site_checked"]).value = comp.site_checked
    return save_workbook(wb, path)


def write_links(path: Path, rows: list[ProductLink]) -> Path:
    """Update tool columns of existing rows and append new rows. Never touches the client's
    Locked and Manual price columns, nor a link marked read-only."""
    wb, ws = _load_or_new(path, "Product links")
    c = _ensure_columns(ws, LINK_COLS)
    next_row = max(ws.max_row, 1) + 1
    for link in rows:
        new = link.row == 0
        if new:
            link.row = next_row
            next_row += 1
            ws.cell(link.row, c["code"]).value = link.item_code
            ws.cell(link.row, c["company"]).value = link.company
        r = link.row
        if not link.url_readonly:
            ws.cell(r, c["url"]).value = link.url or None
            ws.cell(r, c["source"]).value = link.source or None
        ws.cell(r, c["result"]).value = link.last_result or None
        ws.cell(r, c["checked"]).value = link.last_checked
    if ws.max_row > 1:
        _style_header(ws)
    return save_workbook(wb, path)
