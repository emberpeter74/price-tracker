# Price Tracker POC Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A single Windows `.exe` that reads items, competitors and product links from Excel, reads competitor web prices, writes a dated run file plus an issues file, and merges all runs into a comparison workbook with graphs.

**Architecture:** A Python package `pricetracker` with two commands, `scan` and `report`. `scan` is a pipeline: read and validate inputs → check site access → for each item × competitor resolve a product page (saved link, site search, or sitemap) → extract structured price data → match and label → write run file, issues file, and write statuses and links back to the input files. `report` only reads run files. All network access goes through one `Fetcher` that enforces robots rules and per-site pacing, so every other module is testable offline with saved HTML fixtures.

**Tech Stack:** Python 3.13 (`py -3.13`; the machine default 3.15 is too new for PyInstaller wheels), `httpx[http2]`, `selectolax`, `openpyxl`, `pytest`, `PyInstaller`.

**Spec:** [`docs/business-specification.md`](business-specification.md) — read it before any task.

> This is the plan as written before the build. Section 12 of the specification lists what changed during the build.

## Global Constraints

- All paths below are relative to the project root. `example files/` holds the client's original Excel files (not part of the repository): never modify them.
- Runtime folders, next to the exe (or the working directory in development): `input/` (`Items.xlsx`, `Competitors.xlsx`, `Product links.xlsx`) and `output/` (`Prices YYYY-MM-DD HHMM.xlsx`, `Data issues.xlsx`, `Price comparison.xlsx`).
- Excel column headers are exactly those in spec sections 5 and 6. Header matching is case-insensitive and whitespace-trimmed; column order in the file does not matter.
- Money is `decimal.Decimal`, rounded to 2 places only when written. Currency is GBP only; a non-GBP price is an `Error` result with remark `Currency <code>`.
- VAT rate is 20%. Comparison is always ex VAT.
- The tool never changes client-filled columns and never overwrites a link with `Source = Entered by client` or `Locked = Yes`.
- A refused or challenge response is never passed to the extractor.
- User agent string: `PriceTrackerPOC/0.1 (price comparison tool)`. No browser impersonation, no header spoofing, no retries on a refusal. *(Amended 2026-11-16: when checking a site, the tool may try the other connection type, HTTP/2 or HTTP/1.1, once, and remembers which works; see `docs/superpowers/specs/2026-11-16-site-connection-design.md`.)*
- Per-site pacing: at least 5 seconds between requests to one host, or the site's robots `Crawl-delay` if larger, capped at 30. Different hosts run in parallel.
- Robots rules for user agent `*` are obeyed for every request.
- All tests run offline. No test may make a network request.
- User-facing text (console, issues file, user guide) is plain business English, no stack traces.

## Review Focus

1. **An input or output workbook is open in Excel during a run.** Expected: the run does not crash or lose results; the file is written beside the original as `<name> (new).xlsx` and the console says so. → test in Task 2.
2. **The run is interrupted (Ctrl+C) or a site hangs.** Expected: results so far are written to the run file, links found so far are saved, and a hung site times out without blocking the others. → test in Task 7.
3. **A product page with several prices** (variants, pack sizes, "was/now", a related-products block). Expected: never the first price by position; ambiguous means `Unverified` plus a Question. → tests in Task 5.
4. **A code that is a fragment of another code or of an unrelated number** (`ZVA2` in `ZVA25`, `70062` in a phone number or price). Expected: not `Verified`. → tests in Task 5.
5. **Cells Excel has mangled:** codes read as numbers (`305000` → `305000.0`), non-breaking spaces in names, trailing spaces in headers, `#N/A`, blank rows in the middle of a sheet. Expected: codes compare as text, blank rows are skipped silently, the rest become issues. → tests in Task 2.

---

### Task 1: Project skeleton, data model, link cleaning

**Files:**
- Create: `pyproject.toml`, `.gitignore`, `pricetracker/__init__.py`, `pricetracker/models.py`, `pricetracker/urls.py`
- Test: `tests/test_urls.py`

**Interfaces:**
- Produces (`pricetracker/models.py`):

```python
class Access(str, Enum):
    REACHABLE = "Reachable"; REFUSING = "Refusing"; UNREACHABLE = "Unreachable"
    NO_SEARCH = "Search not worked out"

class VatBasis(str, Enum):
    INC = "Inc VAT"; EX = "Ex VAT"; BOTH = "Both shown"; UNKNOWN = "Unknown"

class Label(str, Enum):
    CLIENT_LINK = "Client link"; VERIFIED = "Verified"; UNVERIFIED = "Unverified"
    MANUAL = "Manual price"; NOT_FOUND = "Not found"; BLOCKED = "Blocked"; ERROR = "Error"

CONFIRMED = {Label.CLIENT_LINK, Label.VERIFIED, Label.MANUAL}
SOURCE_TOOL, SOURCE_CLIENT = "Found by tool", "Entered by client"

@dataclass
class Item:
    code: str; name: str; brand: str; group: str
    client_price_ex: Decimal | None; ean: str; search_text: str; row: int

@dataclass
class Competitor:
    group: str; company: str; website: str
    search_status: Access | None; page_status: Access | None
    vat_basis: VatBasis; last_checked: datetime | None
    search_address: str   # tool column: URL template containing "{q}", or ""
    row: int

@dataclass
class ProductLink:
    item_code: str; company: str; url: str; source: str; locked: bool
    last_result: str; last_checked: datetime | None
    manual_price: Decimal | None; manual_vat: VatBasis | None; manual_date: date | None
    row: int

@dataclass
class PriceResult:
    run_at: datetime; item_code: str; name: str; group: str
    client_price_ex: Decimal | None; company: str
    price_ex: Decimal | None; price_inc: Decimal | None
    calculated: str      # "", "Ex VAT" or "Inc VAT": which figure was computed
    label: Label; url: str; remark: str

@dataclass
class Issue:
    kind: str            # "Error" or "Question"
    file: str; row: int | None; subject: str
    problem: str; client_value: str; seen_value: str; fix: str
```

- Produces (`pricetracker/urls.py`): `clean_link(url: str) -> str`, `host_of(url: str) -> str` (lowercase, no `www.`, `""` if not a URL), `same_site(a: str, b: str) -> bool`, `is_product_link(url: str) -> bool`.

`Competitors.xlsx` gets one tool column beyond the spec, `Search address`, so a worked-out search survives between runs.

- [ ] **Step 1: Set up the project**

`git init`; create a venv with `py -3.13 -m venv .venv`; `pyproject.toml` declares package `pricetracker`, runtime deps `httpx[http2]`, `selectolax`, `openpyxl`, dev deps `pytest`, `pyinstaller`, and `[tool.pytest.ini_options] testpaths = ["tests"]`. `.gitignore`: `.venv/`, `__pycache__/`, `build/`, `dist/`, `output/`, `*.spec`, `~$*.xlsx`. Install with `.venv\Scripts\pip install -e .[dev]`.

- [ ] **Step 2: Write the failing tests** (`tests/test_urls.py`)

```python
def test_clean_link_drops_tracking_keeps_variant():
    u = ("https://thetankshop.co.uk/products/harlequin-650?variant=37702586237120"
         "&currency=GBP&utm_source=google&utm_medium=cpc&gad_source=4&gclid=abc&srsltid=x")
    assert clean_link(u) == "https://thetankshop.co.uk/products/harlequin-650?variant=37702586237120"

def test_clean_link_keeps_attribute_selectors():
    u = "https://www.landownerproducts.co.uk/adblue-1000l?attr%5B146%5D=4&ps=MT"
    assert "attr%5B146%5D=4" in clean_link(u) and "ps=" not in clean_link(u)

def test_clean_link_drops_fragment_and_search_param():
    assert clean_link("https://a.co.uk/p/x?search=piusi%20k33#reviews") == "https://a.co.uk/p/x"

def test_host_of():
    assert host_of("https://www.FuelTankShop.co.uk/x") == "fueltankshop.co.uk"
    assert host_of("fueltankshop.co.uk") == "" and host_of("#N/A") == ""

def test_same_site_ignores_www_and_path():
    assert same_site("https://www.manutan.co.uk/en/key", "https://manutan.co.uk/p/1")

@pytest.mark.parametrize("u,ok", [
    ("https://www.chemfastuk.com/products/air1-adblue", True),
    ("https://www.chemfastuk.com/collections/fuelbox", False),
    ("https://www.plumbnation.co.uk/oil-tanks/", False),
    ("https://www.fueltankshop.co.uk/", False),
    ("petroy.com", False), ("#N/A", False), ("", False),
])
def test_is_product_link(u, ok):
    assert is_product_link(u) is ok
```

- [ ] **Step 3: Run** `.venv\Scripts\pytest tests/test_urls.py -v` — expect import failures.

- [ ] **Step 4: Implement `models.py` and `urls.py`**

`clean_link` keeps only query parameters named `variant`, `id`, `product_id`, `sku`, `p`, or starting with `attr`; drops everything else and the fragment. `is_product_link` is false for non-URLs, an empty or root path, and paths containing a `/collections/`, `/category/` or `/c/` segment; `/oil-tanks/`-style single-segment paths ending in `/` are false.

- [ ] **Step 5: Run the tests** — expect PASS.

- [ ] **Step 6: Commit** — `git add -A && git commit -m "feat: project skeleton, data model, link cleaning"`

---

### Task 2: Excel input, validation, issues file

**Files:**
- Create: `pricetracker/excel_io.py`, `pricetracker/issues.py`, `pricetracker/validate.py`
- Test: `tests/test_excel_io.py`, `tests/test_validate.py`, `tests/conftest.py`

**Interfaces:**
- Consumes: Task 1 models, `urls.host_of`, `urls.is_product_link`, `urls.clean_link`.
- Produces:

```python
# issues.py
class IssueLog:
    def add(self, kind, file, row, subject, problem, client_value="", seen_value="", fix="") -> None
    def __len__(self) -> int
    def write(self, path: Path) -> Path       # returns the path actually written
ISSUE_HEADERS = ["Kind", "Input file", "Row", "Item or competitor", "What is wrong",
                 "Value in the client's data", "Value seen", "What to change"]

# excel_io.py
def read_items(path: Path) -> list[Item]
def read_competitors(path: Path) -> list[Competitor]
def read_links(path: Path) -> list[ProductLink]          # missing file -> []
def write_competitors(path: Path, rows: list[Competitor]) -> Path
def write_links(path: Path, rows: list[ProductLink]) -> Path
def save_workbook(wb, path: Path) -> Path                # PermissionError -> "<stem> (new).xlsx"
def write_table(path: Path, sheet: str, headers: list[str], rows: list[list]) -> Path
class InputError(Exception)                              # file or required header missing

# validate.py
def is_family_code(code: str) -> bool                    # ends with lowercase "x" or with "/"
def validate(items, competitors, links, log: IssueLog) -> tuple[list[Item], list[Competitor], list[ProductLink]]
```

- `tests/conftest.py` provides `make_xlsx(tmp_path, name, headers, rows) -> Path`.
- Readers return every non-blank row untouched except for cell cleaning: text is stripped, `\xa0` becomes a space, a numeric cell in a code column becomes text without `.0`, `#N/A` and `NULL` become `""`. Readers do not judge; `validate` does.
- `write_competitors` and `write_links` load the existing workbook and update cells in place by header, so client formatting, extra columns and column order survive. New link rows are appended.

- [ ] **Step 1: Write the failing tests**

```python
# tests/test_excel_io.py
def test_reads_items_with_mangled_cells(tmp_path, make_xlsx):
    p = make_xlsx(tmp_path, "Items.xlsx",
        ["Item code (MPN) ", "name", "Brand", "Competitor group", "Client price ex VAT", "EAN", "Search text"],
        [[305000, "PIUSI E80 Pump,\xa0 230v ", "PIUSI", "Piusi Equipment", 304, 8050031000670, None],
         [None, None, None, None, None, None, None],
         ["ZVA25", "Elaflex ZVA25", "Elaflex", "Piusi Equipment", None, "#N/A", ""]])
    a, b = read_items(p)
    assert a.code == "305000" and a.name == "PIUSI E80 Pump, 230v" and a.client_price_ex == Decimal("304")
    assert a.ean == "8050031000670" and a.row == 2
    assert b.row == 4 and b.client_price_ex is None and b.ean == ""

def test_missing_required_header_names_the_file_and_column(tmp_path, make_xlsx):
    p = make_xlsx(tmp_path, "Items.xlsx", ["Name"], [["x"]])
    with pytest.raises(InputError, match=r"Items\.xlsx.*Item code \(MPN\)"):
        read_items(p)

def test_write_links_preserves_client_columns_and_appends(tmp_path, make_xlsx): ...
    # existing row keeps an extra "My notes" column value; one new row appended; Locked "Yes" read as True

def test_save_falls_back_when_file_is_open(tmp_path, monkeypatch):
    # monkeypatch Workbook.save to raise PermissionError for the original path only
    assert save_workbook(wb, tmp_path / "Product links.xlsx").name == "Product links (new).xlsx"

def test_empty_issue_log_writes_no_issues_found(tmp_path):
    p = IssueLog().write(tmp_path / "Data issues.xlsx")
    ws = openpyxl.load_workbook(p).active
    assert [c.value for c in ws[1]] == ISSUE_HEADERS and ws["A2"].value == "No issues found"
```

```python
# tests/test_validate.py — each asserts both the returned lists and the logged issue
def test_duplicate_item_code_keeps_first_and_logs_error()      # kind "Error", row of the 2nd
def test_missing_item_code_is_skipped_with_error()
def test_unknown_or_missing_group_skips_item_with_error()
def test_missing_client_price_is_kept_with_question()
def test_shared_ean_is_blanked_on_both_items_with_question()
def test_short_ean_is_blanked_with_question()                  # not 13 digits
def test_non_product_link_is_blanked_with_error()              # category page, bare host
def test_link_to_unlisted_site_is_dropped_with_error()
def test_link_for_unknown_item_or_company_is_dropped_with_error()
def test_duplicate_competitor_in_group_keeps_first_with_error()
def test_family_style_code_logs_question()                     # "966130x", "VAC/25/"
def test_manual_price_without_vat_basis_is_ignored_with_error()
def test_clean_inputs_log_nothing()
```

- [ ] **Step 2: Run** `.venv\Scripts\pytest tests/test_excel_io.py tests/test_validate.py -v` — expect FAIL.

- [ ] **Step 3: Implement the three modules**

Rules for `validate`, matching spec section 7:

| Condition | Action | Kind |
|---|---|---|
| Item code empty | drop item | Error |
| Item code repeated (case-insensitive) | keep first | Error |
| Item group empty or not in competitors | drop item | Error |
| Client price empty | keep | Question |
| EAN not 13 digits, or shared by two items | set `ean = ""` | Question |
| Code ends with lowercase `x` or with `/` | keep | Question ("looks like a product family") |
| Competitor website empty or not a URL | drop competitor | Error |
| Same company twice in one group | keep first | Error |
| Link not a product link | keep row, set `url = ""` | Error |
| Link host differs from the competitor's website host | drop row | Error |
| Link's item code or company not in the lists | drop row | Error |
| Manual price without VAT basis or date | ignore the manual price | Error |

Every `fix` text tells the client which file, row and column to change. Links are passed through `clean_link`.

- [ ] **Step 4: Run the tests** — expect PASS.

- [ ] **Step 5: Commit** — `git commit -am "feat: Excel input, validation and issues file"`

---

### Task 3: Fetcher — robots, pacing, response classification

**Files:**
- Create: `pricetracker/fetch.py`
- Test: `tests/test_fetch.py`

**Interfaces:**
- Produces:

```python
class Outcome(str, Enum):
    OK = "ok"; REFUSED = "refused"; NOT_FOUND = "not found"
    UNREACHABLE = "unreachable"; DISALLOWED = "disallowed by site rules"

@dataclass
class Page:
    outcome: Outcome; url: str; status: int; html: str   # url is the final URL after redirects

class Fetcher:
    def __init__(self, transport: httpx.BaseTransport | None = None,
                 clock=time.monotonic, sleep=time.sleep, min_delay: float = 5.0): ...
    def get(self, url: str) -> Page
    def allowed(self, url: str) -> bool            # robots rules for "*"
    def sitemaps(self, site_url: str) -> list[str] # from robots, else ["<root>/sitemap.xml"]
    def close(self) -> None
```

- `get` never raises. Timeout 25 s connect/read. Body capped at 3 MB. Redirects followed.
- Tests use `httpx.MockTransport` and fake `clock`/`sleep`.

- [ ] **Step 1: Write the failing tests**

```python
def test_403_is_refused_and_body_not_returned()            # page.html == ""
def test_200_with_just_a_moment_title_is_refused()
def test_200_page_that_merely_embeds_challenge_script_is_ok()   # normal page containing "challenge-platform"
def test_429_is_refused_and_404_410_are_not_found()
def test_timeout_and_connect_error_and_500_are_unreachable()
def test_robots_disallow_returns_disallowed_without_requesting_the_page()
def test_robots_fetched_once_per_host()
def test_missing_or_refused_robots_means_allowed()
def test_second_request_to_same_host_waits_min_delay()     # fake sleep called with ~5.0
def test_crawl_delay_20_overrides_min_delay_and_60_is_capped_at_30()
def test_different_hosts_do_not_wait_for_each_other()
def test_user_agent_header_is_the_project_string()
def test_sitemaps_from_robots_and_default()
```

- [ ] **Step 2: Run** `.venv\Scripts\pytest tests/test_fetch.py -v` — expect FAIL.

- [ ] **Step 3: Implement `Fetcher`**

Classification: `401/403/429` → REFUSED; `404/410` → NOT_FOUND; other `>=400`, timeouts and transport errors → UNREACHABLE; `200` whose `<title>` is "Just a moment..." or "Attention Required" → REFUSED; else OK. Robots parsing uses `urllib.robotparser.RobotFileParser.parse(lines)` on text fetched through the same client. Pacing is a per-host lock plus last-request timestamp, so it is safe from several threads.

- [ ] **Step 4: Run the tests** — expect PASS.

- [ ] **Step 5: Commit** — `git commit -am "feat: fetcher with robots rules, pacing and refusal detection"`

---

### Task 4: Page extraction and VAT

**Files:**
- Create: `pricetracker/extract.py`, `pricetracker/vat.py`
- Create: `tests/fixtures/pages/` (see Step 1)
- Test: `tests/test_extract.py`, `tests/test_vat.py`

**Interfaces:**
- Produces:

```python
# extract.py
@dataclass
class Offer:
    price: Decimal; currency: str; sku: str; mpn: str; gtin: str; name: str; url: str

@dataclass
class ProductPage:
    is_single_product: bool
    name: str; brand: str; sku: str; mpn: str; gtin: str
    offers: list[Offer]            # [] when no machine-readable price
    text: str                      # visible text, whitespace-collapsed
    labelled_ex: Decimal | None    # visible price labelled ex VAT, if exactly one distinct value
    labelled_inc: Decimal | None

def parse_product_page(html: str, url: str) -> ProductPage

# vat.py
def resolve_vat(price: Decimal, basis: VatBasis, labelled_ex: Decimal | None,
                labelled_inc: Decimal | None) -> tuple[Decimal | None, Decimal | None, str]
    # -> (price_ex, price_inc, calculated)
```

- [ ] **Step 1: Save page fixtures**

With the Fetcher from Task 3 in a throwaway script (not a test), save the HTML of six product pages from six different competitor sites into `tests/fixtures/pages/`, chosen to cover the different price markup patterns: JSON-LD with and without a part number, an offer list, Open Graph price tags only, and a page with no machine-readable price.

Open each saved file and write down, for the tests below, the price, codes and VAT wording actually present. If a page has changed or no longer loads, pick another product page from the same site. Add hand-written fixtures `microdata_only.html`, `meta_only.html`, `listing_page.html` (JSON-LD `ItemList`), `no_price.html`, `usd_price.html`.

- [ ] **Step 2: Write the failing tests**

```python
# tests/test_extract.py
def test_agriemach_jsonld_gives_price_and_mpn():
    p = parse_product_page(load("agriemach_k33.html"), URL)
    assert p.is_single_product and p.mpn == "553000"
    assert [o.price for o in p.offers] == [Decimal("146.14")] and p.offers[0].currency == "GBP"

def test_hytek_gives_sku_without_mpn()                 # sku "K33", mpn ""
def test_multi_offer_page_returns_every_offer()        # thetankshop / fueltankshop: len(offers) as recorded
def test_jsonld_graph_and_list_wrappers_are_unwrapped()
def test_aggregate_offer_yields_low_and_high_as_two_offers()
def test_price_strings_with_pound_sign_and_commas_parse()   # "£1,390.00" -> Decimal("1390.00")
def test_microdata_fallback()
def test_meta_tag_fallback()                           # product:price:amount / og:price:amount
def test_listing_page_is_not_single_product()
def test_page_without_price_has_no_offers()
def test_malformed_jsonld_is_skipped_not_raised()
def test_labelled_prices_found_in_visible_text():
    # "£146.14 Inc. VAT  £121.78 Ex. VAT" -> labelled_inc 146.14, labelled_ex 121.78
def test_two_different_ex_vat_values_give_no_labelled_price()

# tests/test_vat.py
def test_inc_basis_computes_ex():      assert resolve_vat(D("120.00"), VatBasis.INC, None, None) == (D("100.00"), D("120.00"), "Ex VAT")
def test_ex_basis_computes_inc():      assert resolve_vat(D("100.00"), VatBasis.EX, None, None) == (D("100.00"), D("120.00"), "Inc VAT")
def test_labelled_pair_wins_and_nothing_is_calculated():
    assert resolve_vat(D("146.14"), VatBasis.UNKNOWN, D("121.78"), D("146.14")) == (D("121.78"), D("146.14"), "")
def test_unknown_basis_without_labels_returns_no_figures():
    assert resolve_vat(D("50.00"), VatBasis.UNKNOWN, None, None) == (None, None, "")
def test_labelled_pair_that_is_not_20_percent_apart_is_ignored()
def test_both_shown_basis_without_a_readable_pair_is_treated_as_unknown()
```

- [ ] **Step 3: Run** `.venv\Scripts\pytest tests/test_extract.py tests/test_vat.py -v` — expect FAIL.

- [ ] **Step 4: Implement**

Extraction order: JSON-LD `Product` (through `@graph`, arrays, `mainEntity`; `offers` as object, list, or `AggregateOffer`; `hasVariant` products each contribute their offers) → microdata `itemprop` → `product:price:amount` / `og:price:amount` meta. `is_single_product` is true when exactly one top-level `Product` is found, or `og:type` is `product`. A labelled pair counts only when the structured price equals one of the two and `inc` is within 1p of `ex × 1.2`. `gtin` takes the first of `gtin13`, `gtin`, `gtin14`, `gtin12`, `gtin8`.

- [ ] **Step 5: Run the tests** — expect PASS.

- [ ] **Step 6: Commit** — `git commit -am "feat: product page extraction and VAT resolution"`

---

### Task 5: Matching and labelling

**Files:**
- Create: `pricetracker/match.py`
- Test: `tests/test_match.py`

**Interfaces:**
- Consumes: `ProductPage`, `Offer`, `Item`, `Label`, `validate.is_family_code`.
- Produces:

```python
def norm_code(code: str) -> str                       # uppercase, letters and digits only
# is_family_code comes from validate.py (Task 2)
def code_in_text(code: str, text: str) -> bool        # complete code, not a fragment
def name_overlap(item_name: str, page_name: str) -> float   # 0..1 share of significant item tokens

@dataclass
class MatchResult:
    label: Label                    # VERIFIED, UNVERIFIED or NOT_FOUND
    offer: Offer | None
    remark: str
    observed: list[tuple[str, str, str]]   # (field "MPN"/"EAN", client value, value seen on page)

def match_page(item: Item, page: ProductPage, page_url: str, trusted: bool) -> MatchResult
```

`trusted=True` is a client-supplied link: identity is not checked, but the offer still has to be chosen and `observed` is still filled.

- [ ] **Step 1: Write the failing tests**

```python
@pytest.mark.parametrize("code,text,found", [
    ("ZVA2", "Elaflex ZVA25 nozzle", False),
    ("ZVA25", "Elaflex ZVA 25 Automatic Nozzle", True),      # separators inside a code are flexible
    ("F00373000", "Part no: F00373000.", True),
    ("70062", "Call 01234 570062", False),
    ("70062", "Cim-Tek 70062-10 filter", True),
    ("2500FP/230", "Harlequin 2500FP-230 fuel point", True),
    ("2500FP/230", "Harlequin 2500FP 230v", False),          # "230v" is a longer token: name match only
    ("305000", "Price £305000", False),                      # preceded by a currency sign
])
def test_code_in_text(code, text, found): ...

def test_structured_mpn_equal_to_item_code_is_verified()
def test_structured_sku_equal_to_item_code_is_verified()
def test_structured_gtin_equal_to_item_ean_is_verified()
def test_code_only_in_page_text_of_single_product_page_is_verified()
def test_code_on_listing_page_is_not_verified()              # is_single_product False -> NOT_FOUND
def test_family_code_is_never_verified()                     # UNVERIFIED at best, remark says why
def test_name_only_match_is_unverified()                     # overlap >= 0.6, no code
def test_low_name_overlap_is_not_found()
def test_page_without_offers_is_not_found_with_remark_price_not_readable()

def test_single_offer_is_chosen()
def test_offer_whose_sku_matches_item_code_is_chosen_over_first()
def test_offer_matching_variant_id_in_url_is_chosen()
def test_several_offers_with_one_distinct_price_is_chosen()
def test_several_offers_different_prices_no_code_is_unverified_with_ambiguous_remark()
def test_trusted_link_with_ambiguous_offers_is_unverified_not_client_link()

def test_observed_reports_different_mpn_and_ean():
    # item code "57500C", ean "0089470123540"; page mpn "F0057500C", gtin "8050031001234"
    assert r.observed == [("MPN", "57500C", "F0057500C"), ("EAN", "0089470123540", "8050031001234")]
def test_observed_is_empty_when_page_states_no_codes()
```

- [ ] **Step 2: Run** `.venv\Scripts\pytest tests/test_match.py -v` — expect FAIL.

- [ ] **Step 3: Implement `match.py`**

`code_in_text`: split the code at its separators and at every letter/digit boundary (`ZVA25` → `ZVA`, `25`), join the pieces with `[\s\-/._]*`, with `(?<![A-Za-z0-9£$€.,])` before and `(?![A-Za-z0-9])` after, case-insensitive. `name_overlap` ignores tokens shorter than 3 characters and the words `with`, `for`, `and`, `kit`, `c/w`. Identity order: structured code equality → code in `name` → code in `text` (single-product pages only) → name overlap. Offer choice order as the test names state. `match_page` returns `Label.VERIFIED` for a trusted link with a clear offer; the caller turns that into `CLIENT_LINK`.

- [ ] **Step 4: Run the tests** — expect PASS.

- [ ] **Step 5: Commit** — `git commit -am "feat: product matching, offer choice and observed codes"`

---

### Task 6: Site check and product discovery

**Files:**
- Create: `pricetracker/discover.py`
- Create: `tests/fixtures/search/` (hand-written small HTML and XML files)
- Test: `tests/test_discover.py`

**Interfaces:**
- Consumes: `Fetcher`, `Page`, `Outcome`, `Competitor`, `Item`, `Access`, `urls.*`, `match.norm_code`, `match.name_overlap`.
- Produces:

```python
def check_site(fetcher: Fetcher, comp: Competitor, probe_term: str) -> None
    # fills comp.search_status and comp.search_address in place

def find_search_address(home_html: str, home_url: str) -> str      # "" if none; contains "{q}"
def search_terms(item: Item) -> list[str]
def candidates_from_search(html: str, page_url: str, item: Item, limit: int = 3) -> list[str]
def candidates_from_sitemap(fetcher: Fetcher, comp: Competitor, item: Item,
                            cache: dict[str, list[str]], limit: int = 3) -> list[str]
def find_candidates(fetcher: Fetcher, comp: Competitor, item: Item,
                    cache: dict[str, list[str]]) -> list[str]
```

- `comp.page_status` is not set here: Task 7 sets it from the first product page fetched on that site.

- [ ] **Step 1: Write the failing tests**

```python
def test_search_form_with_q_input()            # <form action="/shop/search"><input name="q"> -> "https://h/shop/search?q={q}"
def test_wordpress_search_form()               # name="s" plus hidden post_type=product kept
def test_form_with_post_method_is_ignored()
def test_login_and_newsletter_forms_are_ignored()
def test_no_form_returns_empty()

def test_check_site_home_refused_sets_refusing()
def test_check_site_home_unreachable_sets_unreachable()
def test_check_site_no_form_sets_search_not_worked_out()
def test_check_site_search_disallowed_by_robots_sets_refusing_and_keeps_no_address()
def test_check_site_working_search_sets_reachable_and_address()
def test_check_site_keeps_existing_address_and_only_retests_it()

def test_search_terms_order():
    # item code "F00373000", brand "PIUSI", name "PIUSI EX50 Fuel Transfer Pump, 230v", search_text ""
    assert search_terms(item) == ["F00373000", "PIUSI F00373000", "PIUSI EX50 Fuel Transfer Pump, 230v"]
def test_search_text_column_goes_first()
def test_family_code_is_not_used_as_a_search_term()

def test_candidates_rank_code_in_link_above_name_only()
def test_candidates_exclude_other_hosts_category_links_and_duplicates()
def test_candidates_capped_at_three()

def test_sitemap_index_is_followed_and_urls_ranked()
def test_sitemap_is_read_once_per_site_per_run()          # cache
def test_sitemap_limits_stop_runaway_sites()              # at most 20 files and 100000 URLs
def test_gzipped_sitemap_is_read()

def test_find_candidates_uses_search_when_reachable_else_sitemap()
def test_find_candidates_stops_at_first_term_with_results()
```

- [ ] **Step 2: Run** `.venv\Scripts\pytest tests/test_discover.py -v` — expect FAIL.

- [ ] **Step 3: Implement `discover.py`**

A search form is a GET `<form>` whose action or class/id/role mentions `search`, or that has an `<input type="search">` or an input named one of `q`, `s`, `search`, `query`, `keywords`, `term`, `search_query`. `check_site` fetches the home page, finds or re-uses the address, checks `fetcher.allowed()` on it, then requests it once with `probe_term` and requires `Outcome.OK`. Link scoring: `+10` if `norm_code(item.code)` is in the normalised href or link text, plus `name_overlap(item.name, link_text + slug words)`; links scoring below `0.5` are dropped. Sitemap URLs are scored the same way on their slug.

- [ ] **Step 4: Run the tests** — expect PASS.

- [ ] **Step 5: Commit** — `git commit -am "feat: site check, search discovery and sitemap fallback"`

---

### Task 7: Scan — orchestration, progress, run file, write-back

**Files:**
- Create: `pricetracker/scan.py`, `pricetracker/runfile.py`, `pricetracker/progress.py`
- Test: `tests/test_scan.py`, `tests/test_runfile.py`

**Interfaces:**
- Consumes: everything from Tasks 1–6.
- Produces:

```python
# runfile.py
RUN_HEADERS = ["Run date and time", "Item code", "Name", "Competitor group", "Client price ex VAT",
               "Company", "Price ex VAT", "Price inc VAT", "Calculated", "Match label",
               "Product page link", "Remark"]
def run_file_name(run_at: datetime) -> str            # "Prices 2026-11-20 0930.xlsx"
def write_run(path: Path, results: list[PriceResult]) -> Path
def read_run(path: Path) -> list[PriceResult]

# progress.py
class Progress:
    def __init__(self, total: int, out=sys.stdout): ...
    def line(self, item: Item, comp: Competitor, result: PriceResult) -> None
    def summary(self, results: list[PriceResult], issues: int, issues_path: Path) -> None

# scan.py
def lookup(fetcher, item: Item, comp: Competitor, link: ProductLink | None,
           cache: dict, log: IssueLog, run_at: datetime) -> tuple[PriceResult, ProductLink]
def run_scan(input_dir: Path, output_dir: Path, fetcher: Fetcher | None = None,
             now: datetime | None = None, out=sys.stdout) -> Path     # returns the run file path
```

- [ ] **Step 1: Write the failing tests**

`tests/test_scan.py` drives `lookup` and `run_scan` with a `Fetcher` on `httpx.MockTransport`, `min_delay=0`, serving the Task 4 fixtures.

```python
# lookup — one test per row of the decision table below
def test_client_link_ok_gives_client_link_label()
def test_tool_link_ok_and_code_matches_gives_verified()
def test_tool_link_404_triggers_new_search_and_replaces_link()
def test_client_link_404_gives_error_keeps_link_and_logs_error()
def test_locked_tool_link_404_is_not_replaced()
def test_link_refused_with_manual_price_gives_manual_label_and_date_in_remark()
def test_link_refused_without_manual_price_gives_blocked()
def test_no_link_search_finds_verified_page_and_saves_tool_link()
def test_no_link_only_name_match_gives_unverified_and_saves_link()
def test_no_link_nothing_found_gives_not_found_and_adds_empty_link_row()
def test_site_unreachable_gives_error_with_remark_site_did_not_respond()
def test_unknown_vat_basis_gives_price_in_remark_no_figures_and_question()
def test_non_gbp_price_gives_error_with_currency_remark()
def test_observed_codes_become_questions_with_seen_value()
def test_page_status_is_set_from_first_product_page_outcome()

# run_scan
def test_run_scan_writes_run_file_issues_file_and_updates_inputs(tmp_path)
    # 2 items x 2 competitors; asserts 4 rows, file name from `now`, statuses and links written back
def test_same_site_in_two_groups_is_checked_once()
def test_item_is_only_checked_against_its_groups_competitors()
def test_sites_with_status_already_set_are_not_rechecked_but_empty_ones_are()
def test_ctrl_c_still_writes_partial_run_file_and_links(tmp_path)
    # transport raises KeyboardInterrupt on the 3rd request; run file holds the first results
def test_hung_site_does_not_stop_other_sites()
def test_missing_input_file_prints_plain_message_and_exits_nonzero()
def test_progress_shows_count_item_company_and_result():
    # a line like " 3/4  F00373000  Agriemach  Verified  £121.78 ex VAT"
def test_summary_counts_and_names_the_issues_file()

# tests/test_runfile.py
def test_run_file_name()
def test_write_then_read_round_trips_all_fields()
def test_read_ignores_excel_lock_files()                # "~$Prices ....xlsx"
```

- [ ] **Step 2: Run** `.venv\Scripts\pytest tests/test_scan.py tests/test_runfile.py -v` — expect FAIL.

- [ ] **Step 3: Implement**

`lookup` decision table:

| Link row | Fetch outcome | Result |
|---|---|---|
| has URL | OK | `match_page(trusted = source is client)` → `CLIENT_LINK` / `VERIFIED` / `UNVERIFIED` |
| has URL | REFUSED or DISALLOWED | `MANUAL` if a manual price is present, else `BLOCKED` |
| has URL, tool-found, not locked | NOT_FOUND | discover again; replace the URL if a page is found |
| has URL, client or locked | NOT_FOUND | `ERROR`, remark "Link no longer works", Error issue |
| has URL | UNREACHABLE | `ERROR`, remark "Site did not respond" |
| no URL | – | `MANUAL` if a manual price is present; else `find_candidates`, fetch each, first `VERIFIED` wins, else best `UNVERIFIED`, else `NOT_FOUND` |

After a match with an offer, `resolve_vat(offer.price, comp.vat_basis, page.labelled_ex, page.labelled_inc)`. A manual price goes through `resolve_vat` with its own basis.

`run_scan`: read and validate → `check_site` once per distinct host with empty `search_status` → build the item × competitor pairs → one worker thread per host, each processing its pairs in order → collect results. A `try/finally` writes the run file, `Data issues.xlsx`, `Competitors.xlsx` and `Product links.xlsx` whatever happens. Results are sorted by item code then company before writing.

- [ ] **Step 4: Run the tests** — expect PASS. Then run the whole suite: `.venv\Scripts\pytest -q`.

- [ ] **Step 5: Commit** — `git commit -am "feat: scan command with progress, run file and write-back"`

---

### Task 8: Report workbook

**Files:**
- Create: `pricetracker/report.py`
- Test: `tests/test_report.py`

**Interfaces:**
- Consumes: `runfile.read_run`, `PriceResult`, `CONFIRMED`, `excel_io.save_workbook`.
- Produces: `build_report(output_dir: Path) -> Path` (writes `Price comparison.xlsx`).

Sheets, in this order:

| Sheet | Content |
|---|---|
| `Latest comparison` | One row per item: Item code, Name, Client price ex VAT, one column per company (latest confirmed ex-VAT price), Cheapest competitor, Cheapest price, Difference £, Difference %. |
| `Unverified matches` | Latest `Unverified` rows: item, company, prices, link, remark. |
| `History` | Every row of every run file, with the run headers, as an Excel table with filters. |
| `Graphs` | One line chart per item: x = run date, one series for the client price and one per company with a confirmed price. |
| `Graph data` | The pivoted numbers the charts read. |
| `Coverage` | Per company, counts of the latest run's results by label. |

- [ ] **Step 1: Write the failing tests**

Each test writes two or three small run files with `write_run` into `tmp_path`, calls `build_report`, and reopens the result with `openpyxl`.

```python
def test_latest_uses_newest_run_per_item_and_company()
def test_unverified_blocked_and_not_found_never_appear_as_prices_in_latest()
def test_cheapest_and_difference():
    # client 100.00; A 90.00, B 110.00 -> cheapest "A", 90.00, difference -10.00 and -10.0%
def test_client_more_expensive_than_cheapest_is_filled_red_and_cheaper_green()
def test_item_without_client_price_has_empty_difference()
def test_history_has_one_row_per_result_across_runs()
def test_one_chart_per_item_with_confirmed_prices()       # len(ws._charts)
def test_chart_series_include_client_price_and_each_company()
def test_single_run_still_builds_charts()                 # one point, no error
def test_coverage_counts_by_label_for_latest_run()
def test_unverified_matches_sheet_lists_them()
def test_no_run_files_prints_plain_message_and_raises_input_error()
def test_report_file_open_in_excel_falls_back_to_new_name()
```

- [ ] **Step 2: Run** `.venv\Scripts\pytest tests/test_report.py -v` — expect FAIL.

- [ ] **Step 3: Implement `report.py`**

Colours: more expensive than the cheapest competitor → fill `FFC7CE`; cheaper or equal → `C6EFCE`. Difference = cheapest competitor − client price. Charts are `openpyxl.chart.LineChart`, 16 cm × 8 cm, two per row on `Graphs`, title "`<code>` – `<name>`", y-axis "£ ex VAT", markers on so a single run shows a point. `Graph data` holds one block per item: first column run date, then client price, then one column per company; gaps stay empty.

- [ ] **Step 4: Run the tests** — expect PASS.

- [ ] **Step 5: Commit** — `git commit -am "feat: comparison workbook with history, graphs and coverage"`

---

### Task 9: Clean input files from the client's originals

> The conversion script and its tests are not in this repository: they only work on the client's original files, which are private. Its result is published in `sample-input/`.

**Files:**
- Create: `tools/convert_inputs.py`, `tools/vat_basis.csv`
- Create (generated, committed): `input/Items.xlsx`, `input/Competitors.xlsx`, `input/Product links.xlsx`, `input/Input data findings.xlsx`
- Test: `tests/test_convert_inputs.py`

**Interfaces:**
- Consumes: `excel_io.write_table`, `urls.clean_link`, `urls.host_of`, `urls.is_product_link`, `issues.IssueLog`.
- Produces: `convert(src_dir: Path, dst_dir: Path, vat_csv: Path) -> None`, run as `python tools/convert_inputs.py "example files" input`.

Source layout (verified against the files):

| Client file | Sheet | Columns |
|---|---|---|
| Sample item list | `Sheet2`, header row 2 | B Name, D MPN, E EAN |
| Price list | `Price List`, header row 2 | E Product, G Price (ex VAT), K CODE |
| Price list | `Price List (2)`, header row 2 | L CODE, R to AE competitor links |
| Competitor list | first sheet, header row 1 | B group (first row of block only), C company, D webpage |

- [ ] **Step 1: Write the failing tests** — run `convert` on the real `example files/` into `tmp_path`.

```python
def test_items_are_47_unique_codes()
def test_brand_comes_from_the_product_name():
    # "57500C" -> "PIUSI"; "110700960" -> "Tecalemit"; "11050070X" -> "Tecalemit"; "2500FP/230" -> "Harlequin"
def test_groups(): # Harlequin items -> "Tanks"; every other item -> "Piusi Equipment"
def test_client_price_joined_by_code():   # "57500C" -> 529; "F00373000" -> empty
def test_shared_and_short_eans_are_left_empty()
def test_competitors_have_41_rows_three_groups_no_blanks()    # 16 + 11 + 14
def test_competitor_vat_basis_comes_from_csv_else_unknown()
def test_links_only_for_sample_items_and_listed_competitors_and_are_cleaned()
def test_links_are_marked_entered_by_client()
def test_findings_file_lists_duplicates_shared_eans_missing_prices_and_brand_mismatches()
def test_generated_inputs_pass_validate_with_no_errors():
    # read the three generated files and run validate(): zero issues of kind "Error"
```

- [ ] **Step 2: Run** `.venv\Scripts\pytest tests/test_convert_inputs.py -v` — expect FAIL.

- [ ] **Step 3: Implement `convert_inputs.py`**

Brand is the first word of the name matched case-insensitively against `PIUSI, Tecalemit, Elaflex, Cim-Tek, Harlequin, Snaplock, Vacupress`; `Horn Tecalemit …` → `Tecalemit`. The client's Brand column and `Optimized_Scrape_String` are not used. `Input data findings.xlsx` uses `ISSUE_HEADERS` and lists every problem found in the original files, one row per affected item or row.

- [ ] **Step 4: Fill in `tools/vat_basis.csv` by eye**

Columns `host,vat_basis,checked_on,note`. For each of the 33 distinct competitor hosts, open one product page in a browser, compare the structured price (view source, `"price"`) with the visible price and its VAT wording, and record `Inc VAT`, `Ex VAT` or `Both shown`. Hosts that cannot be opened stay `Unknown` with a note. Do not guess.

- [ ] **Step 5: Run the tests** — expect PASS. Run the converter and open the four generated files to check they read cleanly.

- [ ] **Step 6: Commit** — `git add -A && git commit -m "feat: clean input files and input data findings"`

---

### Task 10: Command line, exe, user guide, live run

**Files:**
- Create: `pricetracker/cli.py`, `pricetracker/__main__.py`, `build.ps1`, `docs/user-guide.md`
- Test: `tests/test_cli.py`

**Interfaces:**
- Consumes: `scan.run_scan`, `report.build_report`, `excel_io.InputError`.
- Produces: `main(argv: list[str] | None = None) -> int`; commands `pricetracker scan` and `pricetracker report`, each with optional `--input <dir>` and `--output <dir>` (defaults `input` and `output` beside the exe, or the current directory when not frozen).

- [ ] **Step 1: Write the failing tests**

```python
def test_no_command_prints_usage_with_both_commands_and_returns_2()
def test_scan_calls_run_scan_with_default_folders(monkeypatch, tmp_path)
def test_report_calls_build_report(monkeypatch, tmp_path)
def test_input_error_prints_message_without_traceback_and_returns_1(capsys)
def test_unexpected_error_prints_one_line_and_writes_error_log(tmp_path, capsys)   # output/error.log
def test_output_folder_is_created_if_missing()
```

- [ ] **Step 2: Run** `.venv\Scripts\pytest tests/test_cli.py -v` — expect FAIL.

- [ ] **Step 3: Implement `cli.py` and `__main__.py`** with `argparse`.

- [ ] **Step 4: Run the whole suite** — `.venv\Scripts\pytest -q` — expect all PASS.

- [ ] **Step 5: Build the exe**

`build.ps1` runs `.venv\Scripts\pyinstaller --onefile --name pricetracker --console pricetracker\__main__.py`, then copies `input\` and the Word user guide into `dist\`. Verify: `dist\pricetracker.exe` with no arguments prints the usage text.

- [ ] **Step 6: Live smoke run**

From `dist\`: `pricetracker.exe scan`. This is the first real run against the competitor sites and takes roughly 15–40 minutes. Then `pricetracker.exe report`. Check and write down:

- the run file has one row for each item × competitor of its group,
- `Competitors.xlsx` now has both statuses for every site,
- five `Verified` and five `Client link` prices, picked at random, equal the price on the web page when opened by hand, on the right VAT basis,
- no row labelled `Verified` points at a different product,
- `Data issues.xlsx` reads sensibly to a non-technical person,
- `Price comparison.xlsx` opens in Excel without a repair prompt and the graphs show.

Record the counts per label per site. Any wrong price is a bug: fix it with a new fixture and test in the owning task's test file before continuing. Do not tune the tool to get past a site that refuses.

- [ ] **Step 7: Write `docs/user-guide.md`**

Plain English, for the pricing manager. Sections: what the tool does; first start (including the Windows "unknown publisher" prompt); running `scan` and `report`; what each input column means and who fills it; adding an item; adding a competitor and what happens on the next run; supplying a product link and locking it; entering a price by hand; reading the match labels; reading `Data issues.xlsx` and how to fix each kind of issue listed in spec section 7; why a file may appear as `(new).xlsx`. Use real screens and file names from the smoke run.

- [ ] **Step 8: Commit** — `git add -A && git commit -m "feat: command line, exe build and user guide"`

---

## After the plan

- Run `scan` on at least three different days before the demonstration so the graphs show lines.
- Report the smoke-run coverage per site to the client together with spec section 11's open questions.
