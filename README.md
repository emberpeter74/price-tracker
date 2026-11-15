# Price Tracker

A command-line tool that reads competitor prices from their websites and compares them with your own prices in Excel.

You give it three Excel files: your items, your competitors, and (optionally) links to product pages. It finds each item on each competitor's site, reads the price, and writes the results back to Excel: one dated file per scan, plus a comparison workbook with price-history graphs.

This is a proof of concept built for a real client in the fuel-handling equipment trade. The sample data in [`sample-input/`](sample-input/) is that client's item and competitor list, published with permission. It was produced from the client's original spreadsheets by a one-off conversion that is not part of this repository.

## What it does

- **Finds product pages itself.** It searches each site by manufacturer part number (MPN), then EAN, then name. Where a site's rules forbid automated searching, it reads the site's sitemap.
- **Remembers what it found.** Product page links are saved to `Product links.xlsx`, so later scans go straight to the page. The user can correct or lock any link.
- **Never guesses a match.** Every price carries a label: `Verified` (the item code is on the page), `Client link`, `Manual price`, or `Unverified` (matched by name only, kept out of the comparison).
- **Handles VAT.** It records prices ex and inc VAT, reading both from the page where shown, and marks any figure it had to calculate.
- **Reports data problems.** Duplicate codes, unusable links, mismatching product codes and similar go to `Data issues.xlsx`, each with the row and what to change. The scan carries on with the data that is usable.
- **Is a polite visitor.** It obeys `robots.txt`, identifies itself honestly, waits between requests to the same site, and does not try to get past sites that refuse automated access. For those, the user can enter a price by hand.

## Quick start

Requires Python 3.11 or newer.

```bash
git clone https://github.com/YOUR-USERNAME/price-tracker.git
cd price-tracker
python -m venv .venv
.venv\Scripts\activate          # Linux/macOS: source .venv/bin/activate
pip install -e .[dev]

# work on a copy of the sample data
xcopy sample-input input\ /E    # Linux/macOS: cp -r sample-input input

python -m pricetracker scan     # writes output\Prices <date> <time>.xlsx and output\Data issues.xlsx
python -m pricetracker report   # writes output\Price comparison.xlsx
```

A first scan of the sample data (47 items, 33 competitor sites) takes 30 to 60 minutes, because of the pauses between requests. Press Ctrl+C to stop early; results so far are saved.

To build a single Windows executable that needs no Python:

```powershell
.\build.ps1          # creates dist\pricetracker.exe
```

## Input and output

| File | Purpose |
|---|---|
| `input/Items.xlsx` | Your items: code (MPN), name, brand, competitor group, your price ex VAT, EAN. |
| `input/Competitors.xlsx` | Competitors per group, with their website. The tool fills in each site's access status. |
| `input/Product links.xlsx` | One row per item per competitor: the product page, and optionally a hand-entered price. |
| `output/Prices <date> <time>.xlsx` | Every price found in one scan, with match label and remarks. |
| `output/Data issues.xlsx` | Input data errors and questions. An empty file is the goal. |
| `output/Price comparison.xlsx` | Latest comparison, unverified matches, full history, graphs, and coverage per competitor. |

The [user guide](docs/user-guide.md) explains every column and how to correct each kind of issue.

## How it works

```
Items.xlsx ─┐
Competitors ─┼─► validate ─► check sites ─► for each item × competitor ─► run file
Product links┘      │                         │                            issues file
                    ▼                         ▼                            links + statuses
               Data issues          link ► fetch ► extract ► match ► VAT   written back
                                     └─ none? search / sitemap ─┘
```

| Module | Responsibility |
|---|---|
| [`excel_io.py`](pricetracker/excel_io.py) | Reads and writes the Excel files; tolerant of what Excel does to codes and numbers. |
| [`validate.py`](pricetracker/validate.py) | Checks the input, keeps what is usable, reports the rest. |
| [`fetch.py`](pricetracker/fetch.py) | The only module that touches the network: `robots.txt`, per-site pacing, refusal detection. |
| [`discover.py`](pricetracker/discover.py) | Works out a site's search form; finds candidate product pages by search or sitemap. |
| [`extract.py`](pricetracker/extract.py) | Reads structured product data (JSON-LD, microdata, Open Graph) from a page. |
| [`match.py`](pricetracker/match.py) | Decides whether a page is the item, which of several prices applies, and how sure we are. |
| [`vat.py`](pricetracker/vat.py) | Works out ex-VAT and inc-VAT figures. |
| [`scan.py`](pricetracker/scan.py) | Orchestrates a scan: one worker per site, progress display, interrupt-safe saving. |
| [`report.py`](pricetracker/report.py) | Builds the comparison workbook and charts from all run files. |
| [`cli.py`](pricetracker/cli.py) | The `scan` and `report` commands. |

Design decisions worth knowing about:

- **One general reader, not a scraper per site.** The tool relies on the structured data shops publish for search engines. This keeps it small and maintainable, at the cost of sites that publish no machine-readable price. Those are reported as such, not guessed.
- **A wrong price is worse than no price.** A code found only in a page's body does not verify a match, because product pages list other products' codes under "related products". Name-only matches must agree on model numbers and voltage, and are still only `Unverified`.
- **VAT basis from arithmetic, not labels.** Shops label prices inconsistently ("£146.14 Inc. VAT: £175.37" can mean the first figure is ex VAT). If the page shows a second figure that is exactly the price × 1.2 or ÷ 1.2, the tool takes the pair as read.
- **The user's data is never changed.** Client-entered links, locks and prices are read-only for the tool; it only writes its own columns.

## Tests

```bash
pytest
```

The suite runs offline: network access is replaced by a mock transport, and page fixtures are small hand-made files that reproduce the price markup patterns found on real shop pages.

## Project documents

- [Business specification](docs/business-specification.md): what was agreed with the client and why.
- [Implementation plan](docs/implementation-plan.md): the task-by-task build plan.
- [User guide](docs/user-guide.md): for the person who runs the tool.

## Limitations

- Sites that publish no machine-readable price, or that block automated visitors, need hand-entered prices.
- Delivery costs, stock levels, quantity discounts and promotions are not compared.
- GBP and UK VAT (20%) only.
- No scheduling or alerts; the user starts each scan.

## Responsible use

The tool follows each site's `robots.txt` and crawl delay and does not disguise itself as a browser. Before pointing it at a website, check that site's terms of use. You are responsible for how you use it.

## Licence

[MIT](LICENSE)
