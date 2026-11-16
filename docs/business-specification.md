# Price Tracker – Proof of Concept

**Business specification**
November 2026

---

## 1. Purpose

The client sells fuel-handling equipment (pumps, tanks, nozzles, filters, AdBlue) and wants to know how its prices compare with competitors' web prices.

This proof of concept (POC) demonstrates that we understood the task and can deliver it. It is a simplified version. Many decisions are deliberately left until the client has seen real results.

**Who uses it:** a pricing manager, about once a week, to decide where to adjust prices.

**What success looks like:** the client opens the result, recognises their products, sees real competitor prices they can verify on the web, and sees clearly where data is missing and why.

## 2. Scope

### In the POC

1. Read items, competitors and product links from Excel files.
2. Find each item on each relevant competitor website, and read its price.
3. Save each run as a dated Excel file, including the client's own price at that time.
4. Merge all run files into one Excel workbook with price comparison and price-history graphs.
5. Keep a list of websites with their access status.
6. Keep a list of product page links that later runs reuse, and that the client can edit.
7. Let the client enter a price by hand where the tool cannot read a site.
8. Write input data problems and questions to an issues file, and continue with the data that is usable.
9. Show progress while a run is in flight.
10. Deliver clean, easy-to-edit input files built from the client's current files.
11. Deliver a short user guide: how to run the tool and how to correct or add data.

### Not in the POC

- Automatic scheduling, alerts, e-mail, a web dashboard, user accounts.
- All 600+ items. The POC covers the agreed sample of about 50.
- Automatic correction of the client's data. The tool reports; the client decides.
- Getting past websites that deliberately block automated access.
- Delivery cost, stock availability, quantity discounts, promotions.
- AdBlue fluid matching by pack size (no AdBlue fluid items are in the sample).

### What would make the POC a failure

- A competitor price shown as a match when it belongs to a different product.
- A price compared on the wrong VAT basis (a silent 20% error).
- Missing data hidden instead of explained.

## 3. Responsibilities

| Who | Responsible for |
|---|---|
| Client | Correctness of the input data: items, codes, own prices, competitors, links and hand-entered prices. |
| Us | A tool that reads that data, reports what it could not use and why, and explains how to correct it. |

## 4. How it works

### 4.1 Two commands

The tool is a single program file (`.exe`) started from a command window. No installation is needed.

| Command | What it does | Output |
|---|---|---|
| **Scan** | Looks up prices for every item on every relevant competitor site. | A run file, e.g. `Prices 2026-11-20 0930.xlsx`, and `Data issues.xlsx` |
| **Report** | Merges all run files in the output folder. | `Price comparison.xlsx` with tables and graphs |

### 4.2 Finding a product page

For each item and competitor, the tool works in this order:

1. **Use the link from the product links file**, if there is one.
2. **Otherwise search the competitor's site** by item code (MPN); if that finds nothing, by EAN; if that finds nothing, by name. Where a site's rules do not allow automated searching, the tool looks through the list of product pages the site publishes for search engines instead.
3. **Save a newly found link** to the product links file, so the next run does not search again.
4. **Search again** only when a saved link no longer works. Links the client entered or locked are never overwritten.

### 4.3 How sure the match is

Every price in the output carries one of these labels:

| Label | Meaning |
|---|---|
| **Client link** | The client supplied the product page. Trusted as given. |
| **Verified** | The tool found a single-product page, and the complete item code is stated on it. A code that is only part of a longer one does not count (`ZVA2` inside `ZVA25`). |
| **Unverified** | The tool found a page by name only. Shown separately, never as a confirmed price. |
| **Manual price** | The client typed the price in. Shown with the date it was entered. |
| **Not found / Blocked / Error** | No price, with the reason. |

Some product pages carry several prices (sizes, variants, pack quantities). The tool takes the price that belongs to the item's code. If it cannot tell which price that is, the result is **Unverified** and a question is written to the issues file.

Item codes that stand for a product family and not one product (for example a code ending in `x`) cannot be verified this way. They rely on client-supplied links.

### 4.4 VAT

- The client's prices are **ex VAT**.
- The tool captures both the ex-VAT and inc-VAT price where the page shows both.
- Where a site shows only one, the other is calculated at 20% and **marked as calculated**.
- Each site's VAT basis is recorded in the competitors file. We fill it in for the sites where we could confirm it from a product page; the others start as `Unknown`. The client can correct it.
- If the basis is unknown (for example, a newly added competitor), the price is shown as read, marked "VAT basis unknown", and is not compared until the basis is filled in. This is written to the issues file.
- Comparison against the client's price is always ex VAT.

### 4.5 Website status

Websites behave differently for a search and for a single product page. For example, one site refused its home page but served a product page normally. The tool therefore records two statuses per site:

| Status | Search | Product page |
|---|---|---|
| Reachable | The tool can search the site. | The tool can read prices from product links. |
| Refusing | The site blocks automated searching. A client-supplied link is needed. | The site blocks automated reading. A hand-entered price is the only option. |
| Unreachable | The site did not respond. | The site did not respond. |
| Search not worked out | The tool could not work out how to search the site. A client-supplied link is needed. | – |

Every run checks the sites that have no status yet, so the first run fills the list and a competitor added later is checked on the next run. A site that is Refusing or Unreachable is checked again when its last check is more than 30 days old. The reason a site does not work (forbidden by its rules, blocked, not found, no search box) is written in plain words in the Status note column. A refusal page is never read as if it were a product page.

### 4.6 Being a polite visitor

The tool identifies itself honestly, follows each site's published rules for automated visitors, and waits between requests to the same site (several sites ask for 10–20 seconds). Different sites are visited in parallel.

Some sites accept only one of the two connection types (HTTP/2 or HTTP/1.1). When the tool checks a site, it tries HTTP/2 and, if the site refuses, tries HTTP/1.1 once, and remembers which one works. If the remembered one stops working, the site is tested again once in that scan. Apart from this, a refused page is never retried.

**Consequence:** a first run on 50 items is expected to take roughly 15–40 minutes. Later runs are faster because links are already known.

### 4.7 Progress display

While scanning, the command window shows:

- the item and competitor currently being checked,
- a running count (e.g. `212 / 640`),
- a one-line result for each lookup (price found, not found, blocked),
- a summary at the end: prices found, not found, blocked, the number of data issues, and where the issues file is.

## 5. Input files

All input files follow the same rules: one header row, one record per row, every row fully filled in, no blank separator rows.

### 5.1 `Items.xlsx` – one row per item

| Column | Filled by | Notes |
|---|---|---|
| Item code (MPN) | Client | Must be unique. The main key for matching. |
| Name | Client | |
| Brand | Client | Used in the search phrase, so it must be the brand competitors use for the product. |
| Competitor group | Client | Decides which competitors are checked. Must match a group in `Competitors.xlsx`. |
| Client price ex VAT | Client | May be empty; the item is then scanned but not compared. |
| EAN | Client | Optional. Used only as a supporting hint. |
| Search text | Client | Optional. Replaces the default search phrase (brand + code). |

### 5.2 `Competitors.xlsx` – one row per competitor per group

| Column | Filled by | Notes |
|---|---|---|
| Competitor group | Client | e.g. Piusi Equipment, Tanks, AdBlue |
| Company | Client | |
| Website | Client | |
| Search status | Tool | Reachable / Refusing / Unreachable / Search not worked out. Empty means not yet checked. Client may overwrite. |
| Product page status | Tool | Reachable / Refusing / Unreachable. Empty means not yet checked. |
| VAT basis | Tool, confirmed by client | Inc VAT / Ex VAT / Both shown / Unknown |
| Last checked | Tool | |
| Connection | Tool | HTTP/2 / HTTP/1.1 / Neither. Empty means not yet tested. |
| Status note | Tool | Plain-words reason the search or the product pages do not work. |
| Site checked | Tool | When the connection and search were last checked. |

### 5.3 `Product links.xlsx` – one row per item per competitor

| Column | Filled by | Notes |
|---|---|---|
| Item code | Tool or client | |
| Company | Tool or client | |
| Product page link | Tool or client | Tracking parts of a link are removed; parts that select a size or variant are kept. |
| Source | Tool | Found by tool / Entered by client |
| Locked | Client | "Yes" stops the tool from ever replacing the link. |
| Last result | Tool | Verified / Unverified / Not found / Blocked / Error |
| Last checked | Tool | |
| Manual price | Client | Only for rows the tool could not read. |
| Manual price VAT basis | Client | Inc VAT / Ex VAT |
| Manual price date | Client | |

The tool adds a row for every item–competitor pair it could not read, so the client sees exactly where a hand-entered link or price would help.

## 6. Output files

### 6.1 Run file – `Prices <date> <time>.xlsx`

One row per item per competitor:

Run date and time · Item code · Name · Competitor group · Client price ex VAT · Company · Price ex VAT · Price inc VAT · Which price was calculated · Match label · Product page link · Remark

### 6.2 Issues file – `Data issues.xlsx`

Written fresh at every scan, for the client to read and act on. The screen shows only the count and where the file is.

One row per issue:

Kind · Input file · Row · Item or competitor · What is wrong · Value in the client's data · Value seen (if any) · What to change

| Kind | Meaning |
|---|---|
| **Error** | The tool had to skip something. |
| **Question** | The tool carried on, but the client should look. |

**An empty issues file is the goal.** When nothing is wrong, the file holds a single line, "No issues found", so a clean run can be told apart from a failed one.

### 6.3 Comparison workbook – `Price comparison.xlsx`

Built from all run files in the folder.

- **Latest comparison:** one row per item, the client's price, each competitor's latest price, the cheapest competitor, and the difference in £ and %. Cells are coloured where the client is more expensive or cheaper.
- **History:** every price from every run, in one table the client can filter and pivot.
- **Graphs:** one price-over-time graph per item, with the client's price and each competitor as lines.
- **Coverage:** per competitor, how many items were found, unverified, not found or blocked.

A graph needs at least two runs on different days to show a line. We will make several runs before the demonstration.

## 7. Handling input data problems

The tool works with the data that is usable and writes everything else to the issues file.

| Situation | What the tool does | Kind |
|---|---|---|
| Item code missing or duplicated | Skips the affected rows (keeps the first of a duplicate). | Error |
| Competitor group missing or unknown | Skips the item. | Error |
| Link is not a product page (a category page, a bare website name, `#N/A`) | Ignores the link and searches instead. | Error |
| Link points to a site that is not in the competitors file | Ignores the link. | Error |
| Client price missing | Scans the item, leaves the comparison empty. | Question |
| Competitor page states a different MPN or EAN | Keeps the client's value and shows the value seen on the page. | Question |
| The same EAN on two different items | Uses neither EAN for matching. | Question |

Values seen on competitor pages are observations for the client to judge. The tool does not claim which value is correct and never changes the client's data.

The user guide explains how to correct each kind of issue.

## 8. Approaches considered

| Approach | Trade-off | Decision |
|---|---|---|
| **A. One general reader for all sites**, using each site's own search and the price information pages publish for search engines | Small single program; covers most sites; cannot read sites that block automated access | **Chosen** |
| B. A hand-written reader per site | Most accurate; 33 separate readers to build and maintain | Too costly for a POC. Possible later for a few key sites. |
| C. Driving a real web browser | Widest coverage; much larger and more fragile program; works around barriers sites put up on purpose | Not chosen. A decision for the client, if ever. |

Approach A is extended with the client-maintained product links and hand-entered prices, which cover most of what A alone would miss.

## 9. Technical choices (ours)

- Written in Python and delivered as a single Windows `.exe`. Python does not need to be installed.
- The program is not digitally signed. Windows may show an "unknown publisher" warning on first start.
- Input and output files sit in plain folders next to the program.
- New competitors are checked automatically on the next run: the tool tests access and works out how to search the site. Where it cannot work out the search, it says so in the competitors file, and that site then relies on client-supplied product links.

## 10. What we checked before writing this

These were small manual tests, not the tool's results. The first real run will produce the actual status list.

- Of the 33 competitor websites, about 21 answered a plain automated request, 9 refused it and 3 did not respond.
- A direct product link worked on 3 of 4 refusing sites tested.
- On the product pages tested, the price was readable in a reliable, structured form. Some pages carry several prices (sizes, variants), so the tool must select the right one.
- Each site's search works differently, so the tool has to work it out per site.
- At least six sites' published rules ask automated visitors not to use their search. For these the tool uses the site's published list of product pages, or client-supplied links.

## 11. Open questions for the client

1. **Which competitors apply to which items?** The competitor list has three groups (Piusi Equipment, Tanks, AdBlue). About 22 sample items are other brands (Elaflex, Cim-Tek, Tecalemit, Snaplock, Vacupress). *Our working assumption:* they are checked against the Piusi Equipment list. The group is an editable column per item.
2. **Codes that look like product families**, not single products. Which exact size or variant should be compared?
3. **Sites found in the client's own links but missing from the competitor list.** Should any be added?
4. **Blocked websites.** Is hand-entering prices acceptable for these, or should they be dropped?

## 12. Changes made during the build

The first live scans against the real sites led to these refinements. Each was recorded as a decision at the time.

| Change | Why |
|---|---|
| Search order is item code (MPN), then EAN, then name. | Requested by the client: the code defines the exact size and variant. |
| The VAT basis is read from the page by arithmetic: if the page also shows the price × 1.2 or ÷ 1.2, both figures are taken as read. | Shops label prices inconsistently, so reading "inc VAT" labels gave wrong results. |
| A code found only in a page's body verifies a match only if the product name also matches. | Product pages list other products' codes under "related products". |
| Name-only matches must agree on model numbers and voltage. | Similar product names (EX50 vs E140, 12 V vs 230 V) produced wrong candidates. |
| A page that looks like a multi-pack is only `Unverified`. | A 12-pack was matched to a single filter. |
| New result: "Product page found, but its price cannot be read automatically". | Several sites publish no machine-readable price. Reporting them as "not found" was misleading. |
| A different MPN on a competitor's page is only reported when the client's code is nowhere on that page. | Shops publish their own stock codes in the MPN field, which flooded the issues file. |
| `Competitors.xlsx` has an extra tool column, `Search address`. | So a worked-out search is remembered between scans. |
| `Competitors.xlsx` has extra tool columns `Connection`, `Status note` and `Site checked`; refusing sites are re-checked after 30 days. | A site refused HTTP/2 requests but served HTTP/1.1, and refusal reasons were otherwise lost. |
| The comparison workbook has an extra sheet, `Unverified matches`. | So name-only matches can be checked by hand without mixing them into the comparison. |
| Test fixtures are small hand-made pages, not saved copies of shop pages. | Other companies' pages should not be redistributed. |
