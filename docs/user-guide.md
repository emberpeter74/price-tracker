# Price Tracker – User guide

Proof-of-concept version.

## 1. What the tool does

The tool looks up your items on your competitors' websites, reads their prices, and compares them with your own prices.

It has two commands:

| Command | What it does | Result |
|---|---|---|
| `pricetracker scan` | Looks up every item on every relevant competitor site. | `output\Prices <date> <time>.xlsx` and `output\Data issues.xlsx` |
| `pricetracker report` | Merges all the `Prices` files. | `output\Price comparison.xlsx`, with comparison tables and price graphs |

Run `scan` whenever you want fresh prices, for example once a week. Run `report` afterwards.

> **Before every scan, update your own prices.**
> The tool does not know your prices. It takes them from the **Client price ex VAT** column of `Items.xlsx` at the moment you start the scan, and stores them in that scan's results.
> If a price of yours has changed since the last scan, change it in `Items.xlsx` first. Otherwise the comparison and the graphs show your old price against the competitors' new ones.

## 2. The folder

Keep the files together like this:

```
pricetracker.exe
input\
    Items.xlsx
    Competitors.xlsx
    Product links.xlsx
    Input data findings.xlsx   (our notes on the original files; the tool does not read it)
output\                        (created by the tool)
```

## 3. Running it

1. Open `input\Items.xlsx` and make sure **Client price ex VAT** holds your current prices. Save and close it.
2. Close the other input files in Excel too.
3. Open a command window in the folder that holds `pricetracker.exe`. In File Explorer, click the address bar, type `cmd` and press Enter.
4. Type `pricetracker scan` and press Enter.
5. When it has finished, type `pricetracker report` and press Enter.
6. Open `output\Price comparison.xlsx`, and look at `output\Data issues.xlsx`.

**First start.** The program is not digitally signed, so Windows may show "Windows protected your PC". Click **More info**, then **Run anyway**.

**How long it takes.** The tool visits websites politely: it follows each site's rules for automated visitors and pauses between requests. Some sites ask for 20 seconds between requests. A first scan of about 50 items takes roughly 30 to 60 minutes. Later scans are quicker, because the tool remembers where each product page is.

**While it runs** you see one line per lookup:

```
 212/717  F00373000         Agriemach                 Verified      £343.00 ex VAT
```

You can stop early with **Ctrl+C**. Everything found so far is saved.

**If a file is open in Excel** when the tool wants to save it, the tool saves a copy named `… (new).xlsx` and tells you. Close Excel, delete the old file and rename the new one.

## 4. The input files

Rules for all three files:

- The first row holds the column names. Do not rename them. You may reorder columns and add columns of your own.
- One record per row. Blank rows are ignored.
- Format code columns (item code, EAN) as **Text**, so Excel does not drop leading zeros or round long numbers.

### 4.1 `Items.xlsx` – your items

| Column | You fill it? | Meaning |
|---|---|---|
| Item code (MPN) | Yes, required | The manufacturer's part number. Must be unique. The tool looks for this code on competitor pages. |
| Name | Yes, required | Product name. |
| Brand | Yes | The brand competitors use for the product. Used when searching. |
| Competitor group | Yes, required | Which competitors to check. Must be a group that exists in `Competitors.xlsx`. |
| Client price ex VAT | Yes, **before every scan** | Your current price, ex VAT. The scan records it as your price on that day. If empty, the item is scanned but not compared. |
| EAN | Optional | 13 digits. Only used as supporting evidence. |
| Search text | Optional | If the tool cannot find an item, type here what you would type into a shop's search box. |

**How the tool searches a site.** It searches by the item code (MPN) first. If that finds nothing, it searches by the EAN, and then by the name. A Search text you fill in is tried before all three.

**To add an item:** add a row and fill in at least the code, name and competitor group.

### 4.2 `Competitors.xlsx` – who to check

| Column | You fill it? | Meaning |
|---|---|---|
| Competitor group | Yes, required | For example Piusi Equipment, Tanks, AdBlue. |
| Company | Yes, required | The competitor's name. |
| Website | Yes, required | The home page, starting with `https://`. |
| Search status | Tool | Whether the tool can search the site (see below). |
| Product page status | Tool | Whether the tool can read product pages on the site. |
| VAT basis | Tool and you | Whether the site's prices are `Inc VAT` or `Ex VAT`. |
| Last checked | Tool | The last scan that visited the site. |
| Search address | Tool | How the tool searches the site. Leave it alone. |
| Connection | Tool | The connection type the site accepts: `HTTP/2`, `HTTP/1.1`, or `Neither` when the site refused both. Leave it alone. |
| Status note | Tool | Why the search or the product pages do not work, in plain words. For example "Search: Search forbidden by the site's rules (robots.txt)" or "Search: Home page not found (404) – check the website address". Empty when everything works. |
| Site checked | Tool | When the tool last checked the site's connection and search. |

A competitor that sells products from two groups gets two rows, one per group.

**To add a competitor:** add a row with the group, company and website. Leave the other columns empty. On the next scan the tool checks the site and fills in the statuses.

**Sites that do not work are checked again by themselves.** If a site is Refusing or Unreachable, or its Connection is `Neither`, the tool checks it again when its Site checked date is more than 30 days old. That re-check replaces a Search status of Refusing or Unreachable, even one you typed yourself; other statuses you type are kept.

**To have a site checked straight away:** clear its Search status and Connection.

| Status | Meaning | What you can do |
|---|---|---|
| Reachable | The tool can use the site. | Nothing. |
| Refusing | The site blocks automated visitors, or its rules forbid them. | For search: supply product page links. For product pages: enter prices by hand. |
| Unreachable | The site did not respond. | Check the website address. |
| Search not worked out | The tool could not work out the site's search. It tries the site's published page list instead. | Supply product page links for items it does not find. |

**VAT basis.** Where a page shows both the ex-VAT and inc-VAT price, the tool reads both and needs nothing from you. Where a page shows one price only, the tool uses this column. If it says `Unknown`, the price is shown but not compared, and `Data issues.xlsx` asks you to fill it in.

### 4.3 `Product links.xlsx` – where each product is

One row per item per competitor. The tool adds rows itself as it finds products, and adds an empty row for every product it could not find, so you can see where help is needed.

| Column | You fill it? | Meaning |
|---|---|---|
| Item code | Tool or you | |
| Company | Tool or you | Exactly as in `Competitors.xlsx`. |
| Product page link | Tool or you | The address of the product's own page. |
| Source | Tool or you | `Found by tool` or `Entered by client`. |
| Locked | You | `Yes` stops the tool from ever replacing the link. |
| Last result | Tool | The result of the last scan. |
| Last checked | Tool | |
| Manual price | You | A price you looked up yourself. |
| Manual price VAT basis | You | `Inc VAT` or `Ex VAT`. Required with a manual price. |
| Manual price date | You | The date you looked it up. Required with a manual price. |

**To give the tool a product page:** open the product on the competitor's site, copy the address from the browser, paste it into Product page link, and set Source to `Entered by client`. The tool trusts your link and never replaces it.

**To correct a link the tool found:** paste the right address and set Source to `Entered by client`.

**To make the tool search again:** clear the Product page link and the Source.

**To enter a price by hand** (for a site that blocks the tool): fill in Manual price, Manual price VAT basis and Manual price date. The tool uses it whenever it cannot read the site itself, and labels it `Manual price` with your date. Update it when the price changes.

## 5. The results

### 5.1 `Prices <date> <time>.xlsx`

One row per item per competitor, for one scan. Keep these files: the report and the graphs are built from them.

Each row has a **match label**:

| Label | Meaning |
|---|---|
| Client link | Read from a page you supplied. |
| Verified | The tool found the page itself, and your item code is stated on it. |
| Unverified | The tool found a likely page by name, but your item code is not on it, or the page has several prices and the tool could not tell which applies. **Check these by hand before relying on them.** |
| Manual price | The price you typed in. |
| Not found | The tool could not find the product on the site. |
| Blocked | The site blocks automated visitors. |
| Error | Something else went wrong. The Remark column says what. The most common: "Product page found, but its price cannot be read automatically". The site does not publish its price in a form the tool can read; the link is saved, and you can enter the price by hand. |

The **Calculated** column says which of the two prices the tool worked out at 20% VAT. If it is empty, both prices were read from the page.

### 5.2 `Price comparison.xlsx`

| Sheet | What it shows |
|---|---|
| Latest comparison | One row per item: your price, each competitor's latest price, the cheapest competitor and the difference. Red means you are more expensive than the cheapest competitor; green means you are cheaper or equal. Only Client link, Verified and Manual prices appear here. |
| Unverified matches | The likely-but-unconfirmed matches, with links, for you to check. |
| History | Every price from every scan. Use Excel's filters or a pivot table. |
| Graphs | One price-over-time graph per item. A line needs at least two scans on different days. |
| Graph data | The numbers behind the graphs. |
| Coverage | Per competitor, how many items were found, not found or blocked in the latest scan. |

If a product is no longer found in the latest scan, its cell in Latest comparison is empty. Its earlier prices remain in History and Graphs.

Your own price in each row is the one that was in `Items.xlsx` when that scan ran. The graphs therefore show your price history as well, as long as you update `Items.xlsx` before each scan.

### 5.3 `Data issues.xlsx`

Every scan writes this file afresh. It lists what the tool could not use in your input files, and questions for you to look at. **An empty file ("No issues found") is the goal.**

- **Error:** the tool had to skip something.
- **Question:** the tool carried on, but please check.

Each row names the file, the row and what to change. The most common ones:

| What is wrong | What to do |
|---|---|
| The item code is missing or already used | Give every item its own code in `Items.xlsx`, or delete the duplicate row. |
| The competitor group is missing or unknown | Set the item's Competitor group to a group that exists in `Competitors.xlsx`. |
| No client price | Fill in Client price ex VAT, or accept that the item is not compared. |
| The EAN does not have 13 digits / is used for another item | Correct the EAN or clear it. The tool works without it. |
| The item code looks like a product family | Enter the exact code of one size or variant, or supply product page links. |
| The link is not a single product page | Paste the address of the product's own page, not a category or home page. |
| The link is not on the company's website | Use a link on that competitor's site, or add the other site to `Competitors.xlsx`. |
| The link no longer works | Paste the new address, or clear the link so the tool searches again. |
| A manual price needs its VAT basis and date | Fill in both columns. |
| The page states a different MPN or EAN | The Value seen column shows what the competitor's page says. Either your code needs correcting, or the page is a different product and the link needs correcting. The tool never changes your data. |
| Several prices on the page | Use a link that selects the exact size, or enter the price by hand. |
| The VAT basis of this site is not known | Open a product on the site, see whether the price is inc or ex VAT, and set VAT basis in `Competitors.xlsx`. |

## 6. What the tool does not do

- It does not get past websites that block automated visitors. For those, enter prices by hand.
- It does not compare delivery costs, stock, quantity discounts or promotions.
- It does not change your items, codes or prices.
- It does not run by itself on a schedule.
