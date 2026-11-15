"""Checks the input data. Usable rows are returned; everything else goes to the issue log."""
from __future__ import annotations

from collections import Counter

from .issues import IssueLog
from .models import ERROR, QUESTION, Competitor, Item, ProductLink
from .urls import clean_link, host_of, is_product_link

ITEMS, COMPS, LINKS = "Items.xlsx", "Competitors.xlsx", "Product links.xlsx"


def is_family_code(code: str) -> bool:
    """Codes such as '966130x' or 'VAC/25/' stand for a product family, not one product."""
    code = (code or "").strip()
    return code.endswith("x") or code.endswith("/")


def _key(s: str) -> str:
    return (s or "").strip().casefold()


def _competitors(competitors: list[Competitor], log: IssueLog) -> list[Competitor]:
    out, seen = [], {}
    for c in competitors:
        if not host_of(c.website):
            log.add(ERROR, COMPS, c.row, c.company, "The website is missing or is not a web address.",
                    c.website, fix=f"In {COMPS} row {c.row}, enter the full address in Website, "
                                   f"starting with https://")
            continue
        if not c.group or not c.company:
            log.add(ERROR, COMPS, c.row, c.company or c.website,
                    "The competitor group or the company name is missing.",
                    fix=f"In {COMPS} row {c.row}, fill in Competitor group and Company.")
            continue
        k = (_key(c.group), _key(c.company))
        if k in seen:
            log.add(ERROR, COMPS, c.row, c.company,
                    f"This company is already listed in group {c.group} (row {seen[k]}).",
                    fix=f"Delete row {c.row} in {COMPS}.")
            continue
        seen[k] = c.row
        out.append(c)
    return out


def _items(items: list[Item], groups: dict[str, str], log: IssueLog) -> list[Item]:
    out, seen = [], {}
    for it in items:
        if not it.code:
            log.add(ERROR, ITEMS, it.row, it.name, "The item code is missing.",
                    fix=f"In {ITEMS} row {it.row}, fill in Item code (MPN).")
            continue
        if _key(it.code) in seen:
            log.add(ERROR, ITEMS, it.row, it.code,
                    f"This item code is already used in row {seen[_key(it.code)]}.", it.code,
                    fix=f"Delete row {it.row} in {ITEMS}, or give it its own item code.")
            continue
        if _key(it.group) not in groups:
            problem = ("The competitor group is missing." if not it.group else
                       f"The competitor group '{it.group}' is not in {COMPS}.")
            log.add(ERROR, ITEMS, it.row, it.code, problem, it.group,
                    fix=f"In {ITEMS} row {it.row}, set Competitor group to one of: "
                        f"{', '.join(sorted(groups.values()))}.")
            continue
        seen[_key(it.code)] = it.row
        it.group = groups[_key(it.group)]
        out.append(it)

    ean_count = Counter(it.ean for it in out if it.ean)
    for it in out:
        if it.client_price_ex is None:
            log.add(QUESTION, ITEMS, it.row, it.code,
                    "No client price, so the item is scanned but not compared.",
                    fix=f"In {ITEMS} row {it.row}, fill in Client price ex VAT.")
        if is_family_code(it.code):
            log.add(QUESTION, ITEMS, it.row, it.code,
                    "The item code looks like a product family, not one product. "
                    "It cannot be verified on a competitor page.", it.code,
                    fix="Enter the exact code of the size or variant to compare, "
                        f"or supply product page links in {LINKS}.")
        if it.ean and not (it.ean.isdigit() and len(it.ean) == 13):
            log.add(QUESTION, ITEMS, it.row, it.code,
                    "The EAN does not have 13 digits, so it is not used.", it.ean,
                    fix=f"In {ITEMS} row {it.row}, correct the EAN or leave it empty. "
                        "Format the column as Text so leading zeros are kept.")
            it.ean = ""
        elif it.ean and ean_count[it.ean] > 1:
            log.add(QUESTION, ITEMS, it.row, it.code,
                    "The same EAN is used for another item, so it is not used for either.", it.ean,
                    fix=f"In {ITEMS} row {it.row}, correct the EAN or leave it empty.")
            it.ean = ""
    return out


def _links(links: list[ProductLink], items: list[Item], competitors: list[Competitor],
           log: IssueLog) -> list[ProductLink]:
    item_by = {_key(i.code): i for i in items}
    comp_by = {}
    for c in competitors:
        comp_by.setdefault(_key(c.company), c)
    out, seen = [], {}
    for ln in links:
        subject = f"{ln.item_code} / {ln.company}"
        it, co = item_by.get(_key(ln.item_code)), comp_by.get(_key(ln.company))
        if it is None or co is None:
            what = "item code" if it is None else "company"
            where = ITEMS if it is None else COMPS
            log.add(ERROR, LINKS, ln.row, subject, f"The {what} is not in {where}, so the row is ignored.",
                    ln.item_code if it is None else ln.company,
                    fix=f"In {LINKS} row {ln.row}, correct the {what}, or add it to {where}.")
            continue
        ln.item_code, ln.company = it.code, co.company
        k = (_key(ln.item_code), _key(ln.company))
        if k in seen:
            log.add(ERROR, LINKS, ln.row, subject,
                    f"This item and company already have a row (row {seen[k]}), so this one is ignored.",
                    fix=f"Delete row {ln.row} in {LINKS}.")
            continue
        if ln.url:
            if not host_of(ln.url) or not is_product_link(ln.url):
                log.add(ERROR, LINKS, ln.row, subject,
                        "The link is not a single product page, so the tool searches instead.", ln.url,
                        fix=f"In {LINKS} row {ln.row}, paste the address of the product's own page, "
                            "or clear the cell.")
                ln.url, ln.url_readonly = "", True
            elif host_of(ln.url) != host_of(co.website):
                log.add(ERROR, LINKS, ln.row, subject,
                        f"The link is not on {co.company}'s website ({host_of(co.website)}), "
                        "so the row is ignored.", ln.url,
                        fix=f"In {LINKS} row {ln.row}, use a link on {host_of(co.website)}, "
                            f"or add the other site to {COMPS}.")
                continue
            else:
                ln.url = clean_link(ln.url)
        if ln.manual_price is not None and (ln.manual_vat is None or ln.manual_date is None):
            log.add(ERROR, LINKS, ln.row, subject,
                    "A manual price needs its VAT basis and date, so it is not used.",
                    str(ln.manual_price),
                    fix=f"In {LINKS} row {ln.row}, fill in Manual price VAT basis "
                        "(Inc VAT or Ex VAT) and Manual price date.")
            ln.manual_price = None
        seen[k] = ln.row
        out.append(ln)
    return out


def validate(items: list[Item], competitors: list[Competitor], links: list[ProductLink],
             log: IssueLog) -> tuple[list[Item], list[Competitor], list[ProductLink]]:
    comps = _competitors(competitors, log)
    groups = {}
    for c in comps:
        groups.setdefault(_key(c.group), c.group)
    good_items = _items(items, groups, log)
    return good_items, comps, _links(links, good_items, comps, log)
