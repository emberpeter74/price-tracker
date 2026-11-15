"""Reads the machine-readable product data from a product page."""
from __future__ import annotations

import html as htmllib
import json
import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from urllib.parse import urljoin

from selectolax.lexbor import LexborHTMLParser as HTMLParser

_PRICE_TEXT = re.compile(r"£\s?(\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+(?:\.\d{1,2})?)")
_PRICE_VALUE = re.compile(r"^£?\s*(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)$")
_GTIN_KEYS = ("gtin13", "gtin", "gtin14", "gtin12", "gtin8")


@dataclass
class Offer:
    price: Decimal
    currency: str = "GBP"
    sku: str = ""
    mpn: str = ""
    gtin: str = ""
    name: str = ""
    url: str = ""


@dataclass
class ProductPage:
    is_single_product: bool = False
    name: str = ""
    brand: str = ""
    sku: str = ""
    mpn: str = ""
    gtin: str = ""
    offers: list[Offer] = field(default_factory=list)  # [] when no machine-readable price
    text: str = ""  # visible text, whitespace-collapsed
    text_prices: frozenset = frozenset()  # every £ amount in the visible text


def parse_price(value) -> Decimal | None:
    """A positive amount from '£1,390.00', '52.74' or 399.6; otherwise None."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        value = repr(value)
    m = _PRICE_VALUE.match(str(value).strip())
    if not m:
        return None
    try:
        amount = Decimal(m.group(1).replace(",", ""))
    except InvalidOperation:
        return None
    return amount if amount > Decimal("0.01") else None  # 0.01 is a 'call for price' placeholder


def _s(value) -> str:
    """Text of a JSON-LD value that may be a string, number, or {'name': ...} object."""
    if isinstance(value, dict):
        value = value.get("name") or value.get("@id") or ""
    if isinstance(value, list):
        value = value[0] if value else ""
        return _s(value)
    if value is None:
        return ""
    return htmllib.unescape(htmllib.unescape(str(value))).strip()


def _types(node: dict) -> set[str]:
    t = node.get("@type")
    return {str(x) for x in t} if isinstance(t, list) else {str(t)} if t else set()


def _gtin(node: dict) -> str:
    for key in _GTIN_KEYS:
        if _s(node.get(key)):
            return _s(node.get(key))
    return ""


def _as_list(value) -> list:
    if value is None:
        return []
    if isinstance(value, list):
        out = []
        for v in value:
            out.extend(_as_list(v))
        return out
    return [value]


def _offers(node: dict, base_url: str, inherit: dict) -> list[Offer]:
    """Offers of one Product node. `inherit` supplies codes the offer does not state itself."""
    out = []
    for raw in _as_list(node.get("offers")):
        if not isinstance(raw, dict):
            continue
        currency = (_s(raw.get("priceCurrency")) or "GBP").upper()
        codes = dict(sku=_s(raw.get("sku")) or inherit.get("sku", ""),
                     mpn=_s(raw.get("mpn")) or inherit.get("mpn", ""),
                     gtin=_gtin(raw) or inherit.get("gtin", ""),
                     name=_s(raw.get("name")) or inherit.get("name", ""),
                     url=urljoin(base_url, _s(raw.get("url"))) if _s(raw.get("url")) else "")
        if "AggregateOffer" in _types(raw) and parse_price(raw.get("price")) is None:
            amounts = []
            for key in ("lowPrice", "highPrice"):
                amount = parse_price(raw.get(key))
                if amount is not None and amount not in amounts:
                    amounts.append(amount)
            out.extend(Offer(a, currency, **codes) for a in amounts)
            nested = _offers(raw, base_url, inherit) if raw.get("offers") else []
            if nested:
                out = [o for o in out if o.price not in amounts] + nested
            continue
        amount = parse_price(raw.get("price"))
        if amount is None:
            for spec in _as_list(raw.get("priceSpecification")):
                if isinstance(spec, dict) and parse_price(spec.get("price")) is not None:
                    amount = parse_price(spec.get("price"))
                    currency = (_s(spec.get("priceCurrency")) or currency).upper()
                    break
        if amount is not None:
            out.append(Offer(amount, currency, **codes))
    return out


def _walk(node, found: list[dict], in_list: bool = False) -> None:
    """Collect top-level Product / ProductGroup nodes (not those inside an ItemList)."""
    if isinstance(node, list):
        for n in node:
            _walk(n, found, in_list)
        return
    if not isinstance(node, dict):
        return
    types = _types(node)
    if types & {"Product", "ProductGroup"}:
        if not in_list:
            found.append(node)
        return
    listing = in_list or bool(types & {"ItemList", "OfferCatalog", "CollectionPage", "SearchResultsPage"})
    for key in ("@graph", "mainEntity", "mainEntityOfPage", "itemListElement", "item"):
        if key in node:
            _walk(node[key], found, listing or key == "itemListElement")


def _from_jsonld(tree: HTMLParser, base_url: str, page: ProductPage) -> int:
    """Fill `page` from JSON-LD. Returns the number of product nodes found."""
    products: list[dict] = []
    for script in tree.css('script[type="application/ld+json"]'):
        try:
            data = json.loads(script.text(deep=True, strip=True) or "", strict=False)
        except (ValueError, TypeError):
            continue
        _walk(data, products)
    if len(products) != 1:
        return len(products)
    node = products[0]
    page.name, page.brand = _s(node.get("name")), _s(node.get("brand"))
    page.sku, page.mpn, page.gtin = _s(node.get("sku")), _s(node.get("mpn")), _gtin(node)
    own = dict(sku=page.sku, mpn=page.mpn, gtin=page.gtin, name=page.name)
    page.offers = _offers(node, base_url, own)
    for variant in _as_list(node.get("hasVariant")):
        if isinstance(variant, dict):
            codes = dict(sku=_s(variant.get("sku")), mpn=_s(variant.get("mpn")),
                         gtin=_gtin(variant), name=_s(variant.get("name")) or page.name)
            page.offers.extend(_offers(variant, base_url, codes))
    return 1


def _attr_or_text(node) -> str:
    return _s(node.attributes.get("content") or node.text(deep=True, strip=True))


def _from_microdata(tree: HTMLParser, page: ProductPage) -> bool:
    scopes = [n for n in tree.css("[itemscope][itemtype]")
              if (n.attributes.get("itemtype") or "").rstrip("/").endswith("/Product")]
    if len(scopes) != 1:
        return False
    scope = scopes[0]

    def prop(name: str) -> str:
        node = scope.css_first(f'[itemprop="{name}"]')
        return _attr_or_text(node) if node else ""

    page.name = page.name or prop("name")
    page.brand = page.brand or prop("brand")
    page.sku, page.mpn = page.sku or prop("sku"), page.mpn or prop("mpn")
    page.gtin = page.gtin or next((prop(k) for k in _GTIN_KEYS if prop(k)), "")
    if not page.offers:
        amount = parse_price(prop("price"))
        if amount is not None:
            page.offers = [Offer(amount, (prop("priceCurrency") or "GBP").upper(), page.sku,
                                 page.mpn, page.gtin, page.name)]
    return True


def _meta(tree: HTMLParser, *names: str) -> str:
    for name in names:
        for attr in ("property", "name"):
            node = tree.css_first(f'meta[{attr}="{name}"]')
            if node and node.attributes.get("content"):
                return _s(node.attributes.get("content"))
    return ""


def parse_product_page(html: str, url: str) -> ProductPage:
    page = ProductPage()
    if not html or not html.strip():
        return page
    try:
        tree = HTMLParser(html)
    except Exception:
        return page

    count = _from_jsonld(tree, url, page)
    micro = _from_microdata(tree, page) if count <= 1 else False
    og_product = _meta(tree, "og:type").lower() in ("product", "og:product", "product.item")
    page.is_single_product = count == 1 or (count == 0 and (micro or og_product))

    if page.is_single_product and not page.offers:
        amount = parse_price(_meta(tree, "product:price:amount", "og:price:amount"))
        if amount is not None:
            currency = _meta(tree, "product:price:currency", "og:price:currency") or "GBP"
            page.offers = [Offer(amount, currency.upper(), page.sku, page.mpn, page.gtin, page.name)]
    if not page.name:
        h1 = tree.css_first("h1")
        title = tree.css_first("title")
        page.name = (_meta(tree, "og:title") or (_s(h1.text(deep=True, strip=True)) if h1 else "")
                     or (_s(title.text(deep=True, strip=True)) if title else ""))

    for tag in tree.css("script, style, noscript, template, svg"):
        tag.decompose()
    body = tree.body or tree.root
    raw = body.text(deep=True, separator=" ") if body else ""
    page.text = re.sub(r"\s+", " ", raw.replace("\xa0", " ")).strip()
    amounts = set()
    for m in _PRICE_TEXT.finditer(page.text):
        try:
            amounts.add(Decimal(m.group(1).replace(",", "")))
        except InvalidOperation:
            pass
    page.text_prices = frozenset(amounts)
    return page
