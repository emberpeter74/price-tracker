"""Decides whether a page is the item, which of its prices applies, and how sure we are."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from urllib.parse import parse_qs, urlsplit

from .extract import Offer, ProductPage
from .models import Item, Label
from .validate import is_family_code

NAME_THRESHOLD = 0.6
PRICE_NOT_READABLE = "Price not readable on the page"
NO_MATCH = "No matching product page"
_STOP = {"with", "for", "and", "kit", "the", "c/w"}
_VOLTAGE = re.compile(r"^\d+v$")
_MULTIPACK = re.compile(r"multi[\s-]?pack|pack of \d+|box of \d+|case of \d+|\b\d+\s?pack\b|\bx\s?\d+\b|\b\d+\s?x\b",
                        re.I)


@dataclass
class MatchResult:
    label: Label  # VERIFIED, UNVERIFIED or NOT_FOUND
    offer: Offer | None
    remark: str = ""
    observed: list[tuple[str, str, str]] = field(default_factory=list)  # (field, client, seen)


def norm_code(code: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", code or "").upper()


def _digits(code: str) -> str:
    return re.sub(r"\D", "", code or "").lstrip("0")


def code_in_text(code: str, text: str) -> bool:
    """True when the complete code is in the text, not as a fragment of a longer code or number."""
    pieces = re.findall(r"[A-Za-z]+|\d+", code or "")
    if not pieces or not text:
        return False
    pattern = (r"(?<![A-Za-z0-9£$€.,])" + r"[\s\-/._]*".join(map(re.escape, pieces))
               + r"(?![A-Za-z0-9]|[.,]\d)")
    return re.search(pattern, text, re.I) is not None


def _words(name: str) -> list[str]:
    return re.findall(r"[a-z0-9]+(?:/[a-z0-9]+)?", (name or "").lower())


def _tokens(name: str) -> set[str]:
    return {w for w in _words(name) if len(w) >= 3 and w not in _STOP}


def name_overlap(item_name: str, page_name: str) -> float:
    """Share of the item's significant words that appear in the page's product name."""
    wanted = _tokens(item_name)
    if not wanted:
        return 0.0
    return len(wanted & _tokens(page_name)) / len(wanted)


def _models_agree(item_name: str, page_name: str) -> bool:
    """Every model number in the item's name (EX50, 56, 2500FP) must be in the page's name,
    and the page must not state a different voltage."""
    ours, theirs = set(_words(item_name)), set(_words(page_name))
    numbered = {w for w in ours if any(ch.isdigit() for ch in w) and len(w) >= 2}
    our_volts = {w for w in numbered if _VOLTAGE.match(w)}
    their_volts = {w for w in theirs if _VOLTAGE.match(w)}
    if our_volts and their_volts and not (our_volts & their_volts):
        return False
    return (numbered - our_volts) <= theirs


def _same_code(item: Item, *values: str) -> bool:
    code = norm_code(item.code)
    return bool(code) and any(norm_code(v) == code for v in values if v)


def _same_ean(item: Item, gtin: str) -> bool:
    return bool(item.ean and gtin) and _digits(item.ean) == _digits(gtin) != ""


def _offer_is_item(item: Item, offer: Offer) -> bool:
    return _same_code(item, offer.sku, offer.mpn) or _same_ean(item, offer.gtin)


def _choose_offer(item: Item, page: ProductPage, page_url: str) -> tuple[Offer | None, bool, str]:
    """Return (offer, offer itself carries the item's code, remark)."""
    offers = page.offers
    if not offers:
        return None, False, PRICE_NOT_READABLE
    coded = [o for o in offers if _offer_is_item(item, o)]
    if coded and len({o.price for o in coded}) == 1:
        return coded[0], True, ""
    if len(offers) == 1:
        return offers[0], False, ""
    variant = parse_qs(urlsplit(page_url).query).get("variant", [""])[0]
    if variant:
        hits = [o for o in offers if parse_qs(urlsplit(o.url).query).get("variant", [""])[0] == variant]
        if hits and len({o.price for o in hits}) == 1:
            return hits[0], False, ""
    if len({o.price for o in offers}) == 1:
        return offers[0], False, ""
    listed = ", ".join(f"£{p:.2f}" for p in sorted({o.price for o in offers})[:6])
    return None, False, f"Several prices on the page ({listed}); could not tell which belongs to the item"


def _observed(item: Item, page: ProductPage, offer: Offer | None, code_on_page: bool) -> list:
    """Codes on the page that differ from the client's. A different MPN is only reported when the
    client's own code is nowhere on the page; otherwise it is just the shop's own stock code."""
    mpn = (offer.mpn if offer and offer.mpn else page.mpn) or ""
    gtin = (offer.gtin if offer and offer.gtin else page.gtin) or ""
    out = []
    if mpn and item.code and not code_on_page and norm_code(mpn) != norm_code(item.code):
        out.append(("MPN", item.code, mpn))
    if gtin and item.ean and _digits(gtin) != _digits(item.ean):
        out.append(("EAN", item.ean, gtin))
    return out


def match_page(item: Item, page: ProductPage, page_url: str, trusted: bool) -> MatchResult:
    """`trusted` = a client-supplied link: identity is not checked, the price still has to be chosen."""
    if not trusted and not page.is_single_product:
        return MatchResult(Label.NOT_FOUND, None, "Not a single product page")

    offer, offer_coded, remark = _choose_offer(item, page, page_url)
    structured = offer_coded or _same_code(item, page.mpn, page.sku) or _same_ean(item, page.gtin)
    in_name = code_in_text(item.code, page.name)
    in_body = code_in_text(item.code, page.text)
    code_on_page = structured or in_name or in_body

    if trusted:
        label, why = Label.VERIFIED, ""
    else:
        family = is_family_code(item.code)
        overlap = name_overlap(item.name, page.name)
        other_mpn = bool(page.mpn) and not _same_code(item, page.mpn)
        in_text = in_body and overlap >= 0.5 and not other_mpn
        name_match = (overlap >= (1.0 if other_mpn else NAME_THRESHOLD)
                      and _models_agree(item.name, page.name))
        if not family and (structured or in_name or in_text):
            label, why = Label.VERIFIED, ""
            if _MULTIPACK.search(page.name) and not _MULTIPACK.search(item.name):
                label, why = Label.UNVERIFIED, "The page looks like a multi-pack, not a single item"
        elif family and (structured or in_name or name_match):
            label, why = Label.UNVERIFIED, "The item code is a product family, so the match cannot be verified"
        elif name_match:
            # a page that states a different code is only offered when the name matches completely
            label, why = Label.UNVERIFIED, "Matched by name only; the item code is not on the page"
        else:
            return MatchResult(Label.NOT_FOUND, None, NO_MATCH)

    if not page.offers:  # the right page, but no machine-readable price
        return MatchResult(Label.NOT_FOUND, None, PRICE_NOT_READABLE)
    if offer is None:  # several prices, none identifiable
        return MatchResult(Label.UNVERIFIED, None, remark, _observed(item, page, None, code_on_page))
    return MatchResult(label, offer, why, _observed(item, page, offer, code_on_page))
