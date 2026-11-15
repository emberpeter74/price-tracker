"""Working out the ex-VAT and inc-VAT figures for a price read from a page."""
from __future__ import annotations

from decimal import ROUND_HALF_UP, Decimal

from .models import VatBasis

RATE = Decimal("1.2")
PENNY = Decimal("0.01")


def pence(value: Decimal) -> Decimal:
    return value.quantize(PENNY, rounding=ROUND_HALF_UP)


def _partner(target: Decimal, prices: frozenset[Decimal]) -> Decimal | None:
    """A price shown on the page within one penny of the target."""
    near = [p for p in prices if abs(p - target) <= PENNY]
    return min(near, key=lambda p: abs(p - target)) if near else None


def resolve_vat(price: Decimal, basis: VatBasis,
                text_prices: frozenset[Decimal]) -> tuple[Decimal | None, Decimal | None, str]:
    """Return (price ex VAT, price inc VAT, which one was calculated).

    If the page itself shows the matching other figure (price x 1.2 or price / 1.2), both
    figures are taken as read and nothing is calculated. Otherwise the site's VAT basis
    decides; with no known basis no figures are returned.
    """
    price = pence(price)
    inc_partner = _partner(pence(price * RATE), text_prices)
    ex_partner = _partner(pence(price / RATE), text_prices)
    if inc_partner is not None and ex_partner is None:
        return price, pence(inc_partner), ""
    if ex_partner is not None and inc_partner is None:
        return pence(ex_partner), price, ""
    if basis == VatBasis.INC:
        return pence(price / RATE), price, "Ex VAT"
    if basis == VatBasis.EX:
        return price, pence(price * RATE), "Inc VAT"
    return None, None, ""
