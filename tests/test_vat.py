from decimal import Decimal as D

from pricetracker.models import VatBasis
from pricetracker.vat import resolve_vat

NONE = frozenset()


def test_inc_basis_computes_ex():
    assert resolve_vat(D("120.00"), VatBasis.INC, NONE) == (D("100.00"), D("120.00"), "Ex VAT")


def test_ex_basis_computes_inc():
    assert resolve_vat(D("100.00"), VatBasis.EX, NONE) == (D("100.00"), D("120.00"), "Inc VAT")


def test_calculated_figures_are_rounded_to_pence():
    assert resolve_vat(D("146.14"), VatBasis.EX, NONE) == (D("146.14"), D("175.37"), "Inc VAT")
    assert resolve_vat(D("52.74"), VatBasis.INC, NONE) == (D("43.95"), D("52.74"), "Ex VAT")


def test_page_showing_the_inc_partner_means_price_is_ex_and_nothing_is_calculated():
    # Agriemach: structured price 146.14, page text also shows 175.37
    prices = frozenset({D("146.14"), D("175.37"), D("404.44")})
    assert resolve_vat(D("146.14"), VatBasis.UNKNOWN, prices) == (D("146.14"), D("175.37"), "")


def test_page_showing_the_ex_partner_means_price_is_inc():
    # The Tank Shop: structured 1820.40, text shows 1517.00
    assert resolve_vat(D("1820.40"), VatBasis.UNKNOWN, frozenset({D("1517.00"), D("1820.40")})) == (
        D("1517.00"), D("1820.40"), "")
    # Chemfast: structured 399.6, text shows 333
    assert resolve_vat(D("399.6"), VatBasis.INC, frozenset({D("333")})) == (D("333.00"), D("399.60"), "")


def test_page_evidence_beats_a_wrong_basis_in_the_competitors_file():
    assert resolve_vat(D("146.14"), VatBasis.INC, frozenset({D("175.37")})) == (
        D("146.14"), D("175.37"), "")


def test_one_penny_rounding_difference_is_accepted():
    assert resolve_vat(D("43.95"), VatBasis.UNKNOWN, frozenset({D("52.75")})) == (D("43.95"), D("52.75"), "")


def test_both_partners_on_page_is_ambiguous_and_falls_back_to_basis():
    prices = frozenset({D("100.00"), D("120.00"), D("144.00")})
    assert resolve_vat(D("120.00"), VatBasis.INC, prices) == (D("100.00"), D("120.00"), "Ex VAT")
    assert resolve_vat(D("120.00"), VatBasis.UNKNOWN, prices) == (None, None, "")


def test_unknown_basis_without_page_evidence_returns_no_figures():
    assert resolve_vat(D("50.00"), VatBasis.UNKNOWN, NONE) == (None, None, "")


def test_both_shown_basis_without_a_readable_pair_is_treated_as_unknown():
    assert resolve_vat(D("50.00"), VatBasis.BOTH, frozenset({D("50.00")})) == (None, None, "")
