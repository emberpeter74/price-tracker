from decimal import Decimal as D

import pytest

from conftest import item
from pricetracker.extract import Offer, ProductPage
from pricetracker.match import code_in_text, match_page, name_overlap, norm_code
from pricetracker.models import Label

URL = "https://shop.example/p/1"


def page(name="PIUSI EX50 Fuel Transfer Pump 230v", offers=None, single=True, text="", **codes):
    offers = [Offer(D("343.00"))] if offers is None else offers
    return ProductPage(is_single_product=single, name=name, offers=offers,
                       text=text or name, **codes)


def test_norm_code():
    assert norm_code(" 2500fp/230 ") == "2500FP230" and norm_code("ZVA-ADBLUE") == "ZVAADBLUE"


@pytest.mark.parametrize("code,text,found", [
    ("ZVA2", "Elaflex ZVA25 nozzle", False),
    ("ZVA25", "Elaflex ZVA 25 Automatic Nozzle", True),      # separators inside a code are flexible
    ("F00373000", "Part no: F00373000.", True),
    ("f00373000", "PIUSI EX50 (F00373000)", True),
    ("70062", "Call 01234 570062", False),
    ("70062", "Cim-Tek 70062-10 filter", True),
    ("2500FP/230", "Harlequin 2500FP-230 fuel point", True),
    ("2500FP/230", "Harlequin 2500FP 230v", False),          # "230v" is a longer token
    ("305000", "Price £305000", False),                      # preceded by a currency sign
    ("146", "Only £99 or 146.14 with VAT", False),           # part of an amount
    ("ZVA2", "", False), ("", "anything", False),
])
def test_code_in_text(code, text, found):
    assert code_in_text(code, text) is found


def test_name_overlap():
    assert name_overlap("PIUSI EX50 Fuel Transfer Pump, 230v", "Piusi EX50 230V Fuel Transfer Pump") == 1.0
    assert name_overlap("PIUSI EX50 Fuel Transfer Pump, 230v", "Piusi EX50 12v pump") == pytest.approx(0.5)
    assert name_overlap("Kit c/w for and", "anything") == 0.0


# ----- identity -----

def test_structured_mpn_equal_to_item_code_is_verified():
    r = match_page(item(), page(name="Pump", mpn="f00373000"), URL, trusted=False)
    assert r.label == Label.VERIFIED and r.offer.price == D("343.00") and r.observed == []


def test_structured_sku_equal_to_item_code_is_verified():
    assert match_page(item(), page(name="Pump", sku="F00373000"), URL, False).label == Label.VERIFIED


def test_structured_gtin_equal_to_item_ean_is_verified():
    it = item(ean="0662256700623")
    assert match_page(it, page(name="Filter", gtin="662256700623"), URL, False).label == Label.VERIFIED


def test_code_in_product_name_is_verified():
    r = match_page(item(), page(name="Piusi EX50 pump F00373000"), URL, False)
    assert r.label == Label.VERIFIED


def test_code_only_in_page_text_of_single_product_page_is_verified():
    r = match_page(item(), page(text="Piusi EX50 Fuel Transfer Pump 230v. Product code F00373000"), URL, False)
    assert r.label == Label.VERIFIED


def test_code_in_text_but_a_different_product_name_is_only_unverified():
    # e.g. the code appears in a "related products" block of another product's page
    r = match_page(item(), page(name="Piusi K33 Flow Meter", text="K33 meter. Related: F00373000 EX50"), URL, False)
    assert r.label == Label.NOT_FOUND


def test_page_stating_a_different_mpn_is_not_verified_by_text():
    r = match_page(item(), page(mpn="F00371000", text="EX50 12v. Also available: F00373000"), URL, False)
    assert r.label == Label.UNVERIFIED          # name still matches, so it is offered as unverified
    assert r.observed == [("MPN", "F00373000", "F00371000")]


def test_code_on_listing_page_is_not_verified():
    r = match_page(item(), page(single=False, text="F00373000"), URL, False)
    assert r.label == Label.NOT_FOUND and r.offer is None


def test_family_code_is_never_verified():
    it = item(code="966130x", name="PIUSI Easyoil Oil Nozzle")
    r = match_page(it, page(name="Piusi Easyoil Oil Nozzle 966130x", mpn="966130x"), URL, False)
    assert r.label == Label.UNVERIFIED and "family" in r.remark


def test_name_only_match_is_unverified():
    r = match_page(item(), page(name="Piusi EX50 230v Fuel Transfer Pump ATEX"), URL, False)
    assert r.label == Label.UNVERIFIED and r.offer.price == D("343.00") and "name" in r.remark


def test_similar_name_on_a_page_stating_a_different_mpn_is_not_found():
    # "K33 fuel" page (mpn 553000) must not be offered for the "K33 oil" item (558000)
    it = item(code="558000", name="Piusi K33 Oil Flow Meter")
    r = match_page(it, page(name="Piusi K33 Fuel Flow Meter", mpn="553000"), URL, False)
    assert r.label == Label.NOT_FOUND


def test_low_name_overlap_is_not_found():
    assert match_page(item(), page(name="Piusi K33 flow meter"), URL, False).label == Label.NOT_FOUND


def test_page_without_offers_is_not_found_with_remark_price_not_readable():
    r = match_page(item(), page(mpn="F00373000", offers=[]), URL, False)
    assert r.label == Label.NOT_FOUND and r.offer is None and r.remark == "Price not readable on the page"
    t = match_page(item(), page(offers=[]), URL, True)
    assert t.label == Label.NOT_FOUND and t.remark == "Price not readable on the page"


# ----- choosing the offer -----

TWO = [Offer(D("2.70"), sku="F100A", url="https://shop.example/p/1?variant=11"),
       Offer(D("4.10"), sku="F150A", url="https://shop.example/p/1?variant=22")]


def test_single_offer_is_chosen():
    assert match_page(item(), page(mpn="F00373000"), URL, False).offer.price == D("343.00")


def test_offer_whose_sku_matches_item_code_is_chosen_over_first():
    it = item(code="F150A", name="Snaplock Type F Camlock")
    r = match_page(it, page(name="Snaplock Type F Camlock", offers=TWO), URL, False)
    assert r.label == Label.VERIFIED and r.offer.price == D("4.10")


def test_offer_matching_item_ean_is_chosen():
    offers = [Offer(D("1.00"), gtin="5056345600001"), Offer(D("2.00"), gtin="5056345600159")]
    it = item(code="XX1", name="Snaplock Camlock", ean="5056345600159")
    r = match_page(it, page(name="Snaplock Camlock", offers=offers), URL, False)
    assert r.label == Label.VERIFIED and r.offer.price == D("2.00")


def test_offer_matching_variant_id_in_url_is_chosen():
    r = match_page(item(), page(offers=TWO), "https://shop.example/p/1?variant=22", trusted=True)
    assert r.label == Label.VERIFIED and r.offer.price == D("4.10")


def test_several_offers_with_one_distinct_price_is_chosen():
    same = [Offer(D("9.99"), sku="A"), Offer(D("9.99"), sku="B")]
    r = match_page(item(), page(mpn="F00373000", offers=same), URL, False)
    assert r.label == Label.VERIFIED and r.offer.price == D("9.99")


def test_several_offers_different_prices_no_code_is_unverified_with_ambiguous_remark():
    r = match_page(item(), page(mpn="F00373000", offers=TWO), URL, False)
    assert r.label == Label.UNVERIFIED and r.offer is None
    assert r.remark == "Several prices on the page (£2.70, £4.10); could not tell which belongs to the item"


def test_trusted_link_with_ambiguous_offers_is_unverified_not_client_link():
    r = match_page(item(), page(name="Something else entirely", offers=TWO), URL, trusted=True)
    assert r.label == Label.UNVERIFIED and r.offer is None


def test_trusted_link_skips_the_identity_check():
    r = match_page(item(), page(name="Something else entirely", single=False), URL, trusted=True)
    assert r.label == Label.VERIFIED and r.offer.price == D("343.00")


# ----- observed codes -----

def test_observed_reports_different_mpn_and_ean():
    it = item(code="57500C", name="PIUSI Cube 56", ean="0089470123540")
    r = match_page(it, page(name="PIUSI Cube 56", mpn="F0057500C", gtin="8050031001234"), URL, True)
    assert r.observed == [("MPN", "57500C", "F0057500C"), ("EAN", "0089470123540", "8050031001234")]


def test_observed_ignores_formatting_differences_and_missing_values():
    it = item(code="2500FP/230", name="Harlequin 2500FP", ean="0662256700623")
    r = match_page(it, page(name="Harlequin 2500FP", mpn="2500fp-230", gtin="662256700623"), URL, True)
    assert r.observed == []
    assert match_page(item(), page(), URL, True).observed == []       # page states no codes
    no_ean = match_page(item(ean=""), page(mpn="F00373000", gtin="8050031000670"), URL, True)
    assert no_ean.observed == []


def test_observed_uses_the_chosen_offers_codes():
    it = item(code="F150A", name="Snaplock Type F Camlock", ean="5056345600159")
    offers = [Offer(D("2.70"), sku="F100A"), Offer(D("4.10"), sku="F150A", mpn="F150A", gtin="5056345600999")]
    r = match_page(it, page(name="Snaplock Type F Camlock", offers=offers), URL, False)
    assert r.observed == [("EAN", "5056345600159", "5056345600999")]
