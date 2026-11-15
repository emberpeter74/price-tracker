from datetime import date

from conftest import comp, item, link
from pricetracker.issues import IssueLog
from pricetracker.models import VatBasis
from pricetracker.validate import is_family_code, validate

AGRI = "https://www.agriemach.com/553000-piusi-k33-fuel-flow-meter"


def run(items=None, comps=None, links=None):
    log = IssueLog()
    out = validate(items if items is not None else [item()],
                   comps if comps is not None else [comp()],
                   links or [], log)
    return out, log


def kinds(log):
    return [(i.kind, i.file, i.row) for i in log.issues]


def test_clean_inputs_log_nothing():
    (items, comps, links), log = run(links=[link(url=AGRI)])
    assert len(log) == 0 and len(items) == 1 and len(comps) == 1 and links[0].url == AGRI


def test_duplicate_item_code_keeps_first_and_logs_error():
    (items, _, _), log = run(items=[item(row=2), item(code="f00373000", name="other", row=7)])
    assert [i.row for i in items] == [2]
    assert kinds(log) == [("Error", "Items.xlsx", 7)]
    assert "row 2" in log.issues[0].problem


def test_missing_item_code_is_skipped_with_error():
    (items, _, _), log = run(items=[item(code="", row=3), item(row=4)])
    assert [i.row for i in items] == [4] and kinds(log) == [("Error", "Items.xlsx", 3)]


def test_unknown_or_missing_group_skips_item_with_error():
    (items, _, _), log = run(items=[item(code="A", group="", row=2),
                                    item(code="B", group="Hoses", row=3),
                                    item(code="C", group="piusi equipment ", row=4)])
    assert [i.code for i in items] == ["C"]
    assert kinds(log) == [("Error", "Items.xlsx", 2), ("Error", "Items.xlsx", 3)]
    assert items[0].group == "Piusi Equipment"


def test_missing_client_price_is_kept_with_question():
    (items, _, _), log = run(items=[item(price=None)])
    assert len(items) == 1 and kinds(log) == [("Question", "Items.xlsx", 2)]


def test_shared_ean_is_blanked_on_both_items_with_question():
    (items, _, _), log = run(items=[item(code="A", ean="8050031002957", row=2),
                                    item(code="B", ean="8050031002957", row=3),
                                    item(code="C", ean="8050031000670", row=4)])
    assert [i.ean for i in items] == ["", "", "8050031000670"]
    assert kinds(log) == [("Question", "Items.xlsx", 2), ("Question", "Items.xlsx", 3)]


def test_short_ean_is_blanked_with_question():
    (items, _, _), log = run(items=[item(ean="89470123540")])
    assert items[0].ean == "" and log.issues[0].kind == "Question"
    assert log.issues[0].client_value == "89470123540"


def test_family_style_code_logs_question():
    assert is_family_code("966130x") and is_family_code("VAC/25/")
    assert not is_family_code("11050070X") and not is_family_code("F00373000")
    (items, _, _), log = run(items=[item(code="966130x")])
    assert len(items) == 1 and kinds(log) == [("Question", "Items.xlsx", 2)]
    assert "family" in log.issues[0].problem


def test_competitor_without_usable_website_is_dropped_with_error():
    (_, comps, _), log = run(comps=[comp(), comp(company="X", website="petroy.com", row=3),
                                    comp(company="Y", website="", row=4)])
    assert [c.company for c in comps] == ["Agriemach"]
    assert kinds(log) == [("Error", "Competitors.xlsx", 3), ("Error", "Competitors.xlsx", 4)]


def test_duplicate_competitor_in_group_keeps_first_with_error():
    (_, comps, _), log = run(comps=[comp(row=2), comp(row=5),
                                    comp(group="Tanks", row=6)])
    assert [c.row for c in comps] == [2, 6] and kinds(log) == [("Error", "Competitors.xlsx", 5)]


def test_non_product_link_is_blanked_with_error():
    (_, _, links), log = run(links=[link(url="https://www.agriemach.com/", source="Entered by client"),
                                    link(item_code="F00373000", company="Agriemach", row=3)])
    assert links[0].url == "" and links[0].url_readonly is True
    assert kinds(log)[0] == ("Error", "Product links.xlsx", 2)


def test_link_to_other_site_is_dropped_with_error():
    (_, _, links), log = run(links=[link(url="https://www.hss.com/buy/p/pc01822/adblue")])
    assert links == [] and kinds(log) == [("Error", "Product links.xlsx", 2)]
    assert "agriemach.com" in log.issues[0].problem


def test_link_for_unknown_item_or_company_is_dropped_with_error():
    (_, _, links), log = run(links=[link(item_code="NOPE", url=AGRI, row=2),
                                    link(company="Nobody Ltd", url=AGRI, row=3)])
    assert links == [] and kinds(log) == [("Error", "Product links.xlsx", 2),
                                          ("Error", "Product links.xlsx", 3)]


def test_duplicate_link_row_keeps_first_with_error():
    (_, _, links), log = run(links=[link(url=AGRI, row=2), link(url=AGRI + "-b", row=3)])
    assert [l.row for l in links] == [2] and kinds(log) == [("Error", "Product links.xlsx", 3)]


def test_links_are_cleaned_and_codes_matched_case_insensitively():
    (_, _, links), log = run(links=[link(item_code="f00373000", company="agriemach",
                                         url=AGRI + "?utm_source=google&gclid=1")])
    assert len(log) == 0 and links[0].url == AGRI
    assert links[0].item_code == "F00373000" and links[0].company == "Agriemach"


def test_manual_price_without_vat_basis_is_ignored_with_error():
    (_, _, links), log = run(links=[link(manual_price="555", manual_date=date(2026, 11, 1), row=2),
                                    link(item_code="F00373000", company="Agriemach",
                                         manual_price="1", manual_vat=VatBasis.EX, row=9)])
    assert links[0].manual_price is None and kinds(log)[0] == ("Error", "Product links.xlsx", 2)


def test_manual_price_with_basis_and_date_is_kept():
    (_, _, links), log = run(links=[link(manual_price="555", manual_vat=VatBasis.EX,
                                         manual_date=date(2026, 11, 1))])
    assert len(log) == 0 and str(links[0].manual_price) == "555"


def test_every_issue_says_what_to_change():
    _, log = run(items=[item(code="", row=2), item(price=None, ean="123", row=3)],
                 links=[link(url="https://elsewhere.co.uk/p/1")])
    assert log.issues and all(i.fix for i in log.issues)
