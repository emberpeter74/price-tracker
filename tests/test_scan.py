import io
from datetime import date, timedelta
from decimal import Decimal as D

import httpx
import openpyxl
import pytest

from conftest import COMP_HEADERS, ITEM_HEADERS, LINK_HEADERS, NOW, comp, item, link
from pricetracker.excel_io import InputError, read_competitors, read_links
from pricetracker.fetch import Fetcher
from pricetracker.issues import IssueLog
from pricetracker.models import Access, Connection, Label, VatBasis
from pricetracker.runfile import read_run
from pricetracker.scan import lookup, run_scan, site_check_due

A = "https://www.agriemach.com"
K33 = A + "/553000-piusi-k33-fuel-flow-meter"
FORM = '<form action="/shop/search"><input name="q"></form>'


def product(name="Piusi K33 Fuel Flow Meter", mpn="553000", price="146.14", currency="GBP", extra=""):
    mpn_json = f'"mpn":"{mpn}",' if mpn else ""
    return (f'<html><head><script type="application/ld+json">{{"@type":"Product","name":"{name}",{mpn_json}'
            f'"offers":{{"@type":"Offer","price":"{price}","priceCurrency":"{currency}"}}}}</script></head>'
            f'<body><h1>{name}</h1>{extra}</body></html>')


def results_page(*hrefs_and_text):
    return "<html><body>" + "".join(f'<a href="{h}">{t}</a>' for h, t in hrefs_and_text) + "</body></html>"


def make_fetcher(pages, raise_on=None):
    """pages: {full url without query, or with}: body | (status, body) | Exception."""
    seen = []

    def handler(request):
        url = str(request.url)
        seen.append(url)
        if raise_on and raise_on(url, len(seen)):
            raise KeyboardInterrupt()
        if request.url.path == "/robots.txt":
            return httpx.Response(404)
        hit = pages.get(url, pages.get(url.split("?")[0]))
        if hit is None:
            return httpx.Response(404)
        if isinstance(hit, Exception):
            raise hit
        status, body = hit if isinstance(hit, tuple) else (200, hit)
        return httpx.Response(status, text=body)

    return Fetcher(transport=httpx.MockTransport(handler), min_delay=0), seen


K33_ITEM = dict(code="553000", name="Piusi K33 Fuel Flow Meter", brand="Piusi", price="150")


def do(pages, lk=None, it=None, co=None):
    f, seen = make_fetcher(pages)
    log = IssueLog()
    it = it or item(**K33_ITEM)
    co = co or comp(search_status=Access.REACHABLE, search_address=A + "/shop/search?q={q}")
    res, out_link = lookup(f, it, co, lk, {}, log, NOW)
    return res, out_link, log, co, seen


# ---------- lookup: a link is present ----------

def test_client_link_ok_gives_client_link_label(load_page):
    res, lk, log, co, _ = do({K33: load_page("agriemach_k33.html")},
                             link(item_code="553000", url=K33, source="Entered by client"))
    assert res.label == Label.CLIENT_LINK and res.url == K33
    assert (res.price_ex, res.price_inc, res.calculated) == (D("146.14"), D("175.37"), "")
    assert res.client_price_ex == D("150") and res.run_at == NOW and res.company == "Agriemach"
    assert lk.last_result == "Client link" and lk.last_checked == NOW and len(log) == 0
    assert co.page_status == Access.REACHABLE


def test_tool_link_ok_and_code_matches_gives_verified():
    res, lk, _, _, _ = do({K33: product()}, link(item_code="553000", url=K33))
    assert res.label == Label.VERIFIED and lk.url == K33
    assert (res.price_ex, res.price_inc, res.calculated) == (D("121.78"), D("146.14"), "Ex VAT")


def test_tool_link_404_triggers_new_search_and_replaces_link():
    new = A + "/piusi-k33-553000-new"
    res, lk, _, _, _ = do({A + "/shop/search": results_page((new, "Piusi K33 Fuel Flow Meter 553000")),
                           new: product()}, link(item_code="553000", url=K33))
    assert res.label == Label.VERIFIED and res.url == new and lk.url == new and lk.source == "Found by tool"


def test_tool_link_that_no_longer_matches_is_searched_again():
    new = A + "/piusi-k33-553000-new"
    res, lk, _, _, _ = do({K33: product(name="Garden hose", mpn="999"),
                           A + "/shop/search": results_page((new, "Piusi K33 Fuel Flow Meter 553000")),
                           new: product()}, link(item_code="553000", url=K33))
    assert res.label == Label.VERIFIED and lk.url == new


def test_client_link_404_gives_error_keeps_link_and_logs_error():
    res, lk, log, _, seen = do({}, link(item_code="553000", url=K33, source="Entered by client", row=7))
    assert res.label == Label.ERROR and res.remark == "Link no longer works" and lk.url == K33
    assert [(i.kind, i.file, i.row) for i in log.issues] == [("Error", "Product links.xlsx", 7)]
    assert not any("search" in u for u in seen)


def test_locked_tool_link_404_is_not_replaced():
    res, lk, _, _, seen = do({}, link(item_code="553000", url=K33, locked=True))
    assert res.label == Label.ERROR and lk.url == K33 and not any("search" in u for u in seen)


def test_link_refused_with_manual_price_gives_manual_label_and_date_in_remark():
    lk = link(item_code="553000", url=K33, source="Entered by client", manual_price="100",
              manual_vat=VatBasis.EX, manual_date=date(2026, 11, 10))
    res, _, _, co, _ = do({K33: (403, "no")}, lk)
    assert res.label == Label.MANUAL and res.remark == "Entered by client on 10 Nov 2026"
    assert (res.price_ex, res.price_inc, res.calculated) == (D("100.00"), D("120.00"), "Inc VAT")
    assert co.page_status == Access.REFUSING


def test_link_refused_without_manual_price_gives_blocked():
    res, lk, _, co, _ = do({K33: (403, "no")}, link(item_code="553000", url=K33))
    assert res.label == Label.BLOCKED and res.price_ex is None and lk.last_result == "Blocked"
    assert "enter a price by hand" in res.remark


def test_site_unreachable_gives_error_with_remark_site_did_not_respond():
    res, _, _, co, _ = do({K33: httpx.ConnectError("down")}, link(item_code="553000", url=K33))
    assert res.label == Label.ERROR and res.remark == "Site did not respond"
    assert co.page_status == Access.UNREACHABLE


def test_client_link_with_unreadable_price_is_an_error():
    res, _, _, _, _ = do({K33: "<html><head><meta property='og:type' content='product'></head><body>call us</body></html>"},
                         link(item_code="553000", url=K33, source="Entered by client"))
    assert res.label == Label.ERROR and res.remark.startswith("Product page found, but its price cannot be read")


# ---------- lookup: no link ----------

def test_no_link_search_finds_verified_page_and_saves_tool_link():
    res, lk, _, _, seen = do({A + "/shop/search": results_page((K33, "Piusi K33 Fuel Flow Meter")),
                              K33: product()})
    assert res.label == Label.VERIFIED and lk.url == K33 and lk.source == "Found by tool"
    assert lk.row == 0 and lk.item_code == "553000" and lk.company == "Agriemach"
    assert seen[0] == A + "/robots.txt" and A + "/shop/search?q=553000" in seen


def test_first_verified_candidate_wins_over_an_earlier_unverified_one():
    other = A + "/piusi-k33-fuel-flow-meter-oil"
    res, lk, _, _, _ = do({A + "/shop/search": results_page((other, "Piusi K33 Fuel Flow Meter Oil"),
                                                            (K33 + "-x", "Piusi K33 Fuel Flow Meter")),
                           other: product(mpn="558000", price="170.53"),
                           K33 + "-x": product()})
    assert res.label == Label.VERIFIED and res.url == K33 + "-x" and res.price_inc == D("146.14")


def test_page_found_without_readable_price_is_reported_as_such_and_link_is_saved():
    hidden = ('<script type="application/ld+json">{"@type":"Product","name":"Piusi K33 Fuel Flow Meter 553000",'
              '"offers":[[]]}</script>')
    res, lk, _, _, _ = do({A + "/shop/search": results_page((K33, "Piusi K33 Fuel Flow Meter")), K33: hidden})
    assert res.label == Label.ERROR and res.url == K33 and lk.url == K33
    assert res.remark.startswith("Product page found, but its price cannot be read")


def test_manual_price_is_used_when_the_page_has_no_readable_price():
    hidden = '<script type="application/ld+json">{"@type":"Product","name":"Piusi K33 553000","offers":[[]]}</script>'
    lk = link(item_code="553000", url=K33, manual_price="100", manual_vat=VatBasis.EX, manual_date=date(2026, 11, 10))
    res, _, _, _, _ = do({K33: hidden}, lk)
    assert res.label == Label.MANUAL and res.price_ex == D("100.00")


def test_no_link_only_name_match_gives_unverified_and_saves_link():
    res, lk, log, _, _ = do({A + "/shop/search": results_page((K33, "Piusi K33 Fuel Flow Meter")),
                             K33: product(mpn="")})
    assert res.label == Label.UNVERIFIED and res.price_inc == D("146.14") and lk.url == K33
    assert "name only" in res.remark


def test_no_link_nothing_found_gives_not_found_and_adds_empty_link_row():
    res, lk, _, _, _ = do({A + "/shop/search": "<html>no results</html>"})
    assert res.label == Label.NOT_FOUND and res.url == "" and lk.url == "" and lk.row == 0
    assert lk.last_result == "Not found" and "supply a product page link" in res.remark


def test_no_link_but_manual_price_is_used_without_any_request():
    lk = link(item_code="553000", manual_price="120", manual_vat=VatBasis.INC, manual_date=date(2026, 11, 1))
    res, _, _, _, seen = do({}, lk)
    assert res.label == Label.MANUAL and res.price_ex == D("100.00") and seen == []


def test_read_only_client_link_is_searched_but_never_overwritten():
    lk = link(item_code="553000", source="Entered by client")
    lk.url_readonly = True
    res, out, _, _, _ = do({A + "/shop/search": results_page((K33, "Piusi K33 Fuel Flow Meter")),
                            K33: product()}, lk)
    assert res.label == Label.VERIFIED and res.url == K33 and out.url == "" and out.url_readonly


def test_all_candidates_refused_gives_blocked():
    res, _, _, co, _ = do({A + "/shop/search": results_page((K33, "Piusi K33 Fuel Flow Meter")),
                           K33: (403, "no")})
    assert res.label == Label.BLOCKED and co.page_status == Access.REFUSING


# ---------- lookup: prices and issues ----------

def test_unknown_vat_basis_gives_price_in_remark_no_figures_and_question():
    res, _, log, _, _ = do({K33: product()}, link(item_code="553000", url=K33),
                           co=comp(vat=VatBasis.UNKNOWN, row=4))
    assert res.label == Label.VERIFIED and res.price_ex is None and res.price_inc is None
    assert res.remark == "Price read: £146.14, VAT basis unknown"
    assert [(i.kind, i.file, i.row) for i in log.issues] == [("Question", "Competitors.xlsx", 4)]


def test_non_gbp_price_gives_error_with_currency_remark():
    res, _, _, _, _ = do({K33: product(currency="EUR")}, link(item_code="553000", url=K33))
    assert res.label == Label.ERROR and res.remark == "Currency EUR" and res.price_ex is None


def test_several_prices_gives_unverified_without_price_and_a_question():
    html = ('<script type="application/ld+json">{"@type":"Product","name":"Piusi K33 Fuel Flow Meter","mpn":"553000",'
            '"offers":[{"price":"146.14","sku":"A"},{"price":"170.53","sku":"B"}]}</script>')
    res, _, log, _, _ = do({K33: html}, link(item_code="553000", url=K33))
    assert res.label == Label.UNVERIFIED and res.price_ex is None and "Several prices" in res.remark
    assert log.issues[0].kind == "Question" and log.issues[0].file == "Product links.xlsx"


def test_observed_codes_become_questions_with_seen_value():
    res, _, log, _, _ = do({K33: product(mpn="F00553000")},
                           link(item_code="553000", url=K33, source="Entered by client", row=5))
    assert res.label == Label.CLIENT_LINK
    i = log.issues[0]
    assert (i.kind, i.file, i.subject, i.client_value, i.seen_value) == (
        "Question", "Items.xlsx", "553000 / Agriemach", "553000", "F00553000")
    assert "MPN" in i.problem and i.fix


def test_page_status_reachable_is_not_downgraded_by_one_refusal():
    co = comp(page_status=Access.REACHABLE)
    do({K33: (403, "no")}, link(item_code="553000", url=K33), co=co)
    assert co.page_status == Access.REACHABLE and co.page_note == ""


def test_page_note_follows_page_status():
    _, _, _, co, _ = do({K33: (403, "no")}, link(item_code="553000", url=K33))
    assert co.page_status == Access.REFUSING and co.page_note == "blocked by the site (403)"
    do({K33: product()}, link(item_code="553000", url=K33), co=co)
    assert co.page_status == Access.REACHABLE and co.page_note == ""


def test_site_check_due_rules():
    ok = comp(search_status=Access.REACHABLE, page_status=Access.REACHABLE)
    ok.connection, ok.site_checked = Connection.HTTP2, NOW - timedelta(days=60)
    assert not site_check_due([ok], NOW)                       # working sites are not re-checked
    bad = comp(search_status=Access.REFUSING)
    bad.connection, bad.site_checked = Connection.HTTP1, NOW - timedelta(days=31)
    assert site_check_due([bad], NOW)
    bad.site_checked = NOW - timedelta(days=29)
    assert not site_check_due([bad], NOW)
    neither = comp(search_status=Access.REACHABLE, page_status=Access.REACHABLE)
    neither.connection = Connection.NEITHER                    # never checked -> due
    assert site_check_due([neither], NOW)
    new = comp(search_status=Access.REACHABLE)                  # file from before this change
    assert site_check_due([new], NOW)


# ---------- run_scan ----------

H = "https://hytekgb.com"


def workspace(tmp_path, make_xlsx, items=None, comps=None, links=None):
    inp, out = tmp_path / "input", tmp_path / "output"
    inp.mkdir()
    make_xlsx(inp, "Items.xlsx", ITEM_HEADERS, items or [
        ["553000", "Piusi K33 Fuel Flow Meter", "Piusi", "Piusi Equipment", 150, None, None],
        ["558000", "Piusi K33 Oil Flow Meter", "Piusi", "Piusi Equipment", 180, None, None]])
    make_xlsx(inp, "Competitors.xlsx", COMP_HEADERS, comps or [
        ["Piusi Equipment", "Agriemach", A + "/", None, None, "Ex VAT", None, None],
        ["Piusi Equipment", "Hytek", H + "/", None, None, "Inc VAT", None, None]])
    if links is not None:
        make_xlsx(inp, "Product links.xlsx", LINK_HEADERS, links)
    return inp, out


SITE = {
    A + "/": FORM,
    A + "/shop/search?q=553000": results_page((K33, "Piusi K33 Fuel Flow Meter")),
    A + "/shop/search?q=558000": results_page((A + "/558000-k33-oil", "Piusi K33 Oil Flow Meter")),
    K33: product(),
    A + "/558000-k33-oil": product(name="Piusi K33 Oil Flow Meter", mpn="558000", price="170.53"),
    H + "/": "<html>no search form</html>",
    H + "/sitemap.xml": f"<urlset><url><loc>{H}/product/piusi-k33-fuel-flow-meter-553000/</loc></url></urlset>",
    H + "/product/piusi-k33-fuel-flow-meter-553000/": product(price="185.15"),
}


def test_run_scan_writes_run_file_issues_file_and_updates_inputs(tmp_path, make_xlsx):
    inp, out = workspace(tmp_path, make_xlsx)
    f, _ = make_fetcher(SITE)
    buf = io.StringIO()
    path = run_scan(inp, out, fetcher=f, now=NOW, out=buf)
    assert path == out / "Prices 2026-11-20 0930.xlsx"
    rows = read_run(path)
    assert [(r.item_code, r.company, r.label) for r in rows] == [
        ("553000", "Agriemach", Label.VERIFIED), ("553000", "Hytek", Label.VERIFIED),
        ("558000", "Agriemach", Label.VERIFIED), ("558000", "Hytek", Label.NOT_FOUND)]
    assert (rows[0].price_ex, rows[0].price_inc) == (D("146.14"), D("175.37"))
    assert (rows[1].price_ex, rows[1].price_inc) == (D("154.29"), D("185.15"))
    assert openpyxl.load_workbook(out / "Data issues.xlsx").active["A2"].value == "No issues found"

    agri, hytek = read_competitors(inp / "Competitors.xlsx")
    assert agri.search_status == Access.REACHABLE and agri.page_status == Access.REACHABLE
    assert agri.search_address == A + "/shop/search?q={q}" and agri.last_checked == NOW
    assert hytek.search_status == Access.NO_SEARCH and hytek.page_status == Access.REACHABLE

    links = {(l.item_code, l.company): l for l in read_links(inp / "Product links.xlsx")}
    assert len(links) == 4 and links[("553000", "Agriemach")].url == K33
    assert links[("558000", "Hytek")].url == "" and links[("558000", "Hytek")].last_result == "Not found"

    text = buf.getvalue()
    assert "4/4" in text and "Prices found:      3" in text and "Data issues.xlsx" in text


def test_second_run_reuses_saved_links_and_statuses(tmp_path, make_xlsx):
    inp, out = workspace(tmp_path, make_xlsx)
    run_scan(inp, out, fetcher=make_fetcher(SITE)[0], now=NOW, out=io.StringIO())
    f, seen = make_fetcher(SITE)
    run_scan(inp, out, fetcher=f, now=NOW.replace(day=21), out=io.StringIO())
    assert A + "/" not in seen and not any("q=553000" in u for u in seen)
    assert len(read_links(inp / "Product links.xlsx")) == 4
    assert (out / "Prices 2026-11-21 0930.xlsx").exists()


def test_same_site_in_two_groups_is_checked_once_and_items_stay_in_their_group(tmp_path, make_xlsx):
    inp, out = workspace(
        tmp_path, make_xlsx,
        items=[["553000", "Piusi K33 Fuel Flow Meter", "Piusi", "Piusi Equipment", 150, None, None],
               ["1400FS", "Harlequin 1400FS", "Harlequin", "Tanks", 1825, None, None]],
        comps=[["Piusi Equipment", "Agriemach", A + "/", None, None, "Ex VAT", None, None],
               ["Tanks", "Agriemach", A + "/", None, None, "Ex VAT", None, None],
               ["Tanks", "Hytek", H + "/", None, None, "Inc VAT", None, None]])
    f, seen = make_fetcher(SITE)
    rows = read_run(run_scan(inp, out, fetcher=f, now=NOW, out=io.StringIO()))
    assert sorted((r.item_code, r.company) for r in rows) == [
        ("1400FS", "Agriemach"), ("1400FS", "Hytek"), ("553000", "Agriemach")]
    assert seen.count(A + "/") == 1
    comps = read_competitors(inp / "Competitors.xlsx")
    assert comps[0].search_status == comps[1].search_status == Access.REACHABLE


def test_issues_from_validation_and_scan_reach_the_issues_file(tmp_path, make_xlsx):
    inp, out = workspace(tmp_path, make_xlsx, items=[
        ["553000", "Piusi K33 Fuel Flow Meter", "Piusi", "Piusi Equipment", None, None, None],
        ["553000", "Duplicate", "Piusi", "Piusi Equipment", 1, None, None]])
    buf = io.StringIO()
    run_scan(inp, out, fetcher=make_fetcher(SITE)[0], now=NOW, out=buf)
    ws = openpyxl.load_workbook(out / "Data issues.xlsx").active
    assert [ws.cell(r, 1).value for r in (2, 3)] == ["Error", "Question"]
    assert "Data issues:       2" in buf.getvalue()


def test_ctrl_c_still_writes_partial_run_file_and_links(tmp_path, make_xlsx):
    inp, out = workspace(tmp_path, make_xlsx, comps=[
        ["Piusi Equipment", "Agriemach", A + "/", "Reachable", None, "Ex VAT", None, A + "/shop/search?q={q}"]])
    f, _ = make_fetcher(SITE, raise_on=lambda url, n: "q=558000" in url)
    buf = io.StringIO()
    with pytest.raises(KeyboardInterrupt):
        run_scan(inp, out, fetcher=f, now=NOW, out=buf)
    rows = read_run(out / "Prices 2026-11-20 0930.xlsx")
    assert [(r.item_code, r.label) for r in rows] == [("553000", Label.VERIFIED)]
    assert read_links(inp / "Product links.xlsx")[0].url == K33
    assert (out / "Data issues.xlsx").exists() and "Stopped early" in buf.getvalue()


def test_ctrl_c_keeps_connection_found_so_far(tmp_path, make_xlsx):
    inp, out = workspace(tmp_path, make_xlsx, comps=[
        ["Piusi Equipment", "Agriemach", A + "/", "Reachable", None, "Ex VAT", None, A + "/shop/search?q={q}"]])
    f, _ = make_fetcher(SITE, raise_on=lambda url, n: "q=558000" in url)
    with pytest.raises(KeyboardInterrupt):
        run_scan(inp, out, fetcher=f, now=NOW, out=io.StringIO())
    saved = read_competitors(inp / "Competitors.xlsx")[0]
    assert saved.connection == Connection.HTTP2 and saved.site_checked == NOW


def test_run_scan_records_connection_note_and_site_checked(tmp_path, make_xlsx):
    inp, out = workspace(
        tmp_path, make_xlsx,
        items=[["553000", "Piusi K33 Fuel Flow Meter", "Piusi", "Piusi Equipment", 150, None, None]],
        comps=[["Piusi Equipment", "Agriemach", A + "/", None, None, "Ex VAT", None, None]],
        links=[["553000", "Agriemach", K33, "Entered by client", None, None, None, None, None, None]])
    f, _ = make_fetcher({A + "/": (403, "no"), K33: product()})
    rows = read_run(run_scan(inp, out, fetcher=f, now=NOW, out=io.StringIO()))
    assert [r.label for r in rows] == [Label.CLIENT_LINK]
    saved = read_competitors(inp / "Competitors.xlsx")[0]
    assert saved.connection == Connection.NEITHER and saved.site_checked == NOW
    assert saved.search_status == Access.REFUSING and saved.page_status == Access.REACHABLE
    ws = openpyxl.load_workbook(inp / "Competitors.xlsx").active
    notes = [ws.cell(2, c).value for c in range(1, ws.max_column + 1)
             if ws.cell(1, c).value == "Status note"]
    assert notes == ["Search: Home page blocked by the site (403)"]


def test_second_run_skips_site_check_but_recheck_after_30_days(tmp_path, make_xlsx):
    inp, out = workspace(tmp_path, make_xlsx, comps=[
        ["Piusi Equipment", "Agriemach", A + "/", None, None, "Ex VAT", None, None]], links=[
        ["553000", "Agriemach", K33, "Entered by client", None, None, None, None, None, None],
        ["558000", "Agriemach", K33, "Entered by client", None, None, None, None, None, None]])
    pages = {A + "/": (403, "no"), K33: product()}
    run_scan(inp, out, fetcher=make_fetcher(pages)[0], now=NOW, out=io.StringIO())
    f, seen = make_fetcher(pages)
    run_scan(inp, out, fetcher=f, now=NOW + timedelta(days=10), out=io.StringIO())
    assert A + "/" not in seen
    f, seen = make_fetcher(pages)
    run_scan(inp, out, fetcher=f, now=NOW + timedelta(days=40), out=io.StringIO())
    assert A + "/" in seen
    assert read_competitors(inp / "Competitors.xlsx")[0].site_checked == NOW + timedelta(days=40)


def test_unreachable_home_is_not_rechecked_before_30_days(tmp_path, make_xlsx):
    inp, out = workspace(tmp_path, make_xlsx, comps=[
        ["Piusi Equipment", "Agriemach", A + "/", None, None, "Ex VAT", None, None]], links=[
        ["553000", "Agriemach", K33, "Entered by client", None, None, None, None, None, None],
        ["558000", "Agriemach", K33, "Entered by client", None, None, None, None, None, None]])
    pages = {A + "/": httpx.ConnectError("down"), K33: product()}
    run_scan(inp, out, fetcher=make_fetcher(pages)[0], now=NOW, out=io.StringIO())
    assert read_competitors(inp / "Competitors.xlsx")[0].search_status == Access.UNREACHABLE
    f, seen = make_fetcher(pages)
    run_scan(inp, out, fetcher=f, now=NOW + timedelta(days=10), out=io.StringIO())
    assert A + "/" not in seen


def test_client_search_status_is_kept_when_only_the_connection_is_untested(tmp_path, make_xlsx):
    inp, out = workspace(tmp_path, make_xlsx, comps=[
        ["Piusi Equipment", "Agriemach", A + "/", "Search not worked out", "Reachable", "Ex VAT",
         None, None]], links=[
        ["553000", "Agriemach", K33, "Entered by client", None, None, None, None, None, None],
        ["558000", "Agriemach", K33, "Entered by client", None, None, None, None, None, None]])
    f, seen = make_fetcher(SITE)
    run_scan(inp, out, fetcher=f, now=NOW, out=io.StringIO())
    saved = read_competitors(inp / "Competitors.xlsx")[0]
    assert saved.connection == Connection.HTTP2 and seen.count(A + "/") == 1
    assert saved.search_status == Access.NO_SEARCH and saved.search_address == ""
    assert not any("/shop/search" in u for u in seen)


def test_hung_site_does_not_stop_other_sites(tmp_path, make_xlsx):
    inp, out = workspace(tmp_path, make_xlsx)
    pages = dict(SITE)
    for url in list(pages):
        if url.startswith(H):
            pages[url] = httpx.ReadTimeout("hung")
    rows = read_run(run_scan(inp, out, fetcher=make_fetcher(pages)[0], now=NOW, out=io.StringIO()))
    by = {(r.item_code, r.company): r for r in rows}
    assert by[("553000", "Agriemach")].label == Label.VERIFIED
    assert by[("553000", "Hytek")].label == Label.ERROR and by[("553000", "Hytek")].remark == "Site did not respond"
    assert read_competitors(inp / "Competitors.xlsx")[1].search_status == Access.UNREACHABLE


def test_unexpected_problem_on_one_lookup_does_not_lose_the_others(tmp_path, make_xlsx, monkeypatch):
    import pricetracker.scan as scan
    real = scan.parse_product_page

    def flaky(html, url):
        if "558000" in url:
            raise RuntimeError("boom")
        return real(html, url)
    monkeypatch.setattr(scan, "parse_product_page", flaky)
    inp, out = workspace(tmp_path, make_xlsx)
    rows = read_run(run_scan(inp, out, fetcher=make_fetcher(SITE)[0], now=NOW, out=io.StringIO()))
    by = {(r.item_code, r.company): r for r in rows}
    assert len(rows) == 4 and by[("558000", "Agriemach")].label == Label.ERROR
    assert by[("553000", "Agriemach")].label == Label.VERIFIED


def test_missing_input_file_raises_input_error(tmp_path):
    (tmp_path / "input").mkdir()
    with pytest.raises(InputError, match="Items.xlsx"):
        run_scan(tmp_path / "input", tmp_path / "output", fetcher=make_fetcher({})[0], now=NOW, out=io.StringIO())


def test_progress_shows_count_item_company_and_result(tmp_path, make_xlsx):
    inp, out = workspace(tmp_path, make_xlsx)
    buf = io.StringIO()
    run_scan(inp, out, fetcher=make_fetcher(SITE)[0], now=NOW, out=buf)
    lines = [l for l in buf.getvalue().splitlines() if "/4" in l]
    assert len(lines) == 4
    line = next(l for l in lines if "553000" in l and "Agriemach" in l)
    assert "Verified" in line and "£146.14 ex VAT" in line


def test_summary_counts_prices_without_a_vat_basis_separately(tmp_path, make_xlsx):
    inp, out = workspace(tmp_path, make_xlsx, comps=[
        ["Piusi Equipment", "Hytek", H + "/", None, None, "Unknown", None, None]])
    buf = io.StringIO()
    run_scan(inp, out, fetcher=make_fetcher(SITE)[0], now=NOW, out=buf)
    assert "Prices found:      0" in buf.getvalue() and "VAT basis unknown: 1" in buf.getvalue()


def test_progress_survives_a_console_that_cannot_show_pound_signs(tmp_path, make_xlsx):
    class Cp1250(io.StringIO):
        encoding = "cp1250"
    inp, out = workspace(tmp_path, make_xlsx)
    buf = Cp1250()
    run_scan(inp, out, fetcher=make_fetcher(SITE)[0], now=NOW, out=buf)
    assert "GBP 146.14 ex VAT" in buf.getvalue() and "£" not in buf.getvalue()
