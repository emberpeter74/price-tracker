import gzip

import httpx

from conftest import comp, item
from pricetracker.discover import (candidates_from_search, candidates_from_sitemap, check_site,
                                   find_candidates, find_search_address, search_terms)
from pricetracker.fetch import Fetcher, Outcome, Page
from pricetracker.models import Access

HOME = "https://www.shop.co.uk/"
FORM_Q = '<html><body><form action="/shop/search" method="get"><input type="text" name="q"></form></body></html>'
RESULTS = '''<html><body>
  <a href="/">Home</a> <a href="/category/pumps">Pumps</a> <a href="https://other.co.uk/f00373000">Elsewhere</a>
  <a href="/piusi-ex50-12v-f00371000">Piusi EX50 Fuel Transfer Pump 12v</a>
  <a href="/piusi-ex50-230v-f00373000">Piusi EX50 Fuel Transfer Pump 230v</a>
  <a href="/piusi-ex50-230v-f00373000?utm_source=x">Piusi EX50 Fuel Transfer Pump 230v</a>
  <a href="/piusi-ex50-atex-fuel-transfer-pump">Piusi EX50 ATEX Fuel Transfer Pump</a>
  <a href="/piusi-ex100-fuel-transfer-pump">Piusi EX100 Fuel Transfer Pump</a>
  <a href="/elaflex-zva-nozzle">Elaflex ZVA nozzle</a>
</body></html>'''


def fetcher(handler):
    seen = []

    def wrapped(request):
        seen.append(str(request.url))
        if request.url.path == "/robots.txt" and "/robots.txt" not in getattr(handler, "paths", ()):
            return httpx.Response(404)
        return handler(request)
    return Fetcher(transport=httpx.MockTransport(wrapped), min_delay=0), seen


def site(pages, robots=None):
    """Handler from {path or path?query: body or (status, body) or bytes}."""
    def handler(request):
        if request.url.path == "/robots.txt":
            return httpx.Response(200, text=robots) if robots else httpx.Response(404)
        key = request.url.path + ("?" + request.url.query.decode() if request.url.query else "")
        hit = pages.get(key, pages.get(request.url.path))
        if hit is None:
            return httpx.Response(404)
        if isinstance(hit, bytes):
            return httpx.Response(200, content=hit)
        status, body = hit if isinstance(hit, tuple) else (200, hit)
        return httpx.Response(status, text=body)
    handler.paths = ("/robots.txt",)
    return handler


# ----- finding the search form -----

def test_search_form_with_q_input():
    assert find_search_address(FORM_Q, HOME) == "https://www.shop.co.uk/shop/search?q={q}"


def test_wordpress_search_form_keeps_hidden_fields():
    html = ('<form role="search" method="get" action="https://www.shop.co.uk/">'
            '<input type="search" name="s"><input type="hidden" name="post_type" value="product"></form>')
    assert find_search_address(html, HOME) == "https://www.shop.co.uk/?post_type=product&s={q}"


def test_form_without_method_or_action_defaults_to_get_on_the_page():
    assert find_search_address('<form class="header-search"><input name="keywords"></form>', HOME) == (
        "https://www.shop.co.uk/?keywords={q}")


def test_search_input_with_bracketed_name_is_recognised():
    html = '<form action="https://www.shop.co.uk/search.html" method="get"><input name="search[keywords]" type="text"></form>'
    assert find_search_address(html, HOME) == "https://www.shop.co.uk/search.html?search%5Bkeywords%5D={q}"


def test_opencart_style_search_box_without_a_form():
    html = '<div id="search"><input type="text" name="search" placeholder="Search"><button></button></div>'
    assert find_search_address(html, HOME) == "https://www.shop.co.uk/index.php?route=product/search&search={q}"


def test_form_with_post_method_is_ignored():
    assert find_search_address('<form action="/search" method="POST"><input name="q"></form>', HOME) == ""


def test_login_and_newsletter_forms_are_ignored():
    html = ('<form action="/login"><input name="email"><input type="password" name="p"></form>'
            '<form action="/newsletter"><input type="email" name="email"></form>')
    assert find_search_address(html, HOME) == ""


def test_no_form_returns_empty():
    assert find_search_address("<html><body>hi</body></html>", HOME) == ""
    assert find_search_address("", HOME) == ""


# ----- checking a site -----

def checked(pages, robots=None, search_note="", **kw):
    f, seen = fetcher(site(pages, robots))
    c = comp(company="Shop", website=HOME, **kw)
    c.search_note = search_note
    check_site(f, c, "F00373000")
    return c, seen


def test_check_site_home_refused_sets_refusing():
    c, _ = checked({"/": (403, "no")})
    assert c.search_status == Access.REFUSING and c.search_address == ""
    assert c.search_note == "Home page blocked by the site (403)"


def test_check_site_home_forbidden_by_robots():
    c, _ = checked({"/": FORM_Q}, robots="User-agent: *\nDisallow: /\n")
    assert c.search_status == Access.REFUSING
    assert c.search_note == "Home page forbidden by the site's rules (robots.txt)"


def test_check_site_home_unreachable_sets_unreachable():
    c, _ = checked({"/": (500, "x")})
    assert c.search_status == Access.UNREACHABLE and c.search_note == "Site did not respond"
    d, _ = checked({})          # home page "not found", as RPM Fuels
    assert d.search_status == Access.UNREACHABLE
    assert d.search_note == "Home page not found (404) – check the website address"


def test_check_site_no_form_sets_search_not_worked_out():
    c, _ = checked({"/": "<html><body>no search here</body></html>"})
    assert c.search_status == Access.NO_SEARCH and c.search_address == ""
    assert c.search_note == "No search box found on the home page"


def test_check_site_search_disallowed_by_robots_sets_refusing_and_keeps_no_address():
    c, seen = checked({"/": FORM_Q, "/shop/search": RESULTS}, robots="User-agent: *\nDisallow: /shop/search\n")
    assert c.search_status == Access.REFUSING and c.search_address == ""
    assert c.search_note == "Search forbidden by the site's rules (robots.txt)"
    assert not any("/shop/search" in u for u in seen)


def test_check_site_working_search_sets_reachable_and_address():
    c, seen = checked({"/": FORM_Q, "/shop/search": RESULTS}, search_note="old")
    assert c.search_status == Access.REACHABLE and c.search_note == ""
    assert c.search_address == "https://www.shop.co.uk/shop/search?q={q}"
    assert seen[-1] == "https://www.shop.co.uk/shop/search?q=F00373000"


def test_check_site_search_that_fails_is_not_worked_out():
    c, _ = checked({"/": FORM_Q, "/shop/search": (404, "x")})
    assert c.search_status == Access.NO_SEARCH and c.search_address == ""
    assert c.search_note == "Search did not return a usable page"
    d, _ = checked({"/": FORM_Q, "/shop/search": (403, "x")})
    assert d.search_status == Access.REFUSING and d.search_note == "Search blocked by the site (403)"


def test_check_site_keeps_existing_address_and_only_retests_it():
    c, seen = checked({"/find": RESULTS}, search_address="https://www.shop.co.uk/find?term={q}")
    assert c.search_status == Access.REACHABLE
    assert c.search_address == "https://www.shop.co.uk/find?term={q}"
    assert not any(u == HOME for u in seen)


def test_check_site_client_address_that_does_not_respond():
    c, _ = checked({"/find": (500, "x")}, search_address="https://www.shop.co.uk/find?term={q}")
    assert c.search_status == Access.UNREACHABLE and c.search_note == "Search did not respond"


def test_check_site_uses_given_home_page_without_fetching_it():
    f, seen = fetcher(site({"/shop/search": RESULTS}))
    c = comp(company="Shop", website=HOME)
    check_site(f, c, "F00373000", home=Page(Outcome.OK, HOME, 200, FORM_Q))
    assert c.search_status == Access.REACHABLE and HOME not in seen


# ----- search terms -----

def test_search_terms_order_is_mpn_then_ean_then_name():
    assert search_terms(item(ean="8050031000670")) == [
        "F00373000", "8050031000670", "PIUSI EX50 Fuel Transfer Pump, 230v"]
    assert search_terms(item()) == ["F00373000", "PIUSI EX50 Fuel Transfer Pump, 230v"]


def test_search_text_column_goes_first():
    assert search_terms(item(search_text="Piusi EX50 230"))[0] == "Piusi EX50 230"


def test_family_code_is_not_used_as_a_search_term():
    assert search_terms(item(code="966130x", name="PIUSI Easyoil Oil Nozzle")) == ["PIUSI Easyoil Oil Nozzle"]


def test_search_terms_have_no_duplicates_or_blanks():
    assert search_terms(item(brand="", name="F00373000")) == ["F00373000"]


# ----- candidates -----

def test_candidates_rank_code_in_link_above_name_only():
    got = candidates_from_search(RESULTS, HOME + "shop/search?q=x", item())
    assert got[0] == "https://www.shop.co.uk/piusi-ex50-230v-f00373000"


def test_candidates_exclude_other_hosts_category_links_and_duplicates():
    got = candidates_from_search(RESULTS, HOME + "shop/search?q=x", item(), limit=10)
    assert all(u.startswith("https://www.shop.co.uk/piusi-") for u in got)
    assert len(got) == len(set(got)) and "https://www.shop.co.uk/elaflex-zva-nozzle" not in got


def test_candidates_capped_at_three():
    assert len(candidates_from_search(RESULTS, HOME, item())) == 3


def test_candidates_empty_for_page_without_matching_links():
    assert candidates_from_search("<a href='/about-us'>About us</a>", HOME, item()) == []


# ----- sitemap -----

INDEX = ('<sitemapindex><sitemap><loc>https://www.shop.co.uk/sm-pages.xml</loc></sitemap>'
         '<sitemap><loc>https://www.shop.co.uk/sm-products.xml.gz</loc></sitemap></sitemapindex>')
PRODUCTS = gzip.compress(b'<urlset><url><loc>https://www.shop.co.uk/piusi-ex50-230v-f00373000</loc></url>'
                         b'<url><loc> https://www.shop.co.uk/piusi-ex50-12v-f00371000 </loc></url>'
                         b'<url><loc>https://www.shop.co.uk/harlequin-1400fs</loc></url></urlset>')
SM_SITE = {"/sitemap.xml": INDEX, "/sm-pages.xml": "<urlset><url><loc>https://www.shop.co.uk/about</loc></url></urlset>",
           "/sm-products.xml.gz": PRODUCTS}


def test_sitemap_index_is_followed_gzip_is_read_and_urls_ranked():
    f, _ = fetcher(site(SM_SITE))
    got = candidates_from_sitemap(f, comp(website=HOME), item(), {})
    assert got[0] == "https://www.shop.co.uk/piusi-ex50-230v-f00373000"
    assert "https://www.shop.co.uk/harlequin-1400fs" not in got and "https://www.shop.co.uk/about" not in got


def test_sitemap_is_read_once_per_site_per_run():
    f, seen = fetcher(site(SM_SITE))
    cache = {}
    candidates_from_sitemap(f, comp(website=HOME), item(), cache)
    n = len(seen)
    candidates_from_sitemap(f, comp(website=HOME), item(code="1400FS", name="Harlequin 1400FS"), cache)
    assert len(seen) == n and len(cache["shop.co.uk"]) == 4


def test_sitemap_limits_stop_runaway_sites():
    children = "".join(f"<sitemap><loc>https://www.shop.co.uk/sm{i}.xml</loc></sitemap>" for i in range(60))
    f, seen = fetcher(site({"/sitemap.xml": f"<sitemapindex>{children}</sitemapindex>"}))
    candidates_from_sitemap(f, comp(website=HOME), item(), {})
    assert sum("/sm" in u for u in seen) <= 20


def test_missing_or_refused_sitemap_gives_no_candidates():
    f, _ = fetcher(site({"/sitemap.xml": (403, "x")}))
    assert candidates_from_sitemap(f, comp(website=HOME), item(), {}) == []


# ----- choosing the route -----

def test_find_candidates_uses_search_when_reachable_and_stops_at_first_term_with_results():
    f, seen = fetcher(site({"/shop/search?q=F00373000": "<html>nothing</html>",
                            "/shop/search?q=8050031000670": RESULTS}))
    c = comp(website=HOME, search_status=Access.REACHABLE,
             search_address="https://www.shop.co.uk/shop/search?q={q}")
    got = find_candidates(f, c, item(ean="8050031000670"), {})
    assert got[0] == "https://www.shop.co.uk/piusi-ex50-230v-f00373000"
    assert [u for u in seen if "search" in u] == [
        "https://www.shop.co.uk/shop/search?q=F00373000",
        "https://www.shop.co.uk/shop/search?q=8050031000670"]


def test_find_candidates_uses_sitemap_when_search_is_not_usable():
    for status in (Access.NO_SEARCH, Access.REFUSING):
        f, seen = fetcher(site(SM_SITE))
        got = find_candidates(f, comp(website=HOME, search_status=status), item(), {})
        assert got and not any("search" in u for u in seen)


def test_find_candidates_for_unreachable_site_makes_no_requests():
    f, seen = fetcher(site(SM_SITE))
    assert find_candidates(f, comp(website=HOME, search_status=Access.UNREACHABLE), item(), {}) == []
    assert seen == []
