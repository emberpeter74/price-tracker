import gzip

import httpx
import pytest

from pricetracker.fetch import USER_AGENT, Fetcher, Outcome
from pricetracker.models import Connection

HTML = "<html><head><title>Piusi K33</title></head><body>ok</body></html>"


class Clock:
    def __init__(self):
        self.t = 1000.0
        self.sleeps = []

    def now(self):
        return self.t

    def sleep(self, seconds):
        self.sleeps.append(round(seconds, 3))
        self.t += seconds


def make(handler, min_delay=0.0):
    """Fetcher on a mock transport. `handler(request)` returns a Response, or raises."""
    clock = Clock()
    seen = []

    def wrapped(request):
        seen.append(request)
        return handler(request)

    f = Fetcher(transport=httpx.MockTransport(wrapped), clock=clock.now, sleep=clock.sleep,
                min_delay=min_delay)
    return f, seen, clock


def routes(table, robots=None):
    """Handler from {path: (status, body)}; robots.txt is 404 unless given."""
    def handler(request):
        if request.url.path == "/robots.txt":
            return httpx.Response(200, text=robots) if robots is not None else httpx.Response(404)
        status, body = table.get(request.url.path, (404, "nope"))
        return httpx.Response(status, text=body, headers={"content-type": "text/html; charset=utf-8"})
    return handler


def paths(seen):
    return [r.url.path for r in seen]


def test_ok_page_returns_html_and_final_url():
    def handler(request):
        if request.url.path == "/old":
            return httpx.Response(301, headers={"location": "https://a.co.uk/new"})
        return routes({"/new": (200, HTML)})(request)
    f, _, _ = make(handler)
    page = f.get("https://a.co.uk/old")
    assert page.outcome == Outcome.OK and page.html == HTML and page.status == 200
    assert page.url == "https://a.co.uk/new"


def test_403_is_refused_and_body_not_returned():
    f, _, _ = make(routes({"/p": (403, "<title>Just a moment...</title>")}))
    page = f.get("https://a.co.uk/p")
    assert page.outcome == Outcome.REFUSED and page.html == "" and page.status == 403


def test_200_with_just_a_moment_title_is_refused():
    f, _, _ = make(routes({"/p": (200, "<html><head><title>Just a moment...</title></head></html>")}))
    page = f.get("https://a.co.uk/p")
    assert page.outcome == Outcome.REFUSED and page.html == ""


def test_200_page_that_merely_embeds_challenge_script_is_ok():
    body = HTML.replace("ok", "<script src='/cdn-cgi/challenge-platform/scripts/jsd/main.js'></script>")
    f, _, _ = make(routes({"/p": (200, body)}))
    assert f.get("https://a.co.uk/p").outcome == Outcome.OK


@pytest.mark.parametrize("status,outcome", [
    (429, Outcome.REFUSED), (401, Outcome.REFUSED),
    (404, Outcome.NOT_FOUND), (410, Outcome.NOT_FOUND),
    (500, Outcome.UNREACHABLE), (503, Outcome.UNREACHABLE),
])
def test_status_classification(status, outcome):
    f, _, _ = make(routes({"/p": (status, "x")}))
    page = f.get("https://a.co.uk/p")
    assert page.outcome == outcome and page.html == ""


@pytest.mark.parametrize("exc", [httpx.ConnectError("dns"), httpx.ReadTimeout("slow"),
                                 httpx.ConnectTimeout("slow"), OSError("tls")])
def test_network_errors_are_unreachable(exc):
    def handler(request):
        raise exc
    f, _, _ = make(handler)
    page = f.get("https://a.co.uk/p")
    assert page.outcome == Outcome.UNREACHABLE and page.status == 0


def test_not_a_web_address_is_unreachable_without_a_request():
    f, seen, _ = make(routes({}))
    assert f.get("petroy.com").outcome == Outcome.UNREACHABLE and seen == []


def test_robots_disallow_returns_disallowed_without_requesting_the_page():
    f, seen, _ = make(routes({"/search": (200, HTML), "/p": (200, HTML)},
                             robots="User-agent: *\nDisallow: /search\n"))
    assert f.get("https://a.co.uk/search?q=k33").outcome == Outcome.DISALLOWED
    assert f.get("https://a.co.uk/p").outcome == Outcome.OK
    assert paths(seen) == ["/robots.txt", "/p"]
    assert f.allowed("https://a.co.uk/p") and not f.allowed("https://a.co.uk/search?q=1")


def test_robots_fetched_once_per_host():
    f, seen, _ = make(routes({"/p": (200, HTML)}, robots="User-agent: *\nDisallow:\n"))
    f.get("https://a.co.uk/p")
    f.get("https://www.a.co.uk/p")
    f.get("https://a.co.uk/p")
    assert paths(seen).count("/robots.txt") == 1


@pytest.mark.parametrize("robots_status", [404, 403, 500])
def test_missing_or_refused_robots_means_allowed(robots_status):
    def handler(request):
        if request.url.path == "/robots.txt":
            return httpx.Response(robots_status)
        return httpx.Response(200, text=HTML)
    f, _, _ = make(handler)
    assert f.get("https://a.co.uk/p").outcome == Outcome.OK


def test_second_request_to_same_host_waits_min_delay():
    f, _, clock = make(routes({"/p": (200, HTML)}), min_delay=5.0)
    f.get("https://a.co.uk/p")      # robots, then the page
    f.get("https://a.co.uk/p")
    assert clock.sleeps == [5.0, 5.0]


def test_no_wait_when_enough_time_has_passed():
    f, _, clock = make(routes({"/p": (200, HTML)}), min_delay=5.0)
    f.get("https://a.co.uk/p")
    clock.sleeps.clear()
    clock.t += 60
    f.get("https://a.co.uk/p")
    assert clock.sleeps == []


@pytest.mark.parametrize("crawl_delay,expected", [("20", 20.0), ("60", 30.0), ("2", 5.0)])
def test_crawl_delay_overrides_min_delay_and_is_capped(crawl_delay, expected):
    f, _, clock = make(routes({"/p": (200, HTML)},
                              robots=f"User-agent: *\nCrawl-delay: {crawl_delay}\nDisallow:\n"),
                       min_delay=5.0)
    f.get("https://a.co.uk/p")
    f.get("https://a.co.uk/p")
    assert clock.sleeps[-1] == expected


def test_different_hosts_do_not_wait_for_each_other():
    f, _, clock = make(routes({"/p": (200, HTML)}), min_delay=5.0)
    f.get("https://a.co.uk/p")
    clock.sleeps.clear()
    f.get("https://b.co.uk/p")
    assert clock.sleeps == [5.0]    # only b's own robots -> page gap


def test_user_agent_header_is_the_project_string():
    f, seen, _ = make(routes({"/p": (200, HTML)}))
    f.get("https://a.co.uk/p")
    assert USER_AGENT == "PriceTrackerPOC/0.1 (price comparison tool)"
    assert all(r.headers["user-agent"] == USER_AGENT for r in seen)


# ----- connection per site -----

CHALLENGE = (403, "<title>Just a moment...</title>")


def make2(h2_table, h1_table, h2_robots=None, h1_robots=None):
    """Fetcher whose HTTP/2 and HTTP/1.1 clients answer from separate tables.
    A table value may be an Exception to raise."""
    seen = {"h2": [], "h1": []}

    def side(name, table, robots):
        def handler(request):
            seen[name].append(request.url.path)
            if request.url.path == "/robots.txt":
                if isinstance(robots, tuple):
                    return httpx.Response(robots[0], text=robots[1])
                return httpx.Response(200, text=robots) if robots is not None else httpx.Response(404)
            hit = table.get(request.url.path, (404, "nope"))
            if isinstance(hit, Exception):
                raise hit
            status, body = hit
            return httpx.Response(status, text=body, headers={"content-type": "text/html"})
        return httpx.MockTransport(handler)

    f = Fetcher(transport=side("h2", h2_table, h2_robots),
                h1_transport=side("h1", h1_table, h1_robots), min_delay=0)
    return f, seen


def pages_only(paths_seen):
    return [p for p in paths_seen if p != "/robots.txt"]


def test_test_connection_keeps_http2_when_it_works():
    f, seen = make2({"/": (200, HTML)}, {"/": (200, HTML)})
    conn, page = f.test_connection("https://a.co.uk/")
    assert conn == Connection.HTTP2 and page.outcome == Outcome.OK
    assert f.connection("https://a.co.uk/") == Connection.HTTP2 and seen["h1"] == []


def test_test_connection_falls_back_to_http1_when_http2_is_refused():
    f, seen = make2({"/": CHALLENGE}, {"/": (200, HTML), "/p": (200, HTML)})
    conn, page = f.test_connection("https://a.co.uk/")
    assert conn == Connection.HTTP1 and page.outcome == Outcome.OK
    assert f.get("https://a.co.uk/p").outcome == Outcome.OK
    assert pages_only(seen["h2"]) == ["/"] and pages_only(seen["h1"]) == ["/", "/p"]


def test_test_connection_neither_when_both_refused_and_pages_still_use_http2():
    f, seen = make2({"/": CHALLENGE, "/p": (200, HTML)}, {"/": CHALLENGE})
    conn, page = f.test_connection("https://a.co.uk/")
    assert conn == Connection.NEITHER and page.outcome == Outcome.REFUSED
    assert f.get("https://a.co.uk/p").outcome == Outcome.OK
    assert pages_only(seen["h2"]) == ["/", "/p"] and pages_only(seen["h1"]) == ["/"]


def test_test_connection_404_counts_as_working():
    f, seen = make2({}, {})
    conn, page = f.test_connection("https://a.co.uk/")
    assert conn == Connection.HTTP2 and page.outcome == Outcome.NOT_FOUND and seen["h1"] == []


def test_test_connection_unreachable_keeps_previous():
    f, _ = make2({"/": httpx.ConnectError("down")}, {})
    f.set_connection("https://a.co.uk/", Connection.HTTP1)
    conn, page = f.test_connection("https://a.co.uk/")
    assert conn == Connection.HTTP1 and page.outcome == Outcome.UNREACHABLE


def test_first_refusal_retests_once_and_refetches_page():
    f, seen = make2({"/p": CHALLENGE, "/q": CHALLENGE, "/": CHALLENGE},
                    {"/": (200, HTML), "/p": (200, HTML), "/q": CHALLENGE})
    assert f.get("https://a.co.uk/p").outcome == Outcome.OK
    assert f.connection("https://a.co.uk/") == Connection.HTTP1
    assert f.get("https://a.co.uk/q").outcome == Outcome.REFUSED   # no second test
    assert pages_only(seen["h2"]) == ["/p", "/"]
    assert pages_only(seen["h1"]) == ["/", "/p", "/q"]


def test_robots_reloaded_after_switch_when_it_was_refused():
    f, seen = make2({"/": CHALLENGE}, {"/": (200, HTML), "/p": (200, HTML)},
                    h2_robots=CHALLENGE, h1_robots="User-agent: *\nDisallow: /p\n")
    assert f.test_connection("https://a.co.uk/")[0] == Connection.HTTP1
    assert f.get("https://a.co.uk/p").outcome == Outcome.DISALLOWED
    assert "/robots.txt" in seen["h1"] and "/p" not in seen["h1"]


def test_preset_connection_is_used_without_testing():
    f, seen = make2({"/p": (200, HTML)}, {"/p": (200, HTML)})
    f.set_connection("https://a.co.uk/", Connection.HTTP1)
    assert f.get("https://a.co.uk/p").outcome == Outcome.OK
    assert seen["h2"] == [] and pages_only(seen["h1"]) == ["/p"]


def test_sitemaps_from_robots_and_default():
    f, _, _ = make(routes({}, robots="Sitemap: https://a.co.uk/sm1.xml\nsitemap: https://a.co.uk/sm2.xml\n"))
    assert f.sitemaps("https://www.a.co.uk/shop/") == ["https://a.co.uk/sm1.xml", "https://a.co.uk/sm2.xml"]
    g, _, _ = make(routes({}))
    assert g.sitemaps("https://www.b.co.uk/shop/") == ["https://www.b.co.uk/sitemap.xml"]


def test_gzipped_body_is_available_as_bytes():
    raw = gzip.compress(b"<urlset><url><loc>https://a.co.uk/p1</loc></url></urlset>")

    def handler(request):
        if request.url.path == "/robots.txt":
            return httpx.Response(404)
        return httpx.Response(200, content=raw, headers={"content-type": "application/x-gzip"})
    f, _, _ = make(handler)
    page = f.get("https://a.co.uk/sitemap.xml.gz")
    assert page.outcome == Outcome.OK and gzip.decompress(page.content).startswith(b"<urlset>")


def test_body_is_capped():
    f, _, _ = make(routes({"/big": (200, "x" * 5000)}))
    assert len(f.get("https://a.co.uk/big", max_bytes=1000).content) <= 1000 + 65536
