"""Works out how to search a competitor's site and finds candidate product pages for an item."""
from __future__ import annotations

import gzip
import re
from urllib.parse import quote_plus, urlencode, urljoin, urlsplit

from selectolax.lexbor import LexborHTMLParser as HTMLParser

from .fetch import Fetcher, Outcome, Page
from .match import name_overlap, norm_code
from .models import Access, Competitor, Item
from .urls import clean_link, host_of, is_product_link, same_site
from .validate import is_family_code

SEARCH_INPUT_NAMES = {"q", "s", "search", "query", "keywords", "keyword", "term", "search_query", "w"}
MAX_SITEMAP_FILES = 20
MAX_SITEMAP_URLS = 100_000
SITEMAP_BYTES = 30_000_000
MIN_SCORE = 0.5
_LOC = re.compile(r"<loc>\s*(?:<!\[CDATA\[)?\s*(.*?)\s*(?:\]\]>)?\s*</loc>", re.I | re.S)


# ---------- the site's own search ----------

def find_search_address(home_html: str, home_url: str) -> str:
    """Address template of the site's search, with '{q}' where the search words go; '' if none."""
    if not home_html:
        return ""
    try:
        tree = HTMLParser(home_html)
    except Exception:
        return ""
    for form in tree.css("form"):
        attrs = form.attributes
        if (attrs.get("method") or "get").lower() != "get":
            continue
        if form.css_first('input[type="password"]'):
            continue
        field = None
        for inp in form.css("input"):
            name = (inp.attributes.get("name") or "").strip()
            kind = (inp.attributes.get("type") or "text").lower()
            if name and kind in ("search", "text") and (kind == "search" or _is_search_name(name)):
                field = name
                break
        if field is None:
            continue
        hidden = [(i.attributes.get("name"), i.attributes.get("value") or "")
                  for i in form.css('input[type="hidden"]') if i.attributes.get("name")]
        action = urljoin(home_url, attrs.get("action") or "")
        parts = urlsplit(action)
        base = f"{parts.scheme}://{parts.netloc}{parts.path or '/'}"
        fixed = "&".join(p for p in (parts.query, urlencode(hidden)) if p)
        return f"{base}?{fixed + '&' if fixed else ''}{quote_plus(field)}={{q}}"
    # OpenCart shops have a search box without a form
    if tree.css_first('input[name="search"]'):
        return urljoin(home_url, "/index.php?route=product/search&search={q}")
    return ""


def _is_search_name(name: str) -> bool:
    name = name.lower()
    return name in SEARCH_INPUT_NAMES or any(w in name for w in ("search", "keyword", "query"))


def _search_url(address: str, term: str) -> str:
    return address.replace("{q}", quote_plus(term))


RULES_NOTE = "forbidden by the site's rules (robots.txt)"


def check_site(fetcher: Fetcher, comp: Competitor, probe_term: str, home: Page | None = None) -> None:
    """Fill comp.search_status, comp.search_address and comp.search_note by trying the site's
    search once. `home` is the home page when the caller already loaded it."""
    address = comp.search_address
    if not address:
        home = home or fetcher.get(comp.website)
        if home.outcome == Outcome.DISALLOWED:
            comp.search_status, comp.search_note = Access.REFUSING, f"Home page {RULES_NOTE}"
            return
        if home.outcome == Outcome.REFUSED:
            comp.search_status = Access.REFUSING
            comp.search_note = f"Home page blocked by the site ({home.status})"
            return
        if home.outcome == Outcome.NOT_FOUND:
            comp.search_status = Access.UNREACHABLE
            comp.search_note = f"Home page not found ({home.status}) – check the website address"
            return
        if home.outcome != Outcome.OK:
            comp.search_status, comp.search_note = Access.UNREACHABLE, "Site did not respond"
            return
        address = find_search_address(home.html, home.url)
        if not address:
            comp.search_status, comp.search_note = Access.NO_SEARCH, "No search box found on the home page"
            return
    result = fetcher.get(_search_url(address, probe_term))
    if result.outcome == Outcome.OK:
        comp.search_status, comp.search_address, comp.search_note = Access.REACHABLE, address, ""
    elif result.outcome == Outcome.DISALLOWED:
        comp.search_status, comp.search_address = Access.REFUSING, ""
        comp.search_note = f"Search {RULES_NOTE}"
    elif result.outcome == Outcome.REFUSED:
        comp.search_status, comp.search_address = Access.REFUSING, ""
        comp.search_note = f"Search blocked by the site ({result.status})"
    elif result.outcome == Outcome.UNREACHABLE and comp.search_address:
        comp.search_status, comp.search_note = Access.UNREACHABLE, "Search did not respond"
    else:
        comp.search_status, comp.search_address = Access.NO_SEARCH, ""
        comp.search_note = "Search did not return a usable page"


def search_terms(item: Item) -> list[str]:
    """What to type into a site's search: the item code (MPN), then the EAN, then the name.
    A 'Search text' the client filled in goes first."""
    terms = [item.search_text]
    if not is_family_code(item.code):
        terms.append(item.code)
    terms.append(item.ean)
    terms.append(item.name)
    out = []
    for t in terms:
        t = (t or "").strip()
        if t and t not in out:
            out.append(t)
    return out


# ---------- scoring candidate links ----------

def _slug_words(url: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", " ", urlsplit(url).path)


def _score(item: Item, url: str, link_text: str) -> float:
    described = f"{link_text} {_slug_words(url)}"
    score = name_overlap(item.name, described)
    code = norm_code(item.code)
    if len(code) >= 4 and not is_family_code(item.code) and code in norm_code(described):
        score += 10
    return score


def _rank(item: Item, scored: dict[str, float], limit: int) -> list[str]:
    good = [(s, u) for u, s in scored.items() if s >= MIN_SCORE]
    good.sort(key=lambda su: (-su[0], len(su[1]), su[1]))
    return [u for _, u in good[:limit]]


def candidates_from_search(html: str, page_url: str, item: Item, limit: int = 3) -> list[str]:
    """Product links on a search results page that could be the item, best first."""
    try:
        tree = HTMLParser(html or "")
    except Exception:
        return []
    scored: dict[str, float] = {}
    for a in tree.css("a[href]"):
        url = clean_link(urljoin(page_url, a.attributes.get("href") or ""))
        if not same_site(url, page_url) or not is_product_link(url):
            continue
        text = " ".join(filter(None, [a.text(deep=True, separator=" ", strip=True),
                                      a.attributes.get("title")]))
        scored[url] = max(scored.get(url, 0.0), _score(item, url, text))
    return _rank(item, scored, limit)


# ---------- the site's published list of pages (sitemap) ----------

def _sitemap_text(content: bytes) -> str:
    if content[:2] == b"\x1f\x8b":
        try:
            content = gzip.decompress(content)
        except (OSError, EOFError):
            return ""
    return content.decode("utf-8", "replace")


def _load_sitemap(fetcher: Fetcher, comp: Competitor) -> list[str]:
    queue = list(fetcher.sitemaps(comp.website))
    urls: list[str] = []
    seen: set[str] = set()
    files = 0
    while queue and files < MAX_SITEMAP_FILES and len(urls) < MAX_SITEMAP_URLS:
        address = queue.pop(0)
        if address in seen:
            continue
        seen.add(address)
        files += 1
        page = fetcher.get(address, max_bytes=SITEMAP_BYTES)
        if page.outcome != Outcome.OK:
            continue
        text = _sitemap_text(page.content)
        locs = [m.strip() for m in _LOC.findall(text)]
        if "<sitemapindex" in text.lower():
            children = [u for u in locs if u not in seen]
            children.sort(key=lambda u: "product" not in u.lower())  # product lists first
            queue = children + queue
        else:
            urls.extend(u for u in locs if same_site(u, comp.website))
    return urls[:MAX_SITEMAP_URLS]


def candidates_from_sitemap(fetcher: Fetcher, comp: Competitor, item: Item,
                            cache: dict[str, list[str]], limit: int = 3) -> list[str]:
    host = host_of(comp.website)
    if host not in cache:
        cache[host] = _load_sitemap(fetcher, comp)
    scored = {}
    for raw in cache[host]:
        if is_product_link(raw):
            url = clean_link(raw)
            scored[url] = max(scored.get(url, 0.0), _score(item, url, ""))
    return _rank(item, scored, limit)


# ---------- choosing the route ----------

def find_candidates(fetcher: Fetcher, comp: Competitor, item: Item,
                    cache: dict[str, list[str]]) -> list[str]:
    """Candidate product pages for the item on this competitor's site, best first."""
    if comp.search_status == Access.UNREACHABLE:
        return []
    if comp.search_status == Access.REACHABLE and comp.search_address:
        for term in search_terms(item):
            url = _search_url(comp.search_address, term)
            page = fetcher.get(url)
            if page.outcome != Outcome.OK:
                continue
            found = candidates_from_search(page.html, page.url, item)
            if found:
                return found
        return []
    return candidates_from_sitemap(fetcher, comp, item, cache)
