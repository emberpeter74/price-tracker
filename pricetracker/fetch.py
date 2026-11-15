"""All network access: site rules (robots), per-site pacing and refusal detection."""
from __future__ import annotations

import re
import threading
import time
from dataclasses import dataclass
from enum import Enum
from urllib.parse import urlsplit
from urllib.robotparser import RobotFileParser

import httpx

from .urls import host_of

USER_AGENT = "PriceTrackerPOC/0.1 (price comparison tool)"
ROBOTS_NAME = "PriceTrackerPOC"
MAX_DELAY = 30.0
DEFAULT_MAX_BYTES = 3_000_000
_REFUSAL_TITLES = ("just a moment", "attention required", "access denied", "security check")
_TITLE = re.compile(r"<title[^>]*>(.*?)</title>", re.I | re.S)


class Outcome(str, Enum):
    OK = "ok"
    REFUSED = "refused"
    NOT_FOUND = "not found"
    UNREACHABLE = "unreachable"
    DISALLOWED = "disallowed by site rules"


@dataclass
class Page:
    outcome: Outcome
    url: str  # final address after redirects
    status: int = 0
    html: str = ""
    content: bytes = b""


class _Host:
    def __init__(self) -> None:
        self.lock = threading.Lock()
        self.last = None  # time of the last request
        self.robots: RobotFileParser | None = None
        self.robots_loaded = False
        self.delay = 0.0


class Fetcher:
    def __init__(self, transport: httpx.BaseTransport | None = None, clock=time.monotonic,
                 sleep=time.sleep, min_delay: float = 5.0, timeout: float = 25.0) -> None:
        self._clock, self._sleep, self._min_delay = clock, sleep, min_delay
        self._hosts: dict[str, _Host] = {}
        self._hosts_lock = threading.Lock()
        self._client = httpx.Client(
            transport=transport, http2=transport is None, follow_redirects=True, timeout=timeout,
            headers={"User-Agent": USER_AGENT,
                     "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                     "Accept-Language": "en-GB,en;q=0.9"})

    def close(self) -> None:
        self._client.close()

    # ----- per host state -----

    def _host(self, url: str) -> _Host:
        key = host_of(url)
        with self._hosts_lock:
            return self._hosts.setdefault(key, _Host())

    def _root(self, url: str) -> str:
        parts = urlsplit(url)
        return f"{parts.scheme}://{parts.netloc}"

    def _request(self, host: _Host, url: str, max_bytes: int) -> Page:
        """One paced request. Never raises."""
        with host.lock:
            delay = max(self._min_delay, host.delay) if self._min_delay > 0 else 0.0
            if host.last is not None:
                wait = host.last + delay - self._clock()
                if wait > 0:
                    self._sleep(wait)
            try:
                page = self._download(url, max_bytes)
            except (httpx.HTTPError, httpx.InvalidURL, OSError, ValueError):
                page = Page(Outcome.UNREACHABLE, url)
            host.last = self._clock()
            return page

    def _download(self, url: str, max_bytes: int) -> Page:
        deadline = time.monotonic() + 90
        with self._client.stream("GET", url) as resp:
            status, final = resp.status_code, str(resp.url)
            if status in (401, 403, 429):
                return Page(Outcome.REFUSED, final, status)
            if status in (404, 410):
                return Page(Outcome.NOT_FOUND, final, status)
            if status >= 400:
                return Page(Outcome.UNREACHABLE, final, status)
            chunks, size = [], 0
            for chunk in resp.iter_bytes():
                chunks.append(chunk)
                size += len(chunk)
                if size >= max_bytes or time.monotonic() > deadline:
                    break
            content = b"".join(chunks)
            encoding = resp.charset_encoding or "utf-8"
        try:
            html = content.decode(encoding, "replace")
        except LookupError:
            html = content.decode("utf-8", "replace")
        title = _TITLE.search(html[:20000])
        if title and any(t in title.group(1).lower() for t in _REFUSAL_TITLES):
            return Page(Outcome.REFUSED, final, status)
        return Page(Outcome.OK, final, status, html, content)

    def _load_robots(self, host: _Host, url: str) -> None:
        if host.robots_loaded:
            return
        page = self._request(host, self._root(url) + "/robots.txt", 500_000)
        if page.outcome == Outcome.OK:
            parser = RobotFileParser()
            parser.parse(page.html.splitlines())
            host.robots = parser
            try:
                wanted = parser.crawl_delay(ROBOTS_NAME)
                host.delay = min(float(wanted), MAX_DELAY) if wanted else 0.0
            except (TypeError, ValueError):
                host.delay = 0.0
        host.robots_loaded = True

    # ----- public -----

    def allowed(self, url: str) -> bool:
        """True when the site's rules for automated visitors permit this address."""
        if not host_of(url):
            return False
        host = self._host(url)
        self._load_robots(host, url)
        return host.robots is None or host.robots.can_fetch(ROBOTS_NAME, url)

    def get(self, url: str, max_bytes: int = DEFAULT_MAX_BYTES) -> Page:
        if not host_of(url):
            return Page(Outcome.UNREACHABLE, url)
        if not self.allowed(url):
            return Page(Outcome.DISALLOWED, url)
        return self._request(self._host(url), url, max_bytes)

    def sitemaps(self, site_url: str) -> list[str]:
        if not host_of(site_url):
            return []
        host = self._host(site_url)
        self._load_robots(host, site_url)
        listed = (host.robots.site_maps() if host.robots else None) or []
        return list(listed) or [self._root(site_url) + "/sitemap.xml"]
