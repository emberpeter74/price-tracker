"""Link cleaning and comparison."""
from __future__ import annotations

from urllib.parse import parse_qsl, quote, urlsplit, urlunsplit

_KEEP_PARAMS = {"variant", "id", "product_id", "sku", "p"}
_CATEGORY_SEGMENTS = {"collections", "collection", "category", "categories", "c"}


def _split(url: str):
    url = (url or "").strip()
    if not url.lower().startswith(("http://", "https://")):
        return None
    parts = urlsplit(url)
    return parts if parts.netloc else None


def host_of(url: str) -> str:
    """Lower-case host without 'www.', or '' when the text is not a web address."""
    parts = _split(url)
    if parts is None:
        return ""
    host = (parts.hostname or "").lower()
    return host[4:] if host.startswith("www.") else host


def same_site(a: str, b: str) -> bool:
    ha, hb = host_of(a), host_of(b)
    return bool(ha) and ha == hb


def clean_link(url: str) -> str:
    """Remove tracking parameters and the fragment; keep parameters that select a variant."""
    parts = _split(url)
    if parts is None:
        return (url or "").strip()
    kept = [
        (k, v) for k, v in parse_qsl(parts.query, keep_blank_values=True)
        if k.lower() in _KEEP_PARAMS or k.lower().startswith("attr")
    ]
    query = "&".join(f"{quote(k, safe='')}={quote(v, safe='')}" for k, v in kept)
    return urlunsplit((parts.scheme, parts.netloc, parts.path, query, ""))


def is_product_link(url: str) -> bool:
    """True when the address looks like a single product page (not a home or category page)."""
    parts = _split(url)
    if parts is None:
        return False
    segments = [s for s in parts.path.split("/") if s]
    if not segments:
        return False
    if any(s.lower() in _CATEGORY_SEGMENTS for s in segments[:-1]):
        return False
    if len(segments) == 1 and parts.path.endswith("/") and not parts.query:
        return False
    return True
