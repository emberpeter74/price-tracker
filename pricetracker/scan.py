"""The scan command: look up every item on every relevant competitor site."""
from __future__ import annotations

import sys
import threading
from concurrent.futures import ThreadPoolExecutor, wait
from datetime import datetime, timedelta
from pathlib import Path

from .discover import RULES_NOTE, check_site, find_candidates
from .excel_io import read_competitors, read_items, read_links, write_competitors, write_links
from .extract import parse_product_page
from .fetch import Fetcher, Outcome, Page
from .issues import IssueLog
from .match import PRICE_NOT_READABLE, match_page
from .models import (ERROR, QUESTION, SOURCE_CLIENT, SOURCE_TOOL, Access, Competitor, Connection,
                     Item, Label, PriceResult, ProductLink, VatBasis)
from .progress import Progress
from .runfile import run_file_name, write_run
from .urls import host_of
from .validate import COMPS, ITEMS, LINKS, is_family_code, validate
from .vat import resolve_vat

MAX_SITES_AT_ONCE = 12
RECHECK_AFTER = timedelta(days=30)
BLOCKED_REMARK = "Site blocks automated reading; enter a price by hand in Product links.xlsx"
RULES_REMARK = "Site rules do not allow automated reading; enter a price by hand in Product links.xlsx"
NOT_FOUND_REMARK = "Not found on the site; supply a product page link in Product links.xlsx"
NO_PRICE_REMARK = ("Product page found, but its price cannot be read automatically; "
                   "enter a price by hand in Product links.xlsx")


_UNUSABLE = (Access.REFUSING, Access.UNREACHABLE)


def _note_page_status(comp: Competitor, page: Page) -> None:
    outcome = page.outcome
    if outcome == Outcome.OK:
        comp.page_status, comp.page_note = Access.REACHABLE, ""
        return
    notes = {Outcome.DISALLOWED: RULES_NOTE, Outcome.REFUSED: f"blocked by the site ({page.status})",
             Outcome.UNREACHABLE: "did not respond"}
    if outcome not in notes:
        return
    if comp.page_status is None:
        comp.page_status = Access.UNREACHABLE if outcome == Outcome.UNREACHABLE else Access.REFUSING
    if comp.page_status in _UNUSABLE:
        comp.page_note = notes[outcome]


def site_check_due(group: list[Competitor], now: datetime) -> bool:
    """Check a site (connection and search) when it was never checked, or when it was not
    usable and the last check is more than RECHECK_AFTER ago."""
    if any(c.search_status is None or c.connection is None for c in group):
        return True
    unusable = any(c.search_status in _UNUSABLE or c.page_status in _UNUSABLE
                   or c.connection == Connection.NEITHER for c in group)
    stale = any(c.site_checked is None or now - c.site_checked > RECHECK_AFTER for c in group)
    return unusable and stale


class _Lookup:
    """One item on one competitor site."""

    def __init__(self, fetcher, item, comp, link, cache, log, run_at):
        self.fetcher, self.item, self.comp = fetcher, item, comp
        self.cache, self.log, self.run_at = cache, log, run_at
        self.link = link or ProductLink(item.code, comp.company, "", SOURCE_TOOL, False, "", None,
                                        None, None, None, 0)
        self.subject = f"{item.code} / {comp.company}"

    # ----- building results -----

    def result(self, label: Label, url: str = "", remark: str = "", ex=None, inc=None, calc=""):
        self.link.last_result, self.link.last_checked = label.value, self.run_at
        it = self.item
        return PriceResult(self.run_at, it.code, it.name, it.group, it.client_price_ex,
                           self.comp.company, ex, inc, calc, label, url, remark)

    def manual(self):
        ln = self.link
        ex, inc, calc = resolve_vat(ln.manual_price, ln.manual_vat, frozenset())
        return self.result(Label.MANUAL, ln.url, f"Entered by client on {ln.manual_date:%d %b %Y}",
                           ex, inc, calc)

    def fetch(self, url: str) -> Page:
        page = self.fetcher.get(url)
        _note_page_status(self.comp, page)
        return page

    def from_page(self, page: Page, url: str, trusted: bool):
        """Result for a page that loaded, or None when the page is not the item."""
        parsed = parse_product_page(page.html, page.url)
        m = match_page(self.item, parsed, url, trusted)
        for name, client_value, seen in m.observed:
            self.log.add(QUESTION, ITEMS, self.item.row, self.subject,
                         f"{self.comp.company}'s page states a different {name}.", client_value, seen,
                         fix=f"Check the {name} of this item in {ITEMS} row {self.item.row}. "
                             f"If the page is a different product, correct the link in {LINKS}.")
        if m.label == Label.NOT_FOUND:
            if m.remark == PRICE_NOT_READABLE:  # the right page, but no machine-readable price
                return self.result(Label.ERROR, url, NO_PRICE_REMARK)
            return self.result(Label.ERROR, url, m.remark) if trusted else None
        if m.offer is None:  # several prices, none identifiable
            self.log.add(QUESTION, LINKS, self.link.row or None, self.subject, m.remark, url,
                         fix="Use a link that selects the exact size or variant, "
                             f"or enter the price by hand in {LINKS}.")
            return self.result(Label.UNVERIFIED, url, m.remark)
        if m.offer.currency != "GBP":
            return self.result(Label.ERROR, url, f"Currency {m.offer.currency}")
        label = Label.CLIENT_LINK if trusted and m.label == Label.VERIFIED else m.label
        ex, inc, calc = resolve_vat(m.offer.price, self.comp.vat_basis, parsed.text_prices)
        if ex is None:
            self.log.add(QUESTION, COMPS, self.comp.row, self.comp.company,
                         "The VAT basis of this site's prices is not known, so its prices are "
                         "shown but not compared.", self.comp.vat_basis.value,
                         fix=f"In {COMPS} row {self.comp.row}, set VAT basis to Inc VAT or Ex VAT.")
            return self.result(label, url, f"Price read: £{m.offer.price:.2f}, VAT basis unknown")
        return self.result(label, url, m.remark, ex, inc, calc)

    # ----- the decision table -----

    def run(self):
        ln = self.link
        has_manual = ln.manual_price is not None
        if ln.url:
            trusted = ln.source == SOURCE_CLIENT
            page = self.fetch(ln.url)
            if page.outcome == Outcome.OK:
                res = self.from_page(page, ln.url, trusted)
                if res is not None:
                    return self.manual() if res.remark == NO_PRICE_REMARK and has_manual else res
                if ln.locked:
                    return self.result(Label.ERROR, ln.url, "The page no longer shows this item")
            elif page.outcome in (Outcome.REFUSED, Outcome.DISALLOWED):
                if has_manual:
                    return self.manual()
                remark = BLOCKED_REMARK if page.outcome == Outcome.REFUSED else RULES_REMARK
                return self.result(Label.BLOCKED, ln.url, remark)
            elif page.outcome == Outcome.UNREACHABLE:
                return self.manual() if has_manual else self.result(Label.ERROR, ln.url, "Site did not respond")
            elif trusted or ln.locked:  # the page no longer exists
                if has_manual:
                    return self.manual()
                self.log.add(ERROR, LINKS, ln.row or None, self.subject, "The link no longer works.", ln.url,
                             fix=f"In {LINKS} row {ln.row}, replace the link, or clear it and set "
                                 "Locked to No so the tool searches again.")
                return self.result(Label.ERROR, ln.url, "Link no longer works")
        elif has_manual:
            return self.manual()
        return self.discover()

    def discover(self):
        ln = self.link
        best = None  # (result, url)
        refused = False
        for url in find_candidates(self.fetcher, self.comp, self.item, self.cache):
            page = self.fetch(url)
            if page.outcome in (Outcome.REFUSED, Outcome.DISALLOWED):
                refused = True
                continue
            if page.outcome != Outcome.OK:
                continue
            res = self.from_page(page, url, trusted=False)
            if res is None:
                continue
            if res.label == Label.VERIFIED:
                best = (res, url)
                break
            if best is None or (best[0].label == Label.ERROR and res.label != Label.ERROR):
                best = (res, url)  # a priced page beats a page whose price cannot be read
        if best is not None:
            res, url = best
            if not ln.url_readonly and not ln.locked:
                ln.url, ln.source = url, SOURCE_TOOL
            ln.last_result = res.label.value
            return res
        if not ln.url_readonly and not ln.locked:
            ln.url = ""
            ln.source = ln.source or SOURCE_TOOL
        if refused:
            return self.result(Label.BLOCKED, "", BLOCKED_REMARK)
        if self.comp.search_status == Access.UNREACHABLE:
            return self.result(Label.ERROR, "", "Site did not respond")
        return self.result(Label.NOT_FOUND, "", NOT_FOUND_REMARK)


def lookup(fetcher: Fetcher, item: Item, comp: Competitor, link: ProductLink | None, cache: dict,
           log: IssueLog, run_at: datetime) -> tuple[PriceResult, ProductLink]:
    job = _Lookup(fetcher, item, comp, link, cache, log, run_at)
    return job.run(), job.link


def _probe_term(items: list[Item]) -> str:
    for it in items:
        if not is_family_code(it.code):
            return it.code
    return items[0].name if items else "pump"


def run_scan(input_dir: Path, output_dir: Path, fetcher: Fetcher | None = None,
             now: datetime | None = None, out=None) -> Path:
    """Run a scan and return the path of the run file. Raises InputError for unusable input."""
    input_dir, output_dir = Path(input_dir), Path(output_dir)
    run_at = (now or datetime.now()).replace(second=0, microsecond=0)
    out = out if out is not None else sys.stdout
    log = IssueLog()

    items, comps, links = validate(read_items(input_dir / ITEMS), read_competitors(input_dir / COMPS),
                                   read_links(input_dir / LINKS), log)
    link_by = {(l.item_code.casefold(), l.company.casefold()): l for l in links}

    sites: dict[str, list[Competitor]] = {}
    for c in comps:
        sites.setdefault(host_of(c.website), []).append(c)
    jobs = {host: [(c, it) for c in group for it in items if it.group == c.group]
            for host, group in sites.items()}
    total = sum(len(j) for j in jobs.values())
    progress = Progress(total, out)
    progress.say(f"Scanning {len(items)} items on {len(sites)} competitor sites ({total} lookups).")
    if total:
        progress.say("Sites are visited politely, with pauses, so this takes a while. Press Ctrl+C to stop early.")

    stop = threading.Event()
    own_fetcher = fetcher is None
    if own_fetcher:
        fetcher = Fetcher(sleep=lambda seconds: stop.wait(seconds))
    results: list[PriceResult] = []
    lock = threading.Lock()

    def work(host: str) -> None:
        group = sites[host]
        lead = group[0]
        fetcher.set_connection(lead.website, next((c.connection for c in group if c.connection), None))
        if jobs[host] and site_check_due(group, run_at):
            lead.search_address = next((c.search_address for c in group if c.search_address), "")
            _, home = fetcher.test_connection(lead.website)
            check_site(fetcher, lead, _probe_term([it for _, it in jobs[host]]), home=home)
            for c in group:
                c.search_status, c.search_address = lead.search_status, lead.search_address
                c.search_note, c.connection = lead.search_note, fetcher.connection(lead.website)
                c.site_checked = c.last_checked = run_at
        cache: dict = {}
        for comp, item in jobs[host]:
            if stop.is_set():
                return
            key = (item.code.casefold(), comp.company.casefold())
            try:
                res, link = lookup(fetcher, item, comp, link_by.get(key), cache, log, run_at)
            except Exception as exc:  # one bad page must not lose the rest of the run
                link = link_by.get(key)
                res = PriceResult(run_at, item.code, item.name, item.group, item.client_price_ex,
                                  comp.company, None, None, "", Label.ERROR, link.url if link else "",
                                  f"Unexpected problem: {type(exc).__name__}")
            for c in group:  # saved even if the scan is stopped early
                c.page_status, c.page_note = comp.page_status, comp.page_note
                c.connection = fetcher.connection(lead.website)
                c.last_checked = run_at
            with lock:
                results.append(res)
                if link is not None and key not in link_by:
                    link_by[key] = link
                    links.append(link)
            progress.line(item, comp, res)

    interrupted = False
    pool = ThreadPoolExecutor(max_workers=max(1, min(len(sites), MAX_SITES_AT_ONCE)))
    try:
        futures = [pool.submit(work, host) for host in sites if jobs[host]]
        pending = set(futures)
        while pending:  # short waits, so Ctrl+C is noticed on Windows
            finished, pending = wait(pending, timeout=0.5)
            for f in finished:
                f.result()
    except KeyboardInterrupt:
        interrupted = True
        stop.set()
        raise
    finally:
        pool.shutdown(wait=True, cancel_futures=True)
        with lock:
            done = sorted(results, key=lambda r: (r.item_code, r.company))
        run_path = write_run(output_dir / run_file_name(run_at), done)
        issues_path = log.write(output_dir / "Data issues.xlsx")
        saved = [write_competitors(input_dir / COMPS, comps), write_links(input_dir / LINKS, links)]
        if own_fetcher:
            fetcher.close()
        if interrupted:
            progress.say("")
            progress.say("Stopped early. The results so far were saved.")
        progress.summary(done, len(log), issues_path)
        progress.say(f"  Run file:          {run_path}")
        for path in saved + [run_path, issues_path]:
            if "(new)" in path.name:
                progress.say(f"  NOTE: {path.name.replace(' (new)', '')} was open in Excel, "
                             f"so it was saved as {path.name}. Close Excel and rename it.")
    return run_path
