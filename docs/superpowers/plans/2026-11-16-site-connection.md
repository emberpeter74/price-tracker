# Site Connection and Status Notes Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Per site, find which connection type (HTTP/2 or HTTP/1.1) works, remember it in `Competitors.xlsx`, re-test when it stops working, write a plain-words status note, and re-check unusable sites every 30 days.

**Architecture:** `Fetcher` owns per-host connection state (two `httpx.Client`s, one per connection type) and does the one-off connection test, including the automatic re-test on the first refusal in a run. `check_site` writes the search note; `scan.run_scan` decides when a site check is due and copies results to every row of the site. `excel_io` reads and writes three new tool columns.

**Tech Stack:** Python 3.13, `httpx[http2]`, `openpyxl`, `pytest`.

**Spec:** `docs/superpowers/specs/2026-11-16-site-connection-design.md`

## Global Constraints

- User agent stays exactly `PriceTrackerPOC/0.1 (price comparison tool)`; no other header changes.
- `robots.txt` is checked before every page request on either connection; a forbidden page is never requested.
- At most one connection test per site per `Fetcher` (= per scan).
- New column titles, exactly: `Connection`, `Status note`, `Site checked`.
- Connection values, exactly: `HTTP/2`, `HTTP/1.1`, `Neither`.
- Re-check gap: 30 days.
- The uncommitted HTTP/1.1-only change in the working tree (`fetch.py`, `pyproject.toml`, `test_real_client_uses_http_1_1`) is replaced by this plan.

## Review Focus

1. **Competitors.xlsx from before this change** (no new columns): reads with Connection empty, every site is checked once on the first scan, columns are added on save. → Task 1 test, Task 4 test.
2. **Site that refuses its home page but serves product pages** (business spec 4.5): connection becomes `Neither`, product pages are still fetched over HTTP/2 and read. → Task 2 test.
3. **`robots.txt` refused over HTTP/2 but readable over HTTP/1.1 and forbidding a page:** after the switch the real rules apply. → Task 2 test.
4. **Junk typed into the Connection or Status note cell:** Connection reads as empty (site re-tested); note lines without a known prefix are ignored. → Task 1 test.
5. **Ctrl+C mid-scan:** connection and notes found so far are still saved. → Task 4 test.

Clarification of the spec (builder's call): in a connection test, a **404 answer counts as "this connection works"** (the server answered over it); only "did not respond" and "forbidden by robots.txt" leave the connection unchanged. Otherwise a site with a wrong address (RPM Fuels) would be re-tested every scan.

---

### Task 1: Model and Excel columns

**Files:**
- Modify: `pricetracker/models.py`
- Modify: `pricetracker/excel_io.py` (`COMP_COLS`, `read_competitors`, `write_competitors`)
- Test: `tests/test_excel_io.py`

**Interfaces:**
- Produces in `models.py`:
  - `class Connection(str, Enum): HTTP2 = "HTTP/2"; HTTP1 = "HTTP/1.1"; NEITHER = "Neither"`
  - `Competitor` gains, after `row`, with defaults: `connection: Connection | None = None`, `search_note: str = ""`, `page_note: str = ""`, `site_checked: datetime | None = None`.
- Produces in `excel_io.py`:
  - `COMP_COLS` adds `"connection": "Connection"`, `"note": "Status note"`, `"site_checked": "Site checked"`.
  - `join_note(search_note: str, page_note: str) -> str` — non-empty parts as lines `"Search: <text>"` and `"Product pages: <text>"`, joined by `"\n"`; `""` when both empty.
  - `split_note(cell: str) -> tuple[str, str]` — inverse; lines without one of the two prefixes are ignored.

- [ ] **Step 1: Write the failing tests**

```python
def test_note_round_trip_and_junk_is_ignored():
    assert join_note("forbidden by the site's rules (robots.txt)", "") == \
        "Search: forbidden by the site's rules (robots.txt)"
    both = join_note("a", "b")
    assert both == "Search: a\nProduct pages: b" and split_note(both) == ("a", "b")
    assert split_note("hello\nSearch: x") == ("x", "") and split_note("") == ("", "")

def test_competitors_new_columns_round_trip(tmp_path, make_xlsx):
    p = make_xlsx(tmp_path, "Competitors.xlsx", COMP_HEADERS,
                  [["Tanks", "T4E", "https://www.t4e.co.uk/", None, None, None, None, None]])
    rows = read_competitors(p)
    assert rows[0].connection is None and rows[0].site_checked is None and rows[0].search_note == ""
    rows[0].connection = Connection.HTTP1
    rows[0].search_note, rows[0].page_note = "x", "y"
    rows[0].site_checked = datetime(2026, 11, 20, 9, 30)
    out = write_competitors(p, rows)
    headers = [c.value for c in openpyxl.load_workbook(out).active[1]]
    assert headers[-3:] == ["Connection", "Status note", "Site checked"]
    again = read_competitors(out)[0]
    assert (again.connection, again.search_note, again.page_note, again.site_checked) == (
        Connection.HTTP1, "x", "y", datetime(2026, 11, 20, 9, 30))

def test_junk_connection_reads_as_untested(tmp_path, make_xlsx):
    p = make_xlsx(tmp_path, "Competitors.xlsx", COMP_HEADERS + ["Connection"],
                  [["Tanks", "T4E", "https://www.t4e.co.uk/", None, None, None, None, None, "fast"]])
    assert read_competitors(p)[0].connection is None
```

- [ ] **Step 2: Run** `.venv\Scripts\python -m pytest tests/test_excel_io.py -q` — Expected: FAIL (import errors for `Connection`, `join_note`).
- [ ] **Step 3: Implement** the model fields, `COMP_COLS`, `join_note`/`split_note`, and read/write of the three columns (`Connection` via `_enum`, `Site checked` via `_datetime`, note via `split_note`/`join_note`; write `None` for empty values).
- [ ] **Step 4: Run** `.venv\Scripts\python -m pytest -q` — Expected: all pass.
- [ ] **Step 5: Commit** — `git commit -m "feat: Connection, Status note and Site checked columns in Competitors.xlsx"`

### Task 2: Fetcher — per-site connection and connection test

**Files:**
- Modify: `pricetracker/fetch.py`, `pyproject.toml` (back to `httpx[http2]>=0.27`)
- Test: `tests/test_fetch.py` (delete `test_real_client_uses_http_1_1`)

**Interfaces:**
- Consumes: `Connection` (Task 1).
- Produces on `Fetcher`:
  - `__init__(self, transport=None, clock=..., sleep=..., min_delay=5.0, timeout=25.0, h1_transport=None)` — two clients: HTTP/2 (`http2=True`, transport `transport`) and HTTP/1.1 (`http2=False`, transport `h1_transport or transport`). (`http2=` only affects httpx's own transport, so a mock transport works for both.)
  - `set_connection(self, site_url: str, conn: Connection | None) -> None`
  - `connection(self, site_url: str) -> Connection | None`
  - `test_connection(self, site_url: str) -> tuple[Connection | None, Page]` — returns the recorded connection and the last home page result.
  - `get()` unchanged in signature; gains the automatic re-test.
- Behaviour:
  - Host requests use the HTTP/1.1 client when the host's connection is `HTTP1`, else the HTTP/2 client (`None`, `HTTP2`, `NEITHER`).
  - `test_connection`: marks the host tested; GET `site_url` over HTTP/2 (through `get`'s robots check). OK or NOT_FOUND → `HTTP2`. REFUSED → switch to HTTP/1.1 and GET again: OK or NOT_FOUND → `HTTP1`; REFUSED → `NEITHER`; UNREACHABLE/DISALLOWED → restore the previous value. UNREACHABLE/DISALLOWED on the first try → previous value kept.
  - Switching the client used for a host: when `robots.txt` was not read successfully (refused/unreachable) it is loaded again, through the new client, before the next request.
  - `get(url)`: when the result is REFUSED and the host is not yet tested in this `Fetcher`, call `test_connection(root of url + "/")`; if that changed the client used and the new connection is not `NEITHER`, request `url` once more and return that result.

- [ ] **Step 1: Write the failing tests** (helper: `make2(h2_handler, h1_handler)` → `Fetcher(transport=MockTransport(h2), h1_transport=MockTransport(h1), min_delay=0)` plus per-transport `seen` lists)

```python
def test_test_connection_keeps_http2_when_it_works():   # h2 "/"=200 -> HTTP2, h1 never called
def test_test_connection_falls_back_to_http1_when_http2_is_refused():
    # h2 "/"=403 challenge, h1 "/"=200 -> (Connection.HTTP1, page OK); later get("/p") goes to h1
def test_test_connection_neither_when_both_refused_and_pages_still_use_http2():
    # h2 "/"=403, h1 "/"=403, h2 "/p"=200 -> NEITHER; get("/p") is OK via h2
def test_test_connection_404_counts_as_working():          # h2 "/"=404 -> HTTP2
def test_test_connection_unreachable_keeps_previous():     # set HTTP1; h1 raises ConnectError -> stays HTTP1
def test_first_refusal_retests_once_and_refetches_page():
    # h2 "/p"=403, h2 "/"=403, h1 "/"=200, h1 "/p"=200 -> get("/p") OK; connection HTTP1
    # a second refused page later does not test again (h2/h1 "/" each requested once)
def test_robots_reloaded_after_switch_when_it_was_refused():
    # h2 robots=403, h2 "/"=403; h1 robots "Disallow: /p", h1 "/"=200 -> after test, get("/p") is DISALLOWED
def test_preset_connection_is_used_without_testing():      # set_connection(HTTP1); get("/p") -> only h1 sees requests
```

- [ ] **Step 2: Run** `.venv\Scripts\python -m pytest tests/test_fetch.py -q` — Expected: FAIL (`h1_transport` unknown).
- [ ] **Step 3: Implement** the interfaces above. `_Host` gains `connection`, `tested`, `robots_ok`.
- [ ] **Step 4: Run** `.venv\Scripts\python -m pytest -q` — Expected: all pass. Existing tests that return 403 now also trigger one test request to `/`; where a test asserts an exact request list, update the expectation to include it (the behaviour is intended), never weaken a refusal assertion.
- [ ] **Step 5: Commit** — `git commit -m "feat: per-site connection with one-off HTTP/2 to HTTP/1.1 test"`

### Task 3: Search check writes the note

**Files:**
- Modify: `pricetracker/discover.py` (`check_site`)
- Test: `tests/test_discover.py`

**Interfaces:**
- Consumes: `Competitor.search_note` (Task 1), `Page.status`.
- Produces: `check_site(fetcher: Fetcher, comp: Competitor, probe_term: str, home: Page | None = None) -> None` — uses `home` instead of fetching the home page when given. Sets `comp.search_note` to exactly one of:

| Situation | `search_note` |
|---|---|
| search reachable | `""` |
| home DISALLOWED | `"Home page forbidden by the site's rules (robots.txt)"` |
| home REFUSED | `f"Home page blocked by the site ({status})"` |
| home NOT_FOUND | `f"Home page not found ({status}) – check the website address"` |
| home UNREACHABLE | `"Site did not respond"` |
| no search form | `"No search box found on the home page"` |
| search DISALLOWED | `"Search forbidden by the site's rules (robots.txt)"` |
| search REFUSED | `f"Search blocked by the site ({status})"` |
| client address UNREACHABLE | `"Search did not respond"` |
| other search failure | `"Search did not return a usable page"` |

- [ ] **Step 1: Write the failing tests** — extend the existing `test_check_site_*` tests with `assert c.search_note == ...` for each row of the table (home 403, home 404 as `checked({})`, home 500 → `"Site did not respond"`, no form, robots-forbidden search, search 403, search 404, working search → `""`), plus:

```python
def test_check_site_uses_given_home_page_without_fetching_it():
    f, seen = fetcher(site({"/shop/search": RESULTS}))
    c = comp(company="Shop", website=HOME)
    check_site(f, c, "F00373000", home=Page(Outcome.OK, HOME, 200, FORM_Q))
    assert c.search_status == Access.REACHABLE and HOME not in seen
```

- [ ] **Step 2: Run** `.venv\Scripts\python -m pytest tests/test_discover.py -q` — Expected: FAIL.
- [ ] **Step 3: Implement** the table in `check_site`.
- [ ] **Step 4: Run** `.venv\Scripts\python -m pytest -q` — Expected: all pass.
- [ ] **Step 5: Commit** — `git commit -m "feat: plain-words search note per site"`

### Task 4: Scan — when to check a site, page note, saving

**Files:**
- Modify: `pricetracker/scan.py`
- Test: `tests/test_scan.py`

**Interfaces:**
- Consumes: Tasks 1–3.
- Produces:
  - `RECHECK_AFTER = timedelta(days=30)`
  - `site_check_due(group: list[Competitor], now: datetime) -> bool` — True when any row has `search_status is None` or `connection is None`, or (any row has search or page status `REFUSING`/`UNREACHABLE` or connection `NEITHER`) and (any row's `site_checked` is None or `now - site_checked > RECHECK_AFTER`).
  - `_note_page_status(comp: Competitor, page: Page) -> None` — as today for the status; additionally `page_note`: OK → `""`; when the status is (or becomes) `REFUSING`/`UNREACHABLE` from this page: DISALLOWED → `"forbidden by the site's rules (robots.txt)"`, REFUSED → `f"blocked by the site ({status})"`, UNREACHABLE → `"did not respond"`. (Joined as "Product pages: …" by Task 1.)
- `run_scan.work(host)`:
  1. `fetcher.set_connection(lead.website, <first non-empty connection in group>)`.
  2. If `site_check_due(group, run_at)` and the host has jobs: `conn, home = fetcher.test_connection(lead.website)`; then `check_site(fetcher, lead, probe, home=home)` (`check_site` ignores `home` when the row already has a search address); copy `search_status`, `search_address`, `search_note`, `connection = fetcher.connection(lead.website)`, `site_checked = run_at`, `last_checked = run_at` to every row of the group.
  3. After every lookup: copy `page_status`, `page_note` and `connection = fetcher.connection(lead.website)` to every row (so Ctrl+C keeps them).

- [ ] **Step 1: Write the failing tests**

```python
def test_site_check_due_rules():
    ok = comp(search_status=Access.REACHABLE, page_status=Access.REACHABLE)
    ok.connection, ok.site_checked = Connection.HTTP2, NOW - timedelta(days=60)
    assert not site_check_due([ok], NOW)                       # working sites are not re-checked
    bad = comp(search_status=Access.REFUSING)
    bad.connection, bad.site_checked = Connection.HTTP1, NOW - timedelta(days=31)
    assert site_check_due([bad], NOW)
    bad.site_checked = NOW - timedelta(days=29)
    assert not site_check_due([bad], NOW)
    new = comp(search_status=Access.REACHABLE)                  # file from before this change
    assert site_check_due([new], NOW)

def test_run_scan_records_connection_note_and_site_checked(tmp_path, make_xlsx):
    # one site, home page refused (403), product link page OK: after run_scan, Competitors.xlsx has
    # Connection "Neither", Status note starting "Search: Home page blocked by the site (403)",
    # Site checked == NOW, and the price was read from the link.

def test_ctrl_c_keeps_connection_found_so_far(tmp_path, make_xlsx):
    # as test_ctrl_c_still_writes_partial_run_file_and_links; Competitors.xlsx Connection is filled.

def test_page_note_follows_page_status():
    # do({K33: (403, "no")}, link(item_code="553000", url=K33)) -> co.page_note == "blocked by the site (403)"
    # then a 200 page on the same comp -> page_status REACHABLE and page_note == ""
```

- [ ] **Step 2: Run** `.venv\Scripts\python -m pytest tests/test_scan.py -q` — Expected: FAIL.
- [ ] **Step 3: Implement** the interfaces above.
- [ ] **Step 4: Run** `.venv\Scripts\python -m pytest -q` — Expected: all pass.
- [ ] **Step 5: Commit** — `git commit -m "feat: scan tests site connections, re-checks monthly and saves status notes"`

### Task 5: Documentation

**Files:**
- Modify: `README.md` (polite-visitor paragraph, line 16 and line 111), `docs/business-specification.md` (4.5, 4.6, 5.2 table), `docs/user-guide.md` (status table, columns), `docs/implementation-plan.md` line 24.

- [ ] **Step 1:** Add the three columns to the 5.2 table and the user guide; in 4.5 add the 30-day re-check and the Status note; in 4.6, README and implementation-plan line 24, replace "no retries on a refusal" with: the tool may test the other connection type (HTTP/2 or HTTP/1.1) once per site when checking it, records which works, and never retries a refused page otherwise.
- [ ] **Step 2: Run** `.venv\Scripts\python -m pytest -q` — Expected: all pass (docs only).
- [ ] **Step 3: Commit** — `git commit -m "docs: connection test, status note and monthly re-check"`
