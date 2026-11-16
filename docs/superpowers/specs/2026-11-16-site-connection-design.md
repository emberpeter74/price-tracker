# Site connection and status notes – design

Date: 2026-11-16. Status: approved in conversation.

## Why

Tests on 2026-11-16 showed Cloudflare refusing the tool's HTTP/2 requests to
Tanks for Everything but serving the same honest request over HTTP/1.1. A site could
just as well accept HTTP/2 and refuse HTTP/1.1. The client's rule: **remember what
works; when it stops working, try the other way.** The findings per site (why a site
refuses, a wrong website address) should be kept in `Competitors.xlsx`, not lost in a
chat.

## Goals

- Find, per site, which connection type works, remember it, and use only that one.
- Re-test a site's connection when the remembered one stops working.
- Record in plain words why a site is not usable.
- Re-check refusing or unreachable sites every 30 days instead of never.

## Non-goals and anti-goals

- No browser impersonation, header spoofing, challenge solving or proxies. The user
  agent string stays `PriceTrackerPOC/0.1 (price comparison tool)`.
- No retry of every refused page. The extra requests are bounded: one connection test per
  site per scan at most.
- `robots.txt` is obeyed first: a forbidden page is never requested by either connection.

## Approaches considered

- **A. Test once per site, remember (chosen).** Extra requests only when a site is
  tested.
- B. Fall back on every request. Simplest, but a blocked request plus a retry on every
  page of a refusing site, every run. Rejected: load, and it looks like evasion.
- C. Separate "test sites" command. Rejected: one more thing to remember to run, and new
  competitors would go untested.

## Design

### New columns in `Competitors.xlsx` (filled by the tool)

| Column | Values |
|---|---|
| Connection | `HTTP/2`, `HTTP/1.1`, `Neither`. Empty = not tested yet. |
| Status note | Plain-words reason, e.g. "Search forbidden by the site's rules (robots.txt)", "Blocked by the site (403)", "Home page not found (404) – check the website address", "No search box found on the home page". Empty when everything works. |
| Site checked | Date and time of the last connection and search check. |

`Last checked` keeps its meaning (last time the tool visited the site in a scan), so it
cannot drive the 30-day rule; that is why `Site checked` is separate. *(Builder's call.)*

### Connection test (one site)

1. Load the home page over HTTP/2.
2. If refused (401/403/429 or a challenge page), load it once over HTTP/1.1.
3. Record the first that works; if both are refused, record `Neither`.
   Not found / unreachable are not refusals: the connection stays as it was and the note
   says what happened.
4. `robots.txt` is read through the connection that works. If the first connection's
   `robots.txt` was refused and the other one works, `robots.txt` is read again through it.
   *(Builder's call: today a refused `robots.txt` counts as "everything allowed", which
   must not survive a switch to a connection that can read the real rules.)*

### When a site is tested

- **Connection empty:** at the site's first visit in a scan, before the search check.
- **30-day re-check:** when Search status or Product page status is `Refusing` or
  `Unreachable`, or Connection is `Neither`, and `Site checked` is empty or more than 30
  days old: test the connection again and re-run the search check.
- **Remembered connection stops working:** the first time in a scan that a request to the
  site is refused, the tool tests the site's connection once, straight away. If the other
  connection works, it is recorded and the refused page is fetched again with it
  *(builder's call)*. At most once per site per scan.
- `Neither` sites still have their pages tried with HTTP/2, because some sites refuse the
  home page but serve product pages (see the business specification, 4.5).

### Status note contents

Written by the tool whenever it sets a site's statuses; summarises the search and product
page findings in one cell, e.g. "Search forbidden by the site's rules (robots.txt).
Product pages readable." The tool distinguishes:

- forbidden by `robots.txt` (DISALLOWED),
- refused by the site, with the status code (REFUSED),
- page not found, with the status code,
- did not respond,
- no search box found.

### Client columns

Search status may still be overwritten by the client, as today. The tool never writes to
client columns.

## Documentation to update

README ("polite visitor" paragraph), `docs/business-specification.md` (4.5, 4.6, 5.2),
`docs/user-guide.md` (status table and columns), and the "no retries on a refusal" line in
`docs/implementation-plan.md`: the tool may test the other connection type once per site
when checking it, and never retries otherwise.

## Left to the builder

Where the per-host connection lives (expected: the `Fetcher`, which already owns per-host
state), exact wording of notes, and test layout.
