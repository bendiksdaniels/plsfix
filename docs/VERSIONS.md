# Versions

One entry per reviewed merge of the hunt loop (26.09.2026 on), newest first: what changed for a
user, the proof, and what is still open. The GitHub Release of each tag lists the commits.

## v2.9.2 (27.09.2026): the relay's revision overflow, 19 relay tests

- Fixed: a push to a link whose revision counter sits at the top of its range (only a restored or
  hand-seeded row gets there) crashed the request in a debug build, and in the release build that
  production runs it wrapped the revision below zero, dropped the push and still answered 200, so
  Excel counted it as pushed. It is now refused with 507 and the store keeps what it had.
- Gate: `npm run check` is green again on main (it was red since the video merges: `video/` now
  sits outside the pane's lint and format, with its own `cargo test`).
- Tests: 19 new relay tests in `server/tests/hunt_*.rs`: link DELETE over HTTP, 32 writers racing
  on one link, the exact TTL boundary second, empty bodies, Content-Type, key length,
  double-encoded static paths, the production write allowance with its Retry-After.
- Proof: cargo test 117 passed (98 before); the overflow pin also passes in a release build; the
  race and rate tests ran 100 times each with no failure (the first cut of the rate test failed
  5 in 500 at a second boundary, caught in review).
- Open: Excel's push path shows the raw relay line for a 429 or a 507 instead of a sentence
  (being fixed in the links hunt).
