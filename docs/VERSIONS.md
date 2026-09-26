# Versions

One entry per reviewed merge of the hunt loop (26.09.2026 on), newest first: what changed for a
user, the proof, and what is still open. The GitHub Release of each tag lists the commits.

## v2.9.3 (27.09.2026): charts on Excel for the web, and at the sheet's edge

- Fixed: Tornado and Football field on Excel for the web. The chart styling Excel for the web
  refuses (the chart-wide font and rounded corners) ran before the chart was placed, so the
  refusal left the chart over the data and put Office's raw sentence in the toast. The chart is
  now placed first and that styling runs in its own tolerated batch, silently, like the
  waterfall's; Chart Smart Format and the waterfall tolerate the same two refusal codes.
- Fixed: a chart made from a selection against the sheet's right edge found no free spot and was
  left over the cells just selected; it now lands below them.
- Tests: 19 chart tests in `test/hunt/charts.*`: the refusal under both codes, three presses in a
  row, two chart tools racing, the sheet edge, an empty sheet.
- Proof: vitest 3070 passed; the web-refusal pins fail on the old tornado code, the double-press
  pin fails with the press queue removed (both shown in review).
- Open: ledger item 9 stays open until the web rig reads the real error code on a tornado.

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
