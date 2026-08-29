# v2.x backlog (loop source of truth, ordered)

Rules per item: worktree agent (sonnet when the change is mechanical or fully specified, opus for
adapters/crypto/server/UI judgment), `git reset --hard <main sha>` first, TDD, `npm run check`
green, Map updated, terse commit; controller reviews small/mechanical diffs, opus reviews the
rest; patch bump per merged item (`npm version patch`, Cargo in lockstep, manifests regenerated);
deploy via `deploy.sh modelis` at each milestone marker; AUTORESUME + this file updated after
each item. Stop when Daniel intervenes or the list is empty.

## Hardening (v2.0.x)

- [x] B1 Split `src/excel/link-anchors.ts` (396 lines): move the relay round trip
      (`pushPayload`, `announce`, `publish` rollback) into `src/excel/link-record.ts`; tests unchanged.
- [x] B2 Dedupe helpers: `relativeTime` (ppt/views.ts + pane/links-tab.ts) -> `src/ui/time.ts`;
      clipboard copy (ui/toast.ts private + pane/links-tab.ts) -> `src/ui/clipboard.ts`; unit tests.
- [x] B3 "Update this slide" uses the ACTIVE slide: `activeSlideId()` in `src/ppt/host.ts`
      (`getSelectedSlides`, 1.5), fake host support, pane uses it (ticked rows no longer needed).
- [x] B4 Server minors: cache rule `starts_with("/assets/")`; split `inbox_lifecycle` test under
      50 lines; startup prints the data path (verify), `cargo clippy` clean.
- [x] MILESTONE v2.0.x deploy (v2.0.4, 29.08).

## Features (v2.1)

- [x] B5 Shapes inside groups: verify `Shape.group` / `ShapeGroup.shapes` availability on
      learn.microsoft.com (requirement set), scan recursively in `scanLinks`, fake-host group
      support, tests; document the floor.
- [x] B6 Brand settings per workbook: `workbook.settings["smt.brand.v1"]` with localStorage
      fallback (user-gap #6); pure settings codec already exists in `src/settings.ts`.
- [x] B7 Ribbon coverage: expose the top 12 Excel actions as ribbon buttons through
      `manifest/spec.ts` (data change + icons reuse), validate both manifests.
- [x] B8 Tornado chart + unpivot selection (docs/FEATURES.md section 11), pure logic + adapter,
      fake-host tests.
- [x] MILESTONE v2.1.0 deploy (29.08); AUTORESUME rewrite.

## Blocked on Daniel

- Spike (plan Task 18), real-Office pass, M365 upload, ROADMAP.md sync sign-off, native PowerPoint
  tables (needs spike + PowerPointApi 1.9 on the team's builds).

## Continuous improvement (after v2.1.0, loop keeps going until Daniel stops it)

- [x] C1 Whole-branch opus review of everything since v1.1.0 (adapters, relay, panes) -> ONE fix
      wave -> v2.1.1 deploy.
- [x] C2 Update-all speed: measure `scanLinks` + `status` on a 60-slide fake deck; batch tag loads
      across slides in one sync where the API allows; target under 3 syncs per update-all.
- [x] C3 Pane UX pass with a headless browser at 320/360/420 px: no clipped controls, badges
      readable, keyboard focus order sane; fix what is found.
- [x] C4 `=SMT.ROUND` research: GO (docs/research/custom-functions.md; shared runtime page reused).
- [x] C8 Implement `SMT.ROUND` / `SMT.ROUNDSUM` custom functions per the research: manifest
      CustomFunctions extension point in `manifest/spec.ts`, functions.json + functions entry via Vite,
      largest-remainder allocation with a range ceiling, ribbon/pane button that writes the formulas.
- [x] C5 FEATURES.md P2 leftovers, one per iteration: border cycles (done v2.1.3), row-height/column-width cycles (done v2.1.5),
      paintbrush slots (done v2.1.7), Super Find (done v2.1.9), unused-style scrubber (done v2.1.11), "prepare for sharing" (done v2.1.13).
- [x] C6 Test debt: split `test/ppt.integration.test.ts` (>400 lines) by feature; keep every
      test file under 400 lines.
- [x] C7 Docs: `docs/FEATURES.md` status column refreshed from what shipped (ROADMAP.md stays
      Daniel's; deltas listed for sign-off).
- [x] Auto-push on edit (spec v2.1 flag; done v2.1.10).
- Milestone after every 3 merged items: deploy + tag + AUTORESUME refresh.

## Next round (D, after v2.1.13; two agents at a time)

- [x] D1 Revert after Update all: server already keeps 2 revs; add `GET /api/links/{id}?rev=<n>` (auth as
      today), client `getLinkRev`, PowerPoint "Revert last update" for the selected rows (repaint the
      previous rev, tag rev rewritten), refuse when no previous rev; tests on both sides.
- [x] D2 Highlight linked cells in Excel (M3): Links tab toggle that paints a subtle brand-tinted fill on
      every anchored range and restores the original fills on toggle-off (snapshot like the audit overlay);
      chart anchors get a one-line hint instead.
- [x] D3 Batched relay fetch: `POST /api/links/fetch` with `[{id, auth, knownRev}]` returning the blobs
      that changed (base64url) in one response (4 MiB total cap, 413 above); client + PowerPoint update
      path use it; per-row fallback to GET stays.
- [x] D4 Super Find phase 4: comments (`worksheet.comments`, ExcelApi 1.10, gated) with a "comment" hit kind.
- [x] D5 Byte-budgeted update batches (flush a repaint batch at 8 MiB of payload).
- [ ] D6 Second whole-branch review (since v2.1.2) -> one fix wave -> deploy.
- [ ] D7 Change source / workbook version resolution (M3): PowerPoint "Change source" that re-keys a
      link to a newer export of the same anchor label from the inbox; design note first, then build.
- Milestone after every 3 merged items: deploy + tag + AUTORESUME refresh.
