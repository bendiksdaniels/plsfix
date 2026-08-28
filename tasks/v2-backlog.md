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
- [ ] B3 "Update this slide" uses the ACTIVE slide: `activeSlideId()` in `src/ppt/host.ts`
      (`getSelectedSlides`, 1.5), fake host support, pane uses it (ticked rows no longer needed).
- [ ] B4 Server minors: cache rule `starts_with("/assets/")`; split `inbox_lifecycle` test under
      50 lines; startup prints the data path (verify), `cargo clippy` clean.
- [ ] MILESTONE v2.0.x deploy.

## Features (v2.1)

- [ ] B5 Shapes inside groups: verify `Shape.group` / `ShapeGroup.shapes` availability on
      learn.microsoft.com (requirement set), scan recursively in `scanLinks`, fake-host group
      support, tests; document the floor.
- [ ] B6 Brand settings per workbook: `workbook.settings["smt.brand.v1"]` with localStorage
      fallback (user-gap #6); pure settings codec already exists in `src/settings.ts`.
- [ ] B7 Ribbon coverage: expose the top 12 Excel actions as ribbon buttons through
      `manifest/spec.ts` (data change + icons reuse), validate both manifests.
- [ ] B8 Tornado chart + unpivot selection (docs/FEATURES.md section 11), pure logic + adapter,
      fake-host tests.
- [ ] MILESTONE v2.1.0 deploy; AUTORESUME rewrite.

## Blocked on Daniel

- Spike (plan Task 18), real-Office pass, M365 upload, ROADMAP.md sync sign-off, native PowerPoint
  tables (needs spike + PowerPointApi 1.9 on the team's builds).
