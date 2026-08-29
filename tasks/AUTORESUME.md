# AUTORESUME - Model Tools (v2.1.20 LIVE 2026-08-29)

## 29.08 evening: demo on the Mac (uncommitted), web pass pending

- Daniel: "launch it now and create a demo where I could play around with it". Done on the Mac:
  `demo/` (Rust, rust_xlsxwriter) builds `demo/out/Demo Model.xlsx` (8 sheets, 435 cells,
  reconciled by `demo/tests/workbook.rs`); `npm run demo` = build + prod manifest into both wef
  folders + open Excel on the workbook + PowerPoint on a new deck. `npm run check` green with
  `test:demo` added. Excel had the OLD 1.0.0.0 manifest until now; both wef files are v2.1.20.
  Nothing committed yet (Daniel commits on request): `git status` shows demo/, scripts/demo-open.sh,
  package.json, .gitignore, CLAUDE.md.
- No desktop control: Daniel will not grant computer-use (macOS perms), so the ribbon on his Mac is
  unverified by me; the "Start here" sheet is his checklist. Verification route instead: Office for
  the web in a CDP Chrome (`scratchpad/driver/drive.mjs`, port 9222, scratch profile) with the
  manifest sideloaded by upload; blocked on Daniel signing in to Microsoft in that window.

## State

v2.1.20 is deployed at dbautomatizacijas.com/modelis/ (`/version` -> 2.1.20): Excel pane with the
new **Links** tab, PowerPoint pane (`pptpane.html`), and the end-to-end encrypted link relay
(`/api/links`, `/api/inbox`, sqlite at `/opt/plsfix/data`). `npm run check` green on
main (700 vitest, 31 cargo, tsc, eslint, prettier, manifest and version gates); GitHub Actions
runs the same gate. Tags `v2.0.0`, `v2.0.4`, `v2.1.0`, `v2.1.2`, `v2.1.5`, `v2.1.8`, `v2.1.11`, `v2.1.16`, `v2.1.19`, `v2.1.20`. Since v2.0.0 (loop, `tasks/v2-backlog.md`):
link-record split, shared time/clipboard helpers, active-slide update, server minors, links inside
groups (PowerPointApi 1.8), brand palette saved in the workbook, 12 more ribbon buttons (5 groups),
tornado chart + unpivot selection; then the whole-branch review fix wave (v2.1.2: stale-rev
status, ETag parsing, inbox squatting, anchor rollback, revoke-first remove, masked key, guards),
update-all in one sync, border + row/column size cycles, Latvian user guide
(`docs/lietotaja-rokasgramata-saites.md`), FEATURES.md status refresh, custom-functions research
(GO). Then: pane UX pass (0 defects at 320/360/420 px, `npm run ux:check`), `=SMT.ROUND`/`=SMT.ROUNDSUM`
custom functions (manifest CustomFunctions extension point, `functions.js`; needs the M365 re-upload),
paintbrush slots, Super Find, unused-style scrubber, auto-push on edit (Links tab checkbox).
Round D shipped: revert after Update all (`?rev=` route), highlight linked cells, batched relay fetch
(`POST /api/links/fetch`), comments in Super Find, byte-budgeted repaint batches, second whole-branch
review + fix wave (registry lock, multi-area edits, chunked status, error bodies, dev functions.js,
`npm run validate` in `check`), change source (`docs/superpowers/specs/2026-08-29-change-source.md`).
The backlog (`tasks/v2-backlog.md`) is empty: the loop stopped here on 29.08; restart it with /loop after
adding items. Your gates below are the next step.

Design: `docs/superpowers/specs/2026-08-28-ppt-links-design.md`. Plans (all tasks done):
`docs/superpowers/plans/2026-08-28-tier1-groundwork.md`, `2026-08-28-ppt-links.md`.
Execution ledgers (git-ignored, rulings + per-task spend): `.superpowers/sdd/*/progress.md`.

## Open gates (Daniel)

1. Real-Office pass (spec section 11, ~20 min): export a range + a chart from a model, insert
   in a deck from the PowerPoint Inbox, move/resize, insert rows above the source, change
   numbers, Push, Update all -> pictures refresh in place; delete the source rows -> "source
   missing"; wrong link key -> empty inbox + clear toast; deck without the add-in -> plain
   pictures. Dev loop: `npm start` (Excel) and `npm run start:ppt` (PowerPoint) against a local
   relay (`MODELIS_DATA=./data cargo run --manifest-path server/Cargo.toml`; vite proxies
   `/api`), or straight against production with the prod manifests in wef.
2. The spike (plan Task 18, ~1 h): tags survive save/reopen, cut/paste, duplicate slide, copy
   to another deck; `Range.getImage` orientation + pixel density on Mac; `OfficeRuntime.storage`
   shared across hosts (would make pairing automatic); team PowerPoint builds vs PowerPointApi
   1.8. Findings go to `docs/research/officejs-feasibility.md`.
3. M365 centralized deployment: upload `manifest.prod.xml` (now two hosts, and since C8 the
   `CustomFunctions` extension point that publishes `=SMT.ROUND` / `=SMT.ROUNDSUM`) in the
   admin center (steps in `tasks/v1-plan.md`); the v1 Excel-only upload never happened, so
   this is the first upload. JS-only updates afterwards need no admin action, except that
   Office caches `functions.js` / `functions.json` separately: per Microsoft those two can
   take up to 24 hours to reach users (docs/research/custom-functions.md, section 5).
4. v1 visual pass in Excel is still unconfirmed (calendar 28.08 19:00).

## Known limits / deferred (from reviews)

- Shapes inside groups are scanned three levels deep, and only on PowerPointApi 1.8
  (PowerPoint 2504 / 16.96): deeper nesting is ignored, and below 1.8 a grouped row
  fails rather than reinsert the picture onto the slide, out of its group.
  Source column ellipsised under ~420 px; no revert of a refresh (relay keeps 2 revs, UI only).
- `src/excel/link-anchors.ts` is at 396/400 lines: split before the next change
  (`link-record.ts` for the relay round trip). `src/main.ts` remains oversized (pre-existing).
- Pairing is a pasted link key; Entra SSO (Milestone 4) can replace it behind `KeyStore`.
- Native PowerPoint tables (`kind: "table"`), change-source, auto-push on edit: v2.1
  candidates (spec section 1).

## Gotchas

- Manifests are generated (`npm run manifest:build`); `npm stop`/`stop:ppt` restore the prod
  manifest into wef. The wef file is a plain copy (no hard link) since 28.08.
- `typescript` is aliased to the TS 6 shim for typescript-eslint; `@typescript/native` is TS 7.
- The gateway's deploy rsync excludes `/data` (relay sqlite); `install -d` creates it.
- Deploy audit shows pre-existing drift in OTHER tools (r2e engine, teaser HANDOFF), not modelis.
