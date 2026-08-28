# AUTORESUME - Model Tools (v2.0.0 LIVE 2026-08-29)

## State

v2.0.0 is deployed at dbautomatizacijas.com/modelis/ (`/version` -> 2.0.0): Excel pane with the
new **Links** tab, PowerPoint pane (`pptpane.html`), and the end-to-end encrypted link relay
(`/api/links`, `/api/inbox`, sqlite at `/opt/plsfix/data`). `npm run check` green on
main (334 vitest, 17 cargo, tsc, eslint, prettier, manifest and version gates); GitHub Actions
runs the same gate. Tag `v2.0.0`.

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
3. M365 centralized deployment: upload `manifest.prod.xml` (now two hosts) in the admin
   center (steps in `tasks/v1-plan.md`); the v1 Excel-only upload never happened, so this is
   the first upload. JS-only updates afterwards need no admin action.
4. v1 visual pass in Excel is still unconfirmed (calendar 28.08 19:00).

## Known limits / deferred (from reviews)

- Shapes inside groups are not scanned; "Update this slide" uses the first ticked row's slide;
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
