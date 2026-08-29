# AUTORESUME - pls,fix (v2.2.3 LIVE 2026-08-29)

## 29.08 late: house number styles per language (v2.2.3)

- Daniel: "take the number styles from the house design for all 3 languages and we always use '.'
  as comma for numbers" -> scope pls,fix (Excel), "." is the decimal everywhere. No written spec
  exists (DESIGN-SYSTEM.md is typography); the house style came from teaser2mail `fmt.rs money`:
  `15 000 000 EUR` (LV, RU), `EUR 15 000 000` (EN). Daniel chose comma grouping for EN.
- Built: `src/numbers.ts` (pure: grouping space/comma, decimal point, currency after/before,
  `formatAmount` for every number the pane writes as text), `language` in BrandSettings (default
  `lv`), currency cycles/buttons place the symbol per language, `src/excel/separators.ts` reads
  Excel's own separators (ExcelApi 1.11) and the Brand tab says what Excel shows and where to
  change it (separators are an Excel setting, format codes cannot force them). Fake host gained
  `application` + `separators` option and a trailing-symbol currency rewrite.
- PowerPoint re-verified on the web after the rebrand (v2.2.3): new key, export, pair, Inbox ->
  Insert (tags `PLSFIX_LINK`/`PLSFIX_KEY`), move + edit + Push all, Update all (rev 2, height
  refit to the kept width because the picture's aspect changed), Revert (rev 1), Break (tags
  gone). A shape still carrying old `SMT_*` tags is ignored, as intended.

## 29.08 night: rebrand to pls,fix (v2.2.0)

- Daniel: drop the old brand, make it his own, name with an IB joke: **pls,fix** (the comma is the
  joke). Colours navy & mint (`#14213D` / `#2EC4B6`, Daniel's tweak of the first slate/teal pick), brand mark a comma, new icons. Custom
  functions `=PLSFIX.ROUND` / `=PLSFIX.ROUNDSUM`; tags `PLSFIX_LINK` / `PLSFIX_KEY`, hidden names
  `PLSFIX_LINK_<id8>`, settings keys `PLSFIX_*` / `plsfix.*`, relay header `X-PLSFIX-Link-Id`,
  HKDF labels `plsfix-*` (no shim: nothing was in production). PowerPoint ribbon group is
  "Links" under the "pls,fix" tab; manifest provider "Daniels Bendiks".
- Repo moved: `~/plsfix` (Desktop symlink `plsfix`), GitHub `bendiksdaniels/plsfix`; gateway
  `deploy.sh` MODELIS path, `workspace/tools.toml` and `tools.json` updated. Hosting identifiers
  unchanged on purpose: key `modelis`, `/modelis/`, binary + unit `plsfix-server`, `MODELIS_*`.
- Desktop Excel/PowerPoint still run the old manifest until quit and reopened (`npm run demo`
  restores the new one into wef); the demo file is `demo/out/pls,fix Demo Model.xlsx`.

## 29.08 evening: demo delivered, v2.1.22 LIVE, web pass done

- Daniel: "launch it now and create a demo where I could play around with it"; then no desktop
  control ("I can not give you computer use", "use a virtual machine or something on the
  server"). Delivered: `demo/` (Rust) builds `demo/out/pls,fix Demo Model.xlsx` (8 sheets,
  435 cells, reconciled by `demo/tests/workbook.rs`, "Start here" checklist of 17 rows);
  `npm run demo` = build + prod manifest into both wef folders + Excel on the workbook +
  PowerPoint on a new deck. Excel and PowerPoint on the Mac were relaunched with the current
  manifest (the Excel wef file had been the old 1.0.0.0 one). Desktop pass itself = Daniel.
- Web verification rig (see tasks/lessons.md 29.08): scratch Chrome on port 9222 driven by
  `scratchpad/driver/drive.mjs` (playwright-core over CDP), Office on the web with the add-in
  registered via `wdaddindevserverport=3001&wdaddinmanifestfile=manifest.prod.xml&
  wdaddinmanifestguid=<id>` from `serve-manifest.mjs` (HTTPS + CORS on the dev certs). Daniel
  signed in with his NDUS account himself. Both hosts show the add-in's tab; both panes load from
  production.
- Verified on Office for the web (Excel): selection inspector (160 cells / 91 formulas),
  Autocolor (links green, inputs blue, F9 partial), Audit overlay on/off, Fill formula right
  (needs neighbouring rows), CAGR, Waterfall, Tornado, Consistent rounding (formulas written;
  values #NAME? because the web custom-functions runtime fails to start), Unpivot, Find, Scan
  broken names (Old_budget), Scan styles, Prepare for sharing, Contents sheet, Precedents /
  Dependents, IFERROR, sign flip, x1000, number cycle, pls,fix Undo. Links: key generated, range
  exported, PowerPoint paired by key, Inbox -> Insert (tagged shape), move/resize, edit + Push
  all, Update available -> Update all (rev 2, geometry kept), Revert (rev 1), Break link (tags
  gone, picture stays). Not automatable on the web: chart export (charts are canvas-drawn,
  `activate()` does not select on the web), custom function values.
- Two fixes found by the web pass, both live: v2.1.21 `src/host-ready.ts` (Office.onReady never
  settles when the custom-functions runtime fails to initialise; the pane now boots off a host
  probe after a 4 s head start, degraded toast); v2.1.22 waterfall applies its chart surface
  (font, corners) in a tolerated batch because Excel for the web rejects both on chartex
  charts. Demo fixes: per-month row between filled rows, SMT hints without "=", source-missing
  row asks for the whole block (deleting rows inside the block only shrinks the anchor).
- Commits 2cd27b1, 4c8c762, d0e903e, 90fba5e (v2.1.22). Not pushed. Backlog candidates from the
  pass: Excel Links list does not re-render on sheet changes (status only refreshes on push);
  chart export on the web needs a UI click.

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
(GO). Then: pane UX pass (0 defects at 320/360/420 px, `npm run ux:check`), `=PLSFIX.ROUND`/`=PLSFIX.ROUNDSUM`
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
   `CustomFunctions` extension point that publishes `=PLSFIX.ROUND` / `=PLSFIX.ROUNDSUM`) in the
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
