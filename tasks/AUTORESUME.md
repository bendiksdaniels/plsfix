# AUTORESUME - pls,fix (v2.3.6 LIVE 2026-08-30)

## 30.08 late: v2.3.6 (values-only chart labels, Templates on top), WP3 native charts next

- WP3 in flight (30.08 late): spec `docs/superpowers/specs/2026-08-30-native-charts-design.md` (spike findings
  inside: pie adjustments = degrees clockwise from 3 o'clock on (-180,180], groups persist with tags
  across reloads, web adds cost more as the slide fills, text boxes ~1 s each on the web, ExcelApi
  1.12/1.15 reads work on the web), plan `docs/superpowers/plans/2026-08-30-native-charts.md` (8
  tasks). Done on main: Task 1 (`src/link/chart-model.ts`, `src/chart-colors.ts`, 83d4a86) and
  Task 3 (payload `chart?`, 5ee404f). Wave 1 dispatched off 5ee404f: Task 2 pure layout (sonnet),
  Task 4 Excel reader (opus), Task 5 fake PowerPoint (opus); then Task 6 adapter (opus), Task 7
  docs (sonnet), Task 8 controller (gate, bump minor v2.4.0, deploy, web proof via the CDP rig).
  Rig: scratch Chrome 9222 + manifest server 3001 still alive; spike scripts in this session's
  scratchpad (`spike-*.js`); the PowerPoint deck is "Presentation 1" (4 slides, slide 3 empty).
- Daniel: "charts always just the values", "I do not see the templates on the side bar", "try
  moving to PP pie charts and graphs". Plan approved (budget 2.3-4.0M tokens):
  ~/.claude/plans/velvet-munching-duckling.md. WP1 done: `src/chart-labels.ts` (label position by
  chart type) + `styleChartLabels` in `src/excel/internal.ts`, applied by the waterfall (in the
  tolerated surface batch), the tornado (helper block in the outcomes' number format) and Format
  chart (a pie keeps its legend for the categories, leader lines on); demo charts `show_value()`;
  the fake host records every label flag. WP2 done: the Templates section moved under the
  Selection inspector as a 2x3 `tool-grid template-grid` (descriptions in the button titles and
  the "?" card); ux:check 0 defects at 320-500; manual text updated, but
  `manual/shots/excel-tools.png` still shows the pre-templates top: retake on the web.
- WP3 (native PowerPoint chart groups): spec + real-host spike first. PowerPointApi 1.10
  `adjustments` (Daniel's 16.107 has it) is the pie route; undrawable types fall back to the
  picture link. Not started.

## 30.08: fleet 2 (Daniel: overlap still seen in Excel; intuitive, advanced, universal, templates, "?")

- 30.08 merged + live v2.3.5: C universal tools + data-aware chart placement
  (`src/excel/chart-place.ts`, `areas.ts` multi-area, `protection.ts`; 881 tests). Fleet 2 done.
  Incident: an agent symlinked `~/plsfix/node_modules` to itself (lessons.md 30.08).
- 30.08 merged + live: B help "?" (v2.3.3, `src/help/copy*.ts`, `src/ui/help.ts`, copy gate over
  both panes) and A templates (v2.3.4, `src/templates.ts` + `src/template-blocks.ts` +
  `src/template-cells.ts` + `src/excel/templates.ts`, Templates section at the end of Tools, manual
  chapter "Veidnes"). First launch of fleet 2 went into the wrong repo (worktree = shell cwd;
  lessons.md 30.08). C (universal tools + data-aware chart placement) still running.
- Three opus worktree agents off main 0378308, budget approved (2-3M): A templates (`src/templates.ts`
  + `src/excel/templates.ts`, Templates section at the END of the Tools tab, manual chapter
  "Veidnes"; report templates-report.md), B help "?" per section in both panes (`src/help/copy.ts`,
  `src/ui/help.ts`, `src/styles/help.css`; help-report.md), C universal tools + chart placement
  that avoids data cells via getUsedRangeOrNullObject candidates (universal-report.md). Merge
  order: C, A, B (A and B both touch taskpane.html/main.ts; B is attributes + install calls);
  gate + ux:check, bump, deploy, `npm run manual`, web check of placement and templates.

## 29.08 late: charts, placement, picker, native tables in flight (v2.2.6)

- Daniel: "actual table not an image", "try charts and pie charts", "data would never overlap
  ... in the PP and excel", "the UI of the tab is not perfect yet". Done solo: `src/layout.ts`
  (pure placement: `placeBeside` for Excel charts, `placeInFreeSpace` for slides); waterfall and
  tornado charts land beside/below their block and never on another chart; "Export active chart"
  falls back to the sheet's only chart or a picker (`#export-chart-pick`, refreshed on tab
  open, sheet activation and the debounced selection change because `worksheets.onActivated`
  never fires on Excel for the web); demo pie on Rounding; Brand preview table fixed layout so
  the pane never overflows (the tab strip was clipped); cool tints; chart-list select styled.
- Fleet 29.08 late (Daniel: "send more agents", budget approved for the UI audit): three opus
  worktree agents off main 6adbc79: (1) native tables (report scratchpad/native-tables-report.md),
  (2) Latvian Word manual `manual/` crate + `manual/out/pls,fix rokasgrāmata.docx`
  (manual-report.md), (3) pane UI audit at 320/360/420/500 px with ux:check overflow gates
  (ui-audit-report.md). Merge order: tables, UI audit (expect taskpane.html/styles.css
  conflicts with tables), manual; gate, bump, deploy after each; backlog agents (Links list
  refresh on sheet edits, chart list label for linked charts, web undo error) only after that.
- MERGED 29.08 night: the UI audit (branch worktree-agent-a28952346a7af3847: 20 defect classes
  fixed, styles split under src/styles/, `npm run ux:check` gates 4 widths x both panes x 5
  states) and the Latvian manual (branch worktree-agent-a8d20f80ecaad887b: `manual/` crate,
  `npm run manual`, `manual/out/pls,fix rokasgrāmata.docx`, screenshots retaken on v2.3.2 from a
  fresh demo upload). Table column widths must be whole points on PowerPoint for the web
  (v2.3.2); a lone object is centred on its slide; the font size always travels.
- MERGED 29.08 night: native tables (branch worktree-agent-a3e01031c747aceb7, 7 commits, merge
  feade84) + numeric right-alignment; v2.3.0 deployed. `src/ppt/tables.ts`, `src/ppt/placement.ts`,
  `src/ppt/picture.ts`, `src/excel/link-table.ts`, fake `test/fakeppt/tables.ts`. Proven on PowerPoint for
  the web 29.08: insert as a native 8x7 table in free space, in-place update (rev 4, geometry
  kept), rebuild as 9x7 at the same corner after a row was inserted into the source (rev 5).
- (was) In flight: opus implementer in a worktree building native PowerPoint tables from
  `docs/superpowers/specs/2026-08-29-native-tables-design.md` (kind `table`, addTable on
  PowerPointApi 1.8, in-place refresh, re-create on size change, free-space placement for
  pictures too); report at scratchpad/native-tables-report.md; merge = review diff, gate, bump.
- Process slip 29.08: v2.2.5 was deployed while five pane tests were red (mock gaps only, fixed
  in 2b18d79); chains now gate the deploy on the check exit code.

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
