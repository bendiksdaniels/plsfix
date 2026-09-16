# AUTORESUME - pls,fix (v2.8.17 tagged 16.09, PUBLIC on GitHub, MIT)

## NEXT SESSION (from 16.09 night): the matrix twice, Windows, the web rig, the folders on his Mac

State at hand-over: **v2.8.19 LIVE 16.09** (main 4acb030+ = origin, live `/version` 2.8.19, release runs
green for v2.8.10 to v2.8.19, manual docx v2.8.019, memory `project_plsfix.md` current), gates green
(2824 vitest, cargo green, ux 0/72, sweep 0/151), no worktree, no branch, no dev server, both wef folders
hold the prod manifest. v2.8.18 is a version-only tag (the 16.09 night section says why). Read
`tasks/lessons.md` 16.09 first, both entries.

1. **Crash matrix, twice per route** (the 14.09 brief, still not run as a matrix): today's 12 inserts, Update
   all, Revert and two pane reloads never crashed PowerPoint 16.107, but no route was run twice and Update
   this slide, Change source and Break link on charts were not driven. The recipe needs no computer-use
   grant: `screencapture` + the Swift CGEvent tool (lessons 16.09); keep the process start time in view.
2. **Slide 1 text-extent proof**: `text-extent.ts` is fake-proven only; on the Mac slide 1 was full (the Big40
   group), so the "chart lands below a short list" case still wants a slide with text only.
3. **Windows and the web rig** are untouched by everything since v2.8.9: `tasks/launch-check.md`
   (Daniel's own Windows pass) and `scripts/rig/` (the web pass) both need a run before the store listing.
4. **Follow-ups seen today, not fixed**: the PowerPoint Inbox needed one manual refresh to show fresh
   exports (pane opened seconds after the pushes); a pie at a quarter-size spot still wraps nothing but
   is small (cosmetic); the Big40 chart on slide 1 keeps its old stacked labels until re-inserted (a
   repaint only redraws on a new revision, and its source chart is gone from the rebuilt workbook).
5. Daniel's gates unchanged: Windows pass, M365 centralized upload (the manifest changed at v2.7.5, not
   since), Cloudflare rate-limit rule on `/modelis/api/*`, the `check` workflow re-enable.
6. **The folders on his Mac**: v2.8.19's deck list (a folder per project, the workbook under every object) is
   proven in the ux gate and the jsdom suites only; the sideloaded pane reloads it from the live server
   (pane reload, or Home > Add-ins > pls,fix once per launch). His computer-use grant or his own look.

## 16.09 night: folders in the deck's Linked objects list (v2.8.19; v2.8.18 is a version-only tag)

- Daniel: "create a folder for linking objects so the IB department could know from which excel each link
  comes and what project does it belong." Projects (= folders, v2.8.6) existed, but the deck's Linked
  objects table was flat, showed the project only in a filter and hid its Source column below 620 px (Office
  gives a pane 320-500). Plan `~/.claude-accounts/work/plans/parsed-frolicking-knuth.md`, approved.
- Shipped (one sonnet worktree agent, 301k, Fable review; gates: check 2824 vitest, ux 0/72, sweep 0/151):
  `groupByProject` in `src/link/project.ts` (named projects in locale order, No project last), used by the
  Excel list, the Inbox and the deck list; `renderLinkRows` in `src/ppt/views.ts` writes a `link-folder`
  header row per project ("Amasty · 3", no checkbox, no `data-key`) and a `link-meta` line "workbook · kind"
  under every object; the Source `<th>` and the 620 px block are gone (five columns); the audits and the
  stress support select `#link-rows tr[data-key]`; the ux fixture rows carry `kind` + `project` (three
  folders in the screenshots); help copy, README, FEATURES, the manual's Projekti section (LV), a
  launch-check row.
- Review rulings: the fourth ux fixture row (k4, the only Wrong link key badge) restored as a Balcia row;
  the sweep's taskpane 404 was `public/functions.js` missing in the fresh worktree (`npm run functions:js`,
  now in the Map); the help sentence trimmed to 139 chars for `HELP_MAX_CHARS`.
- Release slip: the chained merge + release command used `git merge -F -` (reads no stdin); the merge
  failed behind a `tail` pipe and `release.sh patch` still bumped, tagged and auto-pushed v2.8.18 =
  v2.8.17 + bump. Left in place (no history rewrite on the public repo; its release note says so); the
  branch was then merged properly, v2.8.19 released, manual v2.8.019, `deploy.sh modelis` -> modelis OK +
  clean, live `/version` 2.8.19, both release runs green. Lesson in `tasks/lessons.md`.
- The deploy audit's exit 1 came from other components (the tulkojums HOLD and nonbank-review engine
  copies behind the Mac), not modelis.

## 16.09: "most of the features feel sloppy" - the desktop sweep and its seven slices (v2.8.10 to v2.8.14, more to come)

- Session "resume work on this" (plan `~/.claude-accounts/work/plans/jiggly-tickling-raccoon.md`, approved). Daniel
  at 12:00: "Ask for computer use because still most of the features feel sloppy"; at 12:30 the computer-use
  bridge lost its Screen Recording permission mid-run and he could not restart his terminals to re-grant it, so
  the rest of the pass ran WITHOUT the grant: `screencapture` (the Terminal already has Screen Recording) plus a
  40-line Swift CGEvent tool for clicks, scrolls and keys (scratchpad `desk/ui.sh` + `click.swift`; recipe in
  `tasks/lessons.md` 16.09). Excel and PowerPoint 16.107, the demo workbook rebuilt by the demo crate, a scratch
  copy of the demo deck; PowerPoint never crashed (process start time unchanged through 12 inserts, Update all,
  Revert, and two pane reloads).
- Shipped and live, one slice per patch, each a sonnet worktree agent reviewed by Fable, gates green on the rebased
  branch, `deploy.sh modelis` clean after each:
  - v2.8.10 excel-tools: the audit overlay marks a typed number in a formula row (`typed`, solid tint) and compares
    formulas with the nearest formula across, so H19 is no longer striped beside G19; Fill formula right/down
    carries the source's number format; ribbon Precedents/Dependents select every same-sheet area at once;
    Autocolor no longer paints growth formulas purple (0 and 1 are identity constants).
  - v2.8.11 table-header: `TablePayload.h` (first row all bold), `Table.styleSettings.isFirstRowHighlighted`,
    the `PLSFIX_PAINT` tag so a repaint clears only the fills pls,fix painted (`src/link/paint-map.ts`,
    `src/ppt/table-style.ts`). On the Mac the band still did not show: see v2.8.15 below.
  - v2.8.12 excel-polish: Autocolor leaves text labels alone (`text` class); the Links list's second line says
    "Table · P&L!A10:H24" instead of the internal id, a `<colgroup>` fixes the fat tick column, the Project select
    sits on its own row; the demo gains a `CAGR` sheet (Daniel: "Create a sheet where I could test the CAGR?").
  - v2.8.13 table-export: a table is created with `uniformCellProperties.font.size` = the payload's modal size
    (rows were double height until the first Update all); "12,400" is rewritten to Excel's own separators
    ("12 400", `localizeNumberText`); an export made with the audit overlay on ships the model's fill, not the
    tint (`originalFillColor`).
  - v2.8.14 chart-place: chart labels are single-line (`wordWrap` false, AutoSizeNone) with a wider estimate
    (`LABEL_PAD` 15, `CHAR_WIDTH` 0.6); a dense column chart drops its value labels and keeps every k-th category
    label; Free space never shrinks below the caller's floor (a chart `minPlacementScale`, a picture half, a
    text link its natural size) and the overlap note names the largest free spot; a placeholder's occupied box
    is cut to the lines its text holds (`text-extent.ts`), so a chart lands below a short list.
- Diagnosed on the Mac with dev sideloads (the recipe: worktree vite on :3000 with the `/api` proxy pointed at the
  live relay, the dev manifest copied into the wef folder, the app relaunched, the add-in loaded from Add-ins):
  - pls,fix Undo after x1000 toasted a raw InvalidArgument: Excel for Mac returns an unfilled cell's fill as
    `{pattern: null, patternColor: "", color: "#FFFFFF"}` and `setCellProperties` refuses it verbatim (group-by-group
    diagnostic, only the fill group refused; `{pattern: "None"}` accepted). Slice undo-fill (`settableProperties`
    in `src/excel/undo.ts`).
  - The table header band: a table created through `shapes.addTable` has NO style (`<a:tblPr/>` in the saved
    XML) and every `styleSettings.load` on it is refused with GeneralException until a style is written blind;
    after `style = MediumStyle2Accent1`, `isFirstRowHighlighted = true`, `areRowsBanded = false` the loads
    answer, the XML carries `firstRow="1"` + the style id and the band shows. Slice table-style-fix (blind
    `queueTableStyle`, no read).
  - CUSTOM FUNCTIONS DEAD ON DESKTOP: every `=PLSFIX.*` returned #VALUE!. The manifest puts them on the shared
    runtime (`taskpane.html`), which never loaded `functions.js`, so `CustomFunctions.associate` never ran. Proven
    fix: `<script src="./functions.js">` after office.js (CAGR 0.1, ROUNDSUM 331 in the dev pane). Excel stores a
    custom-function formula as `_xldudf_PLSFIX_CAGR(...)`; the demo's un-prefixed cells stayed #NAME?. Slice
    custom-functions (page tag + gate, `cagr.rs` prefixed form, `share.ts` detector).
- Verified on screen, live v2.8.12/13: Autocolor, overlay, fill right, Precedents (D14+D16), ribbon CAGR (10.0 %),
  the Header preset (a grey band: the "top row" complaint was the PowerPoint table), projects (Excel select,
  Move to project, the PowerPoint inbox grouped), all four export kinds, Update all 7/7, Revert, Break link.
- Closed the same evening: v2.8.15 undo-fill, v2.8.16 table-style-fix (blind style write; proven on the
  Mac: the pasted table carries `firstRow="1"` + the style id, band visible, rows at 11 pt, "13 500"), v2.8.17
  custom-functions (proven: a typed `=PLSFIX.CAGR(100,161.051,5)` gives 0.1 on the live pane; the rebuilt demo's
  column I computes 10.0 % / -5.6 % / 0.0 % / 17.1 % / refused / refused at load); v2.8.14 proven on the Mac:
  single-line pie labels, the overlap note "the largest free spot is 888 x 69 pt" on a full slide. Manual
  regenerated at v2.8.017, memory note current. Seven sonnet slices, 2.6M agent tokens, every one reviewed.

## 14.09: v2.8.5 to v2.8.9 - tiers at every level, projects, a stale-payload refusal, strays swept by name

- Written up 16.09 from the commit bodies and the transcript: the 14.09 session shipped five patches and wrote
  no ledger, lessons or memory. Daniel, 14.09 10:49: "some features still work unclearly, for example
  formatting: it does not highlight the top row, and moving things to PP still crashes or just does not
  work." He stopped the desktop automation at 14:38 with PowerPoint on the scratch copy `plsfix-matrix.pptx`
  (a copy, his deck untouched), no insert made; the crash matrix had not started.
- Shipped, one patch each, every one deployed (live `/version` 2.8.9, release runs green):
  - v2.8.5 (af8af6a, 13.09 20:47): `tierUp` in `src/ppt/chart-draw.ts` groups the tier sub-groups again until
    at most six remain: a 12-column chart had become seven sub-groups in one addGroup, a 40-point one would
    have been 21, past the 19 that killed PowerPoint 16.107 on 13.09.
  - v2.8.6 (535b2b7, 13.09 22:21): link projects (item 2 of the 14.09 brief; `src/link/project.ts`): a project
    name on the registry entry, the inbox item and the shape tag; Excel picks or creates it above Linked
    objects, new exports join it, Push all and Update all honour the shown project, "Move to project"; old
    links stay under "No project"; the relay never sees the name. README says so; FEATURES and launch-check
    do not yet; the manual source has "Projekti" but the docx was not regenerated.
  - v2.8.7 (89b2993, 14.09 08:50): `MODELIS_PUBLIC_URL` XML-escaped in `server/src/manifest.rs` (security
    review M1); Update all's toast names the shown project or "No project".
  - v2.8.8 (74079a0, 14.09 10:17): a relay payload whose `pushedAt` is older than the shape tag's is refused
    on Update all (`assertFresh` in `src/link/status.ts`, called from `src/ppt/fetch.ts`; Revert asks for a
    named older rev and is exempt): security review I5.
  - v2.8.9 (e856652 + 40d4aa2, 14.09 14:05-14:11): the root cause of "moving to PP does not work" on the Mac.
    The Bridge EBITDA margin line chart (25 shapes) tiers as 6/6/6/6/1 and PowerPoint refuses `addGroup` of ONE
    shape (InvalidArgument); the refused batch had already landed the four sub-groups, their ids never came
    back, so the id-based cleanup left four orphan untagged groups, no top group, no link, no toast. Fix:
    `tierUp` carries a remainder of one into the next level ungrouped; every primitive and tier sub-group is
    named after its chart in the batch that adds it, and `chart-cleanup.ts` sweeps the slide top level by that
    prefix after the id pass; the fake gained `helpers.refuseNextSync` (a sync that applies its adds, then
    rejects, the way the Mac does). Gate at the time: 2718 vitest, all cargo suites, both manifests valid.
- Open after 14.09: the crash matrix (not started), the slide-3 free-space proof, the projects proof on the
  Mac, and the top row (the session's reading: `Table.styleSettings.isFirstRowHighlighted`, PowerPointApi 1.9,
  is never set and an Update all clears every unfilled cell's fill, which wipes any header band; to confirm on
  screen). Taken up by the 16.09 session (NEXT SESSION above).
## 13.09: the v2.7 wave - comps and hygiene tools, follow-ups closed, stress passes, relay hardening

- 13.09 evening, the desktop pass on Daniel's Mac (computer use for Excel and PowerPoint, granted): Excel
  16.107 passes (pane, Autocolor, audit overlay, four exports); PowerPoint 16.107 pairs and takes the
  P&L table (Slide 2, Whole slide) and a range picture, but a shape chart crashes it twice (Revenue chart
  on Slide 3 / Left half, Segment pie on This slide / Free space): both draw batches return, then the app
  dies on its repaint; no crash report. Branch `mac-charts-as-pictures` (worktree
  `~/.worktrees/plsfix/mac-charts-as-pictures`, base e054652) = v2.8.2: `hostDrawsCharts` in
  `src/ppt/charts.ts` keeps every chart a picture on PlatformType.Mac, on insert and on Update all (the
  decline reason now travels back from `refreshChartGroup`), the fake host defaults to "PC", tests in
  `src/ppt/charts.test.ts`; the pairing copy says where the key really lives (Excel: Links > Link key,
  PowerPoint: Settings); FEATURES, CLAUDE map, manual (LV) and launch-check updated. **v2.8.2 LIVE 13.09**
  (9217b28 + 67096d1 manual label; modelis clean, /version 2.8.2, release run green with 8 assets),
  shipped before the Mac re-check because v2.8.1 crashed every Mac chart insert. OPEN: the desktop
  re-check of the Mac picture route (the Mac was Daniel's live desktop at 16:54, automation paused);
  the bisection on the Mac (E1 = drawGroup without addGroup, then without lines / labels / wedges) to
  bring native charts back there; the ghcr cleanup still waits for the 2FA sudo step in Safari
  (`$S/ghcr-clean.sh` ready; the device code must be re-issued); the rest of the v2.8 desktop section.
  The bisection worktree is gone (branch merged, v2.8.3); the dev-pane-against-the-live-relay recipe
  (vite proxy target `https://dbautomatizacijas.com/modelis`) is in tasks/lessons.md.
- 13.09 night, the bisection on Daniel's Mac (after his GitHub login: ghcr cleanup done, 23 image
  versions -> v2.8.2/latest + v2.8.1): **v2.8.3** = native charts back on the Mac. The crash is one
  `addGroup` of a whole chart (19+ shapes); six group fine; so `groupTier` in `src/ppt/chart-draw.ts`
  groups in sub-groups of `GROUP_TIER_MAC` = 6 on PlatformType.Mac, then the sub-groups into the link's
  group (`tierUp`, cleanup list carries the sub-group ids); the v2.8.2 picture gate (`hostDrawsCharts`,
  `CHARTS_MAC_PICTURE`) is gone again, the decline reason still travels back from `refreshChartGroup`.
  Proven on the Mac with the dev sideload: insert (E15) and an Update all redraw after a pushed
  revision, both alive; tests `src/ppt/charts.test.ts` (tiers on the Mac only, 6/6/6/2 under the link,
  flat on PC, redraw keeps one link, a refused top group takes the sub-groups down). Docs, manual (LV),
  README, FEATURES, CLAUDE map and launch-check say tiers instead of pictures. Experiments E1-E15 in
  `tasks/lessons.md`. The demo workbook on the Mac was edited during the proof (P&L!C11 = 13000, chart
  objects renamed, extra links): `npm run demo` rebuilds it. **v2.8.3 LIVE** (modelis clean, /version
  2.8.3, release run green, 8 assets) and re-proven with the released pane on the Mac: Revenue chart
  on Slide 3 / Left half and Segment pie on Slide 3 / Top right, both shape groups, PowerPoint alive.
  Open: a quarter-size pie wraps its labels (launch-check, cosmetic); Daniel's idea of folders /
  projects for links (asked 13.09, not built, needs his go).
- v2.8.0 + v2.8.1 LIVE 13.09 (session "launch readiness"; plan
  `~/.claude-accounts/work/plans/refactored-painting-yao.md`, approved, budget 2.1-3.1M, spent ~3.6M
  incl. reviews): guided demos + slide/spot placement, four sonnet worktree slices, one opus review + fix
  round + re-review each, merged S3, S1, S2, S4 in that order (all fast-forwards, no conflicts), one
  `release.sh minor`, one deploy, the web proof, then v2.8.1 for the one thing the proof showed.
  S1 demo/: `GUIDE_ROWS = 7`, `sheets/guide.rs` (column A only, row-format mint tint, no wrap: a wrapped
  band was unreadable in a 4-wide column), every row constant `GUIDE_ROWS + n`, tests assert the band shape
  and that every cited button + key exists in shortcuts.json, compile-time task-count asserts; four rig
  snippets and the launch-check moved to the new addresses (P&L!C11, B11:E16, Variance B11:B24 / B13, B17,
  B19, B23). S2 src/: `layout.ts` spotBox + fitInto, `ppt/placement.ts` resolveTarget / finishTarget /
  InsertTarget (a cross-slide selection refused via getParentSlide, consume through getItemOrNullObject,
  stale slide pick -> "Slide N is gone: pick a slide again."), `ppt/target.ts`, `ppt/chart-picture.ts`
  (split out of charts.ts for the cap), the two selects in pptpane.html, help copy, 48 tests. S3 demo/deck/:
  python-pptx generator (venv git-ignored, requirements pinned), tracked reproducible deck (fixed stamp,
  normalised zip times), title band within 0..24 pt so free space starts at the content top, real
  numbering, template thumbnail/printer parts dropped. S4: manual "Slaids un vieta" (LV) + demo files
  section, FEATURES rows, README, Map, launch-check section, AppSource test notes re-addressed. Proof on
  Office for the web (both docs uploaded to the NDUS OneDrive through SharePoint REST from the scratch
  Chrome; both opened by the folder listing's UniqueId, the ListItem GUID is a different id and answers "Item does not exist"; the working URLs sit in the scratch profile's Sessions files): table -> slide 2 whole (368,216,225,108 centred), revenue chart -> slide 3
  left half (36..474), pie -> right half (502..908), Data picture -> slide 4 selected placeholder
  (283,199,395,91, placeholder consumed). v2.8.1: a chosen spot never reports the free-space overlap note
  (the deck's dashed guides made it fire on every insert). Gates at v2.8.1: 2679 vitest / 108 cargo,
  ux:check 0/72, ux:sweep 0/149. Release assets now 8 (the deck). OPEN: old ghcr image versions still
  need Daniel's `gh auth refresh -h github.com -s read:packages,delete:packages`; his desktop pass.
- 13.09 afternoon (session "launch readiness", plan `~/.claude-accounts/work/plans/refactored-painting-yao.md`,
  approved): the old brand's name is gone from everywhere it could still be found, and the app re-proven.
  History: `git filter-repo --replace-text/--replace-message` with a 53-line literal map (product name ->
  "Model Tools", provider -> Daniels Bendiks, identifiers -> today's names, a catch-all that was never
  needed) + `--path-rename` of the old unit file; 709 commits, 80 tags, HEAD tree hash byte-identical
  before and after; zero hits in every tree, diff and message; force-pushed main + tags; GitHub main =
  local HEAD; old SHAs stay fetchable until GitHub's GC. The 20 Office blobs in history were unzipped
  and checked: none carried the name. GitHub: the 18 releases v2.6.16-v2.7.14 deleted (their pane zips
  shipped the two pages that named the bank), v2.7.15 kept with its 7 assets, all tags kept; the old
  container image versions still wait for a token with read:packages + delete:packages. Server: unit
  `plsfix` from `/opt/plsfix` (mv of the whole directory, relay data intact: links 8, revisions 16,
  bytes 313318 before and after), log `/opt/plsfix/logs/build.log`, cache volume `plsfix-cargo-registry`,
  the old unit, directory, log and gateway copy removed; the hosting gateway commit 975dad9. Proven:
  `npm run check` (2631 vitest / 105 cargo), `npm run build`, `ux:check` 0/72, `ux:sweep` 0/147, every
  live page 200 without the name, live manifest = repo, relay refuses unknown keys (400/404), store
  validator valid on the live URLs, Mac installer + uninstaller from the live release into a throwaway
  HOME (manifest 2.7.15.0), Windows installer URLs resolve. The Office-on-the-web rig stopped at the
  Microsoft sign-in + Duo prompt (the NDUS session expired since 09.09): the one check that needs
  Daniel's sign-in in the scratch Chrome, then `register.js` (x, p), `proof-links-excel.js`,
  `proof-insert-one.js`, `proof-update-all.js`.
  DONE after his sign-in (13.09 ~13:30): both hosts registered on the web (the pls,fix tab in 9 s each),
  Excel pane connected v2.7.015, new key generated, Proof sheet + charts built, exports sent (the two
  "ERR" lines are the 08.09 snippet looking for chart names the picker now labels "Linked chart");
  PowerPoint pane opened by the ribbon's Links button, said "New here? paste the key" because the
  Excel proof rotated the key: pasted into #workspace-key + #save-key -> "Paired with Excel.", inbox 3;
  inserts on slides 4-6: table (10 s, native), Doughnut as a picture with its reason, Four series as a
  picture ("55 shapes is over this host's budget of 30", placed over the layout placeholders); update
  round: P&L!C4 +500, Proof!B2 +40, Push all 7 pushed / 4 missing (the 08.09 Proof links whose charts
  the snippet recreated), Update all 160 s: 7 updated, 4 up to date, slide 1 table rev 5 -> 6. Rig
  Chrome + manifest server stopped afterwards. The demo workbook and "Presentation 2" on the NDUS
  OneDrive carry the proof's edits, as after every rig run.
- v2.7.15 13.09 (session "launch readiness"): the former sponsor's name is out of the project (Daniel:
  "remove [the] name from this project everywhere, all md files, all code"; plan
  `~/.claude-accounts/work/plans/refactored-painting-yao.md`). 37 files, 92 hits -> `git grep -i` of the old
  name returns nothing. Server crate + binary `plsfix-server` (lib `plsfix_server`, every test import);
  build variable `PLSFIX_VERSION` (Dockerfile ARG/ENV, release.yml build-arg, `lib.rs` `/version.source`);
  the hosted systemd unit LEFT the repo: it lives in the hosting gateway's `server/systemd/` (same unit
  name on the server, `ExecStart` = the new binary, pushed and installed by its `deploy.sh modelis`,
  which also passes `PLSFIX_VERSION` and removes the old binary); the repo keeps the self-host template
  `deploy/plsfix.service` (`/opt/plsfix` paths). Rig snippet checks for the `pls,fix` tab; support +
  privacy pages say "the pls,fix owner"; 13 md files reworded (in-house, the hosting gateway, the host's
  `MODELIS_DATA` directory). Out of scope, on purpose: git history and the existing tags/releases keep
  the old strings; the server's directory layout and unit name are the host's.
- v2.7.14 LIVE 13.09 12:20 (2c65e47 + manual 627db2a; session "launch readiness"): the ONLY change is
  `package.json` losing `office-js@0.1.0`, an unrelated 2014 package (googleapis, request, hawk, jszip 2,
  two git+ssh GitHub deps) that sat in `dependencies` since the initial import; nothing imports it, the
  types come from `@types/office-js` (tsconfig `types: ["office-js"]` resolves there). 101 packages
  leave the lockfile, `npm audit --omit=dev` 27 findings (7 critical) -> 0; the 13 left are dev-only
  (office-addin-debugging's toolkit). Gate + `npm run build` green before the commit. Deploy: modelis
  clean, live `/version` 2.7.14; the suite audit ends "DEPLOY INCOMPLETE" on OTHER components
  (struktura `samples/amasty.toml`, sejas, aktivitate-nonbank, peers-nonbank: Mac trees ahead of the
  server, their owners' deploys), not modelis. GitHub release: the tag-push run failed at
  `gh release create` with an HTTP 500 from api.github.com (every earlier step green, image pushed);
  `gh run rerun --failed` ended `startup_failure` with zero jobs; `gh workflow run release.yml -f
  tag=v2.7.14` succeeded 09:15 UTC: 7 assets, `latest/download` -> v2.7.14, released manifest
  identical to the repo, note written on the release page. Readiness check before the ship (read-only):
  store validator VALID, `launch-check.md` 0/136 ticked, the Cloudflare zone has NO rate-limit rule
  and no WAF rule (Free plan, one rule available), README `check` badge = the disabled workflow,
  Dependabot alerts off, release bodies = the compare link only.
- State: **v2.7.13 LIVE 13.09** (987ea1c, deploy.sh modelis "OK / clean", live `/version` 2.7.13,
  healthz 200, support.html + privacy.html 200, functions.json with three `helpUrl`s; the store
  validator's support URL and 64 px icon checks pass since v2.7.5). Plan
  `~/.claude/plans/sleepy-mixing-wand.md` (approved 12.09 22:05, budget 3.65-6.25M + the 23:05
  extension 2.7-4.4M); ledger `.superpowers/sdd/sleepy-mixing-wand/progress.md` (git-ignored; every
  ruling, review verdict and merge), briefs, reports, reviews and diffs beside it.
- Shape: worktree agents by file ownership, one opus review per slice, scoped sonnet re-reviews after
  a fix round, the controller merging one branch at a time (rebase onto main, trailers stripped,
  `npm run check` + `ux:check` + `ux:sweep` on the REBASED branch with private dev-server ports, one
  `release.sh patch` per reviewed merge, deploy after each). A second session ("plsfix launch") took
  the repo public in the same checkout earlier on 12.09-13.09; main was handed back and forth by
  message, never touched by both at once. That session is paused at 2be12bf's line of work; its
  handovers to this one (the SELF-HOSTING proxy paragraph, privacy.html as the source of the AppSource
  policy copy) are done.
- Shipped, in merge order (13 slices, each reviewed, gates green on the rebased branch):
  - v2.7.0 M1 comps tools: Comps stats (six rows of live formulas under a comps table), Football field
    (stacked bar with a cleared floor series, first row on top), Pinstripes (every second row/column,
    second press clears; refuses while an overlay owns fills).
  - v2.7.2 M2 hygiene: Indent / Align / Underline cycles (canonicalised read-back), six sheet tools above
    the explorer (Unhide all with the "Include very hidden" tick, Show only this, Bury this, three moves;
    workbook-structure protection answers a sentence), Clean past the data (rows/columns deleted only when
    no chart or shape sits on the sheet AND the host can count them, else formats cleared; outside Undo),
    `=PLSFIX.CAGR` + `helpUrl` on all three functions.
  - v2.7.3 K2 PowerPoint + shell: a chart never lands below 200 x 120 pt, bar labels clamped inside the
    box, stacked bars in one row per category, `chart-guard.ts` keeps the picture when a payload's chart
    is unreadable, roving tabindex + arrow activation on both tab strips, `NEEDS_PANE` for the three
    pane-only shortcuts.
  - v2.7.4 N2 unit-test backfill + four header comments (squashed onto main: K2 had created the same test
    file name).
  - v2.7.5 K3 relay timeouts (client `RELAY_TIMEOUT_MS` 20 s through `relay-timeout.ts`, server 30 s
    `TimeoutLayer`, inbox tie-break by rowid), 64 px icon + `HighResolutionIconUrl`, `SupportUrl` ->
    `public/support.html`, `public/privacy.html`.
  - v2.7.6 N3 audit tools: Select consistent region, Precedents of selection (50 cells, grouped chips).
  - v2.7.7 N5 shortcut manager (Brand tab: remap per signed-in user through `Office.actions`; Apply reads
    before it writes; clashes named; Cmd its own modifier).
  - v2.7.8 N4 paste suite: Paste: duplicate formulas, Paste number formats only, Paste row heights only.
  - v2.7.9 K1 Excel follow-ups: fast-fill cap, `SHEET_SCAN_CELL_CAP` 200 000 for Find / share / model
    check, sheet-scoped names everywhere, series-level pie leader lines, the chart picker's "Linked
    chart" label.
  - v2.7.10 P1 comps polish + stress: `chart-blocks.ts` shared by tornado and football (a refusal never
    spends an Undo slot), `requireNoOverlayOwner` in `fill-store.ts`, stats formats per statistics
    column only, the `stress.comps.*` suites (10 `it.skip` rows with real bodies pin defects outside the
    slice).
  - v2.7.11 P2 hygiene stress: six `stress.hygiene.*` suites; CAGR overflow answered in log space
    (`chartmath.cagr`), size-cycle stage strings, `buriedNote` after Unhide all.
  - v2.7.12 S1 relay hardening (from the 13.09 security review, `security-review.md`): per-client rate
    limits by request AND by bytes keyed by the trusted proxy's header (`MODELIS_TRUSTED_PROXY`, the
    hosted unit sets `cloudflare`; live journal says "client key from Cloudflare"), storage ceilings
    read inside the write's lock, every store call on `spawn_blocking`, one shared permit pool for the
    two write methods (`GlobalConcurrencyLimitLayer` 32) with a 10 s body deadline inside it,
    `MODELIS_RATE_WRITE_PER_MIN` 300 -> 100 so `writes/min x deadline / 60 < permits` holds, a startup
    line naming the trust mode (warning on loopback, note on an open bind) when nothing is trusted.
  - v2.7.13 P3 PowerPoint stress: six `stress.ppt.*` suites over the shipped `pptpane.html`;
    `missing-shape.ts` (deleted/ungrouped objects answered in words), table rebuild keeps the shown
    revision's tag until the last format chunk, `requireTableSize` (60 x 20 on the way in), Select
    similar refuses a grouped child, one action at a time through `guard.ts`'s opt-in `busyMessage`
    latch (ribbon included), `refresh.ts` split out of `host.ts` (447 -> 256), `pane-details.ts`,
    `inbox-queue.ts` (a refused `deleteInbox` no longer duplicates the next paste), `relay-reason.ts`
    (429 / 507 / 500 in the pane's words).
  - Docs on main: manual LV sections for every new tool (charts/tools/workbook.rs) regenerated at
    v2.7.13, README v2.7 lines, `docs/appsource/privacy-policy.md` pointing at `public/privacy.html`
    as its source, `tasks/lessons.md` 13.09 section, CLAUDE.md Map + `tasks/launch-check.md` one
    section per slice, `docs/FEATURES.md` rows marked shipped v2.7.
- Security review 13.09: 1 Critical (one IP could fill the 1 GiB relay in under a minute), 5 Important,
  8 Minor; secrets grep clean. Closed by S1 (C1, I1-I3); I4 release.yml SHA pins done by the launch
  session; I5 closed: a payload whose `pushedAt` is older than the shape tag is
  refused (`assertFresh` in `src/link/status.ts`, called from `src/ppt/fetch.ts`;
  Revert asks for a named older rev and never calls it). M1 closed: `MODELIS_PUBLIC_URL` is XML-escaped
  in `server/src/manifest.rs` (ampersand, angle brackets, quotes).
- Rulings that changed the briefs (details in the ledger): band tint from `settings.primary`; a 4th
  football column is ignored; N2 squashed; the ux gates' default ports 3131/3132 are shared across
  worktrees, so every run passes its own `--port`; 100 writes/min kept (about 50 links per Update all,
  the pane does not retry a 429: if it bites, raise `MODELIS_MAX_INFLIGHT_WRITES` to 64 and restore 300
  in the unit, never lengthen the body deadline).
- Follow-ups, not done: a refused WRITE still spends a pls,fix Undo slot (`src/excel/undo.ts`); a
  refused READ sync leaks office.js's own string pane-wide (`src/ui/report.ts` / `src/pane/shared.ts`);
  "nothing was changed" over a partly unlocked sheet a ctrl-click batch DID paint
  (`src/excel/protection.ts`); "X is now sheet N" counts hidden tabs (`moveSheet`,
  `src/excel/workbook.ts`); the Excel pane could take `guard.ts`'s `busyMessage` latch in one line
  (the double football press; a behaviour change, Daniel's call); P1's three hidden-row skips can be
  un-skipped now that the fake has `helpers.hideRows`; the fake host still has no AutoFilter;
  `inbox-queue.ts`'s pasted set is in-memory (a pane reload forgets it); `model-check-panel.ts` "Too
  large to read" wording; 3-D references judged by their last sheet; `insertToc` 82 lines and the files
  still over the 400-line cap (`src/ppt/main.ts` 530, `src/ppt/links.ts` 432, `src/pane/links-tab.ts`
  429, `src/link/model.ts` 403; launch-week ruling stands); the `check` workflow re-enable.
- CI: the `check` workflow is DISABLED since 12.09 23:43 on Daniel's word (hung-sync flake, last copy
  fixed at v2.7.1); `release` (a Release + ghcr image per tag) stays on and ran for every v2.7.x tag;
  every gate ran locally before each merge.
- Daniel's gates: the desktop launch check (`tasks/launch-check.md` grew by one section per slice, about
  70 rows), Windows once, the M365 centralized upload (the manifest changed: icon-64, SupportUrl), the
  Cloudflare rate-limit rule on `/modelis/api/*`, the Mac Cmd/Ctrl physical-key check for the shortcut
  manager, the chart picker's "Linked chart" label as a UX call, the `check` workflow.
- Branches: none left (every wave/*, merge-* and worktree-agent-* branch deleted after its merge;
  `git cherry` clean; no worktrees under `.claude/worktrees`).

## 12.09 evening: "a package they just install": installers + the store kit (branch `installers`)

- Daniel: pls,fix must feel like an install, no extra steps. The only real install for an
  Office add-in is AppSource; until the listing is live, one line per platform does the
  sideload: `deploy/install/plsfix-install-windows.ps1` (`irm <url> | iex`: manifest to
  `%LOCALAPPDATA%\plsfix`, registered under `HKCU\...\WEF\Developer` like Microsoft's tooling)
  and `plsfix-install-mac.command` (`curl <url> | sh`: manifest into both `wef` folders), both
  with uninstallers, attached to every release by `release.yml`; `PLSFIX_MANIFEST_URL` for a
  self-hosted manifest. Proven: Mac scripts against a fake HOME with the live manifest
  (install, refuse a non-manifest, uninstall, idempotent); Windows scripts parsed by pwsh 7.6
  and dry-run on macOS (download, validate, id extract, `iex` form). The real Windows run is
  Daniel's pass. `docs/INSTALL.md` rewritten around them; README "Get pls,fix" = one line per
  platform, self-host + build a footnote.
- Store kit `docs/appsource/`: `privacy-policy.md` (names the app and the relay; Hetzner
  Helsinki, 14-day logs, relay 30/7 days), `listing.md`, `test-notes.md` (incl. the
  custom-function test the store requires), `checklist.md` (repo side, Partner Center side).
  Screenshots (1366x768, at least one) still to take: Daniel, or the web rig once signed in.
- Store validator 12.09: "Icon URL Unreachable" was the gateway's Latvia-only country gate
  (nginx answered 403 to `MicrosoftOfficeStoreValidationService`), which also meant no Office
  user outside Latvia could load the pane. Fixed in the hosting gateway (056d9c3)
  (`OPEN_COUNTRY_PATHS` = `/modelis/`, apex gates on ``, smoke probes it),
  deployed and proven on the server; the validator now reports only the missing 64 px icon
  (feature wave K3). SupportUrl for the store = `support.html` (K3 sets it).
- `release.yml` pinned after the 12.09 security review: action commit SHAs, job-scoped
  permissions, toolchain 1.98.1; Dockerfile on `rust:1.98-bookworm`.
- Merged 12.09 23:5x as f7da86c -> **v2.7.1 LIVE** (3e88d67; modelis OK + clean, routed smoke OK,
  `/version` 2.7.1); the v2.7.1 release carries the four installers, both proven from the
  real URLs (Mac one-liner into a fake HOME, Windows dry run). Daniel had the `check`
  workflow disabled ("it keeps saying something failed") before the last flaky suite was
  fixed in this merge; re-enabling it is his call.

## 12.09: public on GitHub (plan ~/.claude/plans/validated-juggling-duckling.md, approved)

- State: **PUBLIC, MIT, v2.6.18 LIVE 12.09** (3abdb15; modelis OK + clean in the deploy audit,
  `/version` 2.6.18, `/manifest.xml` byte-equal to the release asset; 1810 vitest; CI green on
  the tag and main). Releases v2.6.16-18 each carry `manifest.prod.xml`, `plsfix-pane-<tag>.zip`
  and the demo workbook; image `ghcr.io/bendiksdaniels/plsfix:{v2.6.16,v2.6.17,v2.6.18,latest}`,
  pullable anonymously. Repo: description, homepage, 7 topics, issues on, private vulnerability
  reporting on. v2.6.17 = a 20 s testTimeout that did NOT cure the CI flake; v2.6.18 = the real
  fix, `test/hung-sync.ts` (lessons 12.09). The deploy audit stays red on OTHER tools (struktura,
  aktivitate-nonbank, sejas, peers-nonbank drift), not modelis. Old pre-rewrite SHAs may stay
  addressable on GitHub until its GC (the repo was private, nobody holds them).
- Daniel: "prepare it on github so it would be easy for anyone to get it for free and set it
  up". Decisions (his): MIT; the UpSlide training dumps purged from tree AND history
  (git filter-repo, 556 commits rewritten, 61 tags re-pointed, force-pushed; everything else
  byte-identical, proven by tree fingerprint); the hosted pane + relay offered publicly; the
  working notes kept with the two VM login fragments redacted.
- Built: `server/src/manifest.rs` + `GET /manifest.xml` (verbatim, or re-pointed at
  `MODELIS_PUBLIC_URL` with a UUID v5 id), `MODELIS_BIND`; `deploy/Dockerfile` + compose +
  `.env.example`; `.github/workflows/release.yml` (assets + GHCR image, smoke-tested);
  `LICENSE`, `CONTRIBUTING.md`, `SECURITY.md`, `docs/INSTALL.md`, `docs/SELF-HOSTING.md`;
  README "Get pls,fix" section; manifest SupportUrl = the GitHub repo; package.json and the
  three Cargo.toml carry license + repository.
- A second session ("plsfix improvments") runs a five-slice feature wave in worktrees off the
  rewritten main; its merges happen in one announced window, versions bumped by whoever holds
  the window; no further history rewrite.

## 09.09: the audit fleet, v2.6.3 to v2.6.15 - every feature debugged on its own

- State: **v2.6.15 LIVE 09.09** (92bb4b5, manual d6315ed after it; SERVER IN SYNC, `/version`
  2.6.15, healthz 200), 1810 vitest / 156 files, ux:check 0/72, ux:sweep 0/124, cargo green.
  Plan `~/.claude-accounts/work/plans/lucky-brewing-journal.md` (its approval was the budget).
- Shape: ten worktree slices by file ownership (A formatting, B formulas, C charts/templates/
  reconcile, D workbook/brand/model check/find/share, E1 Excel links, E2 relay, F PowerPoint
  links, G slide charts, H object tools, I pane shell + the new `ux:sweep` gate) plus J
  (deadline on every PowerPoint batch); opus on the Office.js adapters, sonnet on H, I, J; five
  opus reviewers; one patch per reviewed merge (B, C, E1, E2, D, A, G, F, H, J and three
  controller rounds); subagent trailers stripped before each merge. Spend about 7M tokens
  (agents 4.6M, reviews 0.9M, controller about 1.5M): the top of the 4.3-7.4M estimate, because
  the account session limit killed the fleet once and H and I ran twice.
- Fixed, per slice (mechanism -> file; every fix has a test that fails on 0b330ee):
  - A: cap checked before the undo capture (`format-cycles.ts`); presets, eraser, cycles, Undo,
    row/column sizes and paint slots end on `syncWrite` so a protected sheet gets our sentence;
    `Range.cellCount` answers -1 above 2^31-1, which bypassed every cap and OOM-killed the worker
    on Ctrl+A (`internal.ts` overCap, `areas.ts`, `selection.ts`, `pickScannableSheets`);
    `roundedCorners` 1.9 guard; `parsePalette` null on a non-palette object.
  - B: paste/fill/CAGR/rounding protected-sheet sentences; Fast fill anchored on the selection's
    corner (an upward or leftward drag filled unselected rows); a "[" no longer means external
    (structured refs; `hasWorkbookReference` needs a sheet name and "!"); corrupt overlay JSON
    at boot; caps before loading grids (autocolor, CAGR, paste exact); the overlay owns only
    fills it striped; `fill-store` snapshot shape check.
  - C: chartex surface writes in their own tolerated batch; cap before loading values (waterfall,
    tornado); `chart-place` null row/column sizes -> 15/48 pt; `showConnectorLines` 1.9 guard;
    pie `showLeaderLines` (1.19, not 1.8) dropped from the restyle; reconcile select 1.9 guard,
    stale result line and hint cleared; protected-sheet sentences for templates and tornado.
  - D: brand JSON import of a non-palette file refused (new `src/pane/brand-io.ts`); an
    unreadable file -> sentence; no `setAutocolorOnEdit` before a host; lv/ru separators compare
    U+00A0; model check keeps its other seven kinds when the styles read is refused.
  - E1: the boot touch is awaited before the tab is wired (relay down = half-wired tab); the
    highlight keeps its fills on a protected sheet; ExcelApi 1.9 floor on the push path (Excel
    2019); 413 -> "That export is too big to send. Export a smaller range."
  - E2: inbox delete 404 = done (a working insert reported failed, duplicate inserts); weak ETag
    compare behind Cloudflare; no dotfile (nor its %2e/%2f forms) served on the bypass path.
  - F: `updateLinks` no longer skips rows whose cached status said current (the core loop showed
    "N up to date" after a push); `saveKey` reads before redrawing the inbox; mid-insert buttons
    disabled; change source across kinds refused (`sameFamily`, "Nothing waiting in the Inbox
    fits this link..."); hidden source columns 48 pt; chart refresh reasons surface; a dead relay
    fails rows once (network kind, no deferral).
  - G: THE HANG: every draw sync under `SYNC_TIMEOUT_MS` 60 000 (`withSyncDeadline`,
    `ChartDrawTimeout`) -> cleanup of the confirmed shapes, picture fallback, note "PowerPoint
    stopped answering while drawing the shapes"; placement reads under it; a swallowed refresh
    becomes a picture at the same corner; legend band, waterfall closing total below zero,
    negative bar label drift, a tall chart at top -255 (`onSlide`).
  - H: Smart Painter no longer writes a captured null line colour over a target's own.
  - I: controls wired before host detection (a no-host pane was dead); `cycle-row-height`
    shadowed by the prefix branch; `lastUndoSkipped` drained in `finally`; overlay restore
    toast; `probeExcelHost` loaded a property Excel has not got (the degraded boot never fired);
    tab strip keyboard navigation; `openShortcutCard` without `displayDialogAsync`;
    `registerCommands` at module top level; gate `npm run ux:sweep` (reaction only; routing
    stays `src/pane/dispatch.test.ts`).
  - J: picture, text, table, placement, shape scan, object tools and the group scan under the
    deadline with their own wording; `paintBatch` fails a hung batch once, not row by row;
    `refreshChartGroup` returns the silent note.
  - Controller, v2.6.15 from the rig: table cells written `CELLS_PER_SYNC` 8 per round trip
    (the tag with the last chunk; an insert whose format chunk fails takes the table down).
    On PowerPoint for the web a formatted cell cost about 0.4 s per property (24 cells: 0.7 s
    text-only, 28 s formatted); a 6x4 repaint in one batch ran past 60 s and every batch after
    it queued behind the abandoned one. Proof: Update all 256 s with four deadline hits at
    v2.6.14 -> 128 s, 6 updated, 2 up to date, no deadline at v2.6.15.
- Refuted suspects (the audit tests hold the evidence): the cross-sheet auto-push burst pushes
  once; registry read-modify-write interleavings; touch chunking is client-side; a table push
  below 1.9 already falls back to the plain grid; `renderChartPlain` await-fold; every `.value`
  read across the 23 load sites follows its sync; `hostSupports` stays fail-open (18 call sites
  degrade on false, only `requireImageApi` throws).
- Host-only suspicions the fakes cannot answer: appended to `tasks/launch-check.md` (09.09
  section, Excel and PowerPoint rows) for Daniel's desktop pass.
- Web rig proof 09.09 (Office for the web, NDUS tenant, scratch Chrome): both panes v2.6.015,
  registration, exports, pairing, three inserts (native table, two pictures with their reasons),
  push 6 / 2 missing, Update all as above; rig torn down, ports 9222/3001/3000 free, the link
  key deleted. Two quirks in `tasks/lessons.md` (anchor-named charts, background tab).
- Follow-ups, not done: pie leader lines via `ChartSeries.showLeaderLines` (1.9); a
  `fastFillAuto` cell cap; Super Find / share caps 5 000 / 20 000 against model check's
  200 000; the names scrubber ignores `worksheet.names`; relay `created_at` in whole seconds;
  a relay client request timeout (AbortSignal.timeout is unsafe on old webviews); `model.ts`
  refusing a picture payload on a bad chart; tab strip roving tabindex + `refreshSheets` on
  arrow; `onSlide` below MIN_SIZE; bar-kind label overflow drift; the chart picker listing
  anchor names (UX call); `shortcuts.json` pane-only actions; `src/ppt/host.ts` 424 and
  `links.ts` 440 lines over the cap (launch week, no restructuring).
- Daniel's gates: the desktop launch check with the appended rows, the Windows pass, the M365
  centralized upload (the manifest did not change in this wave), the Cloudflare rate-limit rule
  on `/modelis/api/*`.

## 08.09: "does the Excel -> PowerPoint mover work?" - desktop back on prod, v2.6.2 says why a chart stays a picture

- Finding 1 (environment): both wef folders held the DEV manifest v2.4.9 (`https://localhost:3000`)
  since Codex's `npm start` on 01.09 14:34; two vite processes (PIDs 24071 :3000, 25672 :3131) ran
  for 6 d 17 h and were what the desktop add-in loaded from, with the old ribbon. Fixed:
  `scripts/wef-restore-prod.sh` (v2.6.2 prod manifest in both wef folders, `cmp` proven), both
  processes killed, :3000 free. Daniel must quit + reopen Excel and PowerPoint. The durable fix
  stays his M365 centralized upload.
- Finding 2 (product): the mover recreates natively by design: a table as `addTable`
  (PowerPointApi 1.8), a chart as a tagged group of native shapes (1.8, pie 1.10; shape budget
  desktop 200, web 30). A true chart object is impossible in Office.js (PowerPointApi 1.10 is the
  latest set, no chart API; `insertSlidesFromBase64` = a new slide, unlinked, no refresh). Daniel's
  Mac runs 16.107.2, above every gate. The defect: a chart Excel could not describe (more than 3
  series, an unsupported type, over 40 points, overlapping columns, a refused read) became a
  picture SILENTLY. v2.6.2: `chartIssue` on the picture payload (`readChartData` -> `ChartRead`),
  `chartCapIssue` / `seriesCountIssue` / `pictureNote` in `src/link/chart-model.ts`, the Excel
  toast "(as a picture: ...)", the PowerPoint insert note and `UpdateSummary.notes` under Update
  all; series cap 3 -> 6 (the palette's count). 1152 vitest / 104 files, ux:check green, manual
  v2.6.002 regenerated, stale `docs/lietotaja-rokasgramata-saites.md` corrected. Web budget
  measured: column 5x2 = 31 shapes, 6x2 = 36, line 8x1 = 33 (all pictures on the web, with the note).
- Deploy detour: `db contracts` failed because the bond tool's Desktop repo had been parked by
  the iCloud Desktop toggle (07.09 23:24) in `~/Desktop/Desktop - Dāniels’s MacBook Pro/`; moved
  back with a same-volume rename (repo clean, target.nosync intact), the empty folder to Trash.
- Git: Daniel said "commit to github this version" at 11:15: `git push origin main --follow-tags`
  put 212 commits and the tags v2.5.3 .. v2.6.2 on GitHub; from here plsfix follows the global
  auto-push rule (the earlier commits of the day ran with GIT_AUTO_PUSH=0 while the 02.09 ban
  stood).
- Web proof DONE 08.09 11:0x (scratch Chrome, NDUS tenant, both hosts registered through the
  wdaddin* parameters after the Developer Mode opt-in + reload): Excel pane v2.6.002, five
  exports (Revenue chart, Segment pie, a 4-series column chart = data shipped under the lifted
  cap, a doughnut = toast "(as a picture: Doughnut charts are not drawn as shapes)", P&L!B4:E9
  table); PowerPoint pane paired with the pasted key; slide 1 read back: `Table` native, two
  picture links (doughnut = Excel's reason, four series = web budget), `Group` x 21 shapes
  (pie, 20 s) and `Group` x 20 shapes (column, 30 s); source bump + Push all "5 pushed" +
  Update all "5 updated" in 80 s, every tag rev 1 -> 2 at the same boxes. One pie insert
  jammed after its first chunk of 12 (pane reload + retry healed it). Rig fixes committed
  (drive.mjs semicolon, opt-in tick, README rules, four 08.09 snippets). Screenshots in
  `.superpowers/rig/`; the proof deck "Presentation 2" and "plsfix Demo 0809.xlsx" stay in
  Daniel's NDUS OneDrive.
- NEXT: (1) Daniel's desktop pass after relaunching Excel + PowerPoint (wef = v2.6.2 prod):
  `tasks/launch-check.md` Links rows + a 4-series chart (a shape group on the desktop, budget
  200) + a doughnut (the reason in both toasts); (2) follow-ups not done: a hung web batch
  never rejects, so `chart-cleanup` cannot run - consider a per-sync timeout that cleans up
  and falls back to the picture; drop value labels before the web falls back; more chart
  kinds; an unlinked editable chart on a new slide.


## 02.09 morning: VM parked ("I did not finish building it"; no computer use for now); real-host pass = Daniel's own

- Daniel 02.09 ~07:15: continue without the VM and without computer use. The activation
  monitor is stopped, `claude-vm` is stopped (state kept, manifests still in its wef folders,
  `share/plsfix/` staged for whenever it is finished). The real-Office proof and the demo video
  therefore move to Daniel's own hands: `tasks/launch-check.md` is the 25-minute checklist
  (Mac first, Windows once), and a failing export's "Copy details" paste is the bug report.
  Nothing else in the launch scope is open on the code side.

## 02.09 night: v2.6.1 LIVE (PowerPoint ribbon commands); VM proof and demo video wait for Daniel's sign-in

- v2.6.1 closes Codex's own open note: the PowerPoint ribbon gained an Objects group (Object
  tools = `Office.addin.showAsTaskpane()` + the Tools tab, Match size, Select similar, Swap,
  Capture style, Apply style) and an Arrange group (six alignments, two distributions), 14 new
  icons from `scripts/build-ribbon-icons.ts`, `src/ppt/commands.ts` (`commandTable`,
  `registerCommands(deps)` on the shared runtime, registered at pane load), unit test
  `src/ppt/commands.test.ts`, fake-host proof `test/ppt.commands.integration.test.ts`;
  `manifest/xml.test.ts` holds each host's FunctionNames against its own table. 1137 vitest /
  103 files, manifests valid, `/version` 2.6.1, `/assets/ribbon/swap-32.png` 200. The manifest
  changed, so Daniel's first M365 upload takes this one (`~/claude-vm/share/plsfix/` restaged).

- v2.5.7 = the opus review's three host-only bugs fixed with fake fidelity (rising line segments
  drawn as `GeometricShapeType.lineInverse` from a normalised box with a `rising` flag; the fake
  throws on a negative size via `test/fakeppt/size-guard.ts`; Smart Painter `paintShape` never
  writes a null line or fill value and the fake answers null like the host; line points centred
  on their category slot) plus the controller's reconcile cap/tie/empty-input and paint-slot race
  fixes. v2.6.0 = the minor bump (`release.sh minor`), manual regenerated `manual v2.6.000`
  (d488f92), `deploy.sh modelis` -> `/version` 2.6.0, `/healthz` 200, `ux:check` 0/72,
  1129 vitest / 101 files. Merged today: demo Variance sheet (target 2 230, one 4-cell subset) +
  four checklist rows; Latvian manual chapters; `docs/IT-vienlapa.md`.
- The sonnet fix agent died on the session's rate limit one step from its gate (resets 02:40
  Riga); its three commits were complete, the controller committed its uncommitted tail
  (`test/ppt.charts.line.integration.test.ts`, `addLineShape`) and merged.
- VM: `~/claude-vm/share/plsfix/` holds `manifest.prod.xml` (v2.6.0), the demo workbook and
  `sideload.sh` (`sideload` copies the manifest into both guest wef folders + the workbook to the
  guest Desktop; `record [seconds]` starts `screencapture -v -V<s> -x` into the share; `stop`).
  Probed 02.09 01:30 guest time: `sideload.sh` copied the v2.6.1 manifest into both wef folders
  and the workbook to the guest Desktop; Excel's first-run dialogs (privacy, diagnostics =
  no, experiences) are cleared; in the unlicensed "read-only mode" the demo opens but the
  ribbon shows NO pls,fix tab (add-ins do not load without activation), so nothing more can be
  proven there before the sign-in. Daniel signs in over `vnc://192.168.64.2` (VM login),
  then an opus agent runs
  `vmctl ssh sh "/Volumes/My Shared Files/share/plsfix/sideload.sh"`, opens the demo, exports a
  chart and a table, reads "Copy details" (`vmctl` clipboard), verifies v2.6.0, records the demo
  and the host copies `share/plsfix/pls,fix demo.mov` to `~/Desktop/`.
- Daniel's gates unchanged: Windows pass, M365 centralized upload of `manifest.prod.xml` (icons +
  PowerPoint host + custom functions; first upload ever), Cloudflare rate-limit rule on
  `/modelis/api/*`, IT one-pager review, ROADMAP tick-through from docs/FEATURES.md. (Codex's
  open note about the PowerPoint ribbon is closed by v2.6.1.)

## 01.09: launch slice (plan ~/.claude/plans/zippy-herding-globe.md, approved; goal "zero bugs, all features working")

- Codex (01.09 day) shipped v2.4.9 ribbon icons, v2.5.0 native line charts + link search/status
  filter + paint slots in the workbook, v2.5.1 "Harden Mac link exports" (an unverified
  platform branch), v2.5.2 one-click "Paste latest linked" (LIVE), and left uncommitted the
  reconcile section, the PowerPoint Tools tab, the source/slide filters and the Save/Use labels.
- Controller (Fable) committed that slice as v2.5.3, then three worktree agents: C (opus) replaced
  the Mac platform branch with retry-on-refusal in the same context (`link-render.ts`,
  `link-table.ts`, `bothRefused`; `staged()` keeps `code` + `debugInfo`) = v2.5.4; A (sonnet)
  proved the object tools in the fake (`test/fakeppt/selection.ts`, 1.5 guard, UX gate walks the
  Tools tab) = v2.5.5; B (sonnet) proved reconcile in the fake (`worksheet.getRanges`, four
  syncs), `saveWorkbookPaintSlots` never throws, line charts got value labels in
  `src/chart-shapes-line.ts`, the summary line shows cents = v2.5.6. Gates: 1122 vitest,
  ux:check 0/72, build ok. Docs done: README, FEATURES, Map, lessons (dispatch message = Agent
  calls only). Wave 2 in flight: opus review of the whole day's diff, sonnet Latvian manual
  chapters + `docs/IT-vienlapa.md`, sonnet demo blocks (Variance sheet + checklist rows).
- VM (`~/claude-vm`, `vmctl`): booted 01.09 22:43, Excel shows the M365 activation wall; Daniel
  must sign in over `vnc://192.168.64.2` (VM login) before the VM agent can sideload
  `manifest.prod.xml`, reproduce the Mac export, capture "Copy details", verify v2.6.0 and record
  the demo video (in-VM `screencapture -V`, copy via `share/` to his Desktop).
- NEXT: apply review findings -> merge docs + demo -> `npm run check` -> `sh scripts/release.sh
  minor` (v2.6.0) -> the hosting gateway's `deploy.sh modelis` -> `/version` 2.6.0 -> VM agent
  (after the sign-in) -> memory + this file. Daniel's gates unchanged: Windows pass, M365 upload,
  Cloudflare rate-limit rule, IT one-pager review, ROADMAP tick-through, the Mac paste.

## 31.08: v2.5 slice "ready for other people" (plan ~/.claude/plans/cozy-plotting-honey.md, approved)

- 31.08 afternoon, slice DONE except the bug: v2.4.4 Model check (src/model-check.ts + src/excel/
  model-check.ts + src/pane/model-check-panel.ts), v2.4.5 text links (src/excel/link-export.ts +
  link-render.ts + link-touch.ts, src/ppt/texts.ts; touch on Links-tab boot), v2.4.6 undo stack
  (five deep, 25 000 cells, src/excel/undo-stack.ts), v2.4.7 defects (src/ppt/chart-cleanup.ts;
  src/excel/link-list-watch.ts live Links list), v2.4.8 = first-run cards (src/ui/first-run.ts),
  printable shortcut card (scripts/build-shortcuts-page.ts -> public/shortcuts.html, regenerated
  by bump-patch.sh, `npm run shortcuts:check`), find-a-tool (src/tool-search.ts +
  src/pane/tool-search.ts). ~1 040 vitest. Agents: ~2.6M tokens this slice (relay 245k, split
  304k + review 160k, text links 318k, model check 264k, undo 279k, defects 410k, first-run 361k,
  find-a-tool 345k, probe ~150k killed).
- Daniel 31.08: "Stop trying to open excel or simulate the environment" after a probe agent
  sideloaded a dev manifest into his Excel. Rule in tasks/lessons.md (31.08) and memory: never
  drive his desktop Office; the wef folders were restored to the prod manifest; the probe
  branch is deleted. The Mac export bug therefore waits for HIS "Copy details" paste.

- OPEN BUG (Daniel, 31.08 ~10:00): "exporting charts and tables is not working, there is a
  generation error" on desktop Excel for Mac 16.107.322 (UA in the VPS nginx log). Evidence so
  far: NO request from his Excel reached `/modelis/api` (nginx access log), so the failure is
  inside Excel before the upload (render or anchor); relay write path proven healthy from
  outside (PUT/GET/status/DELETE 200, 300 KB PUT 200, 5 MB 413); "generation" matches no pane
  string, so it is probably Office.js `GeneralException`; docs give NO Mac limitation for
  `Range.getImage`/`Chart.getImage` (office-js #235 is the upside-down Mac picture, the old
  spike item). Waiting for Daniel's "Copy details" text (it carries `code` + `debugInfo`
  with the failing statement, src/ui/report.ts) and whether plain "Export selection" works.
  Desktop Mac never ran the link flow before today (old gate 1). Fallback if no paste: a
  dev-only probe mode in the pane (query flag, runs the exports and posts the result to the
  vite dev server) so desktop can be verified without UI control.
- DONE 31.08: hygiene (18 retro-tags v2.1.21..v2.4.1 from package.json history; 10 stale
  agent worktrees removed, 10 GB; `scripts/rig/` = the web verification rig recovered from
  scratchpads, `npm run rig` / `rig:manifests`; `scripts/release.sh` = bump + commit vX.Y.Z +
  annotated tag, refuses a dirty tree); relay hardening v2.4.2 LIVE (LINK_TTL 30 d, INBOX_TTL
  7 d, `POST /api/links/touch` + client `touchLinks`, `MODELIS_MAX_BYTES` 1 GiB ceiling 507,
  `INBOX_MAX_PER_WS` 500, per-client token buckets `MODELIS_RATE_WRITE_PER_MIN` 300 /
  `MODELIS_RATE_READ_PER_MIN` 1200 keyed CF-Connecting-IP > XFF > peer, 429 + Retry-After,
  `/version` gains `relay` counts; nginx passes CF-Connecting-IP through); `src/main.ts` split
  into `src/pane/*` (sonnet, opus-reviewed, fixes applied) v2.4.3; text links spec
  `docs/superpowers/specs/2026-08-31-text-links-design.md` + plan
  `docs/superpowers/plans/2026-08-31-text-links.md` (7 tasks; Task 1 = `touchWorkbookLinks`
  on Links-tab boot, not yet wired).
- NEXT: (1) root-cause the Mac export bug, fix, deploy; (2) dispatch the text-links plan (opus
  worktree, budget 700k-1.3M); (3) Model Check + undo stack + B5/B6 defects; (4) first-run
  card, find-a-tool, shortcut card. Daniel: Cloudflare rate-limiting rule on
  `/modelis/api/*` in the dashboard (the API token is Access-only), Windows pass, M365 upload.

## 30.08 late: v2.3.6 (values-only chart labels, Templates on top), WP3 native charts next

- WP3 SHIPPED 30.08 late as v2.4.0 + v2.4.1 (LIVE): chart links land as tagged groups of native
  shapes on PowerPointApi 1.8 hosts (pies need 1.10), value labels only, brand colours, rebuilt at
  the same corner on update; picture fallback below the API or over the shape budget (web 30,
  desktop 200, `src/ppt/charts.ts`). Proven on PowerPoint for the web with the demo (pie 18 s,
  column 24 s, Update all 54 s, line chart stays a picture); v2.4.1 fixed sloped baselines (addLine
  treats a zero side as unset). Open: Daniel's desktop pass (desktop budget and text metrics),
  cleanup of drawn shapes when a group sync fails mid-insert (today a hung web sync leaves the
  chunks on the slide, ungrouped), stale agent worktrees under .claude/worktrees (8), the web
  demo doc has P&L!C4 +1000 and Rounding!B4 x1.5 from the proof pushes. Earlier plan text: spec `docs/superpowers/specs/2026-08-30-native-charts-design.md` (spike findings
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
  unchanged on purpose: key `modelis`, `/modelis/`, the binary and unit names (the binary became `plsfix-server` and the unit moved to the host's tree on 13.09), `MODELIS_*`.
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

## State (historic, 29.08: the dated sections above are current)

v2.1.20 is deployed at dbautomatizacijas.com/modelis/ (`/version` -> 2.1.20): Excel pane with the
new **Links** tab, PowerPoint pane (`pptpane.html`), and the end-to-end encrypted link relay
(`/api/links`, `/api/inbox`, sqlite in the host's `MODELIS_DATA` directory). `npm run check` green on
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
