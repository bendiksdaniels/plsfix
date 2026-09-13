# Feature copy-target matrix

Consolidated from: Macabacus help center (~100 features), UpSlide help center (~140 features),
UpSlide sales-team training files (primary source, not kept in the repository), and an Office.js platform feasibility study.
Raw inventories with per-feature source URLs live in `docs/research/`.

Date: 2026-08-27. Owner add-in: pls,fix (Office.js task pane, ExcelApi >= 1.9).

**Status 29.08 (v2.1.1):** everything marked P1 in sections 0-7 shipped by v1.1 — shared
runtime + ribbon + 34 shortcuts, format cycles, autocolor v2 + color key, audit overlay +
Smart Track, pls,fix Undo + paste suite + fast fill + CAGR/sign/decimals, native waterfall +
chart formatter + CAGR label, TOC + sheet explorer + name scrubber. Section 8's core linking
(export, registry, Inbox, Update selected/slide/all, Break, move-resilient anchors) shipped
at v2.0; per-workbook brand palette, grouped-shape links, tornado chart, unpivot selection and
change source followed at v2.1. Open: highlight linked cells, native
PowerPoint tables, link revert (section 8); the PPT companion's own toolset beyond the Links
tab and all of section 10 (enterprise, M4); the P2/P3 long tail in sections 1-7. See
tasks/AUTORESUME.md and tasks/v2-backlog.md for the live backlog.

**Feasibility legend** — `Yes`: doable with documented Office.js APIs. `Yes*`: doable with a
caveat (noted). `Partial`: a reduced version is doable. `Backend`: needs our planned Rust
service or Microsoft Graph relay. `No`: impossible in web add-ins (COM-only).
**Priority** — P1 build next, P2 later milestone, P3 long tail, X skip.
**Status** (replaces the old `Done` marker in this column) — `shipped v1.1` / `shipped v2.0` /
`shipped v2.1`: live in that release. `parked`: ruled out (see "Not worth cloning" below). Any
other Feasibility value on a P1-P3 row (`Yes`/`Yes*`/`Partial`/`Backend`/`No`) is still planned,
not yet built.

## 0. Platform foundations (prerequisites, not user features)

| Item | Why | Feasibility | Priority |
|---|---|---|---|
| Shared runtime (manifest change, lifetime long) | Prerequisite for shortcuts, ribbon-state, pane/ribbon shared memory | shipped v1.1 | P1 |
| Custom keyboard shortcuts | The whole Macabacus/UpSlide UX is shortcut cycles | shipped v1.1 | P1 |
| Ribbon commands (buttons that run code, no pane) | One-keystroke actions without opening the pane | shipped v1.1 | P1 |
| Excel contextual tab | Surface tools next to native tabs | Yes (RibbonApi 1.2, Excel only) | P2 |
| Cell right-click menu items | ContextMenuCell extension point | Partial (add items only, cannot replace native) | P2 |
| Unified + XML dual manifests | Coverage across hosts until unified manifest is universal | Yes | P2 |

## 1. Formatting and styles

| Feature | Source | What it does | Feasibility | Priority |
|---|---|---|---|---|
| Model presets (title/header/input/formula/result) | both | One-click branded cell styles | shipped v1.1 | - |
| Brand palette + font + currency settings | both (MB Color Palettes, US themes) | User-editable brand scheme drives all output | shipped v1.1 (per-workbook persistence v2.1) | - |
| Number format cycles (general/currency/percent/multiple/date) | both | Repeated keystroke cycles curated formats per family | shipped v1.1 | P1 |
| Title / result / item row format cycles | US | Cycles approved row styles | shipped v1.1 | P1 |
| Fill color cycle / font color cycle | both | Cycles brand fills/fonts | shipped v1.1 | P1 |
| Border style/color cycles | both | Cycles approved borders (incl. accounting underlines) | shipped v2.1 | P1 |
| Row height / column width cycles | US | Cycles preset standards | shipped v2.1 | P2 |
| More/fewer decimals | US | Steps decimal places in the active format | shipped v1.1 | P1 |
| Change sign | US | Flips sign of selected constants | shipped v1.1 | P1 |
| Pinstripes (alternate row/col shading) | MB | Odd/even shading via conditional format | Yes (ConditionalFormat API) | shipped v2.7 |
| Custom Styles / Style Cycles (user-defined, 8 slots) | MB | Reusable multi-property styles on one key | Yes (settings-driven) | P2 |
| Paintbrush (multi-slot format painter) | MB | Copy/apply formatting without clipboard, FIFO slots | shipped v2.1; slots saved with the workbook v2.5 | P2 |
| Indent cycles, center cycle, underline cycle | MB | Alignment/underline cycling | Yes | shipped v2.7 |
| Footnote cycle/toggle/hide/checker (in-cell superscripts) | MB | Numbered footnote management in cells | Partial (number-format superscript tricks; font runs limited) | P3 |
| Standard Sizes (conform cells/charts to preset dims) | MB | Preset output dimensions | Yes | P3 |

## 2. AutoColor and model hygiene

| Feature | Source | What it does | Feasibility | Priority |
|---|---|---|---|---|
| Autocolor selection (inputs/formulas/links) | both | Content-based font coloring | shipped v1.1 (basic) | - |
| Autocolor: distinguish same-sheet vs cross-sheet vs external-file vs partial-input formulas | both | Richer classification (MB adds hardcode-in-formula detection) | shipped v1.1 | P1 |
| Autocolor on entry/edit | both | Live coloring as you type | shipped v1.1 | P1 |
| Autocolor legend & customization | US | Legend UI + per-type color overrides | shipped v1.1 | P1 |
| Formula Audit / Formula Flow overlay | both | Striped fill = formula consistent with neighbors, solid = deviation | shipped v1.1 | P1 |
| Uniformulas (select consistent region) | MB | Highlights the consistency region of active formula | Yes (same engine) | shipped v2.7 (Select consistent region) |
| Dependency Density heatmap | MB | Shade by dependent count | Yes* (getDependents per cell is slow at scale; cap range) | P3 |
| Model Check (50+ automated checks) | MB | Error/structure/hidden-data/brand audit with fixes | shipped v2.5 (8 check kinds: formula errors, hardcodes in formulas, inconsistent formulas, volatile functions, broken names, unused styles, hidden sheets, external links) | P2 |

## 3. Paste and fill tools

| Feature | Source | What it does | Feasibility | Priority |
|---|---|---|---|---|
| Fill right/down from edge cell | both | Copy formula across selection | shipped v1.1 (basic) | - |
| Fast Fill with auto-extent (no pre-selection) | US | Detects how far to fill from neighbors | shipped v1.1 | P1 |
| Paste values / formulas / formats / transpose / skip blanks (button-driven) | both | Paste-special suite | shipped v1.1 | P1 |
| Preserve Formulas paste (exact references) | both | Paste keeping original refs | shipped v1.1 | P1 |
| Duplicate Formulas paste | US | In-range refs adapt, external absolutes kept | Yes (formula rewrite) | shipped v2.7 (Paste: duplicate formulas; outside refs keep their cells across sheets) |
| Paste row heights / number formats only | US | Targeted format transfer | Yes | shipped v2.7 (Paste number formats only / Paste row heights only) |
| Intercept native Ctrl+C/Ctrl+V | - | True clipboard interception | parked | X |

## 4. Formula auditing and navigation

| Feature | Source | What it does | Feasibility | Priority |
|---|---|---|---|---|
| Smart Track / Trace In-Out pane (precedents/dependents tree, keyboard nav, color trail) | both | Drill into inputs and back | shipped v1.1 | P1 |
| Show all precedents for multiple cells | MB | Multi-cell trace | Yes* (same; cap cell count) | shipped v2.7 (Precedents of selection, 50 cells) |
| Super Find (values/formulas/comments across workbook) | MB | Better Find with results pane | shipped v2.1 (this workbook only, no other open workbooks; comments and replies searched on ExcelApi 1.10, author included) | P2 |
| Explorer (workbook/sheet tree navigator) | US | Sheet navigation pane with search | shipped v1.1 | P1 |
| Workbook TOC sheet | both | Hyperlinked contents sheet, auto-updating | shipped v1.1 | P1 |
| Sheet tools (move/bury/unhide-multi/activate dialog) | MB | Sheet management | Yes (visibility incl. VeryHidden) | shipped v2.7 (no activate dialog: the explorer activates on click) |
| NavAid (crosshair shading of selection row/col) | MB | Visual navigation aid | Yes* (onSelectionChanged + fills; perf care, restore on move) | P3 |
| Reverse rows/columns preserving formulas | MB | Reorder periods | Yes (careful formula rewrite) | P3 |

## 5. Modeling helpers

| Feature | Source | What it does | Feasibility | Priority |
|---|---|---|---|---|
| IFERROR wrap | both | Wrap/unwrap with custom fallback | shipped v1.1 (wrap/unwrap; custom value still open) | P1 |
| Scale x1000 / /1000 | both | Rescale constants+formulas | shipped v1.1 | - |
| Quick CAGR formula | both | Insert CAGR over range/period | shipped v1.1 | P1 |
| Summary statistics block (min/max/mean/median under data) | MB | Auto stats for comps | Yes | shipped v2.7 |
| Add Scenarios (toggle-driven projection cases) | MB | Scenario switch cells + duplicated rows | Yes (structured but doable) | P3 |
| Replicate Module | MB | Duplicate an analysis block across sheets | Yes | P3 |
| Custom functions (UDFs) e.g. =PLSFIX.CAGR | - | Namespaced functions | Yes (CustomFunctions set; not iPad/perpetual<=2021) | shipped v2.7 (=PLSFIX.CAGR) |

## 6. Charts

| Feature | Source | What it does | Feasibility | Priority |
|---|---|---|---|---|
| Waterfall / bridge builder | both | Build from intuitive table, connectors, brand colors | shipped v1.1 | P1 |
| Stacked waterfall | US | Subcategory pillars | Partial (no native type — compose from stacked columns with helper series) | P2 |
| CAGR arrow on chart | both | Data-driven growth arrow overlay | shipped v1.1 | P1 |
| Chart Smart Format (brand compliance) | both | One-click restyle any chart to brand | shipped v1.1 | P1 |
| Football field (valuation ranges) | MB | Floating-bar range chart | Yes (stacked bar with invisible base) | shipped v2.7 |
| Marimekko / S-curve | US | Width-encoded 100% stacked | Partial (no native type; column-width tricks or scatter-area composition) | P3 |
| XY scatter labels | MB | Correct point labels | Yes (series data labels API) | P3 |
| Gantt (in PPT) | US | Project timelines | Partial, PPT companion scope | P3 |

## 7. Workbook cleanup and publishing

| Feature | Source | What it does | Feasibility | Priority |
|---|---|---|---|---|
| Clean: broken/hidden defined names | both | Name scrubber | shipped v1.1 | P1 |
| Clean: unused styles | both | Style scrubber (style-ceiling fix) | shipped v2.1 | P2 |
| Clean: crop used range | US | Clear stray formatting past data | Yes | shipped v2.7 (rows and columns deleted when no chart or shape sits on the sheet) |
| Prepare for sharing (every visible sheet to A1, report hidden content, external links, broken names, autocolor on edit) | US | Externalize workbook | shipped v2.1 (zoom is not in the Office.js worksheet API) | P2 |
| Smart Print (headers, orientation, print areas) | US | Print prep | Partial (PageLayout API exists; verify coverage) | P3 |
| Send via email (attachment/PDF/image) | both | Outlook handoff | Partial (no Outlook automation; mailto/Graph sendMail via backend) | P3 |
| Workbook performance optimizer | MB | Diagnose slow workbooks | Partial (heuristics: used ranges, volatile fns, styles count) | P3 |

## 8. Excel -> PowerPoint/Word linking (Milestones 2-3)

Architecture verdict from the feasibility study: **no direct add-in-to-add-in channel exists**.
Linking = Excel add-in exports content + link metadata to a relay (our Rust service, or Graph/
OneDrive file), PPT companion add-in reads/refreshes from it. Metadata persists per-file in
custom XML parts (ExcelApi 1.5 / PowerPointApi 1.7) keyed on hidden defined names so links
survive row/column moves (Macabacus mechanism).

| Feature | Source | What it does | Feasibility | Priority |
|---|---|---|---|---|
| Export range/chart as image to PPT | both | Range.getImage/Chart.getImage base64 -> PPT shape | shipped v2.0; column, bar, waterfall and pie charts land as native shape groups (picture fallback; on PowerPoint for Mac the group is built in sub-groups of six since v2.8.3, one flat group of the chart crashes 16.107) v2.4 | P1 (M2) |
| Link registry + refresh one object | both | Stable link IDs, update in place preserving position/size | shipped v2.0 | P1 (M2) |
| Link Manager pane (by slide / by source, batch update, filters) | both | Central link control | shipped v2.0 | P1 (M3) |
| Slide picker for inserts | - | Choose the target slide before an Insert or Paste latest linked, instead of always the active slide | shipped v2.8 | P1 (M3) |
| Spot placement (halves, quarters, whole slide, selected shape) | - | Choose where on the slide the object lands, fitted and centred | shipped v2.8 | P1 (M3) |
| Change source / versioning (Model_v4 detection) | both | Repoint links, newest-file prompts | shipped v2.1 | P2 (M3) |
| Export text (cell -> text box) | both | Live text links | shipped v2.5 (one cell, whole-box; runs inside a sentence deferred) | P2 (M3) |
| Native PPT table export | US | Editable table, format survives refresh | Partial (PPT table API at 1.8; format drift risk) | P3 |
| Preserve cell visibility / image width options | US | Advanced export toggles | Backend | P3 |
| Data Pack (deck's sources -> one workbook) | US | Collect linked content | Backend | P3 |
| Highlight linked cells in Excel | US | Show what is linked out | Yes (once registry exists) | P2 (M3) |

## 9. PowerPoint companion (Milestones 2-4, separate manifest)

Priorities within the PPT add-in once it exists: agenda/TOC engine (sections, dividers,
breadcrumbs, one-click refresh — PowerPointApi supports slide/shape/text manipulation),
Smart Align/Swap/Select Similar/Smart Painter (shipped v2.6 as the PowerPoint Tools tab: align,
distribute, match size, select similar, swap, Smart Painter for solid fill + outline), Slide Check subset (fonts,
placeholders, double spaces, alignment), cross-references + footnotes, library insert
(slides from published decks), templates. Track Changes, AI checks, Dynamic Library,
Logo Finder, Proposal Wizard = P3/enterprise. Deck Check full parity and native-UI
suppression are out of reach; additive checks only.

## 10. Brand, settings, enterprise (Milestone 4)

| Feature | Source | What it does | Feasibility | Priority |
|---|---|---|---|---|
| Shared org settings/themes (publish to team) | both | Admin-published palette/formats | Backend (or Graph file) | P2 |
| Settings export/import | both | XML/JSON round-trip | shipped v1.1 (JSON) | - |
| Shortcut manager (remap, conflicts, print list) | both | User remapping | Yes (replaceShortcuts API, signed-in users) | shipped v2.7 (Excel pane: remap, clashes incl. shipped defaults, printable card) |
| Centralized deployment | both | M365 admin rollout | Yes (Integrated Apps) | P2 |
| Content library (shared slides/ranges/templates) | both | Org content store | Backend | P3 |
| Tombstone generator / dynamic library | both | Deal content from data | Backend | P3 |
| AI assistant (rephrase, consistency check, formula suggestions) | both | LLM layer | Yes (our backend + API) | P3 |
| Track changes / workbook discussions | MB/US | Review workflows | Backend | P3 |
| Corporate dictionary / proofing rules | MB | Terminology enforcement | Partial | P3 |

## 11. Deltas from the wider competitor set (added 27.08; sources in docs/research/competitor-deltas.md and user-gaps.md)

| Feature | Inspired by | Feasibility | Priority |
|---|---|---|---|
| pls,fix Undo: snapshot + restore last bulk action | #1 Macabacus trust complaint (undo broken) | shipped v1.1 | P1 — fold into C5 |
| Palette legend insert (auto color key block) | F1F9 "Keys" | shipped v1.1 | P1 — fold into C3 |
| =PLSFIX.ROUND consistent-rounding custom functions | think-cell TCROUND | shipped v2.1 | P2 (v1.x) |
| Reconciliation solver (subset-sum: which cells make up a variance) | Kutools "Make Up a Number" | shipped v2.6 ("Find a combination", 34 cells, meet-in-the-middle) | - |
| Tornado chart builder | PowerUser | shipped v2.1 | P2 (v1.x) |
| Unpivot selection | PowerUser + Ablebits | shipped v2.1 | P2 (v1.x) |
| Workbook diff vs uploaded version (insertion-aware) | Arixcel | Partial (parse uploaded .xlsx in-pane, diff vs live workbook) | P2/P3 |
| One-click whole-deck link refresh; move-resilient link paths | user gap #7 + empower | shipped v2.0 | P1 within M3 |
| Deck sanitize, stamps, Smart Fields, Gantt | think-cell/PowerUser | PPT companion scope | P3 |
| Circularity logic detection, model risk score | OAK | Partial (heavy) | P3 |
| ERP roll-forward, module library | Modano/bpmToolbox | Backend, different product class | P3/X |
| Number-to-words, fuzzy dedupe, multi goal seek, currency conversion | Kutools/Ablebits/QuickCel | Yes (pure/simple) | P3 |
| First-run card + printable shortcut card | user-gaps #4 discoverability | shipped v2.5 | - |

Positioning note from user research: Macabacus and UpSlide are Windows-COM/VSTO only by their own
docs; we are Office.js = native on Mac, web and Windows. Lead with it. Undo reliability and
large-file performance are the two loudest trust complaints — both are v1 quality bars.

## Not worth cloning / impossible

- Native paste interception (Ctrl+V hooks), app-level events, cross-workbook reads, arbitrary
  context menus, suppressing native UI: **hard blockers** in web add-ins. Design equivalents
  are button/shortcut-driven.
- Undo/redo custom stacks (MB): Office.js writes are not grouped into native undo; a partial
  "revert last SMT action" would need our own state snapshots — costly, low value now. X.
- Airplane mode (logo redaction), disabled-keys, Outlook signature manager: niche or out of scope. X.
- PDF-to-Excel (US): Daniel already owns this problem space in Report2Excel. X here.

## Ranked next builds (proposal — the v0.3 slice)

1. **Shared runtime + keyboard shortcuts + ribbon commands** (foundation for everything).
2. **Format cycles**: number-format families, title/result/item rows, fill/font/border cycles — all palette-driven (theme engine from v0.2 does the heavy lifting).
3. **Autocolor v2**: cross-sheet/external/partial-input classes, legend + custom colors, optional on-edit.
4. **Formula Audit overlay** (consistency striping) + **Smart Track pane** (precedents/dependents).
5. **Fast Fill auto-extent** + **paste-special suite** (values/formats/transpose/preserve-formulas).
6. **Waterfall builder + CAGR arrow + chart Smart Format**.
7. **Quick CAGR, sign flip, decimals steppers, IFERROR unwrap/custom value**.
8. **Workbook TOC + Explorer pane + name scrubber**.

## Proposed ROADMAP.md deltas: awaiting Daniel's sign-off; ROADMAP.md is hand-curated

Milestone 1 is now fully shipped at v1.1. Already-checked M1 lines aren't relisted; every line
still unchecked in `ROADMAP.md` is done:

- "- [ ] Ribbon tab and keyboard shortcuts (shared runtime)" — shipped v1.1
- "- [ ] Palette-driven formatting cycles" — shipped v1.1
- "- [ ] Autocolor v2: external links, partial inputs, legend" — shipped v1.1
- "- [ ] Formula consistency overlay" — shipped v1.1
- "- [ ] Fast fill auto-extent and paste-special suite" — shipped v1.1
- "- [ ] Quick CAGR, sign flip, decimal steppers" — shipped v1.1
- "- [ ] Sheet explorer pane and name scrubber" — shipped v1.1
- "- [ ] Workbook table of contents" — shipped v1.1
- "- [ ] Precedent/dependent navigation" — shipped v1.1 (as the Smart Track pane, single-workbook only)
- "- [ ] Waterfall chart builder" — shipped v1.1

Milestone 2 (linking PoC) — done at v2.0, one line only half done:

- "- [ ] PowerPoint companion manifest and task pane" — shipped v2.0 (`pptpane.html`)
- "- [ ] Export an Excel range as a high-resolution image" — shipped v2.0
- "- [ ] Persist workbook, worksheet, and range link metadata" — shipped v2.0 (hidden-name anchor + workbook registry)
- "- [ ] Choose the link relay: Rust service or Microsoft Graph" — shipped v2.0 (chose the Rust relay, `server/`)
- "- [ ] Update one PowerPoint object from its Excel source" — shipped v2.0
- "- [ ] Detect missing and ambiguous sources" — half done: missing-source detection shipped v2.0
  ("source missing" status); ambiguous/version detection is the still-open M3 item below.

Milestone 3 (link manager) — mostly done at v2.0, three lines still open:

- "- [ ] List and filter all links in a presentation" — listing shipped v2.0 (by slide, source,
  status); filtering was not built.
- "- [ ] Update selected, slide, or all links" — shipped v2.0
- "- [ ] Preserve position and size during refresh" — shipped v2.0
- "- [ ] Change source and resolve workbook versions" — shipped v2.1 (PowerPoint "Change
  source": re-point a tracked picture at another export from the Inbox, slide/position/size kept)
- "- [ ] Highlight linked cells in Excel" — still open
- "- [ ] Performance and failure-isolation testing" — still open (blocked on Daniel's real-Office
  pass, `tasks/AUTORESUME.md` open gate 1)

Proposed new M3 lines (work that exists but has no ROADMAP.md line yet):

- Native PowerPoint table export (editable table, not a picture) — needs the spike
  (`tasks/AUTORESUME.md` open gate 2) and PowerPointApi 1.9 on the team's builds first.
- Revert a link to its previous relay revision — the relay already keeps 2 revisions
  (`server/store.rs`); no UI action uses the older one yet.
