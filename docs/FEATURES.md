# Feature copy-target matrix

Consolidated from: Macabacus help center (~100 features), UpSlide help center (~140 features),
UpSlide sales-team training files (primary source), and an Office.js platform feasibility study.
Raw inventories with per-feature source URLs live in `docs/research/`.

Date: 2026-08-27. Owner add-in: Model Tools (Office.js task pane, ExcelApi >= 1.9).

**Status 27.08 (v0.9.x):** everything marked P1 in sections 0-7 shipped in chunks C1-C7 —
shared runtime + ribbon + 34 shortcuts, ten format cycles, autocolor v2 + color key,
audit overlay + Smart Track, SMT Undo + paste suite + fast fill + CAGR/sign/decimals,
native waterfall + chart formatter + CAGR label, TOC + sheet explorer + name scrubber.
Sections 8-10 (linking, PPT companion, enterprise) remain the M2-M4 roadmap. See
tasks/v1-plan.md for per-chunk detail and open follow-ups.

**Feasibility legend** — `Yes`: doable with documented Office.js APIs. `Yes*`: doable with a
caveat (noted). `Partial`: a reduced version is doable. `Backend`: needs our planned Rust
service or Microsoft Graph relay. `No`: impossible in web add-ins (COM-only).
**Priority** — P1 build next, P2 later milestone, P3 long tail, X skip.
**Done** — already shipped in v0.2.

## 0. Platform foundations (prerequisites, not user features)

| Item | Why | Feasibility | Priority |
|---|---|---|---|
| Shared runtime (manifest change, lifetime long) | Prerequisite for shortcuts, ribbon-state, pane/ribbon shared memory | Yes (SharedRuntime 1.1) | P1 |
| Custom keyboard shortcuts | The whole Macabacus/UpSlide UX is shortcut cycles | Yes* (KeyboardShortcuts 1.1; Excel Win 2102+/Mac 16.55+; some combos reserved on web; users remap) | P1 |
| Ribbon commands (buttons that run code, no pane) | One-keystroke actions without opening the pane | Yes | P1 |
| Excel contextual tab | Surface tools next to native tabs | Yes (RibbonApi 1.2, Excel only) | P2 |
| Cell right-click menu items | ContextMenuCell extension point | Partial (add items only, cannot replace native) | P2 |
| Unified + XML dual manifests | Coverage across hosts until unified manifest is universal | Yes | P2 |

## 1. Formatting and styles

| Feature | Source | What it does | Feasibility | Priority |
|---|---|---|---|---|
| Model presets (title/header/input/formula/result) | both | One-click branded cell styles | Done | - |
| Brand palette + font + currency settings | both (MB Color Palettes, US themes) | User-editable brand scheme drives all output | Done (v0.2 Brand tab) | - |
| Number format cycles (general/currency/percent/multiple/date) | both | Repeated keystroke cycles curated formats per family | Yes (Range.numberFormat + shortcuts) | P1 |
| Title / result / item row format cycles | US | Cycles approved row styles | Yes | P1 |
| Fill color cycle / font color cycle | both | Cycles brand fills/fonts | Yes | P1 |
| Border style/color cycles | both | Cycles approved borders (incl. accounting underlines) | Yes | P1 |
| Row height / column width cycles | US | Cycles preset standards | Yes (Range format) | P2 |
| More/fewer decimals | US | Steps decimal places in the active format | Yes | P1 |
| Change sign | US | Flips sign of selected constants | Yes (like our x1000) | P1 |
| Pinstripes (alternate row/col shading) | MB | Odd/even shading via conditional format | Yes (ConditionalFormat API) | P2 |
| Custom Styles / Style Cycles (user-defined, 8 slots) | MB | Reusable multi-property styles on one key | Yes (settings-driven) | P2 |
| Paintbrush (multi-slot format painter) | MB | Copy/apply formatting without clipboard, FIFO slots | Yes (read format -> store -> apply) | P2 |
| Indent cycles, center cycle, underline cycle | MB | Alignment/underline cycling | Yes | P2 |
| Footnote cycle/toggle/hide/checker (in-cell superscripts) | MB | Numbered footnote management in cells | Partial (number-format superscript tricks; font runs limited) | P3 |
| Standard Sizes (conform cells/charts to preset dims) | MB | Preset output dimensions | Yes | P3 |

## 2. AutoColor and model hygiene

| Feature | Source | What it does | Feasibility | Priority |
|---|---|---|---|---|
| Autocolor selection (inputs/formulas/links) | both | Content-based font coloring | Done (basic) | - |
| Autocolor: distinguish same-sheet vs cross-sheet vs external-file vs partial-input formulas | both | Richer classification (MB adds hardcode-in-formula detection) | Yes (parse formula text) | P1 |
| Autocolor on entry/edit | both | Live coloring as you type | Yes* (onChanged event; perf care, off by default) | P1 |
| Autocolor legend & customization | US | Legend UI + per-type color overrides | Yes (Brand tab extension) | P1 |
| Formula Audit / Formula Flow overlay | both | Striped fill = formula consistent with neighbors, solid = deviation | Yes (R1C1 compare + fills; store/restore original formats) | P1 |
| Uniformulas (select consistent region) | MB | Highlights the consistency region of active formula | Yes (same engine) | P2 |
| Dependency Density heatmap | MB | Shade by dependent count | Yes* (getDependents per cell is slow at scale; cap range) | P3 |
| Model Check (50+ automated checks) | MB | Error/structure/hidden-data/brand audit with fixes | Partial (subset: errors, hardcodes in formulas, inconsistent rows, hidden sheets; no full parity) | P2 |

## 3. Paste and fill tools

| Feature | Source | What it does | Feasibility | Priority |
|---|---|---|---|---|
| Fill right/down from edge cell | both | Copy formula across selection | Done (basic) | - |
| Fast Fill with auto-extent (no pre-selection) | US | Detects how far to fill from neighbors | Yes (scan neighbor rows/cols for extent) | P1 |
| Paste values / formulas / formats / transpose / skip blanks (button-driven) | both | Paste-special suite | Yes (Range.copyFrom modes) | P1 |
| Preserve Formulas paste (exact references) | both | Paste keeping original refs | Yes (read formulasR1C1/A1 text, write verbatim) | P1 |
| Duplicate Formulas paste | US | In-range refs adapt, external absolutes kept | Yes (formula rewrite) | P2 |
| Paste row heights / number formats only | US | Targeted format transfer | Yes | P2 |
| Intercept native Ctrl+C/Ctrl+V | - | True clipboard interception | No (hard blocker; button/shortcut-driven instead) | X |

## 4. Formula auditing and navigation

| Feature | Source | What it does | Feasibility | Priority |
|---|---|---|---|---|
| Smart Track / Trace In-Out pane (precedents/dependents tree, keyboard nav, color trail) | both | Drill into inputs and back | Yes* (getDirectPrecedents/getDependents; single workbook only — cross-workbook impossible) | P1 |
| Show all precedents for multiple cells | MB | Multi-cell trace | Yes* (same; cap cell count) | P2 |
| Super Find (values/formulas/comments across workbook) | MB | Better Find with results pane | Partial (this workbook only; no other open workbooks) | P2 |
| Explorer (workbook/sheet tree navigator) | US | Sheet navigation pane with search | Partial (sheets of THIS workbook only — Office.js cannot see other open workbooks) | P1 |
| Workbook TOC sheet | both | Hyperlinked contents sheet, auto-updating | Yes (worksheets + hyperlinks; refresh on onAdded/onNameChanged) | P1 |
| Sheet tools (move/bury/unhide-multi/activate dialog) | MB | Sheet management | Yes (visibility incl. VeryHidden) | P2 |
| NavAid (crosshair shading of selection row/col) | MB | Visual navigation aid | Yes* (onSelectionChanged + fills; perf care, restore on move) | P3 |
| Reverse rows/columns preserving formulas | MB | Reorder periods | Yes (careful formula rewrite) | P3 |

## 5. Modeling helpers

| Feature | Source | What it does | Feasibility | Priority |
|---|---|---|---|---|
| IFERROR wrap | both | Wrap/unwrap with custom fallback | Done (basic; add unwrap + custom value) | P1 |
| Scale x1000 / /1000 | both | Rescale constants+formulas | Done | - |
| Quick CAGR formula | both | Insert CAGR over range/period | Yes | P1 |
| Summary statistics block (min/max/mean/median under data) | MB | Auto stats for comps | Yes | P2 |
| Add Scenarios (toggle-driven projection cases) | MB | Scenario switch cells + duplicated rows | Yes (structured but doable) | P3 |
| Replicate Module | MB | Duplicate an analysis block across sheets | Yes | P3 |
| Custom functions (UDFs) e.g. =SMT.CAGR | - | Namespaced functions | Yes (CustomFunctions set; not iPad/perpetual<=2021) | P2 |

## 6. Charts

| Feature | Source | What it does | Feasibility | Priority |
|---|---|---|---|---|
| Waterfall / bridge builder | both | Build from intuitive table, connectors, brand colors | Yes (native Excel.ChartType.waterfall at 1.9) | P1 |
| Stacked waterfall | US | Subcategory pillars | Partial (no native type — compose from stacked columns with helper series) | P2 |
| CAGR arrow on chart | both | Data-driven growth arrow overlay | Yes* (chart shapes/annotation via series or floating shape; verify API surface) | P1 |
| Chart Smart Format (brand compliance) | both | One-click restyle any chart to brand | Yes (chart format API + palette) | P1 |
| Football field (valuation ranges) | MB | Floating-bar range chart | Yes (stacked bar with invisible base) | P2 |
| Marimekko / S-curve | US | Width-encoded 100% stacked | Partial (no native type; column-width tricks or scatter-area composition) | P3 |
| XY scatter labels | MB | Correct point labels | Yes (series data labels API) | P3 |
| Gantt (in PPT) | US | Project timelines | Partial, PPT companion scope | P3 |

## 7. Workbook cleanup and publishing

| Feature | Source | What it does | Feasibility | Priority |
|---|---|---|---|---|
| Clean: broken/hidden defined names | both | Name scrubber | Yes (NamedItemCollection) | P1 |
| Clean: unused styles | both | Style scrubber (style-ceiling fix) | Yes (StyleCollection; verify add/delete surface) | P2 |
| Clean: crop used range | US | Clear stray formatting past data | Yes | P2 |
| Prepare for sharing (formulas->values, strip comments, hidden content, reset zoom) | US | Externalize workbook | Yes | P2 |
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
| Export range/chart as image to PPT | both | Range.getImage/Chart.getImage base64 -> PPT shape | Yes (snippet-confirmed APIs; deep-verify first) | P1 (M2) |
| Link registry + refresh one object | both | Stable link IDs, update in place preserving position/size | Backend | P1 (M2) |
| Link Manager pane (by slide / by source, batch update, filters) | both | Central link control | Backend | P1 (M3) |
| Change source / versioning (Model_v4 detection) | both | Repoint links, newest-file prompts | Backend | P2 (M3) |
| Export text (cell -> placeholder) | both | Live text links | Backend | P2 (M3) |
| Native PPT table export | US | Editable table, format survives refresh | Partial (PPT table API at 1.8; format drift risk) | P3 |
| Preserve cell visibility / image width options | US | Advanced export toggles | Backend | P3 |
| Data Pack (deck's sources -> one workbook) | US | Collect linked content | Backend | P3 |
| Highlight linked cells in Excel | US | Show what is linked out | Yes (once registry exists) | P2 (M3) |

## 9. PowerPoint companion (Milestones 2-4, separate manifest)

Priorities within the PPT add-in once it exists: agenda/TOC engine (sections, dividers,
breadcrumbs, one-click refresh — PowerPointApi supports slide/shape/text manipulation),
Smart Align/Swap/Select Similar/Smart Painter (shape APIs), Slide Check subset (fonts,
placeholders, double spaces, alignment), cross-references + footnotes, library insert
(slides from published decks), templates. Track Changes, AI checks, Dynamic Library,
Logo Finder, Proposal Wizard = P3/enterprise. Deck Check full parity and native-UI
suppression are out of reach; additive checks only.

## 10. Brand, settings, enterprise (Milestone 4)

| Feature | Source | What it does | Feasibility | Priority |
|---|---|---|---|---|
| Shared org settings/themes (publish to team) | both | Admin-published palette/formats | Backend (or Graph file) | P2 |
| Settings export/import | both | XML/JSON round-trip | Done (JSON) | - |
| Shortcut manager (remap, conflicts, print list) | both | User remapping | Yes (replaceShortcuts API, signed-in users) | P2 |
| Centralized deployment | both | M365 admin rollout | Yes (Integrated Apps) | P2 |
| Content library (shared slides/ranges/templates) | both | Org content store | Backend | P3 |
| Tombstone generator / dynamic library | both | Deal content from data | Backend | P3 |
| AI assistant (rephrase, consistency check, formula suggestions) | both | LLM layer | Yes (our backend + API) | P3 |
| Track changes / workbook discussions | MB/US | Review workflows | Backend | P3 |
| Corporate dictionary / proofing rules | MB | Terminology enforcement | Partial | P3 |

## 11. Deltas from the wider competitor set (added 27.08; sources in docs/research/competitor-deltas.md and user-gaps.md)

| Feature | Inspired by | Feasibility | Priority |
|---|---|---|---|
| SMT Undo: snapshot + restore last bulk action | #1 Macabacus trust complaint (undo broken) | Yes (we read state before writing anyway) | P1 — fold into C5 |
| Palette legend insert (auto color key block) | F1F9 "Keys" | Yes (trivial once autocolor v2 exists) | P1 — fold into C3 |
| =SMT.ROUND consistent-rounding custom functions | think-cell TCROUND | Yes (CustomFunctions set) | P2 (v1.x) |
| Reconciliation solver (subset-sum: which cells make up a variance) | Kutools "Make Up a Number" | Yes (pure TS) | P2 (v1.x) |
| Tornado chart builder | PowerUser | Yes (composed bar chart) | P2 (v1.x) |
| Unpivot selection | PowerUser + Ablebits | Yes (pure transform) | P2 (v1.x) |
| Workbook diff vs uploaded version (insertion-aware) | Arixcel | Partial (parse uploaded .xlsx in-pane, diff vs live workbook) | P2/P3 |
| One-click whole-deck link refresh; move-resilient link paths | user gap #7 + empower | Backend (M2/M3 design requirements) | P1 within M3 |
| Deck sanitize, stamps, Smart Fields, Gantt | think-cell/PowerUser | PPT companion scope | P3 |
| Circularity logic detection, model risk score | OAK | Partial (heavy) | P3 |
| ERP roll-forward, module library | Modano/bpmToolbox | Backend, different product class | P3/X |
| Number-to-words, fuzzy dedupe, multi goal seek, currency conversion | Kutools/Ablebits/QuickCel | Yes (pure/simple) | P3 |

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

## Proposed ROADMAP.md deltas (need Daniel's sign-off — his file)

- M1: add items 1-8 above; retitle "Precedent/dependent navigation" to "Smart Track pane
  (single-workbook)"; mark palette item done (already ticked).
- M2 (linking PoC): add "decide relay: Rust service vs Graph/OneDrive"; keep image-export PoC
  scope; note custom-XML-part link registry keyed on hidden defined names.
- M3 (link manager): add link versioning + highlight-linked-cells + Data Pack (stretch).
- M4 (enterprise): add shared settings publish, centralized deployment package, shortcut
  manager UI; AI layer explicitly via our own backend.
