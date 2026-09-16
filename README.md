# pls,fix

Named after the email every analyst knows.

[![check](https://github.com/bendiksdaniels/plsfix/actions/workflows/check.yml/badge.svg)](https://github.com/bendiksdaniels/plsfix/actions/workflows/check.yml) [![release](https://img.shields.io/github/v/release/bendiksdaniels/plsfix)](https://github.com/bendiksdaniels/plsfix/releases/latest) [![license: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

pls,fix is an Excel productivity add-in for financial modelling teams. It is a web add-in, so unlike the COM/VSTO incumbents it runs natively on Windows, Mac and Excel on the web. The first milestone covers fast, consistent workbook formatting, model auditing and common formula operations.

## Get pls,fix

Free, MIT-licensed, nothing to install but one line that tells Office where to load it from ([docs/INSTALL.md](docs/INSTALL.md) has every platform and the manual steps):

- **Windows**, in PowerShell: `irm https://github.com/bendiksdaniels/plsfix/releases/latest/download/plsfix-install-windows.ps1 | iex`
- **Mac**, in Terminal: `curl -fsSL https://github.com/bendiksdaniels/plsfix/releases/latest/download/plsfix-install-mac.command | sh`
- **Excel on the web**: Home > Add-ins > More Settings > Upload My Add-in, with [`manifest.prod.xml`](https://github.com/bendiksdaniels/plsfix/releases/latest/download/manifest.prod.xml).

Reopen Excel and PowerPoint: the **pls,fix** tab is on the ribbon, Ctrl+Shift+M opens the pane. `pls,fix Demo Model.xlsx` from the same release has a "Start here" sheet that walks through every tool, and `pls,fix Demo Deck.pptx` demos the PowerPoint Slide and Where pickers. The Microsoft AppSource listing (Excel > Add-ins > search "pls,fix") is in preparation.

For IT teams and developers: [docs/SELF-HOSTING.md](docs/SELF-HOSTING.md) (your own server, `docker compose up`, a manifest re-pointed at your address) and [CONTRIBUTING.md](CONTRIBUTING.md) (build from source).

## What works in v1.0

- Live selection summary: cells, formulas, blanks, and formula errors
- Five financial-model formatting presets
- Whole-number, decimal, euro, and percentage number formats
- Format cycles: press the same key again to step through number formats (general, date, currency, percent, multiple), title, result and item row styles, brand fill or font colors, borders (bottom rule, total, result line, box, grid), and row heights or column widths
- Hygiene cycles (v2.7): indent 0 / 1 / 2 / 3, alignment left / center / right / general and underline single / double / none, each stepping from the cell's read-back state; Pinstripes band every second row (or column) of the selection in the brand tint and a second press clears them, refusing while an audit or autocolor overlay owns the fills
- Paintbrush: three slots that capture the active cell's number format, font, fill, alignment and border rules and paint any of them over a selection (Save 1-3 / Use 1-3); the slots are saved with the workbook, so a shared model carries its formatting kit, and a workbook that never saved any falls back to the slots kept on the machine
- Fill, paste and undo: fills sized by the neighbouring column or row, a marked copy source pasted as values, formats, exact formulas or transposed, quick CAGR, sign flip and decimal steppers, and **Undo last pls,fix action**, which puts back the formulas, number formats and colors the add-in last wrote (Office.js writes never reach Excel's own undo stack)
- Three narrow pastes (v2.7): Paste: duplicate formulas (references inside the copied block move with it, references outside it stay on their cells, across sheets too), Paste number formats only, Paste row heights only
- Reversible `IFERROR(..., 0)` guard: the same action strips it again
- Financial-model autocoloring: hardcodes blue, formulas black, cross-sheet links green
- Autocolor v2: external-file links and numbers hardcoded inside formulas get their own palette colors, an optional on-edit toggle recolors as you type, and "Insert color key" drops the legend on the sheet
- Multiply or divide selected constants and formulas by 1,000
- Audit: a reversible formula-consistency overlay whose snapshot travels with the file, so reopening a workbook saved mid-audit restores your original fills (striped where a formula matches its neighbours, soft red where one breaks the pattern) and a Smart Track panel that walks direct precedents and dependents
- Select consistent region and Precedents of selection (v2.7): the rectangle whose every cell carries the active cell's formula in R1C1 form, and the direct precedents of up to 50 selected cells grouped by cell, the same-sheet ones selected
- Charts: a native waterfall built from a two-column bridge table with branded opening, closing, up and down columns and a reconciliation of the deltas against the closing total, a one-click brand restyle of any selected chart, and a CAGR callout beside the selected series
- Tornado chart: a driver, low and high table becomes a sensitivity bar chart ranked by swing, plotted as deltas from the base stated above the outcome columns (or their mean when none is)
- Football field (v2.7): a method, low and high table becomes the valuation-range chart, stacked bars with a cleared floor series, first row on top, the axis number format taken from the low column, low above high swapped and reported
- Comps stats (v2.7): six rows of live formulas (min, 25th percentile, median, average, 75th percentile, max) one blank row under the selected comps table, the number format copied per numeric column, text columns left alone
- Unpivot selection: a cross-tab block is rewritten as Row / Column / Value lines on a new sheet, blanks skipped and the source untouched
- Consistent rounding: `=PLSFIX.ROUND` and `=PLSFIX.ROUNDSUM` custom functions round a row or column so the rounded numbers still add up to the rounded total (largest remainder, think-cell TCROUND style), written beside the selection by one button in the Model tools list
- `=PLSFIX.CAGR(first, last, periods)` (v2.7): the compound annual growth rate as a custom function, `#VALUE!` on a non-positive value, fewer than one period or an overflowing ratio; all three functions carry a help link to the support page
- Templates: six ready calculation blocks written at the active cell in your brand styles (annuity debt schedule, DCF, NPV / IRR with a formula-based payback, working-capital days, a two-way sensitivity grid and an EBITDA bridge shaped for the waterfall), each refused rather than written where something already stands
- Brand tab: company palette (pickers, hex entry, or logo upload with local color extraction), font, language and currency settings (the language sets the house number style: thousands with a space in Latvian and Russian, a comma in English, the decimal always a point, the currency after the amount or before it; the pane also reads which separators Excel is set to and says where to change them), JSON import/export; all presets and autocolor follow the palette; persisted in the task pane and saved with the workbook, so a model keeps its brand when it is opened on another computer
- pls,fix ribbon tab with one-click commands (autocolor, fills, IFERROR), each with its own icon, and customizable keyboard shortcuts via the shared runtime (`public/shortcuts.json`)
- Keyboard shortcuts section (v2.7, Brand tab): remap any pls,fix key per signed-in user through the Office shortcut runtime, clashes with the effective keys named before anything is written, Reset all restores the defaults
- Workbook tab: a sheet explorer that jumps to, hides and shows sheets (very hidden ones are listed but never touched), a hyperlinked contents sheet rebuilt on demand, and a scrubber that finds and deletes defined names left pointing at `#REF!`
- Sheet tools and Clean past the data (v2.7): Unhide all (very hidden sheets only with the tick), Show only this, Bury this, Move up / down / to end above the explorer, and one button that deletes the formatted rows and columns past a sheet's values (formats only cleared when a chart or shape sits on the sheet, or the host cannot count them), with the counts in the toast
- Super Find: one search across every sheet (hidden ones included) over values, formula text, workbook-level defined names, sheet names and cell comments, listed in workbook order with one click to jump to the hit
- Style scrubber: the custom cell styles no cell in the workbook wears, listed with a count and deleted on a confirmed second click; a sheet too large to scan is named and blocks the delete, so a partial answer never removes a style still in use
- Prepare for sharing: one pass that puts every visible sheet back at A1 and leaves the workbook on the first of them, then reports what a reader would still find - hidden sheets, links to other workbooks, names left on `#REF!` and autocolor still running on every edit; nothing is deleted, hidden sheets are untouched, and zoom cannot be reset because Office.js does not expose it
- Model check: one pass that lists formula errors, hardcodes inside formulas, inconsistent formulas, volatile functions, broken names, unused styles, hidden sheets and external links, each with a jump to the cell, and a copyable report
- Find a combination: which numbers in the selected block add up to a target, within a tolerance (up to 34 numeric cells); the matching cells become the selection and the pane reports their sum and the remaining variance

## Keyboard shortcuts

All shortcuts run without the pane open (shared runtime). Excel shows a one-time
conflict dialog where a combination shadows a native one; the number-format cycles
deliberately sit on the native format keys so existing muscle memory lands on the
branded equivalent. Users can remap under Office add-in shortcut preferences.

The pane's Tools tab opens a printable version of this table (Shortcut card).

| Action | Keys |
|---|---|
| Open pls,fix | Ctrl+Shift+M |
| Autocolor selection | Ctrl+Shift+K |
| Fill formula right | Ctrl+Alt+R |
| Fill formula down | Ctrl+Alt+D |
| Wrap with IFERROR | Ctrl+Shift+I |
| Multiply by 1,000 | Ctrl+Shift+Alt+8 |
| Divide by 1,000 | Ctrl+Shift+Alt+9 |
| Cycle general number format | Ctrl+Shift+1 |
| Cycle date format | Ctrl+Shift+2 |
| Cycle currency format | Ctrl+Shift+4 |
| Cycle percent format | Ctrl+Shift+5 |
| Cycle multiple format | Ctrl+Shift+6 |
| Cycle title row style | Ctrl+Shift+H |
| Cycle result row style | Ctrl+Shift+R |
| Cycle item row style | Ctrl+Shift+E |
| Cycle fill color | Ctrl+Shift+F |
| Cycle font color | Ctrl+Shift+G |
| Cycle borders | Ctrl+Shift+7 |
| Cycle row height | Ctrl+Shift+Y |
| Cycle column width | Ctrl+Shift+X |
| Apply paintbrush slot 1 | Ctrl+Shift+D |
| Apply paintbrush slot 2 | Ctrl+Shift+N |
| Apply paintbrush slot 3 | Ctrl+Shift+O |
| Toggle audit overlay | Ctrl+Shift+A |
| Trace precedents | Ctrl+Alt+Q |
| Trace dependents | Ctrl+Alt+W |
| Undo last pls,fix action | Ctrl+Shift+Z |
| Mark copy source | Ctrl+Shift+C |
| Paste values | Ctrl+Shift+V |
| Paste formats | Ctrl+Alt+F |
| Paste formulas exactly | Ctrl+Alt+P |
| Paste transposed | Ctrl+Alt+T |
| Insert CAGR | Ctrl+Shift+Q |
| Flip sign | Ctrl+Shift+J |
| One more decimal | Ctrl+Shift+Alt+3 |
| One less decimal | Ctrl+Shift+Alt+7 |
| Waterfall from selection | Ctrl+Shift+B |
| Brand-format chart | Ctrl+Alt+G |
| CAGR label | Ctrl+Alt+K |
| Insert contents sheet | Ctrl+Alt+O |
| Find in workbook | Ctrl+Shift+Alt+F |

## Known limits

- Row style cycles apply per row and support up to 500 rows at once.
- Undo restores the full range state of the last five actions, newest first, up to 25,000 cells across them (formulas, number formats,
  fills, fonts, borders, alignment, wrapping and indent) up to 5,000 cells; larger
  actions are refused or run without undo and say so in the toast. Row heights,
  column widths (including their cycles) and chart, shape, sheet and defined-name
  operations sit outside undo.
- The waterfall's closing total needs one manual right-click > Set as Total: Office.js
  exposes no API for it.
- Smart Track tracing needs ExcelApi 1.12 (precedents) / 1.13 (dependents); older
  builds get a clear message. Everything else runs on the manifest floor (1.9).
- Tracing and auditing work within the open workbook; Office.js cannot cross into
  other files.

The task pane runs locally and reads the workbook in the Excel process. The only data that
leaves the machine is a linked object's rendered picture, encrypted in the pane before it is
uploaded (see *Linked objects in PowerPoint*); a workbook with no links sends nothing anywhere.
Every section of both panes carries a `?` beside its heading that opens a short explanation of
what that section is for and what each of its buttons does.

## Run locally

Requirements: Microsoft Excel with Microsoft 365, Node.js 20+, and npm.

```bash
npm install
npm run validate
npm start
```

The first start installs a trusted local HTTPS development certificate and sideloads `manifest.xml` into desktop Excel. Use `npm stop` when finished. To sideload the same manifest into desktop PowerPoint instead, use `npm run start:ppt` and `npm run stop:ppt`.

For browser-only UI development, use `npm run dev` and open `https://localhost:3000/taskpane.html`. Excel actions only work when the page is hosted inside Excel.

## Quality checks

```bash
npm test
npm run build
npm run validate
npm run ux:check  # needs `npm run dev` running first: both panes at 320/360/420px, screenshots + defect list
npm run ux:sweep  # both panes with no Office host at all: every control clicked, dead ones and broken tab or help states fail (reaction only; routing is the dispatch tests)
```

## Deploy

`manifest.xml` is the localhost development manifest; `manifest.prod.xml` is the same
add-in (same GUID) served live from `https://dbautomatizacijas.com/modelis/` (suite key
`modelis`: Rust static host in `server/`, deployed with the gateway's
`deploy.sh modelis`). The path carries a Cloudflare Access bypass - an access-gated
pane cannot load inside Office webviews - so availability is restricted through
Microsoft 365 centralized deployment group assignment instead. Hosted JS updates need
no admin action; manifest changes need a re-upload. Details and sources:
`docs/research/launch-path.md`. `npm stop` restores `manifest.prod.xml` into Excel's
local sideload folder (`scripts/wef-restore-prod.sh`), so ending a dev session never
leaves the desktop add-in pointed at `localhost`. Any other host runs the same server from
the container image (`docs/SELF-HOSTING.md`); its `/manifest.xml` hands out a manifest
re-pointed at `MODELIS_PUBLIC_URL` with an add-in id of its own.

## Linked objects in PowerPoint (v2)

Export a range or a chart from Excel and keep it fresh in a deck without re-pasting.

- Excel, tab **Links**: *Export selection* or *Export chart* renders a picture, anchors the
  source with a hidden defined name (`PLSFIX_LINK_<id>`, so rows can be inserted above it) or the
  chart's name, and sends the picture to the relay. *Export as table* sends the same range as an
  editable PowerPoint table instead (up to 60 rows and 20 columns, PowerPoint 2021 or Microsoft 365).
  A table whose first row is all bold keeps PowerPoint's header row; *Update all* leaves the deck's
  table style alone and clears only the fills it painted before.
  *Export as text* sends one cell's displayed text as a text box that keeps its place, size and
  font when it refreshes (up to 500 characters). A column, bar, line, waterfall or pie
  chart lands as a group of editable shapes with value labels (PowerPoint 2504/16.96 and newer; pie wedges 2601/16.105; on the Mac the group holds sub-groups of six, the shape PowerPoint for Mac can take), a
  picture elsewhere or past 40 points, 6 series or 12 slices. *Push all* re-renders every link through
  its anchor. Generate the **link key** once under Links > Link key and paste it into
  PowerPoint once. **Auto-push on edit** (a tick box under the list) re-pushes a link
  three seconds after the last edit inside it, so a deck's *Update all* always finds the
  current picture; it runs only while the pane is open and is remembered per workbook.
  **Highlight linked cells** (a second tick box) tints every linked range so you can see what
  feeds the deck, and puts the original fills back when it is switched off; charts are not tinted.
  **Projects** group the list: pick or create a project above Linked objects, new exports join it,
  *Push all* covers the shown project, and *Move to project* reassigns ticked links. Old links
  without a name stay under No project.
- PowerPoint, tab **pls,fix > Links**: the **Inbox** lists exports waiting to be placed;
  *Insert* puts a picture on the selected slide with a tracker in the shape's tags, and
  *Paste latest linked* does the same for the newest export in one click. The
  **Links** list shows every tracked picture in the deck (slide, source, status), can be
  searched and filtered by status, source workbook, slide and project, and the Inbox is grouped
  by project. *Update selected / slide / all* repaints them in place (*Update all* covers the
  filtered set): position and size are kept, only
  the height follows when the picture's aspect ratio changed. A round trip PowerPoint never
  answers ends after a minute with a sentence naming what stopped, and a chart insert or
  refresh that stops mid-draw becomes the picture with the reason under *Update all*. *Revert last update* puts the
  ticked rows back to the previous render, the one revision the relay still holds. *Change
  source* points the one ticked picture at another export waiting in the Inbox - the same table
  from a newer workbook, say - keeping its slide, position and size. *Break
  link* removes the tracker and leaves the picture.
- Insert placement (v2.8): two pickers above the Inbox list choose where *Insert* and *Paste
  latest linked* land - **Slide** (the active slide, or any slide by number) and **Where**
  (*Free space* by default, the selected shape, a half, a quarter or the whole slide). Halves,
  quarters and the whole slide are fitted to the slide's content area with the object's aspect
  kept and centred, and an empty layout placeholder is removed once something is fitted into it.
- PowerPoint, tab **Tools**: object tools for the shapes selected on the slide - align (left,
  centre, right, top, middle, bottom), distribute across or down, match size to the first
  selected object, select similar (same type and size on the slide), swap two positions - and a
  **Smart Painter** that captures one object's solid fill and outline and applies them to the
  objects you select next (PowerPoint 2021 or Microsoft 365). The same tools sit on the
  pls,fix ribbon in PowerPoint, each with its own icon: an **Objects** group (Object tools,
  Match size, Select similar, Swap, Capture style, Apply style) and an **Arrange** group (the
  six alignments and both distributions), so a deck can be tidied without opening the pane.
- The relay stores only encrypted blobs for 30 days; anyone holding the deck can pull a
  linked picture for that long, so **break links before sending a deck outside**.
- Every relay call gives up after 20 s on the client and 30 s on the server (v2.7): the pane
  says "The link relay did not answer in time." instead of hanging, and a retry is one press away.

## Architecture

- `taskpane.html` and `src/main.ts`: task-pane UI and action routing
- `src/excel/`: Office.js integration, one file per feature area; workbook data remains in the Excel process
- `src/model.ts`, `src/cycles.ts`, `src/classify.ts`, `src/audit.ts`, `src/paste.ts`, `src/chartmath.ts`, `src/workbook.ts`: pure, tested logic (cycling, classification, auditing, paste math, bridge math, TOC/name hygiene)
- `src/settings.ts`: brand palette model, theme derivation, logo color extraction, persistence helpers
- `manifest.xml`: Excel add-in identity, permissions, and local development URL

TypeScript is used because an Office.js task pane is a web front end. A future cloud service - for shared brand libraries, authentication and link metadata - should be implemented in Rust.

## Roadmap

The Excel-to-PowerPoint link milestone shipped in v2 (*Linked objects in PowerPoint* above);
what comes next is in `ROADMAP.md`.

## License

MIT, see [LICENSE](LICENSE). Copyright (c) 2026 Daniels Bendiks.
