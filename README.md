# Model Tools

Model Tools is an Excel productivity add-in for financial modelling teams. It is a web add-in, so unlike the COM/VSTO incumbents it runs natively on Windows, Mac and Excel on the web. The first milestone covers fast, consistent workbook formatting, model auditing and common formula operations.

## What works in v1.0

- Live selection summary: cells, formulas, blanks, and formula errors
- Five financial-model formatting presets
- Whole-number, decimal, euro, and percentage number formats
- Format cycles: press the same key again to step through number formats (general, date, currency, percent, multiple), title, result and item row styles, brand fill or font colors, borders (bottom rule, total, result line, box, grid), and row heights or column widths
- Paintbrush: three slots that capture the active cell's number format, font, fill, alignment and border rules and paint any of them over a selection; the slots stay on the machine, so they survive closing the pane
- Fill, paste and undo: fills sized by the neighbouring column or row, a marked copy source pasted as values, formats, exact formulas or transposed, quick CAGR, sign flip and decimal steppers, and **Undo last Model Tools action**, which puts back the formulas, number formats and colors the add-in last wrote (Office.js writes never reach Excel's own undo stack)
- Reversible `IFERROR(..., 0)` guard: the same action strips it again
- Financial-model autocoloring: hardcodes blue, formulas black, cross-sheet links green
- Autocolor v2: external-file links and numbers hardcoded inside formulas get their own palette colors, an optional on-edit toggle recolors as you type, and "Insert color key" drops the legend on the sheet
- Multiply or divide selected constants and formulas by 1,000
- Audit: a reversible formula-consistency overlay whose snapshot travels with the file, so reopening a workbook saved mid-audit restores your original fills (striped where a formula matches its neighbours, soft red where one breaks the pattern) and a Smart Track panel that walks direct precedents and dependents
- Charts: a native waterfall built from a two-column bridge table with branded opening, closing, up and down columns and a reconciliation of the deltas against the closing total, a one-click brand restyle of any selected chart, and a CAGR callout beside the selected series
- Tornado chart: a driver, low and high table becomes a sensitivity bar chart ranked by swing, plotted as deltas from the base stated above the outcome columns (or their mean when none is)
- Unpivot selection: a cross-tab block is rewritten as Row / Column / Value lines on a new sheet, blanks skipped and the source untouched
- Consistent rounding: `=SMT.ROUND` and `=SMT.ROUNDSUM` custom functions round a row or column so the rounded numbers still add up to the rounded total (largest remainder, think-cell TCROUND style), written beside the selection by one button in the Model tools list
- Brand tab: company palette (pickers, hex entry, or logo upload with local color extraction), font and currency settings, JSON import/export; all presets and autocolor follow the palette; persisted in the task pane and saved with the workbook, so a model keeps its brand when it is opened on another computer
- the add-in's ribbon tab with one-click commands (autocolor, fills, IFERROR) and customizable keyboard shortcuts via the shared runtime (`public/shortcuts.json`)
- Workbook tab: a sheet explorer that jumps to, hides and shows sheets (very hidden ones are listed but never touched), a hyperlinked contents sheet rebuilt on demand, and a scrubber that finds and deletes defined names left pointing at `#REF!`
- Super Find: one search across every sheet (hidden ones included) over values, formula text, defined names and sheet names, listed in workbook order with one click to jump to the hit
- Prepare for sharing: one pass that puts every visible sheet back at A1 and leaves the workbook on the first of them, then reports what a reader would still find - hidden sheets, links to other workbooks, names left on `#REF!` and autocolor still running on every edit; nothing is deleted, hidden sheets are untouched, and zoom cannot be reset because Office.js does not expose it

## Keyboard shortcuts

All shortcuts run without the pane open (shared runtime). Excel shows a one-time
conflict dialog where a combination shadows a native one; the number-format cycles
deliberately sit on the native format keys so existing muscle memory lands on the
branded equivalent. Users can remap under Office add-in shortcut preferences.

| Action | Keys |
|---|---|
| Open Model Tools | Ctrl+Shift+M |
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
| Undo last Model Tools action | Ctrl+Shift+Z |
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
- Undo restores the full range state of the last action (formulas, number formats,
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

The task pane runs locally. It has no backend and sends no workbook data anywhere.

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
leaves the desktop add-in pointed at `localhost`.

## Linked objects in PowerPoint (v2)

Export a range or a chart from Excel and keep it fresh in a deck without re-pasting.

- Excel, tab **Links**: *Export selection* or *Export chart* renders a picture, anchors the
  source with a hidden defined name (`SMT_LINK_<id>`, so rows can be inserted above it) or the
  chart's name, and sends the picture to the relay. *Push all* re-renders every link through
  its anchor. Generate the **link key** once under Links > Settings and paste it into
  PowerPoint once. **Auto-push on edit** (a tick box under the list) re-pushes a link
  three seconds after the last edit inside it, so a deck's *Update all* always finds the
  current picture; it runs only while the pane is open and is remembered per workbook.
- PowerPoint, tab **Model Tools > Links**: the **Inbox** lists exports waiting to be placed;
  *Insert* puts a picture on the selected slide with a tracker in the shape's tags. The
  **Links** list shows every tracked picture in the deck (slide, source, status) and
  *Update selected / slide / all* repaints them in place: position and size are kept, only
  the height follows when the picture's aspect ratio changed. *Break link* removes the
  tracker and leaves the picture.
- The relay stores only encrypted blobs for 7 days; anyone holding the deck can pull a
  linked picture for that long, so **break links before sending a deck outside**.

## Architecture

- `taskpane.html` and `src/main.ts`: task-pane UI and action routing
- `src/excel.ts`: Office.js integration; workbook data remains in the Excel process
- `src/model.ts`, `src/cycles.ts`, `src/classify.ts`, `src/audit.ts`, `src/paste.ts`, `src/chartmath.ts`, `src/workbook.ts`: pure, tested logic (cycling, classification, auditing, paste math, bridge math, TOC/name hygiene)
- `src/settings.ts`: brand palette model, theme derivation, logo color extraction, persistence helpers
- `manifest.xml`: Excel add-in identity, permissions, and local development URL

TypeScript is used because an Office.js task pane is a web front end. A future cloud service - for shared brand libraries, authentication and link metadata - should be implemented in Rust.

## Next milestone

The next high-value slice is an Excel-to-PowerPoint link proof of concept:

1. Export a selected range or chart as an image.
2. Store a stable link identifier and source metadata.
3. Add a PowerPoint companion add-in that refreshes one linked object.
4. Test renamed, moved, and duplicated workbooks before expanding to bulk refresh.

See `ROADMAP.md` for the staged product plan.
