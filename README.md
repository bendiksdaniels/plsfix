# Model Tools

Model Tools is an Excel productivity add-in for financial modelling teams. It is a web add-in, so unlike the COM/VSTO incumbents it runs natively on Windows, Mac and Excel on the web. The first milestone covers fast, consistent workbook formatting, model auditing and common formula operations.

## What works in v1.0

- Live selection summary: cells, formulas, blanks, and formula errors
- Five financial-model formatting presets
- Whole-number, decimal, euro, and percentage number formats
- Format cycles: press the same key again to step through number formats (general, date, currency, percent, multiple), title, result and item row styles, and brand fill or font colors
- Fill, paste and undo: fills sized by the neighbouring column or row, a marked copy source pasted as values, formats, exact formulas or transposed, quick CAGR, sign flip and decimal steppers, and **Undo last Model Tools action**, which puts back the formulas, number formats and colors the add-in last wrote (Office.js writes never reach Excel's own undo stack)
- Reversible `IFERROR(..., 0)` guard: the same action strips it again
- Financial-model autocoloring: hardcodes blue, formulas black, cross-sheet links green
- Autocolor v2: external-file links and numbers hardcoded inside formulas get their own palette colors, an optional on-edit toggle recolors as you type, and "Insert color key" drops the legend on the sheet
- Multiply or divide selected constants and formulas by 1,000
- Audit: a reversible formula-consistency overlay whose snapshot travels with the file, so reopening a workbook saved mid-audit restores your original fills (striped where a formula matches its neighbours, soft red where one breaks the pattern) and a Smart Track panel that walks direct precedents and dependents
- Charts: a native waterfall built from a two-column bridge table with branded opening, closing, up and down columns and a reconciliation of the deltas against the closing total, a one-click brand restyle of any selected chart, and a CAGR callout beside the selected series
- Brand tab: company palette (pickers, hex entry, or logo upload with local color extraction), font and currency settings, JSON import/export; all presets and autocolor follow the palette; persisted in the task pane
- the add-in's ribbon tab with one-click commands (autocolor, fills, IFERROR) and customizable keyboard shortcuts via the shared runtime (`public/shortcuts.json`)
- Workbook tab: a sheet explorer that jumps to, hides and shows sheets (very hidden ones are listed but never touched), a hyperlinked contents sheet rebuilt on demand, and a scrubber that finds and deletes defined names left pointing at `#REF!`

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

## Known limits

- Row style cycles apply per row and support up to 500 rows at once.
- Undo restores the full range state of the last action (formulas, number formats,
  fills, fonts, borders, alignment, wrapping and indent) up to 5,000 cells; larger
  actions are refused or run without undo and say so in the toast. Row heights and
  chart, shape, sheet and defined-name operations sit outside undo.
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

The first start installs a trusted local HTTPS development certificate and sideloads `manifest.xml` into desktop Excel. Use `npm stop` when finished.

For browser-only UI development, use `npm run dev` and open `https://localhost:3000/taskpane.html`. Excel actions only work when the page is hosted inside Excel.

## Quality checks

```bash
npm test
npm run build
npm run validate
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
