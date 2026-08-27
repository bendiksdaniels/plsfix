# Model Tools

Model Tools is an early Excel productivity add-in for financial modelling teams. The first milestone focuses on fast, consistent workbook formatting and common formula operations.

## What works in v0.9

- Live selection summary: cells, formulas, blanks, and formula errors
- Five financial-model formatting presets
- Whole-number, decimal, euro, and percentage number formats
- Format cycles: press the same key again to step through number formats (general, date, currency, percent, multiple), title, result and item row styles, and brand fill or font colors
- Fill, paste and undo: fills sized by the neighbouring column or row, a marked copy source pasted as values, formats, exact formulas or transposed, quick CAGR, sign flip and decimal steppers, and **Undo last Model Tools action**, which puts back the formulas, number formats and colors the add-in last wrote (Office.js writes never reach Excel's own undo stack)
- Reversible `IFERROR(..., 0)` guard: the same action strips it again
- Financial-model autocoloring: hardcodes blue, formulas black, cross-sheet links green
- Autocolor v2: external-file links and numbers hardcoded inside formulas get their own palette colors, an optional on-edit toggle recolors as you type, and "Insert color key" drops the legend on the sheet
- Multiply or divide selected constants and formulas by 1,000
- Audit: a reversible formula-consistency overlay (striped where a formula matches its neighbours, soft red where one breaks the pattern) and a Smart Track panel that walks direct precedents and dependents
- Charts: a native waterfall built from a two-column bridge table with branded opening, closing, up and down columns and a reconciliation of the deltas against the closing total, a one-click brand restyle of any selected chart, and a CAGR callout beside the selected series
- Brand tab: company palette (pickers, hex entry, or logo upload with local color extraction), font and currency settings, JSON import/export; all presets and autocolor follow the palette; persisted in the task pane
- the add-in's ribbon tab with one-click commands (autocolor, fills, IFERROR) and customizable keyboard shortcuts via the shared runtime (`public/shortcuts.json`)
- Workbook tab: a sheet explorer that jumps to, hides and shows sheets (very hidden ones are listed but never touched), a hyperlinked contents sheet rebuilt on demand, and a scrubber that finds and deletes defined names left pointing at `#REF!`

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

## Architecture

- `taskpane.html` and `src/main.ts`: task-pane UI and action routing
- `src/excel.ts`: Office.js integration; workbook data remains in the Excel process
- `src/model.ts`: pure formula and grid transformations
- `src/settings.ts`: brand palette model, theme derivation, logo color extraction, persistence helpers
- `manifest.xml`: Excel add-in identity, permissions, and local development URL

TypeScript is used because an Office.js task pane is a web front end. A future cloud service—for shared brand libraries, authentication and link metadata—should be implemented in Rust.

## Next milestone

The next high-value slice is an Excel-to-PowerPoint link proof of concept:

1. Export a selected range or chart as an image.
2. Store a stable link identifier and source metadata.
3. Add a PowerPoint companion add-in that refreshes one linked object.
4. Test renamed, moved, and duplicated workbooks before expanding to bulk refresh.

See `ROADMAP.md` for the staged product plan.
