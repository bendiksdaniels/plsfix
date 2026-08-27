# Model Tools

Model Tools is an early Excel productivity add-in for financial modelling teams. The first milestone focuses on fast, consistent workbook formatting and common formula operations.

## What works in v0.2

- Live selection summary: cells, formulas, blanks, and formula errors
- Five financial-model formatting presets
- Whole-number, decimal, euro, and percentage number formats
- Fill-formula-right and fill-formula-down actions
- Safe `IFERROR(..., 0)` wrapping
- Financial-model autocoloring: hardcodes blue, formulas black, external links green
- Multiply or divide selected constants and formulas by 1,000
- Brand tab: company palette (pickers, hex entry, or logo upload with local color extraction), font and currency settings, JSON import/export; all presets and autocolor follow the palette; persisted in the task pane

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
