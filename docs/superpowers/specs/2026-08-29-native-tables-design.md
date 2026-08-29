# Native PowerPoint tables for pls,fix links (design)

Date: 2026-08-29. Status: approved by Daniel ("Can we have it added as an actual table
not an image"). Implements the deferred `kind: "table"` from the 2026-08-28 links design.

## What it does

"Export as table" in the Excel Links tab sends the selected range to PowerPoint as an
editable native table (`shapes.addTable`, PowerPointApi 1.8), tracked exactly like a
picture link: the same hidden-name anchor in Excel, the same registry entry (kind
`table`), the same relay payload envelope and inbox item, the same shape tags in
PowerPoint. Update rewrites the table's cells in place; when the source grew or shrank the
table is re-created at the same left/top/width and re-tagged. Pictures stay as they are.

## Model (`src/link/model.ts`)

- `LinkKind = "range" | "chart" | "table"`. A table link's source is a range; its anchor is
  the hidden name, its label `sourceLabel(src, "table")` = `${sheet}!${ref} table`.
- Payload becomes a union: `PicturePayload` (today's shape, `kind: "picture"`) and
  `TablePayload`:
  ```ts
  interface TableCell { t: string; b?: true; i?: true; c?: string; f?: string; a?: "l" | "c" | "r"; z?: number }
  interface TablePayload { v: 1; kind: "table"; rows: number; cols: number; cells: TableCell[][]; widths: number[]; src: Source; pushedAt: string; hash: string }
  ```
  `t` is the cell's displayed text (Excel `text`), `b`/`i` bold/italic, `c` font colour,
  `f` fill colour (absent = no fill), `a` horizontal alignment, `z` font size in points,
  `widths` column widths in points. Optional keys are omitted when default, so a plain
  table stays small. Caps: `TABLE_MAX_ROWS = 60`, `TABLE_MAX_COLS = 20`; over the cap the
  export throws `Tables go up to 60 rows and 20 columns; export a picture for more.`
- `decodePayload` validates both kinds field by field; `hash` = the existing hash helper
  over the JSON of `cells`.

## Excel (`src/excel/`)

- `links.ts`: `exportSelectionAsTable(ws, relay)` mirrors `exportSelection`: same anchor,
  registry entry with kind `table`, `renderAnchored` with a resolved source of kind
  `table`, same publish/rollback and inbox item.
- `link-anchors.ts`: `renderSource` gains the table branch: one `getCellProperties` with
  `{ format: { font: { bold, italic, color, size }, fill: { color }, horizontalAlignment } }`
  plus `range.load("text")` and each column's `format.columnWidth` in one sync, mapped
  to `TableCell[][]`. The fake host already answers `getCellProperties` and column widths;
  add `text` there if it does not.
- Push, auto-push, remove and `resolveSources` treat kind `table` like `range` (the anchor
  is a range). The table's `hash` compares payloads the way the picture hash does.

## PowerPoint (`src/ppt/`)

- `host.ts`: `insertLink` dispatches on `payload.kind`. Tables: require
  `PowerPointApi 1.8` (else throw `Tables need PowerPoint 2021 or Microsoft 365.`), then
  `slide.shapes.addTable(rows, cols, { left, top, width, height, values })` with
  `values = cells.map(row => row.map(c => c.t))`; per cell `getCellOrNullObject(r, c)`:
  `text` is already set, then `font.bold/italic/color/size` when present, `fill.color`
  when present, `horizontalAlignment` mapped from `a`. Name `pls,fix table ${label}`,
  tags as for pictures. Width = the sum of `widths` capped to the slide width minus
  margins; height = rows * 18 pt (PowerPoint grows rows to fit); placement through
  `placeInFreeSpace` from `src/layout.ts` (slide 960x540, margin 36, gap 12) over the
  boxes of the slide's other shapes, skipping empty placeholders
  (`textFrame.hasText === false`); pictures use the same placement from now on, and the
  pane toasts `Placed over other objects: no free space on this slide` when
  `overlapping` comes back true.
- `refreshLink` for a table: `shape.getTable()`; load `rowCount, columnCount`; equal to the
  payload -> write every cell's text and formats in place, geometry untouched; different
  -> delete the shape, `addTable` at the old left/top/width, write cells, re-tag (the
  reinsertion path already used below 1.8 for pictures). `refreshLinks` (the batched
  repaint) keeps handling pictures only; table rows fall to `refreshLink` through the
  existing per-row fallback in `paintBatch`. `splitByBytes` counts a table payload by its
  JSON length.
- `status.ts`: `aspectChanged` is picture-only; tables never touch height on refresh.
- `views.ts`: the inbox row and the links table show the kind (`table`) in the meta text.
- Revert and Change source go through the same refresh path unchanged.

## Fake PowerPoint host (`test/fakeppt/`)

- `addTable(rows, cols, options)` creates a shape of type `Table` with a `table` model
  `{ rowCount, columnCount, cells: {text, font: {bold, italic, color, size}, fill: {color}, horizontalAlignment}[][] }`.
- `ShapeProxy.getTable()` -> `TableProxy` (`rowCount`, `columnCount`, `getCellOrNullObject`)
  -> `TableCellProxy` with settable `text`, `font.*`, `fill.color`, `horizontalAlignment`.
  Strict-load rules in `strict.ts` for the new objects, like the existing ones. A slide's
  placeholder shapes expose `textFrame.hasText`.
- `isSetSupported("PowerPointApi", "1.8")` decides the table path, as for groups today.

## Pane

- Excel Links tab: third export button `#export-table` ("Export as table", "The selected
  range as an editable table"); `links-tab.ts` wires it like the other two.
- PowerPoint pane: no new controls; Insert/Update/Revert/Break work on tables as they do
  on pictures.

## Tests (mirror the modules)

- `src/link/model.test.ts`: table payload round trip, validation of every field, the caps.
- `test/links.integration.test.ts`: export as table registers a table entry, the payload
  carries text/bold/fill/alignment/widths from a formatted fake range, push re-renders.
- `test/ppt.*.integration.test.ts` (new file `ppt.table.integration.test.ts`): insert
  creates a `Table` shape with tags and cell formats in free space; update with equal
  dimensions rewrites cells and leaves left/top/width/height alone; update after the
  source grew re-creates at the same left/top/width and re-tags; below PowerPointApi 1.8
  insert throws the message; revert repaints the older cells; a second insert lands beside
  the first, not on it.
- `src/layout.test.ts` exists; extend only if the placement helpers change.

## Constraints (from CLAUDE.md)

Files at most 400 lines, functions at most 50, one concept per file, header comment on
every new file, no `any`, every scalar read preceded by `load()` + `sync()` (the strict
fakes throw otherwise), manifests untouched (PowerPointApi 1.8 is checked at runtime),
`npm run check` green, CLAUDE.md Map updated in the same change, README and the Latvian
guide get one surgical sentence each. No version bump: the controller bumps after merge.
