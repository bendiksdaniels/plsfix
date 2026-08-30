# Native chart groups on slides (design, 2026-08-30)

## Goal

A chart link ("Export active chart" in Excel, opened from the PowerPoint inbox) lands on the
slide as a native, editable object with brand fonts and colours and labels that are just the
values, and an update rebuilds it in place. Today every chart link is a PNG.

## The fact that shapes the design

Office.js cannot create a chart object in PowerPoint. `PowerPoint.ShapeType.chart` and
`PlaceholderType.chart` exist read-only (`@types/office-js` 1.0.606, checked 30.08); OOXML
coercion (`Office.CoercionType.Ooxml`) is Word only. What PowerPoint does give an add-in:

| Capability | API set | Daniel's hosts (PowerPoint 16.107 Mac, PowerPoint for the web) |
| --- | --- | --- |
| `shapes.addGeometricShape` (`Rectangle`, `Pie`, `BlockArc`, ...), `addTextBox`, `addLine`, fill/line/text formatting | 1.4 | yes |
| `shapes.addGroup(shapes)`, `Shape.group.shapes`, `fill.setImage` | 1.8 | yes (Mac 16.96+) |
| `Shape.adjustments.get/set` (a `Pie` wedge's start and end angles) | 1.10 | yes (Mac 16.105+, web supported, Windows 2601+) |
| `Office.CoercionType.XmlSvg` picture | ImageCoercion 1.2 | yes, not used (see non-goals) |

So a "native chart" is a tagged GROUP of native shapes: rectangles for bars, `Pie` shapes with
adjusted angles for slices, text boxes for the title, the values, the categories and the
legend, straight lines for the baseline and the waterfall connectors. Real Excel-style chart
objects are out of reach for any add-in; the manual and the pane say so in one sentence.

## Scope

Drawn as shape groups (v1): clustered column, stacked column, clustered bar, stacked bar, the
bridge/waterfall, the tornado (a clustered bar whose two series overlap), pie. Everything
else - line, area, scatter, combo, doughnut, the other chartex charts, bubbles, 3-D, stock,
surface - keeps arriving as the picture it is today, with no change in behaviour.

Caps, past which the link is a picture: 40 points per series, 3 series, 12 slices, a title of
80 characters. A chart with fewer than 2 points is a picture too.

## Payload

`src/link/model.ts`: `PicturePayload` gains an optional `chart?: ChartData` on links of kind
`chart`. The picture stays in every payload as the universal fallback, so a deck on an old
host, a chart type the renderer does not know and every deck inserted before this version
behave exactly as today. `payloadBytes` is unchanged (the PNG dominates).

`src/link/chart-model.ts` (core, pure, validated codec like the rest of `model.ts`):

```ts
export type ChartKind =
  | "column" | "stackedColumn" | "bar" | "stackedBar" | "waterfall" | "pie";

export interface ChartSeries {
  name: string;
  values: number[];   // one per category
  labels: string[];   // the text each value shows, decided in Excel
  colors: string[];   // one per point: brand colour of that bar, segment or slice
}

export interface ChartData {
  v: 1;
  kind: ChartKind;
  title: string | null;
  categories: string[];
  series: ChartSeries[];        // 1-3; a pie has exactly one
  font: string;                 // brand font
  ink: string;                  // label and axis text colour (theme formulaFont)
  titleColor: string;           // brand primary
  overlap?: true;               // tornado: both series share one row, drawn longest first
}
```

Labels are strings so the slide side never formats a number: the house number style, the
language and the source cell's own format are all Excel-side decisions, made once.

## Excel side: reading a chart (`src/excel/link-chart.ts`)

Mirrors `link-table.ts`. Given the resolved chart of a link:

1. `chart.chartType`, `chart.title.text`, `chart.series.items[*].name`, and for bar charts
   `series.overlap` (ExcelApi 1.8) to spot the tornado, in one load.
2. Per series `getDimensionValues("Categories")` and `getDimensionValues("Values")` (ExcelApi
   1.12; below it, no `chart` in the payload).
3. Label text: `getDimensionDataSourceString("Values")` (ExcelApi 1.15) names the source
   range; its `text` grid gives the value exactly as the model shows it. Without 1.15, or when
   the source is not a range, `formatChartAmount` (house style, pane language).
4. Colours by the rules the builders already use: `chartSeriesColors()` per series for
   column and bar charts (every point of a series the same colour); the waterfall's
   total/rise/fall by position and sign through `bridgeSeries` (`src/chartmath.ts`); a pie
   cycles the series palette per slice.
5. Chart type mapping: `*Clustered` column types -> `column`; `*Stacked`/`*Stacked100`
   column -> `stackedColumn`; bar likewise -> `bar` / `stackedBar` (overlap 100 -> `bar` +
   `overlap`); `Waterfall` -> `waterfall`; `Pie`, `PieExploded`, `3DPie`, `3DPieExploded` ->
   `pie`; anything else -> no `chart`.

`renderSource` in `src/excel/link-anchors.ts` attaches the result to the picture render: one
extra sync per chart link, none for range or table links. The chart on the sheet is never
written to; export stays read-only apart from the anchor it already creates.

## Pure layout (`src/chart-shapes.ts`, core: typed in, typed out, no I/O)

```ts
export type Primitive =
  | { kind: "rect"; box: Box; color: string; name: string }
  | { kind: "wedge"; box: Box; start: number; end: number; color: string; name: string }
  | { kind: "text"; box: Box; text: string; size: number; bold: boolean; color: string;
      align: "l" | "c" | "r"; name: string }
  | { kind: "line"; box: Box; color: string; weight: number; name: string };

export function chartSize(data: ChartData, source: Size): Size;      // points, capped
export function layoutChart(data: ChartData, box: Box): Primitive[];
```

Rules (all in points; text width estimated as 0.55 x size x characters):

- Bands: title 18 (12 pt bold `titleColor`, left), plot, category labels 14 (column) or a
  left column of 28 % of the width (bar, tornado), legend 16 when there is more than one
  series or the chart is a pie. Text is 9 pt `ink`, the brand font.
- Value scale: from the smallest negative reach to the largest positive reach (stacked: the
  sums of each sign), zero baseline drawn as a 0.75 pt line; negatives hang below it.
- Column: slot = plot width / categories; bar width 0.6 x slot / series (clustered, side by
  side) or 0.6 x slot (stacked). Value labels 12 pt high outside the end (above a positive
  bar, below a negative one); a stacked segment carries its label centred inside, in white
  when the fill is dark by luminance, and drops it when the segment is under 12 pt.
- Bar and tornado: the same, turned. Tornado rows draw the longer of the two bars first so the
  shorter stays visible; labels outside the far ends; the zero line is vertical.
- Waterfall: floating bars from `bridgeSeries` (totals from the baseline), a 0.75 pt
  connector from the top edge of each bar to the next at the running level, labels above a
  rise or a total and below a fall.
- Pie: radius 0.8 x half the shorter plot side; slices from 12 o'clock clockwise, angle =
  value / total x 360, zero and negative values skipped; labels at 1.18 r on the slice's
  mid-angle, anchored left or right of centre; legend swatches (8 x 8) plus category text in
  rows wrapped under the plot.
- Category labels truncated to the slot with an ellipsis; a legend that overflows wraps.
- Every primitive carries a `name`, which becomes the PowerPoint shape name
  ("pls,fix chart <label>: Volume 2026E"), so a user inspecting the group sees what each
  shape is.

Tests: sums and geometry (bars inside the plot, no two bars of one category overlapping unless
`overlap`, labels outside the bar they belong to, wedge angles closing the circle), text
passthrough (labels appear verbatim), caps and skips.

## Slide side (`src/ppt/charts.ts`, mirrors `tables.ts`)

- `insertChart(stage, item, payload, tag)`: `placeOnSlide(chartSize(...))` (free space, the
  existing rule), then for each primitive one `addGeometricShape(rectangle | pie)` +
  `fill.setSolidColor` + `lineFormat.visible = false` (+ `adjustments.set` for a wedge),
  `addTextBox` with `autoSizeSetting = AutoSizeNone`, `wordWrap = false`, zero margins,
  vertical middle, the font, size, colour and alignment, `fill.clear()`, no line, or
  `addLine("Straight", box)` with colour and weight; then `shapes.addGroup(created)`, the
  group named `pls,fix chart <label>` and carrying BOTH tags (children carry none); one
  sync. The group's box is what `scanLinks` reports as the link's geometry.
- `refreshChart(found, payload, tag)`: `isGrouped(found)` (the user grouped our group with
  something) -> "ungroup the chart before it can update", exactly like a table. Otherwise the
  new group is drawn at the found corner and width, the height following the chart's aspect,
  and the old group is deleted in the same batch after the add is queued (the table's
  `recreate` order: a refused add never takes the old object down).
- Capability: `canDrawChart(data)` = PowerPointApi 1.8 (groups, text) and, for a pie, 1.10.
  A host without it inserts the picture. The representation chosen at insert is the link's
  for life: a found shape of type `Group` is a chart group and refreshes as one; a rectangle
  refreshes as a picture even when the new payload carries chart data. A chart group whose
  new payload has no chart data (the modeller switched the chart to a line) is rebuilt as a
  picture at the same corner, and the row says so.
- Batching: `refreshLinks` already hands anything that is not a picture back to the
  row-by-row path; a chart repaint is one `PowerPoint.run`, one sync, like a table.

## Host dispatch (`src/ppt/host.ts`)

`insertLink`: `payload.chart && canDrawChart(payload.chart)` -> `insertChart`, else the
picture route. `refreshLink`: by the found shape's type (`Group` -> chart route). `FoundLink`
gains the shape `type` (already loaded by `SHAPE_PROPERTIES`). `scanLinks` keeps the group
entry itself (`expandGroups` appends children after the top-level list, and every entry's tags
are read), so a tagged group is found as a top-level link with an empty `groupPath`.

## Fakes and tests

- `test/fakeppt/objects.ts`: `addTextBox`, `addLine`, `addGroup` (members leave the slide's
  top level for the group's `shapes`; `delete()` on the group cascades), `adjustments`
  (count by geometry: `Pie` 2, `Rectangle` 0; `get` as a `ClientResult`, `set`), text frame
  settings and `textRange.font` / `paragraphFormat`; strict rules in `strict.ts`.
- `test/fakehost.ts`: chart `title.text` and `chartType` loadable, `series.items` with `name`
  and `overlap`, `getDimensionValues`, `getDimensionDataSourceString` as `ClientResult`s.
- Suites: `src/chart-shapes.test.ts`, `src/link/chart-model.test.ts`,
  `test/links.chart.integration.test.ts` (Excel: a bridge chart exports data and a picture;
  a line chart exports the picture alone; the caps), `test/ppt.charts.integration.test.ts`
  (insert as a tagged group in free space, refresh rebuilds at the same corner, a grouped
  chart refuses, no 1.10 -> a pie is a picture, a chart-to-line change -> picture at the
  corner), `ppt.perf`: one sync per chart insert and one per repaint.

## Spike (before the plan): what only a real host can answer

Run on PowerPoint for the web through the CDP rig (lessons 2026-08-29) and on Daniel's
desktop pass; the findings are appended to this section and the plan is written after them.

1. `Pie` adjustments: which index is the start and which the end, the unit (degrees,
   60000ths of a degree, or a 0-1 fraction), the direction and the zero (3 o'clock in OOXML).
2. `addGroup` of ~60 shapes: tags on the group are read back by `scanLinks`; `delete()` on
   the group removes the members; the group's left/top/width/height are the union.
3. A 9 pt text box with autosize off and zero margins renders one line, no wrap, no clip.
4. Time to insert a 12-point, 2-series column chart (about 60 shapes) in one sync: must be
   under 2 seconds on the web.
5. `isSetSupported("PowerPointApi", "1.10")` is true on the web and on 16.107.

A failed item narrows scope (a pie stays a picture; a slower insert lowers the caps) rather
than being worked around.

## Docs and copy

Help sentences (`src/help/copy.ppt.ts` for the inbox/insert rows, `copy.excel.ts` for
"Export active chart"), the manual (`manual/src/content/links.rs`, `charts.rs`, Latvian),
`docs/FEATURES.md` one delta line, README one surgical line, CLAUDE.md Map lines for
`src/link/chart-model.ts`, `src/chart-shapes.ts`, `src/excel/link-chart.ts`,
`src/ppt/charts.ts` (and the Map deduplicated back under 40 lines).

## Non-goals

Real PowerPoint chart objects (no API). SVG pictures through `XmlSvg` as a middle way (a
vector picture is still a picture; groups already give editability). Doughnut, line, area
and scatter drawing. Restyling the modeller's Excel chart on export. Reading a chart's own
colours (Office.js cannot; brand colours are the product's point).
