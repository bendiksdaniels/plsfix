# Native chart groups on slides: implementation plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A chart link opened from the PowerPoint inbox lands as a tagged group of native shapes (bars, pie wedges, value labels, title, legend) on hosts that can draw it, refreshes by rebuilding in place, and stays the picture it is today everywhere else.

**Architecture:** Excel reads the chart's data and decides every label string and colour (`src/excel/link-chart.ts`) and ships it inside the existing picture payload (`PicturePayload.chart`). A pure layout module (`src/chart-shapes.ts`) turns `ChartData` + a box into primitives (rect, wedge, text, line). The slide adapter (`src/ppt/charts.ts`) draws primitives in syncs of twelve, groups them, tags the group; `src/ppt/host.ts` dispatches on the payload and on the found shape's type. The picture stays in every payload as the fallback.

**Tech Stack:** TypeScript, Office.js (ExcelApi 1.12/1.15 reads, PowerPointApi 1.8 groups, 1.10 adjustments), Vitest with the strict fakes in `test/fakehost.ts` and `test/fakeppt/`, Rust manual crate.

**Spec:** `docs/superpowers/specs/2026-08-30-native-charts-design.md` (read it first; its "Spike findings" section is the ground truth for the host behaviour).

## Global Constraints

- Every source file opens with a 2-4 line header comment (purpose, what it owns, key invariant); files at most 400 lines, functions at most 50 lines; no `utils`/`helpers` files.
- Every Office.js scalar read is preceded by `load()` + `sync()`; the strict fakes throw otherwise.
- Pure modules (`src/link/chart-model.ts`, `src/chart-shapes.ts`, `src/chart-colors.ts`) import nothing from Office.js and nothing from `src/excel/` or `src/ppt/`.
- Constants live at the top of the module that owns them; no magic numbers in logic.
- Tests mirror modules by name; TDD: write the failing test, run it, implement, run it green, commit.
- Commit messages in Daniel's terse voice: short imperative subject, plain 3-6 line body, no trailers.
- Labels are strings decided in Excel; the slide side never formats a number.
- Caps (spec): `CHART_MAX_POINTS` 40 per series, `CHART_MAX_SERIES` 3, `CHART_MAX_SLICES` 12, `CHART_MIN_POINTS` 2, `CHART_TITLE_MAX` 80 characters.
- Host budget (spec): `SHAPES_PER_SYNC` 12, `SHAPE_BUDGET_WEB` 30, `SHAPE_BUDGET_DESKTOP` 200.
- Pie wedge angles: degrees, clockwise, zero at 3 o'clock, normalised to (-180, 180]; a 100 % slice is an ellipse; adjustments are set in the sync after the wedge's add.
- Gates before any merge: `npm run check` and `npm run ux:check` green; agent worktrees run `git reset --hard <main sha>` first and never touch the main checkout's `node_modules`.

---

## File map

| File | Responsibility | Task |
| --- | --- | --- |
| `src/link/chart-model.ts` (new, pure) | `ChartData` types, caps, `isChartData` validator, `normalizeAngle` | 1 |
| `src/chart-colors.ts` (new, pure) | brand colour rules for series, waterfall points and pie slices (moved out of `src/excel/charts.ts`) | 1 |
| `src/chart-shapes.ts` (new, pure) | `layoutChart(data, box)`, `chartSize`, `Primitive` | 2 |
| `src/link/model.ts` | `PicturePayload.chart?: ChartData`, validator | 3 |
| `src/excel/link-record.ts` | spreads `render.chart` into the payload | 3 |
| `src/excel/link-anchors.ts` | `Render` gains `chart?`; `renderSource` calls the reader for chart sources | 4 |
| `src/excel/link-chart.ts` (new) | `readChartData(context, chart)`: type, title, series, labels, colours | 4 |
| `test/fakehost.ts` | chart title/series readable, `getDimensionValues`, `getDimensionDataSourceString`, `overlap`, `helpers.addChart` series seeding | 4 |
| `test/fakeppt/model.ts`, `objects.ts`, `strict.ts` | text boxes, lines, groups by id, adjustments, fills, line format, text formatting | 5 |
| `src/ppt/charts.ts` (new) | `chartPlan`, `canDrawChart`, `insertChart`, `refreshChart`, chunked drawing | 6 |
| `src/ppt/host.ts` | `FoundLink.type`, insert/refresh dispatch, group-to-picture rebuild, `InsertResult.note` | 6 |
| `src/ppt/links.ts` | toasts `InsertResult.note` | 6 |
| `src/help/copy.ppt.ts`, `copy.excel.ts`, `manual/src/content/links.rs`, `docs/FEATURES.md`, `README.md`, `CLAUDE.md` | words | 7 |

Dispatch: the controller does Tasks 1 and 3 in the main checkout; wave 1 = Task 2 (sonnet), Task 4 (opus), Task 5 (opus) in parallel worktrees off the Task 1+3 commit; wave 2 = Task 6 (opus) off the merged wave 1; Task 7 (sonnet) last; Task 8 is the controller's.

---

### Task 1: Chart data model and colour rules (pure)

**Files:**
- Create: `src/link/chart-model.ts`, `src/link/chart-model.test.ts`
- Create: `src/chart-colors.ts`, `src/chart-colors.test.ts`
- Modify: `src/excel/charts.ts` (use `src/chart-colors.ts` instead of its private `chartSeriesColors`; behaviour unchanged, `test/host.integration.test.ts` "restyles the selected chart and brands every series" and the waterfall `pointColors` test stay green)

**Interfaces (produces):**

```ts
// src/link/chart-model.ts
export type ChartKind = "column" | "stackedColumn" | "bar" | "stackedBar" | "waterfall" | "pie";
export interface ChartSeries { name: string; values: number[]; labels: string[]; colors: string[] }
export interface ChartData {
  v: 1; kind: ChartKind; title: string | null; categories: string[]; series: ChartSeries[];
  font: string; ink: string; titleColor: string; overlap?: true;
}
export const CHART_MAX_POINTS = 40, CHART_MAX_SERIES = 3, CHART_MAX_SLICES = 12,
  CHART_MIN_POINTS = 2, CHART_TITLE_MAX = 80;
export function isChartData(value: unknown): value is ChartData;   // shape + consistency + caps
export function normalizeAngle(degrees: number): number;            // to (-180, 180]
// src/chart-colors.ts
export function seriesPalette(settings: Pick<BrandSettings, "primary" | "accent">): string[]; // 6 colours, the order charts.ts used
export function waterfallColors(values: number[], settings: Pick<BrandSettings, "primary" | "accent" | "external">): string[]; // totals first/last primary, fall external, rise accent
export function pieColors(count: number, settings: Pick<BrandSettings, "primary" | "accent">): string[]; // palette cycled
```

- [ ] **Step 1: failing tests for the model**

```ts
// src/link/chart-model.test.ts
import { describe, expect, it } from "vitest";
import { isChartData, normalizeAngle, type ChartData } from "./chart-model";

const column: ChartData = {
  v: 1, kind: "column", title: "Revenue", categories: ["2024A", "2025E"],
  series: [{ name: "Revenue", values: [100, 120], labels: ["100", "120"], colors: ["#2EC4B6", "#2EC4B6"] }],
  font: "Arial", ink: "#333333", titleColor: "#14213D",
};

describe("isChartData", () => {
  it("accepts a consistent column chart", () => expect(isChartData(column)).toBe(true));
  it("refuses a series whose labels do not match its values", () => {
    expect(isChartData({ ...column, series: [{ ...column.series[0]!, labels: ["100"] }] })).toBe(false);
  });
  it("refuses more categories than the cap, fewer than two, and a pie with two series", () => {
    const many = Array.from({ length: 41 }, (_, i) => String(i));
    expect(isChartData({ ...column, categories: many, series: [{ name: "s", values: many.map(Number), labels: many, colors: many.map(() => "#000000") }] })).toBe(false);
    expect(isChartData({ ...column, categories: ["a"], series: [{ name: "s", values: [1], labels: ["1"], colors: ["#000000"] }] })).toBe(false);
    expect(isChartData({ ...column, kind: "pie", series: [column.series[0]!, column.series[0]!] })).toBe(false);
  });
  it("refuses an unknown kind, a bad colour and a title over 80 characters", () => {
    expect(isChartData({ ...column, kind: "line" })).toBe(false);
    expect(isChartData({ ...column, series: [{ ...column.series[0]!, colors: ["red", "#2EC4B6"] }] })).toBe(false);
    expect(isChartData({ ...column, title: "x".repeat(81) })).toBe(false);
  });
  it("allows overlap only as true", () => {
    expect(isChartData({ ...column, kind: "bar", overlap: true })).toBe(true);
    expect(isChartData({ ...column, kind: "bar", overlap: false })).toBe(false);
  });
});

describe("normalizeAngle", () => {
  it("maps onto (-180, 180] the way PowerPoint reads angles back", () => {
    expect(normalizeAngle(200)).toBe(-160);
    expect(normalizeAngle(269.9)).toBeCloseTo(-90.1);
    expect(normalizeAngle(360)).toBe(0);
    expect(normalizeAngle(-90)).toBe(-90);
    expect(normalizeAngle(180)).toBe(180);
    expect(normalizeAngle(-180)).toBe(180);
  });
});
```

- [ ] **Step 2: run, expect FAIL (module missing)**: `npx vitest run src/link/chart-model.test.ts`
- [ ] **Step 3: implement `src/link/chart-model.ts`** (validator checks: `v === 1`, kind in the six, `title` null or string within `CHART_TITLE_MAX`, categories array of strings with `CHART_MIN_POINTS <= n <= CHART_MAX_POINTS`, 1 to `CHART_MAX_SERIES` series, a pie exactly one series and `n <= CHART_MAX_SLICES`, every series' `values` finite numbers and `labels`/`colors` the same length as the categories, colours `/^#[0-9A-Fa-f]{6}$/`, `font`/`ink`/`titleColor` strings, `overlap` absent or `true`). `normalizeAngle`: `let n = ((d + 180) % 360 + 360) % 360 - 180; return n === -180 ? 180 : n;`.
- [ ] **Step 4: run green**, then prettier and eslint on the two files.
- [ ] **Step 5: failing tests for the colours** (`src/chart-colors.test.ts`): `seriesPalette` returns six entries starting `[accent, primary]`; `waterfallColors([100, 20, -30, 90], s)` equals `[primary, accent, external, primary]` (the rule the waterfall test in `test/host.integration.test.ts` already asserts through `pointColors`); `pieColors(8, s)` cycles the palette (`[0]` equals `[6]`).
- [ ] **Step 6: implement `src/chart-colors.ts`** by moving `chartSeriesColors` out of `src/excel/charts.ts` (`tint` from `src/settings.ts`), and the waterfall rule from `insertWaterfall` (`index === 0 || last -> primary`, `bridge.fall[index] > 0 -> external`, else `accent`, with `bridgeSeries` from `src/chartmath.ts`); make `charts.ts` call them.
- [ ] **Step 7: run** `npx vitest run src/chart-colors.test.ts test/host.integration.test.ts test/chart-placement.integration.test.ts` green; `npx tsc --noEmit`.
- [ ] **Step 8: commit** "Chart data model and colour rules as pure modules".

---

### Task 2: Pure chart layout (`src/chart-shapes.ts`)

**Files:**
- Create: `src/chart-shapes.ts`, `src/chart-shapes.test.ts` (split the module into `src/chart-shapes.ts` (entry, bands, size, bars) and `src/chart-shapes-parts.ts` (the shared primitive builders, pie wedges, legend) if it passes 400 lines)

**Interfaces (consumes):** `ChartData`, `ChartSeries`, `normalizeAngle` from `src/link/chart-model.ts`; `Box`, `Size` from `src/layout.ts`; `bridgeSeries` from `src/chartmath.ts`.

**Interfaces (produces):**

```ts
export interface Rect { kind: "rect"; box: Box; color: string; name: string }
export interface Wedge { kind: "wedge"; box: Box; start: number; end: number; color: string; name: string } // start/end already normalised
export interface Ellipse { kind: "ellipse"; box: Box; color: string; name: string }                             // the one 100 % slice
export interface Text { kind: "text"; box: Box; text: string; size: number; bold: boolean; color: string; align: "l" | "c" | "r"; name: string }
export interface Line { kind: "line"; box: Box; color: string; weight: number; name: string }
export type Primitive = Rect | Wedge | Ellipse | Text | Line;
export const LABEL_SIZE = 9, TITLE_SIZE = 12, TITLE_BAND = 18, LEGEND_BAND = 16, CATEGORY_BAND = 14,
  BAR_FILL = 0.6, LABEL_HEIGHT = 18, LABEL_PAD = 7, LINE_WEIGHT = 0.75, PIE_RADIUS = 0.8,
  PIE_LABEL_RADIUS = 1.18, CHAR_WIDTH = 0.55, SWATCH = 8, BAR_LABEL_COLUMN = 0.28, MIN_SEGMENT = 12,
  MIN_SIZE: Size = { width: 200, height: 120 };
export function textWidth(text: string, size: number): number;          // CHAR_WIDTH * size * text.length + LABEL_PAD
export function darkFill(hex: string): boolean;                          // relative luminance < 0.5
export function chartSize(source: Size, maxWidth: number): Size;         // source in points, scaled down to maxWidth keeping the aspect, never under MIN_SIZE
export function layoutChart(data: ChartData, box: Box): Primitive[];
```

Rules the implementation must follow (from the spec): bands title / plot / category (column) or a left label column of `BAR_LABEL_COLUMN` of the width (bar, tornado) / legend when `series.length > 1` or a pie; value scale from the most negative reach to the most positive reach (stacked: signed sums), zero baseline as a line of `LINE_WEIGHT`; column slot = plot width / categories, bar width `BAR_FILL` x slot / series (clustered) or `BAR_FILL` x slot (stacked); value labels `LABEL_HEIGHT` high, centred on the bar, outside its end (above a positive bar, below a negative one); a stacked segment's label centred inside it in white when `darkFill(color)`, dropped when the segment is under `MIN_SEGMENT`; bar charts turned; tornado (`overlap`) draws the longer bar of a row first; waterfall floating bars from `bridgeSeries(values)` (first and last from the baseline), a horizontal connector from each bar's top edge to the next at the running level, labels above a rise or total and below a fall; pie radius `PIE_RADIUS` x half the shorter plot side centred in the plot, slices from 12 o'clock clockwise (`start = normalizeAngle(-90 + a)`, `end = normalizeAngle(-90 + a + span)`), zero and negative values skipped, one slice of 100 % as an `Ellipse`, labels centred at `PIE_LABEL_RADIUS` x r along the mid-angle; legend swatches `SWATCH` square + name, rows wrapped under the plot; category labels truncated with "…" when `textWidth` exceeds the slot; every primitive named `"<what> <index or category>"` (e.g. `bar 0.2`, `label 0.2`, `category 2`, `slice 1`, `title`, `baseline`, `legend 1`), unique within one chart.

- [ ] **Step 1: failing tests** (`src/chart-shapes.test.ts`):

```ts
import { describe, expect, it } from "vitest";
import type { ChartData } from "./link/chart-model";
import { chartSize, layoutChart, type Primitive, type Rect, type Text, type Wedge } from "./chart-shapes";

const box = { left: 100, top: 100, width: 480, height: 270 };
const base = { v: 1 as const, font: "Arial", ink: "#333333", titleColor: "#14213D", title: "Revenue" };
const rects = (p: Primitive[]) => p.filter((x): x is Rect => x.kind === "rect");
const texts = (p: Primitive[]) => p.filter((x): x is Text => x.kind === "text");
const inside = (b: { left: number; top: number; width: number; height: number }) =>
  b.left >= box.left - 0.01 && b.top >= box.top - 0.01 && b.left + b.width <= box.left + box.width + 0.01 && b.top + b.height <= box.top + box.height + 0.01;

const column: ChartData = { ...base, kind: "column", categories: ["A", "B", "C"],
  series: [{ name: "s", values: [100, -50, 200], labels: ["100", "(50)", "200"], colors: ["#2EC4B6", "#2EC4B6", "#2EC4B6"] }] };

describe("layoutChart column", () => {
  const out = layoutChart(column, box);
  it("draws one bar per category inside the box, tallest for the largest value", () => {
    const bars = rects(out);
    expect(bars).toHaveLength(3);
    expect(bars.every((b) => inside(b.box))).toBe(true);
    expect(bars[2]!.box.height).toBeGreaterThan(bars[0]!.box.height);
  });
  it("hangs a negative bar below the baseline", () => {
    const [a, b] = rects(out);
    const baseline = out.find((x) => x.kind === "line" && x.name === "baseline")!;
    expect(a!.box.top + a!.box.height).toBeCloseTo(baseline.box.top, 0);
    expect(b!.box.top).toBeCloseTo(baseline.box.top, 0);
  });
  it("labels every bar with its text, above a positive and below a negative bar, and names the categories", () => {
    const labels = texts(out).filter((t) => t.name.startsWith("label"));
    expect(labels.map((t) => t.text)).toEqual(["100", "(50)", "200"]);
    const [a, b] = rects(out);
    expect(labels[0]!.box.top + labels[0]!.box.height).toBeLessThanOrEqual(a!.box.top + 0.01);
    expect(labels[1]!.box.top).toBeGreaterThanOrEqual(b!.box.top + b!.box.height - 0.01);
    expect(texts(out).filter((t) => t.name.startsWith("category")).map((t) => t.text)).toEqual(["A", "B", "C"]);
  });
  it("titles the chart in the title colour and draws no legend for one series", () => {
    const title = texts(out).find((t) => t.name === "title")!;
    expect(title).toMatchObject({ text: "Revenue", bold: true, color: "#14213D", size: 12 });
    expect(out.some((x) => x.name.startsWith("legend"))).toBe(false);
  });
  it("gives every primitive a unique name", () => {
    expect(new Set(out.map((x) => x.name)).size).toBe(out.length);
  });
});

describe("layoutChart stacked and bars", () => {
  const stacked: ChartData = { ...column, kind: "stackedColumn", series: [
    { name: "a", values: [60, 40, 80], labels: ["60", "40", "80"], colors: ["#14213D", "#14213D", "#14213D"] },
    { name: "b", values: [40, 60, 20], labels: ["40", "60", "20"], colors: ["#2EC4B6", "#2EC4B6", "#2EC4B6"] } ] };
  it("stacks two series into one column per category with white labels inside the dark segments", () => {
    const out = layoutChart(stacked, box);
    const bars = rects(out);
    expect(bars).toHaveLength(6);
    expect(bars[0]!.box.left).toBeCloseTo(bars[3]!.box.left);
    const inA = texts(out).find((t) => t.name === "label 0.0")!;
    expect(inA.color).toBe("#FFFFFF");
    expect(out.filter((x) => x.name.startsWith("legend")).length).toBeGreaterThan(0);
  });
  it("turns a bar chart on its side: bars grow to the right from a vertical baseline", () => {
    const out = layoutChart({ ...column, kind: "bar" }, box);
    const [a, , c] = rects(out);
    expect(c!.box.width).toBeGreaterThan(a!.box.width);
    expect(a!.box.top).toBeLessThan(c!.box.top);
    const baseline = out.find((x) => x.name === "baseline")!;
    expect(baseline.box.width).toBe(0);
  });
  it("draws a tornado row with the longer bar first so the shorter stays visible", () => {
    const tornado: ChartData = { ...column, kind: "bar", overlap: true, categories: ["Volume", "Price"], series: [
      { name: "Low", values: [-20, -5], labels: ["(20)", "(5)"], colors: ["#B00020", "#B00020"] },
      { name: "High", values: [15, 25], labels: ["15", "25"], colors: ["#2EC4B6", "#2EC4B6"] } ] };
    const out = layoutChart(tornado, box);
    const row = rects(out).filter((r) => r.name.endsWith(".1"));
    expect(row).toHaveLength(2);
    expect(row[0]!.box.width).toBeGreaterThan(row[1]!.box.width);
  });
});

describe("layoutChart waterfall", () => {
  const bridge: ChartData = { ...column, kind: "waterfall", categories: ["Open", "Price", "Cost", "Close"],
    series: [{ name: "s", values: [100, 20, -30, 90], labels: ["100", "20", "(30)", "90"], colors: ["#14213D", "#2EC4B6", "#B00020", "#14213D"] }] };
  it("floats the steps, connects the bars at the running level and labels a fall below its bar", () => {
    const out = layoutChart(bridge, box);
    const bars = rects(out);
    const baseline = out.find((x) => x.name === "baseline")!;
    expect(bars[0]!.box.top + bars[0]!.box.height).toBeCloseTo(baseline.box.top, 0);
    expect(bars[1]!.box.top + bars[1]!.box.height).toBeCloseTo(bars[0]!.box.top, 0);
    expect(bars[2]!.box.top).toBeCloseTo(bars[1]!.box.top, 0);
    expect(out.filter((x) => x.name.startsWith("connector"))).toHaveLength(3);
    const fall = texts(out).find((t) => t.name === "label 0.2")!;
    expect(fall.box.top).toBeGreaterThanOrEqual(bars[2]!.box.top + bars[2]!.box.height - 0.01);
  });
});

describe("layoutChart pie", () => {
  const pie: ChartData = { ...column, kind: "pie", categories: ["A", "B", "C", "D"],
    series: [{ name: "s", values: [1, 1, 2, 0], labels: ["1", "1", "2", "0"], colors: ["#14213D", "#2EC4B6", "#B27E54", "#000000"] }] };
  it("cuts wedges clockwise from 12 o'clock, skips a zero slice and closes the circle", () => {
    const out = layoutChart(pie, box);
    const wedges = out.filter((x): x is Wedge => x.kind === "wedge");
    expect(wedges.map((w) => [w.start, w.end])).toEqual([[-90, 0], [0, 90], [90, -90]]);
    expect(wedges.every((w) => w.box.width === w.box.height)).toBe(true);
    expect(texts(out).filter((t) => t.name.startsWith("label")).map((t) => t.text)).toEqual(["1", "1", "2"]);
    expect(texts(out).filter((t) => t.name.startsWith("legend")).map((t) => t.text)).toEqual(["A", "B", "C", "D"]);
  });
  it("draws a lone 100 % slice as an ellipse", () => {
    const one: ChartData = { ...pie, categories: ["A", "B"], series: [{ name: "s", values: [5, 0], labels: ["5", "0"], colors: ["#14213D", "#000000"] }] };
    expect(layoutChart(one, box).some((x) => x.kind === "ellipse")).toBe(true);
  });
});

describe("chartSize", () => {
  it("scales a wide chart down to the content width and never under the minimum", () => {
    expect(chartSize({ width: 1200, height: 600 }, 888)).toEqual({ width: 888, height: 444 });
    expect(chartSize({ width: 100, height: 50 }, 888)).toEqual({ width: 200, height: 120 });
    expect(chartSize({ width: 420, height: 225 }, 888)).toEqual({ width: 420, height: 225 });
  });
});
```

- [ ] **Step 2: run, expect FAIL**: `npx vitest run src/chart-shapes.test.ts`
- [ ] **Step 3: implement** in small named functions (`bands`, `valueScale`, `columnBars`, `barBars`, `waterfallBars`, `pieWedges`, `legend`, `categoryLabels`), each under 50 lines; the module header states the invariant "every primitive lies inside the box it was given".
- [ ] **Step 4: run green**; prettier, eslint, `npx tsc --noEmit`.
- [ ] **Step 5: commit** "Pure chart layout: primitives for bars, waterfalls and pies".

---

### Task 3: The payload carries chart data

**Files:**
- Modify: `src/link/model.ts` (`PicturePayload` gains `chart?: ChartData`; `isPicturePayload` accepts a missing `chart` or one that passes `isChartData`; export `ChartData` types from `./chart-model` for callers), `src/link/model.test.ts`
- Modify: `src/excel/link-anchors.ts:50-53` (`Render` picture variant gains `chart?: ChartData`), `src/excel/link-record.ts:60-70` (spreads `render.chart` into the payload when present)

- [ ] **Step 1: failing test** in `src/link/model.test.ts`:

```ts
it("keeps chart data on a picture payload and refuses a malformed one", () => {
  const chart = { v: 1, kind: "column", title: null, categories: ["a", "b"],
    series: [{ name: "s", values: [1, 2], labels: ["1", "2"], colors: ["#000000", "#000000"] }],
    font: "Arial", ink: "#333333", titleColor: "#14213D" };
  const withChart = { ...picture, chart };                 // `picture` = the valid PicturePayload fixture the file already uses
  expect(decodePayload(encodePayload(withChart))).toEqual(withChart);
  expect(() => decodePayload(encodePayload({ ...picture, chart: { ...chart, kind: "line" } }))).toThrow();
});
```

- [ ] **Step 2: run, expect FAIL**; **Step 3: implement**; **Step 4: run** `npx vitest run src/link test/links.integration.test.ts test/links.table.integration.test.ts` green.
- [ ] **Step 5: commit** "Picture payload carries optional chart data".

---

### Task 4: Excel reads a chart's data (`src/excel/link-chart.ts`)

**Files:**
- Create: `src/excel/link-chart.ts`, `src/excel/link-chart.test.ts` (pure parts: `chartKind`, `valuesOf`)
- Modify: `src/excel/link-anchors.ts:334-351` (`renderSource`: for `resolved.kind === "chart"` also `readChartData`; the image sync and the reads share syncs where the fake allows)
- Modify: `test/fakehost.ts` (see below), `test/links.chart.integration.test.ts` (new)

**Interfaces (produces):**

```ts
// src/excel/link-chart.ts
export function chartKind(chartType: string, overlap: number | null): ChartKind | null;
//   /Clustered$/ column types -> "column"; /^(3D)?(Column|Cylinder|Cone|Pyramid)Col?.*Stacked/ -> "stackedColumn";
//   /^(3D)?(Bar|CylinderBar|ConeBar|PyramidBar)Clustered$/ -> "bar" (overlap === 100 -> overlap flag set by the caller);
//   /^(3D)?(Bar|CylinderBar|ConeBar|PyramidBar)Stacked/ -> "stackedBar"; "Waterfall" -> "waterfall";
//   /^(3D)?Pie(Exploded)?$/ -> "pie"; anything else -> null
export function valuesOf(raw: (string | number)[]): number[] | null;     // Number() each; null when any is not finite
export async function readChartData(context: Excel.RequestContext, chart: Excel.Chart): Promise<ChartData | null>;
//   null when: the host lacks ExcelApi 1.12, the kind is unknown, a value is not a number, or the caps fail (isChartData false)
```

Reading order (three syncs at most): (1) `chart.load("chartType,name")`, `chart.title.load("text")`, `chart.series.load("items/name" + (hostSupports("1.8") ? ",items/overlap" : ""))`; (2) per series `getDimensionValues("Categories")`, `getDimensionValues("Values")`, and when `hostSupports("1.15")` `getDimensionDataSourceString("Values")`; (3) when the source string parses with `parseAddress` to a sheet and an A1 range, `context.workbook.worksheets.getItem(sheet).getRange(address).load("text")` for every series, else labels = `formatChartAmount(value)`. Colours: `seriesPalette(settings)[i]` repeated per point for column/bar kinds, `waterfallColors(values, settings)` for the waterfall, `pieColors(n, settings)` for a pie; `font = settings.font`, `ink = activeTheme().formulaFont`, `titleColor = settings.primary`. Categories: the strings from `getDimensionValues("Categories")`, or `"1".."n"` when empty.

Fake host (`test/fakehost.ts`): `FakeSeries` gains `name?: string; categories?: string[]; values?: number[]; valuesSource?: string`; `helpers.addChart(sheet, { chartType, title, series: [...] })` seeds them; the chart proxy's `title` gains a `text` getter, `series` gains `items` (proxies in order) and each series proxy gains `name`/`overlap` getters, `getDimensionValues(dim)` -> `{ value }` result (`Categories` -> the strings, `Values` -> the numbers), `getDimensionDataSourceString(dim)` -> `{ value: valuesSource ?? "" }`; strict rules: `chartTitle` already lists `text`; `chartSeries` gains `items: "chartSeriesItem"`; `chartSeriesItem.returns` gains `getDimensionValues: "clientResult"`, `getDimensionDataSourceString: "clientResult"`.

- [ ] **Step 1: failing pure tests** (`src/excel/link-chart.test.ts`): `chartKind("ColumnClustered", 0)` -> "column"; `("BarClustered", 100)` -> "bar"; `("ColumnStacked100", null)` -> "stackedColumn"; `("Waterfall", null)` -> "waterfall"; `("3DPieExploded", null)` -> "pie"; `("Line", null)`, `("Doughnut", null)`, `("Area", null)` -> null; `valuesOf(["1", 2.5])` -> [1, 2.5]; `valuesOf(["x"])` -> null.
- [ ] **Step 2: failing integration test** (`test/links.chart.integration.test.ts`, boot exactly as `test/links.table.integration.test.ts` does):

```ts
it("exports a column chart with its data, the cells' text as labels and brand colours", async () => {
  helpers.seed("Model!B3", [["2024A", "2025E", "2026E"], [1240, 1302, 1400]]);
  helpers.setNumberFormat("Model!B4:D4", "#,##0");
  const chart = helpers.addChart("Model", { name: "Revenue chart", chartType: "ColumnClustered", title: "Revenue",
    series: [{ name: "Revenue", categories: ["2024A", "2025E", "2026E"], values: [1240, 1302, 1400], valuesSource: "Model!$B$4:$D$4" }] });
  helpers.setActiveChart(chart);
  const { id } = await links.exportChart(ws, relay, null);          // the signature links.ts already has
  const payload = decodePayload(await openLatest(id));                // the helper the table test uses to read the relay
  expect(payload.kind).toBe("picture");
  expect(payload.chart).toMatchObject({ kind: "column", title: "Revenue", categories: ["2024A", "2025E", "2026E"] });
  expect(payload.chart?.series[0]).toMatchObject({ name: "Revenue", values: [1240, 1302, 1400], labels: ["1,240", "1,302", "1,400"] });
  expect(payload.chart?.series[0]?.colors).toEqual([palette.accent, palette.accent, palette.accent]);
});
it("sends a line chart as a picture alone", ...);                  // chartType "Line" -> payload.chart undefined
it("marks the tornado's overlapped bars", ...);                     // BarClustered, two series with overlap 100 -> kind "bar", overlap true
it("falls back to the picture past forty points", ...);             // 41 categories -> payload.chart undefined
it("formats labels in the house style when the values have no cells", ...); // no valuesSource -> labels via formatChartAmount
```

(If the fake's `text` does not honour `#,##0`, seed `numberFormat` "General" and expect `["1240", "1302", "1400"]`; say so in the commit.)

- [ ] **Step 3: run, expect FAIL**; **Step 4: implement** reader + fake; **Step 5: run** `npx vitest run src/excel test/links.chart.integration.test.ts test/links.integration.test.ts test/host.integration.test.ts` green; `npx tsc --noEmit`.
- [ ] **Step 6: commit** "Excel reads a chart's data into the link payload".

---

### Task 5: The fake PowerPoint learns shapes, groups and adjustments

**Files:**
- Modify: `test/fakeppt/model.ts` (`FakePptShape` gains `geometry: string | null`, `fillColor: string | null`, `fillCleared: boolean`, `lineColor: string | null`, `lineWeight: number | null`, `adjustments: number[]`, `text: string | null`, `font: { name?: string; size?: number; color?: string; bold?: boolean }`, `alignment: string | null`, `autoSize: string | null`, `wordWrap: boolean | null`, `synced: boolean`; `type` union gains `"TextBox" | "Line"`), `test/fakeppt/objects.ts`, `test/fakeppt/strict.ts`
- Create: `test/fakeppt/shapes.test.ts`

**Produces (fake API, mirrors Office.js):** `shapes.addTextBox(text, box)`, `shapes.addLine(connectorType, box)`, `shapes.addGroup(ids: string[])` (throws "InvalidArgument" on an empty list; uses `FakePresentation.groupShapes`), `shape.adjustments` with `count` (2 for `"Pie"`, 0 otherwise), `get(i)` returning a result, `set(i, v)` storing `normalizeAngle(v)` and throwing `"InvalidParam passed to GetItem(id)"` while `shape.synced` is false (a shape becomes synced on the first `context.sync()` after its add; the runtime marks every shape added in the batch), `shape.fill.setSolidColor(hex)`, `shape.fill.clear()`, `shape.lineFormat.color/weight`, `shape.textFrame.autoSizeSetting/wordWrap/leftMargin/rightMargin/topMargin/bottomMargin/verticalAlignment`, `shape.textFrame.textRange.font.name/size/color/bold`, `shape.textFrame.textRange.paragraphFormat.horizontalAlignment`, `shape.delete()` on a group removing it with its members; `helpers.setPlatform(platform: string)` on `FakePptHelpers` (sets `Office.context.platform`, default `"Mac"`). Strict rules: `shapeCollection.returns` += `addTextBox`, `addLine`, `addGroup` -> `"shape"`; `shape.children` += `adjustments: "adjustments"`; `adjustments: { scalars: ["count"], returns: { get: "clientResult" } }`; `lineFormat.scalars` += `color`, `weight`; `textFrame` children `textRange: "textRange"`; `textRange: { children: { font: "textFont", paragraphFormat: "paragraphFormat" } }`; `textFont: { scalars: ["name", "size", "color", "bold"] }`; `paragraphFormat: { scalars: ["horizontalAlignment"] }`.

- [ ] **Step 1: failing tests** (`test/fakeppt/shapes.test.ts`, using `installFakePpt({ slides: 1 })` and `PowerPoint.run` like `test/ppt.*.integration.test.ts`):

```ts
it("groups shapes by id, tags the group and deletes the members with it", async () => {
  await PowerPoint.run(async (c) => {
    const shapes = c.presentation.slides.getItemAt(0).shapes;
    const a = shapes.addGeometricShape("Rectangle", { left: 0, top: 0, width: 10, height: 10 });
    const b = shapes.addTextBox("1 234", { left: 20, top: 0, width: 30, height: 18 });
    a.load("id"); b.load("id"); await c.sync();
    const g = shapes.addGroup([a.id, b.id]); g.name = "pls,fix chart x"; g.tags.add("PLSFIX_LINK", "t"); g.load("id,type,left,top,width,height"); await c.sync();
    expect(g.type).toBe("Group"); expect([g.left, g.top, g.width, g.height]).toEqual([0, 0, 50, 18]);
    const all = c.presentation.slides.getItemAt(0).shapes; all.load("items/type"); await c.sync();
    expect(all.items.map((s) => s.type)).toEqual(["Group"]);
    g.delete(); await c.sync();
    const after = c.presentation.slides.getItemAt(0).shapes; after.load("items/id"); await c.sync();
    expect(after.items).toHaveLength(0);
  });
});
it("refuses an empty group", ...);                                   // addGroup([]) throws /InvalidArgument/
it("shapes a pie only after its first sync and normalises the angles", async () => {
  await PowerPoint.run(async (c) => {
    const pie = c.presentation.slides.getItemAt(0).shapes.addGeometricShape("Pie", { left: 0, top: 0, width: 100, height: 100 });
    expect(() => pie.adjustments.set(0, -90)).toThrow(/InvalidParam/);
    await c.sync();
    pie.adjustments.set(0, -90); pie.adjustments.set(1, 200); await c.sync();
    const a = pie.adjustments.get(0), b = pie.adjustments.get(1); pie.adjustments.load("count"); await c.sync();
    expect([pie.adjustments.count, a.value, b.value]).toEqual([2, -90, -160]);
  });
});
it("records text formatting, fills and line formats", ...);          // font/alignment/fillColor/lineColor on the model
```

- [ ] **Step 2: run, expect FAIL**; **Step 3: implement**; **Step 4: run** `npx vitest run test/fakeppt test/ppt.*.integration.test.ts test/links.table.integration.test.ts` green (the existing suites must not change).
- [ ] **Step 5: commit** "Fake PowerPoint: text boxes, lines, groups by id, adjustments".

---

### Task 6: The slide adapter and the host dispatch

**Files:**
- Create: `src/ppt/charts.ts` (insert, refresh, budget) and `src/ppt/chart-draw.ts` (primitive -> Office.js writes, chunked syncs) so each stays under 400 lines
- Modify: `src/ppt/host.ts` (`FoundLink.type`, `asFound`, `insertLink`, `refreshLink`, `replaceGroupWithPicture`, `InsertResult.note`), `src/ppt/links.ts:322-335` (toast the note like `OVERLAP_NOTE`)
- Create: `test/ppt.charts.integration.test.ts`; Modify: `test/ppt.perf.integration.test.ts`, `test/ppt.support.ts` (`seedChart(data, png)` seeding a picture payload with `chart`)

**Interfaces (produces):**

```ts
// src/ppt/charts.ts
export const SHAPES_PER_SYNC = 12, SHAPE_BUDGET_WEB = 30, SHAPE_BUDGET_DESKTOP = 200;
export const CHARTS_NEED_1_8 = "as a picture: shape charts need PowerPoint 2504/16.96 or newer";
export const PIES_NEED_1_10 = "as a picture: pie shapes need PowerPoint 2601/16.105 or newer";
export function shapeBudget(): number;                                 // Office.context.platform === Office.PlatformType.OfficeOnline ? WEB : DESKTOP
export function overBudgetNote(count: number, budget: number): string; // `as a picture: ${count} shapes is over this host's budget of ${budget}`
export interface ChartPlan { data: ChartData; size: Size; primitives: Primitive[] } // size = chartSize({ width: payload.width * 0.75, height: payload.height * 0.75 }, CONTENT_WIDTH); primitives = layoutChart(data, { left: 0, top: 0, ...size })
export function chartPlan(payload: PicturePayload): ChartPlan | null;
export function declineReason(plan: ChartPlan): string | null;        // null = draw it; else one of the three notes
export async function insertChart(stage: string, item: InboxItem, plan: ChartPlan, tag: LinkTag): Promise<InsertResult>;
export async function refreshChart(found: FoundLink, plan: ChartPlan, tag: LinkTag): Promise<void>;
// src/ppt/chart-draw.ts
export async function drawGroup(context: PowerPoint.RequestContext, shapes: PowerPoint.ShapeCollection, primitives: Primitive[], box: Box, name: string, tag: LinkTag, token: string, before?: () => void): Promise<string>;
//   adds primitives offset by box.left/top in syncs of SHAPES_PER_SYNC (each shape.load("id")), sets wedge adjustments in the sync after their add,
//   then addGroup(ids) + name + both tags (+ `before()` queued in that last sync: the caller's delete of the old group), returns the group id
```

Behaviour: `insertLink` in `host.ts`: `const plan = payload.kind === "picture" && payload.chart ? chartPlan(payload) : null; const reason = plan ? declineReason(plan) : null; if (plan && reason === null) return insertChart(...)`; otherwise the picture route as today, with `note: reason ?? undefined` in the result. `refreshLink`: `found.type === "Group"` -> plan present and drawable -> `refreshChart` (box = found left/top/width, height = width x size.height / size.width; `isGrouped(found)` throws `"${stage}: ungroup the chart before it can update"`); plan absent or declined -> `replaceGroupWithPicture(found, payload, tag)` (a rectangle with `fill.setImage` at the found box, tags, old group deleted in the same sync, exactly the order `recreate` in `tables.ts` uses); `found.type !== "Group"` -> the existing picture path unchanged. `refreshLinks` batches: unchanged (a group request is not a picture request: add `found.type === "Group"` to the `isPicture` exclusion). `src/ppt/links.ts`: after an insert, toast `placed.note` when present (the overlap note first if both).

- [ ] **Step 1: failing tests** (`test/ppt.charts.integration.test.ts`, booted with `bootPpt()`; `seedChart` publishes a picture payload with `chart`):

```ts
const column = /* the ChartData fixture from Task 2 with 6 categories, 1 series */;
it("inserts a chart link as a tagged group of native shapes in free space", async () => {
  const item = await seedChart(column, fakePng(800, 400));
  await links.insertFromInbox(item);                                   // the pane action links.ts already exposes for the inbox row
  const slide = presentation.slides[0]!;
  const group = slide.shapes.find((s) => s.type === "Group")!;
  expect(group.name).toBe("pls,fix chart " + item.label);
  expect(group.tags.get("PLSFIX_LINK")).toBeDefined();
  expect(group.group!.shapes.map((s) => s.type)).toContain("TextBox");
  expect(group.group!.shapes.filter((s) => s.geometry === "Rectangle")).toHaveLength(6);
  expect(group.group!.shapes.every((s) => s.tags.size === 0)).toBe(true);
});
it("draws in syncs of twelve and one more for the group", async () => {  // 20 primitives -> 2 + 1 syncs beyond placement's
  const before = helpers.syncCount(); await links.insertFromInbox(await seedChart(column, png));
  expect(helpers.syncCount() - before).toBeLessThanOrEqual(6);
});
it("rebuilds the group at the same corner and width on update", async () => {
  ... insert; move the group (left 80, top 320, width 380) through the fake; push new data (7 categories); refresh all;
  const group = ...; expect([group.left, group.top, group.width]).toEqual([80, 320, 380]); expect(rects).toHaveLength(7); expect(old group id gone)
});
it("keeps a picture a picture when a later push carries chart data", ...);   // seedLink picture -> insert -> pushAgain with chart -> refresh -> still Image/rectangle, no Group
it("turns a chart group into a picture when the source chart stopped being drawable", ...); // push a payload without chart -> refresh -> rectangle with fillImage at the old corner, group gone
it("refuses to update a chart the user grouped with something else", ...);  // helpers.group([our group, another shape]) -> refresh -> row error /ungroup the chart/
it("inserts the picture and says why when the host lacks 1.10 for a pie", async () => {
  helpers.setSupported((_set, v) => v !== "1.10");                       // the fake's requirement hook (test/fakeppt/index.ts)
  const placed = await links.insertFromInbox(await seedChart(pie, png)); // insertFromInbox returns the InsertResult
  expect(presentation.slides[0]!.shapes.some((s) => s.type === "Group")).toBe(false);
  expect(placed.note).toMatch(/pie shapes need PowerPoint/);
});
it("inserts the picture and says why over the host's shape budget", ...);   // helpers.setPlatform("OfficeOnline") (added to the fake in Task 5: sets Office.context.platform) + 3 series x 12 points -> picture + note "over this host's budget of 30"
it("shapes pie wedges in the sync after their add", ...);                     // wedges' adjustments equal layoutChart's start/end
```

- [ ] **Step 2: run, expect FAIL**; **Step 3: implement** `chart-draw.ts` then `charts.ts` then the host dispatch, then the toast; every function under 50 lines; **Step 4: run** `npx vitest run test/ppt.charts.integration.test.ts test/ppt.*.integration.test.ts test/links.*.integration.test.ts` green; `npm run check`.
- [ ] **Step 5: commit** "Chart links land as native shape groups; picture below the budget".

---

### Task 7: Words: help, manual, features, README, Map

**Files:**
- Modify: `src/help/copy.ppt.ts` (`inbox-heading.about` + the insert row sentence: "A chart from Excel lands as native shapes you can edit, labelled with its values; where this PowerPoint cannot draw it, or the chart is too big, it lands as a picture and the pane says so."), `src/help/copy.excel.ts` (the Links tab's export-chart sentence: "... on the slide it becomes editable shapes on PowerPoint 2504/16.96 and newer, a picture elsewhere.")
- Modify: `manual/src/content/links.rs` (one Latvian paragraph in the chapter that describes inserting from the inbox: "Diagramma no Excel slaidā nonāk kā rediģējamu figūru grupa ar zīmola krāsām un tikai vērtību etiķetēm; ja PowerPoint to nevar uzzīmēt vai diagramma ir par lielu (vairāk nekā 40 punkti vai 3 sērijas), tā tiek ievietota kā attēls, un panelis to pasaka."), regenerate with `npm run manual`
- Modify: `docs/FEATURES.md:145` (status "shipped v2.0; native shape groups for column, bar, waterfall and pie charts v2.4"), `README.md` (one line in the feature list, surgical), `CLAUDE.md` Map (lines for `src/link/chart-model.ts`, `src/chart-colors.ts`, `src/chart-shapes.ts`, `src/excel/link-chart.ts`, `src/ppt/charts.ts` + `chart-draw.ts`; the data-flow paragraph gains "a chart also ships its data, drawn as a group by `src/ppt/charts.ts`"; and dedupe: the `src/excel/` bullet appears four times and "Where bugs live" five times from merges, keep the newest tokens of each, the Map back under 40 lines)
- Test: `npx vitest run src/help/copy.test.ts`; `cargo check --manifest-path manual/Cargo.toml`; `npm run check`

- [ ] **Step 1: run `src/help/copy.test.ts`** to see which keys the new sentences must keep; **Step 2: edit**; **Step 3: run** the tests, `npm run manual`; **Step 4: commit** "Native charts: help, manual, features, Map".

---

### Task 8: Controller: merge, gate, bump, deploy, prove

- [ ] Merge order: Task 1+3 (main), wave 1 branches (Task 2, Task 4, Task 5), Task 6, Task 7. After each merge: `test -d node_modules && ! test -L node_modules`, `npm run check`, `npm run ux:check`.
- [ ] `sh scripts/bump-patch.sh minor` (v2.4.0: a new capability), `db deploy modelis`, `curl https://dbautomatizacijas.com/modelis/version`.
- [ ] Web proof through the CDP rig (`scratchpad/driver/drive.mjs`, lessons 30.08): export the demo "Revenue chart" (Bridge sheet) and "Segment pie" (Rounding) from Excel for the web, insert both from the PowerPoint inbox on slide 3, read back one `Group` per link with both tags and the expected child counts, change a source value, "Update all", read back the rebuilt group at the same corner; export "EBITDA margin chart" (Line) and confirm it arrives as a picture. Screenshots into the scratchpad; Daniel's desktop pass proves the desktop budget and the text metrics.
- [ ] `tasks/AUTORESUME.md` status, memory update, `npm run manual` if Task 7 did not.

## Self-review

- Spec coverage: payload (3), reader (4), layout (2), adapter + dispatch + budget + batching (6), fakes (5), colours (1), docs (7), proof (8). The spec's "chart-to-line -> picture at the corner" and "grouped -> refuse" are Task 6 tests; "labels are strings decided in Excel" is Task 4; "a full slice is an ellipse" is Task 2 + Task 6 draw.
- Names used consistently: `ChartData`, `ChartSeries`, `Primitive` (`Rect`, `Wedge`, `Ellipse`, `Text`, `Line`), `layoutChart`, `chartSize`, `normalizeAngle`, `chartKind`, `readChartData`, `chartPlan`, `declineReason`, `insertChart`, `refreshChart`, `drawGroup`, `seedChart`, `InsertResult.note`, `FoundLink.type`.
