// Stress pass on what Excel can put on the relay: a chart payload at and past
// every cap the slide draws to, values a chart cannot be drawn from (all zero,
// negative, NaN) and a text link at and past its own cap.
// Invariant: a payload the slide cannot draw arrives as the picture beside it
// with the reason said in words, and no shape is ever given a negative side.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  CHART_MAX_POINTS,
  CHART_MAX_SERIES,
  CHART_MAX_SLICES,
  CHART_MIN_POINTS,
  type ChartData,
} from "../src/link/chart-model";
import { TEXT_MAX_CHARS } from "../src/link/model";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import { fakePng } from "./fakepng";
import type { FakeRelay } from "./fakerelay";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import {
  bootPpt,
  columnChart,
  memoryStore,
  seedChart,
  seedText,
} from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

const UNREADABLE = "as a picture: chart data unreadable";
const PNG = fakePng(600, 300);

let links: typeof LinksModule;
let presentation: FakePresentation;
let relay: FakeRelay;
let ws: Workspace;

beforeEach(async () => {
  ({ links, presentation, relay } = await bootPpt());
  ws = await createWorkspace(memoryStore());
});
afterEach(() => {
  uninstallFakePpt();
});

// A chart of `series` series over `points` categories, of any kind: the shape
// every cap in chart-model is measured on.
function chart(kind: ChartData["kind"], points: number, series = 1): ChartData {
  const base = columnChart(Math.max(points, 1));
  const categories = Array.from({ length: points }, (_one, index) =>
    String(2000 + index),
  );
  return {
    ...base,
    kind,
    categories,
    series: Array.from({ length: series }, (_one, index) => ({
      name: `Series ${String(index + 1)}`,
      values: categories.map((_c, at) => 10 + at),
      labels: categories.map((_c, at) => String(10 + at)),
      colors: categories.map(() => "#B27E54"),
    })),
  };
}

function withValues(kind: ChartData["kind"], values: number[]): ChartData {
  const data = chart(kind, values.length);
  return {
    ...data,
    series: [{ ...data.series[0]!, values }],
  };
}

function shapesOn(slide = 0): FakePptShape[] {
  return presentation.slides[slide]!.shapes;
}

function lastShape(): FakePptShape {
  const shapes = shapesOn();
  return shapes[shapes.length - 1]!;
}

// Every shape a draw put on the slide, groups walked: the fake refuses a
// negative side on the way in, and this catches one that arrived as NaN.
function allShapes(shapes: FakePptShape[]): FakePptShape[] {
  return shapes.flatMap((shape) => [
    shape,
    ...(shape.group ? allShapes(shape.group.shapes) : []),
  ]);
}

function expectDrawable(): void {
  for (const shape of allShapes(shapesOn())) {
    expect(Number.isFinite(shape.width)).toBe(true);
    expect(Number.isFinite(shape.height)).toBe(true);
    expect(shape.width).toBeGreaterThanOrEqual(0);
    expect(shape.height).toBeGreaterThanOrEqual(0);
  }
}

async function insert(
  data: ChartData | null,
  issue?: string,
): Promise<string | undefined> {
  const item = await seedChart(data, PNG, issue);
  const placed = await links.insertFromInbox(item, ws, relay);
  return placed.note;
}

describe("a chart payload at the cap and one past it", () => {
  it.each([
    ["points", CHART_MAX_POINTS, () => chart("column", CHART_MAX_POINTS)],
    ["series", CHART_MAX_SERIES, () => chart("bar", 3, CHART_MAX_SERIES)],
    ["slices", CHART_MAX_SLICES, () => chart("pie", CHART_MAX_SLICES)],
    ["the fewest points", CHART_MIN_POINTS, () => chart("line", 2)],
  ] as [string, number, () => ChartData][])(
    "draws a chart with exactly the %s cap (%i) as a group",
    async (_name, _cap, make) => {
      expect(await insert(make())).toBeUndefined();

      expect(lastShape().type).toBe("Group");
      expectDrawable();
    },
  );

  it.each([
    ["one point too many", () => chart("column", CHART_MAX_POINTS + 1)],
    ["one series too many", () => chart("bar", 3, CHART_MAX_SERIES + 1)],
    ["one slice too many", () => chart("pie", CHART_MAX_SLICES + 1)],
    ["a single point", () => chart("column", CHART_MIN_POINTS - 1)],
    ["a value that is not a number", () => withValues("column", [10, NaN, 12])],
    ["a value that is infinite", () => withValues("line", [1, Infinity])],
  ] as [string, () => ChartData][])(
    "takes the picture and says why for %s",
    async (_name, make) => {
      expect(await insert(make())).toBe(UNREADABLE);

      const shape = lastShape();
      expect(shape.type).toBe("GeometricShape");
      expect(shape.fillImage).toBe(PNG);
      expectDrawable();
      // Still a link: the row updates like any other picture.
      expect(await links.listLinks(relay)).toHaveLength(1);
    },
  );

  it("keeps Excel's own reason when the payload carries one", async () => {
    const note = await insert(
      chart("column", CHART_MAX_POINTS + 1),
      "41 points; shapes draw up to 40",
    );

    expect(note).toBe("as a picture: 41 points; shapes draw up to 40");
  });
});

describe("values no chart can be drawn from", () => {
  it("draws a pie whose values are all zero without a wedge or a bad box", async () => {
    expect(await insert(withValues("pie", [0, 0, 0, 0]))).toBeUndefined();

    const group = lastShape();
    expect(group.type).toBe("Group");
    expect(
      group.group!.shapes.filter((shape) => shape.geometry === "Pie"),
    ).toHaveLength(0);
    expectDrawable();
  });

  it("draws only the positive slices of a pie that holds negatives", async () => {
    expect(await insert(withValues("pie", [-4, -3, 10, 2]))).toBeUndefined();

    const group = lastShape();
    const wedges = group.group!.shapes.filter(
      (shape) => shape.geometry === "Pie",
    );
    expect(wedges).toHaveLength(2);
    expectDrawable();
  });

  it("draws a single positive slice as a full circle", async () => {
    expect(await insert(withValues("pie", [0, 12, 0]))).toBeUndefined();

    const group = lastShape();
    expect(
      group.group!.shapes.filter((shape) => shape.geometry === "Ellipse"),
    ).toHaveLength(1);
    expectDrawable();
  });

  it("draws a column chart whose values are all zero", async () => {
    expect(await insert(withValues("column", [0, 0, 0]))).toBeUndefined();

    expect(lastShape().type).toBe("Group");
    expectDrawable();
  });

  it("draws a waterfall that closes below zero", async () => {
    expect(await insert(withValues("waterfall", [10, -40, 5]))).toBeUndefined();

    expect(lastShape().type).toBe("Group");
    expectDrawable();
  });
});

describe("a text link at its cap", () => {
  it("lands the longest text Excel will export on one line", async () => {
    const item = await seedText("9".repeat(TEXT_MAX_CHARS));

    await links.insertFromInbox(item, ws, relay);

    const shape = lastShape();
    expect(shape.type).toBe("TextBox");
    expect(shape.text).toHaveLength(TEXT_MAX_CHARS);
    // Wide text is clamped to the slide's content width, never past the edge.
    expect(shape.width).toBeLessThanOrEqual(960 - 2 * 36);
    expect(shape.wordWrap).toBe(false);
  });

  it("still lands one Excel would have refused, rather than throwing", async () => {
    const item = await seedText("9".repeat(TEXT_MAX_CHARS + 1));

    await links.insertFromInbox(item, ws, relay);

    expect(lastShape().text).toHaveLength(TEXT_MAX_CHARS + 1);
    expect(lastShape().width).toBeLessThanOrEqual(960 - 2 * 36);
  });

  it("lands an empty text without a box of no width", async () => {
    const item = await seedText("");

    await links.insertFromInbox(item, ws, relay);

    const shape = lastShape();
    expect(shape.type).toBe("TextBox");
    expect(shape.width).toBeGreaterThan(0);
    expect(shape.height).toBeGreaterThan(0);
  });
});
