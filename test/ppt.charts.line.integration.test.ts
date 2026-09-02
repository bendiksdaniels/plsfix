// A native line chart with a rising segment against the fake host: a rising
// connector cannot be drawn as ConnectorType.straight (see chart-draw.ts), so
// it lands as the host's LineInverse geometric shape instead. Split out of
// ppt.charts.integration.test.ts, already at the file's line cap, the same
// way ppt.support.ts is shared across the whole ppt.*.integration.test.ts
// family so each file stays under it.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { ChartData } from "../src/link/chart-model";
import { createWorkspace } from "../src/link/workspace";
import type * as LinksModule from "../src/ppt/links";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import { bootPpt, memoryStore, seedChart } from "./ppt.support";

enableStrictLoadSemantics();

const PNG = fakePng(800, 400);

// 100, -50, 200: point 0 to 1 falls, point 1 to 2 rises. A rising
// connector's naive box has a negative height (valueY maps a bigger value to
// a smaller top), which the real host refuses on both the add and the
// width/height writes after it.
const LINE_RISE: ChartData = {
  v: 1,
  kind: "line",
  title: null,
  categories: ["A", "B", "C"],
  series: [
    {
      name: "s",
      values: [100, -50, 200],
      labels: ["100", "(50)", "200"],
      colors: ["#2EC4B6", "#2EC4B6", "#2EC4B6"],
    },
  ],
  font: "Aptos Narrow",
  ink: "#282623",
  titleColor: "#14213D",
};
// 3 markers, 2 connectors, 3 point labels, 1 baseline, 3 category labels, no
// title.
const LINE_RISE_SHAPES = 12;

let links: typeof LinksModule;
let presentation: FakePresentation;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, presentation, relay } = await bootPpt());
});
afterEach(() => {
  uninstallFakePpt();
});

function children(): FakePptShape[] {
  const group = presentation.slides[0]!.shapes.find(
    (shape) => shape.type === "Group",
  )!;
  return group.group!.shapes;
}

function geometries(geometry: string): FakePptShape[] {
  return children().filter((shape) => shape.geometry === geometry);
}

describe("insert a chart link: a line chart with a rising segment", () => {
  it("draws the rising connector as a LineInverse geometric shape, with no negative geometry anywhere", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(LINE_RISE, PNG);
    await links.insertFromInbox(item, ws, relay);

    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(children()).toHaveLength(LINE_RISE_SHAPES);
    expect(geometries("Ellipse")).toHaveLength(3);
    // The falling segment and the flat baseline stay native lines; only the
    // rising segment becomes the geometric shape.
    expect(geometries("Straight")).toHaveLength(2);
    const rising = geometries("LineInverse");
    expect(rising).toHaveLength(1);
    expect(rising[0]!.lineColor).toBe("#2EC4B6");
    // A connector carries no fill, whichever shape kind draws it.
    expect(rising[0]!.fillColor).toBeNull();
    expect(rising[0]!.fillCleared).toBe(false);
    expect(
      children().every((shape) => shape.width >= 0 && shape.height >= 0),
    ).toBe(true);
  });
});
