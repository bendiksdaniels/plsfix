// A chart the slide has to scale down to fit its own height can end up
// smaller than MIN_SIZE - a group of shapes too small to read. It declines to
// the picture with its own note, the way every other decline does, while a
// chart that still clears the minimum draws as a group. Strict load semantics
// are on, so a batch that reads a scalar it never loaded fails here.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { MIN_SIZE } from "../src/chart-shapes";
import { createWorkspace } from "../src/link/workspace";
import type { InboxItem } from "../src/link/model";
import type { InsertResult } from "../src/ppt/host";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import { bootPpt, columnChart, memoryStore, seedChart } from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

// 200 x 2000 pixels are 150 x 1500 points: a chart three times taller than the
// slide's content height, so the fit scales it to 46.8 x 468 - a fifth of the
// 200 pt MIN_SIZE asks for across.
const TALL = fakePng(200, 2000);
// 800 x 400 pixels are 600 x 300 points: the ordinary export, untouched by the
// fit, and the shape the rest of the chart suites use.
const NORMAL = fakePng(800, 400);
// 4000 x 100 pixels are 3000 x 75 points. chartSize floors a chart's height at
// MIN_SIZE before the slide ever sees it, so the flattest export there is
// still lands at 888 x 120 and draws: only a TALL one is ever scaled down.
const FLAT = fakePng(4000, 100);

const TOO_SMALL =
  "as a picture: the chart would be smaller than 200 x 120 pt on this slide";

const COLUMN = columnChart(6);

let links: typeof LinksModule;
let presentation: FakePresentation;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, presentation, relay } = await bootPpt());
});
afterEach(() => {
  uninstallFakePpt();
});

async function insert(
  png: string,
): Promise<{ item: InboxItem; placed: InsertResult }> {
  const ws = await createWorkspace(memoryStore());
  const item = await seedChart(COLUMN, png);
  const placed = await links.insertFromInbox(item, ws, relay);
  return { item, placed };
}

function shapes(): FakePptShape[] {
  return presentation.slides[0]!.shapes;
}

describe("a chart the slide would shrink below the minimum", () => {
  it("inserts the picture and says the chart would be too small", async () => {
    const { placed } = await insert(TALL);
    expect(shapes()).toHaveLength(1);
    expect(shapes()[0]!.type).not.toBe("Group");
    expect(shapes()[0]!.fillImage).toBe(TALL);
    expect(placed.note).toBe(TOO_SMALL);
  });

  it("names the minimum the layout actually enforces", () => {
    expect(TOO_SMALL).toContain(
      `${String(MIN_SIZE.width)} x ${String(MIN_SIZE.height)} pt`,
    );
  });

  it("draws the group when the fitted chart still clears the minimum", async () => {
    const { placed } = await insert(NORMAL);
    expect(shapes()[0]!.type).toBe("Group");
    expect(placed.note).toBeUndefined();
  });

  it("draws a very wide, very flat chart: the fit never shrinks one", async () => {
    const { placed } = await insert(FLAT);
    expect(shapes()[0]!.type).toBe("Group");
    expect(shapes()[0]!.width).toBe(888);
    expect(shapes()[0]!.height).toBe(MIN_SIZE.height);
    expect(placed.note).toBeUndefined();
  });
});
