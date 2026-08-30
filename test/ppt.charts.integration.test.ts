// Native chart groups against the fake host and the fake relay: an insert
// draws the layout's primitives as one tagged group in the slide's free space,
// an update rebuilds it at the corner and width the user left it at, and a
// host that cannot draw the chart - or would spend more shapes than it can
// afford - gets the picture and a note saying why. Strict load semantics are
// on, so a batch that reads a scalar it never loaded fails here.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { layoutChart, type Wedge } from "../src/chart-shapes";
import { overlaps } from "../src/layout";
import type { ChartData } from "../src/link/chart-model";
import { TAG_KEY, TAG_LINK, type InboxItem } from "../src/link/model";
import { createWorkspace } from "../src/link/workspace";
import { SHAPES_PER_SYNC } from "../src/ppt/charts";
import type { InsertResult } from "../src/ppt/host";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import {
  bootPpt,
  columnChart,
  memoryStore,
  pushChart,
  seedChart,
  seedLink,
} from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

const PNG = fakePng(800, 400);
// 800 x 400 pixels are 600 x 300 points, and the chart is laid out at that
// size: what "Model: B4:F12" occupies, centred on an empty slide.
const SIZE = { width: 600, height: 300 };
const BOX = { left: 180, top: 120, ...SIZE };
const LABEL = "Model: B4:F12";
const GROUP_NAME = `pls,fix chart ${LABEL}`;

const COLUMN = columnChart(6);
const GREW = columnChart(7);
// One title, six bars, six value labels, a baseline and six category labels.
const COLUMN_SHAPES = 20;

const PIE: ChartData = {
  v: 1,
  kind: "pie",
  title: "Mix",
  categories: ["North", "South", "West"],
  series: [
    {
      name: "Mix",
      values: [1, 1, 2],
      labels: ["1", "1", "2"],
      colors: ["#14213D", "#2EC4B6", "#B27E54"],
    },
  ],
  font: "Aptos Narrow",
  ink: "#282623",
  titleColor: "#14213D",
};

// Three series over twelve points: 92 primitives, inside the desktop budget
// and far past the web's.
const WIDE: ChartData = {
  ...COLUMN,
  categories: Array.from({ length: 12 }, (_, i) => `C${String(i)}`),
  series: [0, 1, 2].map((j) => {
    const values = Array.from({ length: 12 }, (_, i) => 10 + i + j);
    return {
      name: `S${String(j)}`,
      values,
      labels: values.map(String),
      colors: values.map(() => "#B27E54"),
    };
  }),
};
const WIDE_SHAPES = 92;

let links: typeof LinksModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, presentation, helpers, relay } = await bootPpt());
});
afterEach(() => {
  uninstallFakePpt();
});

async function insert(
  data: ChartData = COLUMN,
): Promise<{ item: InboxItem; placed: InsertResult }> {
  const ws = await createWorkspace(memoryStore());
  const item = await seedChart(data, PNG);
  const placed = await links.insertFromInbox(item, ws, relay);
  return { item, placed };
}

function shapes(): FakePptShape[] {
  return presentation.slides[0]!.shapes;
}

function group(): FakePptShape {
  return shapes().find((shape) => shape.type === "Group")!;
}

function children(): FakePptShape[] {
  return group().group!.shapes;
}

function geometries(geometry: string): FakePptShape[] {
  return children().filter((shape) => shape.geometry === geometry);
}

function box(one: FakePptShape) {
  return { left: one.left, top: one.top, width: one.width, height: one.height };
}

async function updateAll(): Promise<LinksModule.UpdateSummary> {
  return links.updateLinks(await links.listLinks(relay), relay);
}

describe("insert a chart link", () => {
  it("draws it as one tagged group of native shapes in the free space", async () => {
    const { item } = await insert();
    expect(shapes()).toHaveLength(1);
    const drawn = group();
    expect(drawn.name).toBe(GROUP_NAME);
    // The layout fills the box it was given, so the group's own box - the one
    // scanLinks reports as the link's geometry - is the placement's.
    expect(box(drawn)).toEqual(BOX);

    expect(children()).toHaveLength(COLUMN_SHAPES);
    // Six bars, and no legend swatch beside them: one series needs no legend.
    const bars = geometries("Rectangle");
    expect(bars).toHaveLength(6);
    expect(bars.map((bar) => bar.fillColor)).toEqual(
      new Array<string>(6).fill("#B27E54"),
    );
    expect(bars.every((bar) => !bar.lineVisible)).toBe(true);
    // The title, one value label per bar and one category label per bar; the
    // fake records a text box by the text in it, not by a host shape type.
    const written = children().filter((shape) => shape.text !== null);
    expect(written).toHaveLength(13);
    expect(written.map((shape) => shape.text)).toContain("Revenue");
    // The label is the string Excel decided; the slide never formats a number.
    expect(written.map((shape) => shape.text)).toContain("220");
    expect(written.map((shape) => shape.text)).toContain("2026");
    expect(geometries("Straight")).toHaveLength(1);

    expect(drawn.tags.get(TAG_KEY)).toBe(item.token);
    expect(JSON.parse(drawn.tags.get(TAG_LINK)!)).toMatchObject({
      id: item.id,
      kind: "chart",
      rev: 1,
    });
    // The tags live on the group alone: a child carrying them would be found
    // as a second link on the same chart.
    expect(children().every((shape) => shape.tags.size === 0)).toBe(true);
    expect(
      await links.listInbox(await createWorkspace(memoryStore()), relay),
    ).toHaveLength(0);
  });

  it("names every shape after the part of the chart it draws", async () => {
    await insert();
    const names = children().map((shape) => shape.name);
    expect(names[0]).toBe(`${GROUP_NAME}: title`);
    expect(names).toContain(`${GROUP_NAME}: bar 0.0`);
    expect(names).toContain(`${GROUP_NAME}: label 0.5`);
    expect(names).toContain(`${GROUP_NAME}: baseline`);
    expect(names).toContain(`${GROUP_NAME}: category 5`);
    // A label carries the brand font, its size, its colour and its alignment,
    // and nothing else: every extra property write costs a round trip's worth
    // of time on the web.
    const title = children()[0]!;
    expect(title.font).toEqual({
      name: "Aptos Narrow",
      size: 12,
      color: "#14213D",
      bold: true,
    });
    expect(title.alignment).toBe("Left");
    expect(title.margins).toEqual({
      left: null,
      right: null,
      top: null,
      bottom: null,
    });
    expect([title.autoSize, title.wordWrap, title.fillColor]).toEqual([
      null,
      null,
      null,
    ]);
  });

  it("lands clear of what the slide already holds", async () => {
    const caption = { left: 5, top: 5, width: 60, height: 40 };
    presentation.addShape(presentation.slides[0]!, caption);
    await insert();
    // Free space is found for the whole chart, not for a shape at a time: the
    // group's box is the placement's, and nothing of it lands on the caption.
    expect(overlaps(box(group()), caption)).toBe(false);
    expect(box(group())).toMatchObject({ ...SIZE });
  });

  it("draws in syncs of twelve and one more for the group", async () => {
    const before = helpers.syncCount();
    await insert();
    // Two for the placement (the selected slide, then the boxes already on
    // it), two chunks of at most twelve for twenty shapes, and one that
    // groups, names and tags them: five, and never one per shape.
    const chunks = Math.ceil(COLUMN_SHAPES / SHAPES_PER_SYNC);
    expect(chunks).toBe(2);
    expect(helpers.syncCount() - before).toBe(2 + chunks + 1);
  });

  it("shapes pie wedges in the sync after their add", async () => {
    await insert(PIE);
    // The fake refuses adjustments on a shape it has not synced yet, exactly
    // as PowerPoint does ("InvalidParam passed to GetItem(id)"), so an insert
    // that got this far shaped its wedges a sync later than it added them.
    const wedges = geometries("Pie");
    expect(wedges).toHaveLength(3);
    const planned = layoutChart(PIE, { left: 0, top: 0, ...SIZE }).filter(
      (one): one is Wedge => one.kind === "wedge",
    );
    expect(wedges.map((one) => one.adjustments)).toEqual(
      planned.map((one) => [one.start, one.end]),
    );
    expect(wedges.map((one) => one.fillColor)).toEqual([
      "#14213D",
      "#2EC4B6",
      "#B27E54",
    ]);
  });

  it("inserts the picture and says why when the host lacks 1.10 for a pie", async () => {
    helpers.setSupported((_set, version) => version !== "1.10");
    const { placed } = await insert(PIE);
    expect(shapes().some((shape) => shape.type === "Group")).toBe(false);
    expect(shapes()[0]!.fillImage).toBe(PNG);
    expect(placed.note).toMatch(/pie shapes need PowerPoint/);
  });

  it("inserts the picture and says why over the host's shape budget", async () => {
    helpers.setPlatform("OfficeOnline");
    const { placed } = await insert(WIDE);
    expect(shapes().some((shape) => shape.type === "Group")).toBe(false);
    expect(placed.note).toBe(
      `as a picture: ${String(WIDE_SHAPES)} shapes is over this host's budget of 30`,
    );
    // The same chart is inside the desktop budget and draws there.
    helpers.setPlatform("Mac");
    helpers.selectSlide(presentation.slides[1]!.id);
    const second = await insert(WIDE);
    expect(second.placed.note).toBeUndefined();
    expect(presentation.slides[1]!.shapes[0]!.type).toBe("Group");
  });
});

describe("update a chart link", () => {
  // What the user does with a chart once it is on the slide: drags it into a
  // corner and pulls it narrower.
  function moveTo(left: number, top: number, width: number): void {
    const drawn = group();
    drawn.left = left;
    drawn.top = top;
    drawn.width = width;
    drawn.height = Math.round(width / 2);
  }

  it("rebuilds the group at the corner and width the user left it at", async () => {
    const { item } = await insert();
    const first = group().id;
    moveTo(80, 320, 380);
    await pushChart(item, GREW, PNG);

    expect(await updateAll()).toMatchObject({ updated: 1, failed: 0 });
    expect(shapes()).toHaveLength(1);
    const drawn = group();
    expect(drawn.id).not.toBe(first);
    expect(presentation.peekShape(first)).toBeNull();
    // Left, top and width are the user's; only the height follows the chart's
    // own aspect at that width.
    expect(box(drawn)).toEqual({ left: 80, top: 320, width: 380, height: 190 });
    expect(geometries("Rectangle")).toHaveLength(7);
    expect(children()).toHaveLength(23);
    expect(drawn.name).toBe(GROUP_NAME);
    expect(drawn.tags.get(TAG_KEY)).toBe(item.token);
    expect(JSON.parse(drawn.tags.get(TAG_LINK)!)).toMatchObject({
      id: item.id,
      kind: "chart",
      rev: 2,
    });
    expect(await links.listLinks(relay)).toHaveLength(1);
  });

  it("repaints a chart in one sync per twelve shapes and one for the group", async () => {
    const { item } = await insert();
    await pushChart(item, GREW, PNG);
    const rows = await links.listLinks(relay);
    const before = helpers.syncCount();
    await links.updateLinks(rows, relay);
    // Twenty-three shapes: two chunks and the grouping sync. Nothing is placed
    // and nothing is read, so the insert's two placement syncs are not spent.
    expect(helpers.syncCount() - before).toBe(3);
  });

  it("turns a chart group into a picture when the source stopped being drawable", async () => {
    const { item } = await insert();
    moveTo(80, 320, 380);
    await pushChart(item, null, PNG);

    expect(await updateAll()).toMatchObject({ updated: 1, failed: 0 });
    expect(shapes()).toHaveLength(1);
    const picture = shapes()[0]!;
    expect(picture.type).not.toBe("Group");
    expect(picture.geometry).toBe("Rectangle");
    expect(picture.fillImage).toBe(PNG);
    // The corner and width the user chose, the height from the picture.
    expect(box(picture)).toEqual({
      left: 80,
      top: 320,
      width: 380,
      height: 190,
    });
    expect(picture.tags.get(TAG_KEY)).toBe(item.token);
    expect(JSON.parse(picture.tags.get(TAG_LINK)!).rev).toBe(2);
  });

  it("keeps a picture a picture when a later push carries chart data", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(PNG);
    await links.insertFromInbox(item, ws, relay);
    const first = shapes()[0]!.id;
    await pushChart(item, COLUMN, PNG);

    expect(await updateAll()).toMatchObject({ updated: 1, failed: 0 });
    expect(shapes()).toHaveLength(1);
    // The representation chosen at insert is the link's for life.
    expect(shapes()[0]!.id).toBe(first);
    expect(shapes()[0]!.type).not.toBe("Group");
    expect(shapes()[0]!.setImageCalls).toBe(2);
  });

  it("refuses to update a chart the user grouped with something else", async () => {
    const { item } = await insert();
    const chart = group().id;
    const slide = presentation.slides[0]!;
    const caption = presentation.addShape(slide, {
      left: 5,
      top: 5,
      width: 60,
    });
    presentation.groupShapes([chart, caption.id], slide.id);
    await pushChart(item, GREW, PNG);

    const summary = await updateAll();
    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expect(summary.failures[0]).toMatch(/ungroup the chart/);
    // Nothing was drawn and nothing was taken down.
    expect(presentation.findShape(chart).shape.group!.shapes).toHaveLength(
      COLUMN_SHAPES,
    );
    expect(shapes()).toHaveLength(1);
  });
});
