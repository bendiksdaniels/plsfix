// Audit suite for native slide charts: what happens when the host takes a
// draw batch and never answers it, what the shape budget does exactly at its
// edge, and what a host without the group or pie APIs inserts instead.
// Strict load semantics are on.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { ChartData } from "../src/link/chart-model";
import { TAG_KEY, TAG_LINK, type InboxItem } from "../src/link/model";
import { createWorkspace } from "../src/link/workspace";
import { cleanupShapes } from "../src/ppt/chart-cleanup";
import { SYNC_TIMEOUT_MS } from "../src/ppt/chart-draw";
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
} from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

const PNG = fakePng(800, 400);
const SIZE = { width: 600, height: 300 };
const BOX = { left: 180, top: 120, ...SIZE };
const COLUMN = columnChart(6);
// The insert's syncs: 1 the selected slide, 2 the boxes already on it, 3 the
// first chunk of twelve, 4 the second chunk of eight, 5 the group.
const SECOND_CHUNK_SYNC = 3;
const SILENT_NOTE =
  "as a picture: PowerPoint stopped answering while drawing the shapes";

let links: typeof LinksModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, presentation, helpers, relay } = await bootPpt());
});
afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
});

function shapes(): FakePptShape[] {
  return presentation.slides[0]!.shapes;
}

function box(one: FakePptShape) {
  return { left: one.left, top: one.top, width: one.width, height: one.height };
}

// Drives the fake clock until the work under test settles, so a per-sync
// deadline can be proven without the test waiting a real minute. A promise
// that never settles leaves this loop and fails on vitest's own timeout,
// which is exactly what a pane stuck busy for ever looks like.
async function settle<T>(work: Promise<T>): Promise<T> {
  let done = false;
  void work.then(
    () => (done = true),
    () => (done = true),
  );
  for (let step = 0; step < 10 && !done; step += 1) {
    await vi.advanceTimersByTimeAsync(SYNC_TIMEOUT_MS);
  }
  return work;
}

async function insert(
  data: ChartData = COLUMN,
): Promise<{ item: InboxItem; placed: InsertResult }> {
  const ws = await createWorkspace(memoryStore());
  const item = await seedChart(data, PNG);
  const placed = await links.insertFromInbox(item, ws, relay);
  return { item, placed };
}

describe("a draw batch the host never answers", () => {
  it("cleans up, lands the picture and says why, instead of leaving the pane busy", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(COLUMN, PNG);
    helpers.hangNextSync(SECOND_CHUNK_SYNC);
    vi.useFakeTimers();

    const placed = await settle(links.insertFromInbox(item, ws, relay));

    // One shape on the slide: the picture. The twelve shapes the first chunk
    // had already committed are gone, and no half-built group survived.
    expect(shapes()).toHaveLength(1);
    const picture = shapes()[0]!;
    expect(picture.type).not.toBe("Group");
    expect(picture.fillImage).toBe(PNG);
    expect(box(picture)).toEqual(BOX);
    expect(picture.tags.get(TAG_KEY)).toBe(item.token);
    expect(JSON.parse(picture.tags.get(TAG_LINK)!)).toMatchObject({
      id: item.id,
      kind: "chart",
      rev: 1,
    });
    expect(placed.note).toBe(SILENT_NOTE);
    expect(placed.shapeId).toBe(picture.id);
  });

  it("draws the next chart normally: one hung batch is not a jammed pane", async () => {
    const ws = await createWorkspace(memoryStore());
    const first = await seedChart(COLUMN, PNG);
    helpers.hangNextSync(SECOND_CHUNK_SYNC);
    vi.useFakeTimers();
    await settle(links.insertFromInbox(first, ws, relay));
    vi.useRealTimers();

    helpers.selectSlide(presentation.slides[1]!.id);
    const second = await seedChart(COLUMN, PNG);
    const placed = await links.insertFromInbox(second, ws, relay);

    const drawn = presentation.slides[1]!.shapes;
    expect(drawn).toHaveLength(1);
    expect(drawn[0]!.type).toBe("Group");
    expect(placed.note).toBeUndefined();
  });

  it("fails the insert when the picture batch stops answering too", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(COLUMN, PNG);
    // The draw's second chunk, then the cleanup's two syncs, then the one the
    // picture fallback spends: a host that has stopped answering altogether.
    helpers.hangNextSync(SECOND_CHUNK_SYNC);
    helpers.hangNextSync(SECOND_CHUNK_SYNC + 1);
    helpers.hangNextSync(SECOND_CHUNK_SYNC + 2);
    vi.useFakeTimers();

    // Settling with a rejection is still settling: the pane's guard releases
    // the busy flag and the row says the host stopped answering.
    await expect(
      settle(links.insertFromInbox(item, ws, relay)),
    ).rejects.toThrow(/stopped answering/);
  });

  it("puts the picture where the group was when a refresh stops answering", async () => {
    const { item } = await insert();
    const group = shapes()[0]!;
    group.left = 80;
    group.top = 320;
    group.width = 380;
    group.height = 190;
    await pushChart(item, columnChart(7), PNG);
    const rows = await links.listLinks(relay);

    // A group leaves the batched repaint before it spends a round trip, so
    // the update's own first sync is the draw's first chunk of twelve and its
    // second is the chunk of eleven that the host swallows.
    helpers.hangNextSync(1);
    vi.useFakeTimers();
    const summary = await settle(links.updateLinks(rows, relay));

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(shapes()).toHaveLength(1);
    const picture = shapes()[0]!;
    expect(picture.type).not.toBe("Group");
    expect(picture.fillImage).toBe(PNG);
    expect(box(picture)).toEqual({
      left: 80,
      top: 320,
      width: 380,
      height: 190,
    });
    expect(JSON.parse(picture.tags.get(TAG_LINK)!).rev).toBe(2);
  });
});

describe("the host's shape budget at its edge", () => {
  // Two series over five categories with no title: ten bars, ten value
  // labels, a baseline, five category labels and a two-item legend of four -
  // thirty shapes exactly, the web's whole budget.
  function twoSeries(points: number): ChartData {
    const categories = Array.from({ length: points }, (_, i) =>
      String(2021 + i),
    );
    return {
      v: 1,
      kind: "column",
      title: null,
      categories,
      series: [0, 1].map((j) => {
        const values = categories.map((_, i) => 100 + 10 * i + j);
        return {
          name: `S${String(j)}`,
          values,
          labels: values.map(String),
          colors: values.map(() => "#B27E54"),
        };
      }),
      font: "Aptos Narrow",
      ink: "#282623",
      titleColor: "#14213D",
    };
  }

  it("draws a chart that spends the budget exactly and refuses the one past it", async () => {
    helpers.setPlatform("OfficeOnline");
    const { placed } = await insert(twoSeries(5));
    expect(placed.note).toBeUndefined();
    expect(shapes()[0]!.type).toBe("Group");
    expect(shapes()[0]!.group!.shapes).toHaveLength(30);

    helpers.selectSlide(presentation.slides[1]!.id);
    const over = await insert(twoSeries(6));
    expect(over.placed.note).toBe(
      "as a picture: 35 shapes is over this host's budget of 30",
    );
    expect(presentation.slides[1]!.shapes[0]!.type).not.toBe("Group");
  });
});

describe("a host without the shape APIs", () => {
  it("inserts the picture and says shape charts need a newer PowerPoint", async () => {
    helpers.setSupported((_set, version) => version !== "1.8");
    const { item, placed } = await insert();

    expect(placed.note).toBe(
      "as a picture: shape charts need PowerPoint 2504/16.96 or newer",
    );
    expect(shapes()).toHaveLength(1);
    // Below 1.8 there is no fill.setImage: the picture goes in through the
    // selection API and the tags are written to it afterwards.
    const inserted = helpers.insertedViaSelection();
    expect(inserted).toHaveLength(1);
    expect(inserted[0]!.png).toBe(PNG);
    expect(shapes()[0]!.tags.get(TAG_KEY)).toBe(item.token);
  });

  it("repaints that picture in place on the next update", async () => {
    helpers.setSupported((_set, version) => version !== "1.8");
    const { item } = await insert();
    const first = shapes()[0]!;
    const geometry = box(first);
    await pushChart(item, columnChart(7), PNG);

    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(shapes()).toHaveLength(1);
    // Reinserted at the same box, still carrying both tags at the new rev.
    expect(box(shapes()[0]!)).toEqual(geometry);
    expect(shapes()[0]!.tags.get(TAG_KEY)).toBe(item.token);
    expect(JSON.parse(shapes()[0]!.tags.get(TAG_LINK)!).rev).toBe(2);
  });

  it("repaints a pie that stayed a picture below 1.10 in place", async () => {
    helpers.setSupported((_set, version) => version !== "1.10");
    const pie: ChartData = {
      ...COLUMN,
      kind: "pie",
      categories: ["North", "South", "West"],
      series: [
        {
          name: "Mix",
          values: [1, 1, 2],
          labels: ["1", "1", "2"],
          colors: ["#14213D", "#2EC4B6", "#B27E54"],
        },
      ],
    };
    const { item, placed } = await insert(pie);
    expect(placed.note).toMatch(/pie shapes need PowerPoint/);
    const first = shapes()[0]!;
    expect(first.type).not.toBe("Group");
    await pushChart(item, pie, PNG);

    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    // The same shape, repainted where it stands: a picture stays a picture.
    expect(shapes()).toHaveLength(1);
    expect(shapes()[0]!.id).toBe(first.id);
    expect(shapes()[0]!.setImageCalls).toBe(2);
  });
});

// A chart the layout has to squeeze, and one whose last bar hangs below the
// baseline: both used to reach the host with a negative width or height,
// which PowerPoint refuses outright (the fake's size guard throws the same
// InvalidArgument), so the insert failed and left the drawn chunks behind.
describe("shapes the host would refuse", () => {
  // 200 x 100 pixels are 150 x 75 points, which chartSize lifts to its own
  // minimum of 200 x 120: the smallest box a chart is ever drawn in.
  const SMALL_PNG = fakePng(200, 100);

  function seriesOf(values: number[], name: string) {
    return {
      name,
      values,
      labels: values.map(String),
      colors: values.map(() => "#B27E54"),
    };
  }

  it("draws a six-series chart in the smallest box a chart gets", async () => {
    const data: ChartData = {
      ...COLUMN,
      title: "Revenue by segment",
      categories: ["2024A", "2025E", "2026E"],
      series: [0, 1, 2, 3, 4, 5].map((j) =>
        seriesOf([10 + j, 20 + j, 30 + j], `Segment number ${String(j)}`),
      ),
    };
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(data, SMALL_PNG);
    const placed = await links.insertFromInbox(item, ws, relay);

    expect(placed.note).toBeUndefined();
    const group = shapes()[0]!;
    expect(group.type).toBe("Group");
    const children = group.group!.shapes;
    expect(children.every((one) => one.width >= 0 && one.height >= 0)).toBe(
      true,
    );
    // The group is the union of its children, so its box staying inside the
    // 200 x 120 the placement reserved (centred at 380, 210) is the proof
    // that nothing was laid out past it.
    expect(group.left).toBeGreaterThanOrEqual(380);
    expect(group.top).toBeGreaterThanOrEqual(210);
    expect(group.left + group.width).toBeLessThanOrEqual(580.01);
    expect(group.top + group.height).toBeLessThanOrEqual(330.01);
  });

  it("draws a waterfall whose closing total is below zero", async () => {
    const data: ChartData = {
      ...COLUMN,
      kind: "waterfall",
      title: "Bridge",
      categories: ["Open", "Cost", "Close"],
      series: [seriesOf([100, -150, -50], "Bridge")],
    };
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(data, PNG);
    const placed = await links.insertFromInbox(item, ws, relay);

    expect(placed.note).toBeUndefined();
    const group = shapes()[0]!;
    expect(group.type).toBe("Group");
    expect(box(group)).toEqual(BOX);
    expect(
      group.group!.shapes.every((one) => one.width >= 0 && one.height >= 0),
    ).toBe(true);
  });
});

// A chart taller than the slide: the picture route fits one to the slide
// before it places it (fitToSlide), and a drawn chart has to do the same or
// the group hangs off the top and bottom of the slide.
describe("a chart taller than the slide", () => {
  it("fits the group to the slide, as the picture would have been", async () => {
    const ws = await createWorkspace(memoryStore());
    // 400 x 1400 pixels are 300 x 1050 points on a 540 pt slide.
    const item = await seedChart(COLUMN, fakePng(400, 1400));
    const placed = await links.insertFromInbox(item, ws, relay);

    expect(placed.note).toBeUndefined();
    const group = shapes()[0]!;
    expect(group.type).toBe("Group");
    // Inside the slide, margin and all, with the chart's own aspect kept.
    expect(group.top).toBeGreaterThanOrEqual(36);
    expect(group.top + group.height).toBeLessThanOrEqual(504);
    expect(group.height / group.width).toBeCloseTo(3.5, 1);
  });
});

// The teardown a failed or swallowed draw runs: it has to survive the two
// things a real slide does to it - being handed nothing to delete, and being
// handed a shape the user removed while the draw was failing.
describe("cleaning up after a draw that never finished", () => {
  it("spends no round trip when the first sync was the one that failed", async () => {
    const before = helpers.syncCount();
    await cleanupShapes(presentation.slides[0]!.id, []);
    expect(helpers.syncCount()).toBe(before);
  });

  it("skips a shape that is already gone and deletes the rest", async () => {
    const slide = presentation.slides[0]!;
    const kept = presentation.addShape(slide, { left: 10, top: 10, width: 20 });
    const drawn = presentation.addShape(slide, {
      left: 40,
      top: 10,
      width: 20,
    });
    await cleanupShapes(slide.id, [drawn.id, "shape-does-not-exist"]);

    expect(slide.shapes.map((one) => one.id)).toEqual([kept.id]);
  });
});

// The other end of the cleanup rule: a redraw whose very first chunk is
// refused has nothing to take back, and must leave the chart the deck
// already has exactly where it is.
describe("a refresh the host refuses outright", () => {
  it("leaves the old group standing and reports the row", async () => {
    const { item } = await insert();
    const before = shapes()[0]!;
    await pushChart(item, columnChart(7), PNG);
    const rows = await links.listLinks(relay);

    // The update's own first sync is the draw's first chunk; nothing of the
    // new group has reached the host when it is refused.
    helpers.failNextSync(new Error("the host refused the shapes"));
    const summary = await links.updateLinks(rows, relay);

    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expect(summary.failures[0]).toMatch(/refused the shapes/);
    expect(shapes()).toHaveLength(1);
    expect(shapes()[0]!.id).toBe(before.id);
    expect(shapes()[0]!.group!.shapes).toHaveLength(20);
  });
});
