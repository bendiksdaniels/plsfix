// Stress pass on the deck moving under the pane's feet: a row whose shape the
// modeller deleted or ungrouped since the list was drawn, a tag edited by
// hand, a deck with no slides and a slide with three hundred shapes.
// Invariant: every one of them ends in a sentence naming the link, never an
// office.js code, and the deck is left as the modeller had it.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { TAG_KEY, TAG_LINK, type TableCell } from "../src/link/model";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import { fakePng } from "./fakepng";
import type { FakeRelay } from "./fakerelay";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import {
  bootPpt,
  columnChart,
  memoryStore,
  pushAgain,
  pushChart,
  pushTable,
  pushText,
  seedChart,
  seedLink,
  seedTable,
  seedText,
} from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";
import type * as HostModule from "../src/ppt/host";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let host: typeof HostModule;
let presentation: FakePresentation;
let relay: FakeRelay;
let ws: Workspace;

beforeEach(async () => {
  ({ links, presentation, relay } = await bootPpt());
  host = await import("../src/ppt/host");
  ws = await createWorkspace(memoryStore());
});
afterEach(() => {
  uninstallFakePpt();
});

const CELLS: TableCell[][] = [[{ t: "Revenue" }, { t: "1 000" }]];

// Every kind of link the deck can hold, seeded and then pushed again so an
// update has something to repaint.
const KINDS = [
  {
    name: "a picture",
    async seed(): Promise<void> {
      const item = await seedLink(fakePng(200, 100));
      await links.insertFromInbox(item, ws, relay);
      await pushAgain(item, fakePng(200, 100));
    },
  },
  {
    name: "a text box",
    async seed(): Promise<void> {
      const item = await seedText("1 000");
      await links.insertFromInbox(item, ws, relay);
      await pushText(item, "1 200");
    },
  },
  {
    name: "a table",
    async seed(): Promise<void> {
      const item = await seedTable(CELLS, [80, 60]);
      await links.insertFromInbox(item, ws, relay);
      await pushTable(item, [[{ t: "Revenue" }, { t: "1 200" }]], [80, 60]);
    },
  },
  {
    name: "a chart group",
    async seed(): Promise<void> {
      const item = await seedChart(columnChart(3), fakePng(600, 300));
      await links.insertFromInbox(item, ws, relay);
      await pushChart(item, columnChart(4), fakePng(600, 300));
    },
  },
] as const;

function shapes(slide = 0): FakePptShape[] {
  return presentation.slides[slide]!.shapes;
}

// No office.js code, no shape id, and the link the row named.
function expectSentence(line: string | undefined, label: string): void {
  expect(line).toBeDefined();
  expect(line).toContain(label);
  expect(line).not.toMatch(/ItemNotFound|GeneralException|InvalidArgument/);
  expect(line).not.toMatch(/shape-\d+/);
}

describe("a row whose object the modeller deleted", () => {
  it.each(KINDS)(
    "reports $name in words instead of the host's code",
    async ({ seed }) => {
      await seed();
      const rows = await links.listLinks(relay);
      presentation.deleteShape(rows[0]!.found.shapeId);

      const summary = await links.updateLinks(rows, relay);

      expect(summary).toMatchObject({ updated: 0, failed: 1 });
      expectSentence(summary.failures[0], "Model");
      // The rescan the pane runs after every action drops the row.
      expect(await links.listLinks(relay)).toHaveLength(0);
    },
  );

  it("says so when Break is pressed on it", async () => {
    const item = await seedLink(fakePng(200, 100));
    await links.insertFromInbox(item, ws, relay);
    const rows = await links.listLinks(relay);
    presentation.deleteShape(rows[0]!.found.shapeId);

    let message = "";
    await host.breakLink(rows[0]!.found).catch((error: unknown) => {
      message = error instanceof Error ? error.message : String(error);
    });

    expectSentence(message, "Model!B4:F12");
  });

  it("says so on a revert, and the rows beside it still revert", async () => {
    const revert = await import("../src/ppt/revert");
    const gone = await seedLink(fakePng(200, 100));
    const kept = await seedLink(fakePng(200, 100));
    for (const item of [gone, kept]) {
      await links.insertFromInbox(item, ws, relay);
      await pushAgain(item, fakePng(240, 120));
    }
    await links.updateLinks(await links.listLinks(relay), relay);
    const rows = await links.listLinks(relay);
    presentation.deleteShape(rows[0]!.found.shapeId);

    const summary = await revert.revertLinks(rows, relay);

    expect(summary).toMatchObject({ reverted: 1, failed: 1 });
    expectSentence(summary.failures[0], "Model!B4:F12");
  });

  it("names the one row the user pulled out of its group, not the deck", async () => {
    const item = await seedLink(fakePng(200, 100));
    await links.insertFromInbox(item, ws, relay);
    const picture = shapes()[0]!;
    const caption = presentation.addShape(presentation.slides[0]!, {
      left: 5,
      top: 5,
      width: 60,
    });
    const group = presentation.groupShapes(
      [picture.id, caption.id],
      presentation.slides[0]!.id,
    );
    const stale = await links.listLinks(relay);
    await pushAgain(item, fakePng(200, 100));
    // Ctrl+Shift+G, after the list was drawn.
    const slide = presentation.slides[0]!;
    slide.shapes.splice(slide.shapes.indexOf(group), 1, ...group.group!.shapes);

    const summary = await links.updateLinks(stale, relay);

    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expectSentence(summary.failures[0], "Model!B4:F12");
    // The fresh scan finds it again, and the next press updates it.
    const fresh = await links.listLinks(relay);
    expect(await links.updateLinks(fresh, relay)).toMatchObject({ updated: 1 });
  });
});

describe("a tag edited by hand", () => {
  it.each([
    ["broken JSON", (value: string) => `${value.slice(0, 12)}`],
    [
      "a revision that is not a number",
      (v: string) => v.replace(/"rev":\d+/, '"rev":"3"'),
    ],
    ["another schema", () => JSON.stringify({ v: 2, id: "x" })],
  ] as [string, (value: string) => string][])(
    "leaves a shape carrying %s out of the list, and touches nothing",
    async (_name, edit) => {
      const item = await seedLink(fakePng(200, 100));
      await links.insertFromInbox(item, ws, relay);
      const shape = shapes()[0]!;
      shape.tags.set(TAG_LINK, edit(shape.tags.get(TAG_LINK)!));
      await pushAgain(item, fakePng(200, 100));

      const rows = await links.listLinks(relay);

      expect(rows).toHaveLength(0);
      expect(links.summarize(await links.updateLinks(rows, relay))).toBe(
        "No links found",
      );
      expect(shape.setImageCalls).toBe(1);
    },
  );

  it("leaves a shape whose key tag is gone out of the list", async () => {
    const item = await seedLink(fakePng(200, 100));
    await links.insertFromInbox(item, ws, relay);
    shapes()[0]!.tags.delete(TAG_KEY);

    expect(await links.listLinks(relay)).toHaveLength(0);
  });

  it("heals a revision typed far ahead of the relay", async () => {
    const item = await seedLink(fakePng(200, 100));
    await links.insertFromInbox(item, ws, relay);
    const shape = shapes()[0]!;
    shape.tags.set(
      TAG_LINK,
      shape.tags.get(TAG_LINK)!.replace(/"rev":\d+/, '"rev":99'),
    );

    const rows = await links.listLinks(relay);
    expect(rows).toMatchObject([{ status: "updateAvailable" }]);

    expect(await links.updateLinks(rows, relay)).toMatchObject({ updated: 1 });
    expect(await links.listLinks(relay)).toMatchObject([{ status: "current" }]);
  });
});

describe("a deck with nowhere to put anything", () => {
  it("asks for a slide instead of throwing when the deck has none", async () => {
    uninstallFakePpt();
    installFakePpt({ slides: 0 });
    const item = await seedLink(fakePng(200, 100));

    let message = "";
    await links.insertFromInbox(item, ws, relay).catch((error: unknown) => {
      message = error instanceof Error ? error.message : String(error);
    });

    expect(message).toBe("insert Model!B4:F12: select a slide first.");
  });

  it("lands on a slide already holding three hundred shapes, and says it overlaps", async () => {
    const slide = presentation.slides[0]!;
    for (let index = 0; index < 300; index += 1) {
      presentation.addShape(slide, {
        left: (index % 20) * 48,
        top: Math.floor(index / 20) * 36,
        width: 46,
        height: 34,
      });
    }
    const item = await seedLink(fakePng(400, 200));

    const placed = await links.insertFromInbox(item, ws, relay);

    expect(placed.overlapping).toBe(true);
    expect(links.insertNote(placed)).toContain("no free space on this slide");
    expect(shapes()).toHaveLength(301);
    expect(await links.listLinks(relay)).toHaveLength(1);
  });

  it("scans an empty deck and updates nothing", async () => {
    uninstallFakePpt();
    installFakePpt({ slides: 0 });

    const rows = await links.listLinks(relay);

    expect(rows).toEqual([]);
    expect(links.summarize(await links.updateLinks(rows, relay))).toBe(
      "No links found",
    );
  });
});

describe("the same export inserted twice in a row", () => {
  it("tracks both copies and updates them together", async () => {
    const item = await seedLink(fakePng(200, 100));
    await links.insertFromInbox(item, ws, relay);
    await links.insertFromInbox(item, ws, relay);
    await pushAgain(item, fakePng(200, 100));

    const rows = await links.listLinks(relay);

    expect(rows).toHaveLength(2);
    expect(await links.updateLinks(rows, relay)).toMatchObject({ updated: 2 });
    expect(shapes().every((shape) => shape.setImageCalls === 2)).toBe(true);
  });
});

describe("a deck the modeller rearranged", () => {
  it("repaints a link three groups deep where it sits", async () => {
    const item = await seedLink(fakePng(200, 100));
    await links.insertFromInbox(item, ws, relay);
    const slide = presentation.slides[0]!;
    let inner = shapes()[0]!.id;
    for (let level = 0; level < 3; level += 1) {
      const neighbour = presentation.addShape(slide, {
        left: 400 + level * 20,
        top: 10,
        width: 20,
        height: 20,
      });
      inner = presentation.groupShapes([inner, neighbour.id], slide.id).id;
    }
    await pushAgain(item, fakePng(240, 120));

    const rows = await links.listLinks(relay);

    expect(rows[0]!.found.groupPath).toHaveLength(3);
    expect(await links.updateLinks(rows, relay)).toMatchObject({ updated: 1 });
    // Still inside the group the user built, repainted in place.
    expect(shapes()).toHaveLength(1);
    expect(shapes()[0]!.type).toBe("Group");
  });

  it("inserts an inbox item whose kind does not match its payload as the payload", async () => {
    // A relay row nobody in this deck wrote: the announcement says table, the
    // sealed payload is a picture. The shape follows the payload, so the deck
    // never ends up with a rectangle something later reads as a table.
    const item = {
      ...(await seedLink(fakePng(200, 100))),
      kind: "table" as const,
    };

    const placed = await links.insertFromInbox(item, ws, relay);

    const shape = shapes().find((one) => one.id === placed.shapeId)!;
    expect(shape.type).toBe("GeometricShape");
    expect(shape.table).toBeNull();
    expect(await links.listLinks(relay)).toHaveLength(1);
  });
});

// Placeholder so vi is used when a suite variant drops its fake timers.
afterEach(() => {
  vi.useRealTimers();
});
