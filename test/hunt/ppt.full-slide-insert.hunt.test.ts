// Hunt pass 1: inserting onto a slide with no free space left, for the table
// and text kinds test/ppt.insert.integration.test.ts's own case never tried.
// The scan (src/free-space.ts) is already proven; this pins the adapter
// wiring - insertFromInbox lands at the exact centred box for every kind.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { TablePayload, TextPayload } from "../../src/link/model";
import { createWorkspace, type Workspace } from "../../src/link/workspace";
import type { FakeRelay } from "../fakerelay";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePresentation,
} from "../fakeppt";
import { bootPpt, memoryStore, seedTable, seedText } from "../ppt.support";
import type * as LinksModule from "../../src/ppt/links";
import { SLIDE } from "../../src/ppt/placement";
import { tableSize } from "../../src/ppt/tables";
import { textSize } from "../../src/ppt/texts";

enableStrictLoadSemantics();

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

function coverWholeSlide(): void {
  presentation.addShape(presentation.slides[0]!, {
    left: 0,
    top: 0,
    width: 960,
    height: 540,
  });
}

// Alone on a fully covered slide the fallback in src/free-space.ts centres
// the object full size over what is there (fitHole's own hole-of-nothing
// case) - the box every "centred" title below is actually pinned to, not
// merely a claim the assertions never checked.
function centredBox(size: { width: number; height: number }): {
  left: number;
  top: number;
  width: number;
  height: number;
} {
  return {
    left: (SLIDE.width - size.width) / 2,
    top: (SLIDE.height - size.height) / 2,
    width: size.width,
    height: size.height,
  };
}

describe("a table inserted onto a slide with no free space", () => {
  it("still lands, centred over what is there, and says it overlaps", async () => {
    coverWholeSlide();
    const cells = [
      [{ t: "Q1" }, { t: "Q2" }],
      [{ t: "100" }, { t: "120" }],
    ];
    const widths = [80, 80];
    const item = await seedTable(cells, widths);

    const placed = await links.insertFromInbox(item, ws, relay);

    expect(placed.overlapping).toBe(true);
    expect(links.insertNote(placed)).toContain("no free space on this slide");
    const shapes = presentation.slides[0]!.shapes;
    expect(shapes).toHaveLength(2);
    const size = tableSize({
      v: 1,
      kind: "table",
      rows: cells.length,
      cols: widths.length,
      cells,
      widths,
      src: { workbook: "", sheet: "", ref: "", anchor: "" },
      pushedAt: "",
      hash: "",
    } satisfies TablePayload);
    const table = shapes[1]!;
    expect({
      left: table.left,
      top: table.top,
      width: table.width,
      height: table.height,
    }).toEqual(centredBox(size));
  });
});

describe("a text box inserted onto a slide with no free space", () => {
  it("still lands, centred over what is there, and says it overlaps", async () => {
    coverWholeSlide();
    const text = "EUR 15.7m";
    const item = await seedText(text);

    const placed = await links.insertFromInbox(item, ws, relay);

    expect(placed.overlapping).toBe(true);
    expect(links.insertNote(placed)).toContain("no free space on this slide");
    const shapes = presentation.slides[0]!.shapes;
    expect(shapes).toHaveLength(2);
    const size = textSize({
      v: 1,
      kind: "text",
      text,
      src: { workbook: "", sheet: "", ref: "", anchor: "" },
      pushedAt: "",
      hash: "",
    } satisfies TextPayload);
    const box = shapes[1]!;
    expect({
      left: box.left,
      top: box.top,
      width: box.width,
      height: box.height,
    }).toEqual(centredBox(size));
  });
});
