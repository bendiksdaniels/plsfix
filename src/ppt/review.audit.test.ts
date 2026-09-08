// What the slice F review added: a hidden source column beside visible ones
// keeps a place in the deck's table, and a relay that cannot be reached fails
// every row of an update at once instead of asking the relay once per row.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { RelayError } from "../link/relay";
import { createWorkspace } from "../link/workspace";
import type { FakeRelay } from "../../test/fakerelay";
import { fakePng } from "../../test/fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
} from "../../test/fakeppt";
import { bootPpt, memoryStore, seedLink } from "../../test/ppt.support";
import { columnWidths, tableSize } from "./tables";
import type * as LinksModule from "./links";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, relay } = await bootPpt());
});

afterEach(() => {
  uninstallFakePpt();
});

describe("a table with a hidden column beside visible ones", () => {
  const payload = {
    v: 1 as const,
    kind: "table" as const,
    rows: 1,
    cols: 3,
    cells: [[{ t: "a" }, { t: "b" }, { t: "c" }]],
    // Excel reports zero for a hidden column; two of three are hidden here.
    widths: [0, 0, 80],
    src: {
      workbook: "Model.xlsx",
      sheet: "Model",
      ref: "A1:C1",
      anchor: "PLSFIX_LINK_00000000",
    },
    pushedAt: new Date().toISOString(),
    hash: "0".repeat(64),
  };

  it("gives the hidden columns Excel's default width", () => {
    expect(tableSize(payload).width).toBe(48 + 48 + 80);
    expect(columnWidths(payload, tableSize(payload).width)).toEqual([
      48, 48, 80,
    ]);
  });
});

describe("a relay that cannot be reached", () => {
  it("fails every row once, without a GET per row", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedLink(fakePng(800, 400));
    await links.insertFromInbox(item, ws, relay);
    const rows = await links.listLinks(relay);

    vi.spyOn(relay, "fetchLinks").mockRejectedValue(
      new RelayError("network", "relay unreachable"),
    );
    const getLink = vi.spyOn(relay, "getLink");

    const summary = await links.updateLinks(rows, relay);

    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expect(summary.failures[0]).toContain("relay unreachable");
    expect(getLink).not.toHaveBeenCalled();
  });
});
