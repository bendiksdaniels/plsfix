import { describe, expect, it } from "vitest";
import type { LinkStatus } from "../link/status";
import {
  pruneSelection,
  requireSelection,
  rowKey,
  selectedRows,
  slideRows,
  toRowViews,
  updateDetails,
} from "./actions";
import type { LinkRow, UpdateSummary } from "./links";

function row(
  slideId: string,
  shapeId: string,
  status: LinkStatus = "current",
  slideIndex = 0,
): LinkRow {
  return {
    found: {
      slideId,
      slideIndex,
      shapeId,
      tag: {
        v: 1,
        id: "a".repeat(32),
        kind: "range",
        rev: 2,
        src: {
          workbook: "Model_v4.xlsx",
          sheet: "Model",
          ref: "B4:F12",
          anchor: "SMT_LINK_aaaaaaaa",
        },
        pushedAt: "2026-08-28T10:00:00.000Z",
      },
      token: "t",
      left: 10,
      top: 10,
      width: 400,
      height: 200,
    },
    status,
    relayRev: 2,
    pushedAt: 1000,
  };
}

const emptySummary: UpdateSummary = {
  updated: 0,
  current: 0,
  missing: 0,
  wrongKey: 0,
  failed: 0,
  sourceChanges: [],
  failures: [],
};

describe("selection", () => {
  it("keys a row by slide and shape, not by link id", () => {
    expect(rowKey(row("s1", "sh1").found)).toBe("s1/sh1");
    expect(rowKey(row("s2", "sh1").found)).toBe("s2/sh1");
  });

  it("draws the ticked rows as selected and numbers slides from one", () => {
    const rows = [row("s1", "sh1"), row("s2", "sh2", "missing", 1)];
    const views = toRowViews(rows, new Set(["s2/sh2"]));
    expect(views.map((view) => view.selected)).toEqual([false, true]);
    expect(views.map((view) => view.slide)).toEqual([1, 2]);
    expect(views[1]!.label).toBe("Model!B4:F12");
    expect(views[1]!.source).toBe("Model_v4.xlsx");
    expect(views[1]!.status).toBe("missing");
  });

  it("takes the slide from the first ticked row, and none when nothing is", () => {
    const rows = [row("s1", "sh1"), row("s2", "sh2"), row("s2", "sh3")];
    expect(
      selectedRows(rows, new Set(["s2/sh3"])).map((r) => r.found.shapeId),
    ).toEqual(["sh3"]);
    expect(
      slideRows(rows, new Set(["s2/sh3"])).map((r) => r.found.shapeId),
    ).toEqual(["sh2", "sh3"]);
    expect(slideRows(rows, new Set())).toEqual([]);
  });

  it("drops keys whose shape is gone after a rescan", () => {
    const selected = new Set(["s1/sh1", "s9/gone"]);
    pruneSelection([row("s1", "sh1")], selected);
    expect([...selected]).toEqual(["s1/sh1"]);
  });

  it("refuses an action with nothing ticked", () => {
    expect(() => requireSelection([])).toThrow("Tick a link");
    expect(requireSelection([row("s1", "sh1")])).toHaveLength(1);
  });
});

describe("updateDetails", () => {
  it("has no details when nothing went wrong or moved", () => {
    expect(updateDetails(emptySummary)).toBeUndefined();
  });

  it("lists failures first, then every source change", () => {
    expect(
      updateDetails({
        ...emptySummary,
        failures: ["Model!B4:F12: relay GET /api/links/x: 500"],
        sourceChanges: ["Model_v4.xlsx -> Model_v5.xlsx"],
      }),
    ).toBe(
      "Model!B4:F12: relay GET /api/links/x: 500\n" +
        "Source changed: Model_v4.xlsx -> Model_v5.xlsx",
    );
  });
});
