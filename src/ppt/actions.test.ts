import { describe, expect, it } from "vitest";
import type { LinkStatus } from "../link/status";
import {
  filterRows,
  linkSlides,
  linkSources,
  pruneSelection,
  requireSelection,
  rowKey,
  scopedSummary,
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
          anchor: "PLSFIX_LINK_aaaaaaaa",
        },
        pushedAt: "2026-08-28T10:00:00.000Z",
      },
      token: "t",
      type: "GeometricShape",
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
  notPasted: 0,
  failed: 0,
  sourceChanges: [],
  failures: [],
  notes: [],
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
    expect(views[1]!.kind).toBe("range");
    expect(views[1]!.status).toBe("missing");
  });

  it("selects rows by shape key from the ticked set", () => {
    const rows = [row("s1", "sh1"), row("s2", "sh2"), row("s2", "sh3")];
    expect(
      selectedRows(rows, new Set(["s2/sh3"])).map((r) => r.found.shapeId),
    ).toEqual(["sh3"]);
  });

  it("returns every row on the given slide, ticked or not, and none for a slide with no links", () => {
    const rows = [row("s1", "sh1"), row("s2", "sh2"), row("s2", "sh3")];
    expect(slideRows(rows, "s2").map((r) => r.found.shapeId)).toEqual([
      "sh2",
      "sh3",
    ]);
    expect(slideRows(rows, "s1").map((r) => r.found.shapeId)).toEqual(["sh1"]);
    expect(slideRows(rows, "s9")).toEqual([]);
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

describe("link filters", () => {
  it("filters by status without changing the row identity", () => {
    const rows = [
      row("s1", "sh1", "current", 0),
      row("s2", "sh2", "missing", 1),
      row("s3", "sh3", "updateAvailable", 2),
    ];
    expect(
      filterRows(rows, { query: "", status: "missing" }).map((item) =>
        rowKey(item.found),
      ),
    ).toEqual(["s2/sh2"]);
  });

  it("searches visible source, object type, status and slide number", () => {
    const rows = [
      row("s1", "sh1", "current", 0),
      row("s2", "sh2", "missing", 4),
    ];
    expect(filterRows(rows, { query: "model_v4", status: "all" })).toHaveLength(
      2,
    );
    expect(filterRows(rows, { query: "range", status: "all" })).toHaveLength(2);
    expect(filterRows(rows, { query: "missing", status: "all" })).toEqual([
      rows[1],
    ]);
    expect(filterRows(rows, { query: "5", status: "all" })).toEqual([rows[1]]);
  });

  it("combines workbook, slide and status views", () => {
    const rows = [
      row("s1", "sh1", "current", 0),
      row("s2", "sh2", "missing", 4),
    ];
    rows[1]!.found.tag.src.workbook = "Budget.xlsx";
    expect(
      filterRows(rows, {
        query: "",
        status: "missing",
        source: "Budget.xlsx",
        slide: 5,
      }),
    ).toEqual([rows[1]]);
    expect(linkSources(rows)).toEqual(["Budget.xlsx", "Model_v4.xlsx"]);
    expect(linkSlides(rows)).toEqual([1, 5]);
  });

  it("filters by project, treating a missing name as no project", () => {
    const rows = [
      row("s1", "sh1", "current", 0),
      row("s2", "sh2", "current", 1),
    ];
    rows[1]!.found.tag.project = "Amasty";
    expect(
      filterRows(rows, { query: "", status: "all", project: "Amasty" }),
    ).toEqual([rows[1]]);
    expect(filterRows(rows, { query: "", status: "all", project: "" })).toEqual(
      [rows[0]],
    );
  });
});

describe("scopedSummary", () => {
  it("leaves All projects unnamed", () => {
    expect(scopedSummary("3 updated", "all")).toBe("3 updated");
    expect(scopedSummary("3 updated")).toBe("3 updated");
  });

  it("names the shown project, and No project when the filter is empty", () => {
    expect(scopedSummary("3 updated", "Amasty")).toBe("3 updated (Amasty)");
    expect(scopedSummary("No links found", "")).toBe(
      "No links found (No project)",
    );
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

// A chart that arrived as a picture, on insert or on update, says why under
// the counts, after the failures and the source changes.
describe("updateDetails: charts that arrived as pictures", () => {
  it("lists the notes last", () => {
    expect(
      updateDetails({
        ...emptySummary,
        sourceChanges: ["Model_v4.xlsx -> Model_v5.xlsx"],
        notes: [
          "Model!Revenue chart as a picture: 7 series; shapes draw up to 6",
        ],
      }),
    ).toBe(
      "Source changed: Model_v4.xlsx -> Model_v5.xlsx\n" +
        "Model!Revenue chart as a picture: 7 series; shapes draw up to 6",
    );
  });
});
