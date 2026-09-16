// Unit tests for undo.ts's non-Office.js surface: settableProperties, the pure
// sanitiser undoLastAction runs a captured cell through before it reaches
// setCellProperties, and the pending-entry bookkeeping (commitUndo/discardUndo)
// that keeps a write the host refuses from spending a real Undo slot.
// Fixtures for settableProperties shape their input the way Excel for Mac's
// getCellProperties actually answers (null fill pattern, empty patternColor,
// @odata.type on every nested object), proven on Excel for Mac 16.107 (16.09).

import { beforeEach, describe, expect, it, vi } from "vitest";
import { settableProperties } from "./undo";
import type * as UndoModule from "./undo";

// A fully-populated capture (as requestFormats asks for it) with an unfilled
// cell's fill, so a test only has to override the one branch it cares about.
function macCell(
  overrides: {
    fill?: Record<string, unknown>;
    font?: Record<string, unknown>;
    borders?: Record<string, unknown>;
    horizontalAlignment?: string;
    verticalAlignment?: string;
    wrapText?: boolean;
    indentLevel?: number;
  } = {},
): Excel.CellProperties {
  return {
    format: {
      "@odata.type": "#Microsoft.Graph.Excel.CellPropertiesFormat",
      fill: {
        "@odata.type": "#Microsoft.Graph.Excel.CellPropertiesFill",
        pattern: null,
        patternColor: "",
        color: "#FFFFFF",
      },
      font: {
        bold: false,
        color: "#000000",
        italic: false,
        name: "Calibri",
        size: 11,
        underline: "None",
      },
      borders: {
        top: { color: "#000000", style: "None", weight: "Thin" },
        bottom: { color: "#000000", style: "None", weight: "Thin" },
      },
      horizontalAlignment: "General",
      verticalAlignment: "Bottom",
      wrapText: false,
      indentLevel: 0,
      ...overrides,
    },
  } as unknown as Excel.CellProperties;
}

describe("settableProperties: fill", () => {
  it("turns a null pattern into a plain None fill, dropping the moot colour", () => {
    const result = settableProperties(macCell());
    expect(result.format?.fill).toEqual({ pattern: "None" });
  });

  it("treats an undefined or an already-tidy None pattern the same way", () => {
    for (const pattern of [undefined, "None"]) {
      const result = settableProperties(
        macCell({
          fill: { pattern, patternColor: "#FFFFFF", color: "#FFFFFF" },
        }),
      );
      expect(result.format?.fill).toEqual({ pattern: "None" });
    }
  });

  it("keeps a real pattern's colour, dropping only an empty pattern colour", () => {
    const result = settableProperties(
      macCell({
        fill: { pattern: "Solid", color: "#FF0000", patternColor: "" },
      }),
    );
    expect(result.format?.fill).toEqual({ pattern: "Solid", color: "#FF0000" });
  });

  it("keeps a striped pattern's own pattern colour when Excel set one", () => {
    const result = settableProperties(
      macCell({
        fill: {
          pattern: "LightUp",
          color: "#0057B8",
          patternColor: "#EEDDCC",
        },
      }),
    );
    expect(result.format?.fill).toEqual({
      pattern: "LightUp",
      color: "#0057B8",
      patternColor: "#EEDDCC",
    });
  });

  it("drops an empty colour on a real pattern the same way as an empty pattern colour", () => {
    const result = settableProperties(
      macCell({
        fill: { pattern: "Gray50", color: "", patternColor: "#EEDDCC" },
      }),
    );
    expect(result.format?.fill).toEqual({
      pattern: "Gray50",
      patternColor: "#EEDDCC",
    });
  });
});

describe("settableProperties: font", () => {
  it("keeps a captured false or zero instead of treating it as unset", () => {
    const result = settableProperties(
      macCell({
        font: {
          bold: false,
          color: "#000000",
          italic: false,
          name: "Calibri",
          size: 0,
          underline: "None",
        },
      }),
    );
    expect(result.format?.font).toEqual({
      bold: false,
      color: "#000000",
      italic: false,
      name: "Calibri",
      size: 0,
      underline: "None",
    });
  });

  it("skips a font key Excel answered as undefined or null", () => {
    const result = settableProperties(
      macCell({ font: { bold: true, color: undefined, name: null } }),
    );
    expect(result.format?.font).toEqual({ bold: true });
  });

  it("drops an @odata.type annotation on the fill and the font", () => {
    const result = settableProperties(
      macCell({
        fill: {
          "@odata.type": "#Microsoft.Graph.Excel.CellPropertiesFill",
          pattern: "Solid",
          color: "#FF0000",
        },
        font: {
          "@odata.type": "#Microsoft.Graph.Excel.CellPropertiesFont",
          bold: true,
          color: "#000000",
          italic: false,
          name: "Calibri",
          size: 11,
          underline: "None",
        },
      }),
    );
    expect(result.format?.fill).toEqual({ pattern: "Solid", color: "#FF0000" });
    expect(result.format?.font).toEqual({
      bold: true,
      color: "#000000",
      italic: false,
      name: "Calibri",
      size: 11,
      underline: "None",
    });
  });
});

describe("settableProperties: borders", () => {
  it("keys borders by whichever edges the capture carries", () => {
    const result = settableProperties(
      macCell({
        borders: {
          top: { color: "#B27E54", style: "Continuous", weight: "Medium" },
          bottom: { color: "#000000", style: "None", weight: "Thin" },
        },
      }),
    );
    expect(result.format?.borders).toEqual({
      top: { color: "#B27E54", style: "Continuous", weight: "Medium" },
      bottom: { color: "#000000", style: "None", weight: "Thin" },
    });
    expect(Object.keys(result.format?.borders ?? {})).toEqual([
      "top",
      "bottom",
    ]);
  });

  it("drops an empty field on one border edge without dropping the edge", () => {
    const result = settableProperties(
      macCell({
        borders: {
          left: { color: "", style: "None", weight: "Thin" },
        },
      }),
    );
    expect(result.format?.borders).toEqual({
      left: { style: "None", weight: "Thin" },
    });
  });
});

describe("settableProperties: the rest of the format", () => {
  it("passes alignment, wrap and indent through as captured", () => {
    const result = settableProperties(
      macCell({
        horizontalAlignment: "Right",
        verticalAlignment: "Center",
        wrapText: true,
        indentLevel: 2,
      }),
    );
    expect(result.format).toMatchObject({
      horizontalAlignment: "Right",
      verticalAlignment: "Center",
      wrapText: true,
      indentLevel: 2,
    });
  });

  it("returns an empty settable object for a cell with no captured format", () => {
    expect(settableProperties({} as Excel.CellProperties)).toEqual({});
  });
});

// A minimal Office.js double: every property a capture reads is already on
// the object (no deferred load/sync semantics to fake), which is all
// captureUndoAreas's own bookkeeping needs - the shape of what it reads is
// covered elsewhere (the integration suites over test/fakehost.ts).
function fakeContext(): Excel.RequestContext {
  return {
    sync: () => Promise.resolve(),
  } as unknown as Excel.RequestContext;
}

function fakeRange(address: string): Excel.Range {
  return {
    load: () => undefined,
    address,
    rowCount: 1,
    columnCount: 1,
    formulas: [["1"]],
    numberFormat: [["General"]],
    worksheet: { load: () => undefined, id: address },
    getCellProperties: () => ({ value: [[{}]] }),
  } as unknown as Excel.Range;
}

describe("captureUndoAreas: the pending entry a refused write must not keep", () => {
  // A fresh module instance per test: undoStack and pendingUndo are private
  // module state, the same way every stress rig resets ../src/excel.
  let undo: typeof UndoModule;

  beforeEach(async () => {
    vi.resetModules();
    undo = await import("./undo");
  });

  it("commit keeps the entry", async () => {
    await undo.captureUndoAreas(fakeContext(), [fakeRange("Model!A1")]);
    undo.commitUndo();

    // Committed, so no longer pending: a later discard must leave it alone.
    undo.discardUndo();
    expect(undo.undoTarget()).toBe("Model!A1");
  });

  it("discard removes only the pending top entry and leaves older ones", async () => {
    await undo.captureUndoAreas(fakeContext(), [fakeRange("Model!A1")]);
    undo.commitUndo();
    await undo.captureUndoAreas(fakeContext(), [fakeRange("Model!B1")]);

    undo.discardUndo();
    expect(undo.undoTarget()).toBe("Model!A1");
  });

  it("a new capture drops a stale pending entry", async () => {
    await undo.captureUndoAreas(fakeContext(), [fakeRange("Model!A1")]);
    // A1 is never committed or discarded: still pending when B1 is captured.
    await undo.captureUndoAreas(fakeContext(), [fakeRange("Model!B1")]);

    // B1 is also still pending. Discarding it should leave nothing - proving
    // A1 was already dropped, not sitting underneath it.
    undo.discardUndo();
    expect(undo.undoTarget()).toBeNull();
  });

  it("discard with nothing pending is a no-op", () => {
    expect(() => undo.discardUndo()).not.toThrow();
    expect(undo.undoTarget()).toBeNull();
  });
});
