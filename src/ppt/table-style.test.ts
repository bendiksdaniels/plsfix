// Unit coverage for queueTableStyle against a hand-built stub, not the
// fakeppt harness: the one place the exact order of its seven styleSettings
// writes is pinned. Owns nothing beyond that; the fake-host behaviour this
// order exists for lives in test/ppt.table-style.integration.test.ts.

import { afterEach, describe, expect, it } from "vitest";
import { queueTableStyle } from "./table-style";

const STYLE_PROPERTIES = [
  "style",
  "isFirstRowHighlighted",
  "areRowsBanded",
  "areColumnsBanded",
  "isFirstColumnHighlighted",
  "isLastRowHighlighted",
  "isLastColumnHighlighted",
] as const;

// A table whose styleSettings only records what is written to it, in order:
// queueTableStyle never reads styleSettings back, so there is nothing for a
// stub to answer reads with.
function stubTable(): {
  table: PowerPoint.Table;
  writes: [string, unknown][];
} {
  const writes: [string, unknown][] = [];
  const styleSettings = {} as PowerPoint.TableStyleSettings;
  for (const property of STYLE_PROPERTIES) {
    Object.defineProperty(styleSettings, property, {
      set: (value: unknown) => writes.push([property, value]),
      enumerable: true,
      configurable: true,
    });
  }
  return { table: { styleSettings } as PowerPoint.Table, writes };
}

// hasTableStyle() reads Office.context.requirements.isSetSupported: the only
// two globals queueTableStyle ever touches, minimal enough to keep this file
// independent of the fakeppt harness.
function setPowerPointApi(supported: boolean): void {
  const scope = globalThis as unknown as Record<string, unknown>;
  scope.PowerPoint = {
    TableStyle: { mediumStyle2Accent1: "MediumStyle2Accent1" },
  };
  scope.Office = {
    context: { requirements: { isSetSupported: () => supported } },
  };
}

afterEach(() => {
  const scope = globalThis as unknown as Record<string, unknown>;
  delete scope.PowerPoint;
  delete scope.Office;
});

describe("queueTableStyle", () => {
  it("writes the default style, the header flag and a plain unbanded body, in order", () => {
    setPowerPointApi(true);
    const { table, writes } = stubTable();

    queueTableStyle(table, true);

    expect(writes.map(([property]) => property)).toEqual(STYLE_PROPERTIES);
    expect(writes[0]).toEqual(["style", "MediumStyle2Accent1"]);
    expect(writes[1]).toEqual(["isFirstRowHighlighted", true]);
    expect(writes.slice(2).map(([, value]) => value)).toEqual([
      false,
      false,
      false,
      false,
      false,
    ]);
  });

  it("clears the header flag when the payload has none", () => {
    setPowerPointApi(true);
    const { table, writes } = stubTable();

    queueTableStyle(table, false);

    expect(writes[1]).toEqual(["isFirstRowHighlighted", false]);
  });

  it("writes nothing on a host below PowerPointApi 1.9", () => {
    setPowerPointApi(false);
    const { table, writes } = stubTable();

    queueTableStyle(table, true);

    expect(writes).toEqual([]);
  });
});
