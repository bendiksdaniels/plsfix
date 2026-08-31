import { describe, expect, it } from "vitest";
import {
  aspectChanged,
  deriveStatus,
  fitToSlide,
  sourceChanged,
} from "./status";

describe("deriveStatus", () => {
  it("compares the relay rev with the tag rev", () => {
    expect(deriveStatus(3, { id: "a", rev: 3, pushedAt: 1 })).toBe("current");
    expect(deriveStatus(3, { id: "a", rev: 5, pushedAt: 1 })).toBe(
      "updateAvailable",
    );
    expect(deriveStatus(3, { id: "a", rev: null, pushedAt: null })).toBe(
      "missing",
    );
    expect(deriveStatus(3, undefined)).toBe("missing");
    expect(
      deriveStatus(3, { id: "a", rev: null, pushedAt: null, error: "auth" }),
    ).toBe("wrongKey");
  });

  it("treats a relay rev below the tag rev as an update", () => {
    // The store sweeps a link thirty days after its last push, and the next PUT
    // starts again at rev 1 while the deck's tag still holds the old rev. A
    // "greater than" test would call that stale deck current for ever.
    expect(deriveStatus(12, { id: "a", rev: 1, pushedAt: 1 })).toBe(
      "updateAvailable",
    );
    expect(deriveStatus(1, { id: "a", rev: 1, pushedAt: 1 })).toBe("current");
  });
});

describe("sourceChanged", () => {
  const src = {
    workbook: "Model_v4.xlsx",
    sheet: "Model",
    ref: "B4",
    anchor: "PLSFIX_LINK_1",
  };
  it("flags a different workbook only", () => {
    expect(sourceChanged(src, { ...src, workbook: "model_V4.XLSX" })).toBe(
      false,
    );
    expect(sourceChanged(src, { ...src, workbook: "Model_v5.xlsx" })).toBe(
      true,
    );
    expect(sourceChanged(src, { ...src, ref: "B5" })).toBe(false);
  });
});

describe("aspectChanged", () => {
  it("ignores sub-tolerance drift", () => {
    expect(aspectChanged(400, 200, 800, 400)).toBe(false);
    expect(aspectChanged(400, 200, 801, 400)).toBe(false);
    expect(aspectChanged(400, 200, 800, 450)).toBe(true);
  });
});

describe("fitToSlide", () => {
  it("converts 96 dpi pixels to points and centres", () => {
    expect(fitToSlide(800, 400)).toEqual({
      left: 180,
      top: 120,
      width: 600,
      height: 300,
    });
  });
  it("caps to the slide width minus margins, keeping the aspect ratio", () => {
    const box = fitToSlide(4000, 1000);
    expect(box.width).toBe(888);
    expect(box.height).toBe(222);
    expect(box.left).toBe(36);
  });
  it("caps to the height too", () => {
    const box = fitToSlide(1000, 4000);
    expect(box.height).toBe(468);
    expect(box.width).toBe(117);
  });
});
