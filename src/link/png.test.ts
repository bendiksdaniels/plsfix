import { describe, expect, it } from "vitest";
import { base64ToBytes, pngSize } from "./png";
import { fakePng } from "../../test/fakepng";

// 1x1 transparent PNG
const ONE_BY_ONE =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

describe("pngSize", () => {
  it("reads IHDR", () => {
    expect(pngSize(base64ToBytes(ONE_BY_ONE))).toEqual({ width: 1, height: 1 });
  });
  it("accepts a data: prefix", () => {
    expect(
      pngSize(base64ToBytes(`data:image/png;base64,${ONE_BY_ONE}`)),
    ).toEqual({ width: 1, height: 1 });
  });
  it("rejects non-PNG bytes", () => {
    expect(() => pngSize(new Uint8Array([1, 2, 3]))).toThrow(/PNG/);
  });
  it("reads a header-only fake PNG from the shared test double", () => {
    expect(pngSize(base64ToBytes(fakePng(800, 400)))).toEqual({
      width: 800,
      height: 400,
    });
  });
});
