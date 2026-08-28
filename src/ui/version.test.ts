import { describe, expect, it } from "vitest";
import { formatVersion } from "./version";

describe("formatVersion", () => {
  it("shows a three-digit patch", () => {
    expect(formatVersion("1.1.0")).toBe("v1.1.000");
    expect(formatVersion("2.0.12")).toBe("v2.0.012");
  });
  it("rejects anything that is not MAJOR.MINOR.PATCH", () => {
    expect(() => formatVersion("2.0")).toThrow(/semver/);
  });
});
