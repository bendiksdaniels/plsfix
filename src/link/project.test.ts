// Project names on a link: cap, trim, drop control characters, and treat an
// empty result as "no project". The codecs live in model.ts; this file is the
// name rule.

import { describe, expect, it } from "vitest";
import { cleanProject, NO_PROJECT, PROJECT_MAX, projectLabel } from "./project";

describe("cleanProject", () => {
  it("trims and keeps a short name", () => {
    expect(cleanProject("  Amasty  ")).toBe("Amasty");
  });

  it("caps at 40 characters", () => {
    const name = "x".repeat(PROJECT_MAX + 8);
    expect(cleanProject(name)).toBe("x".repeat(PROJECT_MAX));
  });

  it("drops control characters", () => {
    expect(cleanProject("Ama\nsty")).toBe("Amasty");
    expect(cleanProject("Ama\u0000sty")).toBe("Amasty");
  });

  it("treats a blank or controls-only name as missing", () => {
    expect(cleanProject("")).toBeNull();
    expect(cleanProject("   ")).toBeNull();
    expect(cleanProject("\n\t")).toBeNull();
  });
});

describe("projectLabel", () => {
  it("says No project when the name is missing", () => {
    expect(projectLabel(undefined)).toBe(NO_PROJECT);
    expect(projectLabel("Amasty")).toBe("Amasty");
  });
});
