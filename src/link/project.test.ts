// Project names on a link: cap, trim, drop control characters, and treat an
// empty result as "no project". The codecs live in model.ts; this file is the
// name rule.

import { describe, expect, it } from "vitest";
import {
  cleanProject,
  groupByProject,
  NO_PROJECT,
  PROJECT_MAX,
  projectLabel,
} from "./project";

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

describe("groupByProject", () => {
  interface Item {
    id: string;
    project?: string;
  }
  const byProject = (item: Item): string | undefined => item.project;

  it("orders named groups by locale and puts No project last", () => {
    const items: Item[] = [
      { id: "1", project: "Balcia" },
      { id: "2", project: undefined },
      { id: "3", project: "Amasty" },
    ];
    const names = groupByProject(items, byProject).map(([name]) => name);
    expect(names).toEqual(["Amasty", "Balcia", NO_PROJECT]);
  });

  it("keeps items in their input order inside a group", () => {
    const items: Item[] = [
      { id: "1", project: "Amasty" },
      { id: "2", project: "Amasty" },
      { id: "3", project: "Amasty" },
    ];
    expect(groupByProject(items, byProject)).toEqual([["Amasty", items]]);
  });

  it("puts one group per distinct name", () => {
    const items: Item[] = [
      { id: "1", project: "Amasty" },
      { id: "2", project: "Balcia" },
      { id: "3", project: "Amasty" },
    ];
    const groups = groupByProject(items, byProject);
    expect(groups.map(([name, group]) => [name, group.length])).toEqual([
      ["Amasty", 2],
      ["Balcia", 1],
    ]);
  });

  it("returns an empty array for empty input", () => {
    expect(groupByProject([], byProject)).toEqual([]);
  });

  it("groups an undefined project under NO_PROJECT", () => {
    const items: Item[] = [{ id: "1", project: undefined }];
    expect(groupByProject(items, byProject)).toEqual([[NO_PROJECT, items]]);
  });
});
