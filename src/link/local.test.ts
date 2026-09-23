// src/link/local.test.ts
// The local revision space and the fixed workspace every device derives.
import { describe, expect, it } from "vitest";
import { isLocalRev, LOCAL_REV_BASE, localWorkspace, nextLocalRev, previousRevOf } from "./local";

describe("local revisions", () => {
  it("start above the relay's space and count up from there", () => {
    expect(nextLocalRev(0)).toBe(LOCAL_REV_BASE + 1);
    expect(nextLocalRev(7)).toBe(LOCAL_REV_BASE + 1);
    expect(nextLocalRev(LOCAL_REV_BASE + 4)).toBe(LOCAL_REV_BASE + 5);
  });

  it("tell the two spaces apart", () => {
    expect(isLocalRev(12)).toBe(false);
    expect(isLocalRev(LOCAL_REV_BASE + 1)).toBe(true);
  });

  it("give Revert the rev below, never under the first of its space", () => {
    expect(previousRevOf(5)).toBe(4);
    expect(previousRevOf(1)).toBeNull();
    expect(previousRevOf(LOCAL_REV_BASE + 3)).toBe(LOCAL_REV_BASE + 2);
    expect(previousRevOf(LOCAL_REV_BASE + 1)).toBeNull();
  });
});

describe("localWorkspace", () => {
  it("is the same on every call and every device", async () => {
    const a = await localWorkspace();
    const b = await localWorkspace();
    expect(a.id).toBe(b.id);
    expect(a.id).toMatch(/^[A-Za-z0-9_-]+$/);
  });
});
