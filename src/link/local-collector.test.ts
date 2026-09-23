// src/link/local-collector.test.ts
// What one Excel action records instead of sending.
import { describe, expect, it } from "vitest";
import { LOCAL_REV_BASE } from "./local";
import { LocalCollector } from "./local-collector";

const ID = "0123456789abcdef0123456789abcdef";

describe("LocalCollector", () => {
  it("answers the next local rev after the registry's", async () => {
    const collector = new LocalCollector(() => 1000);
    expect(await collector.putLink(ID, "auth", new Uint8Array([1]), 3)).toEqual(
      { rev: LOCAL_REV_BASE + 1 },
    );
    expect(
      await collector.putLink(
        ID,
        "auth",
        new Uint8Array([2]),
        LOCAL_REV_BASE + 1,
      ),
    ).toEqual({ rev: LOCAL_REV_BASE + 2 });
  });

  it("keeps the last put per link and every inbox row", async () => {
    const collector = new LocalCollector(() => 1000);
    await collector.putLink(ID, "a", new Uint8Array([1]));
    await collector.putLink(ID, "a", new Uint8Array([2]));
    await collector.postInbox("ws", "auth", ID, new Uint8Array([7]));
    expect(collector.bundle()).toEqual({
      links: [
        {
          id: ID,
          rev: LOCAL_REV_BASE + 2,
          sentAt: 1000,
          blob: new Uint8Array([2]),
        },
      ],
      inbox: [{ id: ID, createdAt: 1000, blob: new Uint8Array([7]) }],
    });
  });

  it("has nothing to revoke or touch, and refuses every read", async () => {
    const collector = new LocalCollector();
    await expect(collector.deleteLink(ID, "a")).resolves.toBeUndefined();
    await expect(collector.touchLinks([{ id: ID, auth: "a" }])).resolves.toBe(
      1,
    );
    await expect(collector.status([])).rejects.toThrow(/not used in Excel/);
    await expect(collector.listInbox("ws", "a")).rejects.toThrow(
      /not used in Excel/,
    );
    expect(collector.isEmpty()).toBe(true);
  });
});
