// Unit test for serialised(), the queue insertFootballField and insertTornado
// share: no Excel host needed, since the queue itself is plain promise
// chaining over whatever work function each caller hands it.

import { describe, expect, it } from "vitest";
import { serialised } from "./chart-blocks";

describe("serialised", () => {
  it("runs a second call only after the first settles", async () => {
    const order: string[] = [];
    let releaseFirst: () => void = () => undefined;
    const first = serialised(async () => {
      order.push("first-start");
      await new Promise<void>((resolve) => {
        releaseFirst = resolve;
      });
      order.push("first-end");
    });
    const second = serialised(async () => {
      order.push("second-start");
    });

    await Promise.resolve();
    await Promise.resolve();
    // The second call is queued behind the first: it has not run yet.
    expect(order).toEqual(["first-start"]);

    releaseFirst();
    await first;
    await second;

    expect(order).toEqual(["first-start", "first-end", "second-start"]);
  });

  it("keeps the queue running after a call rejects", async () => {
    const order: string[] = [];
    const failing = serialised(async () => {
      order.push("failing");
      throw new Error("boom");
    });
    const next = serialised(async () => {
      order.push("next");
    });

    await expect(failing).rejects.toThrow("boom");
    await next;

    expect(order).toEqual(["failing", "next"]);
  });
});
