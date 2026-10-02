// Direct unit coverage of summarize's own line, which the integration suites
// only exercise indirectly through a real updateLinks summary.
import { describe, expect, it } from "vitest";
import { summarize } from "./links";

describe("summarize", () => {
  it("counts links this computer holds no copy of", () => {
    expect(
      summarize({
        updated: 1,
        current: 0,
        notPasted: 2,
        missing: 0,
        wrongKey: 0,
        failed: 0,
        sourceChanges: [],
        failures: [],
        notes: [],
      }),
    ).toBe("1 updated, 2 not pasted yet");
  });
});
