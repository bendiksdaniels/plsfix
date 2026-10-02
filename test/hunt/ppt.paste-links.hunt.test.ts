// Hunt pass 1: src/ppt/paste-links.ts's pure functions - requireBundle's two
// messages (a foreign copy vs. a newer add-in's bundle) and summarizePaste's
// remaining branches (an empty no-op paste, and a paste that did nothing but
// fail). No fakes: both functions take plain data in, a string or a throw out.

import { describe, expect, it } from "vitest";
import { requireBundle, summarizePaste } from "../../src/ppt/paste-links";

describe("requireBundle", () => {
  it("names a newer add-in's bundle apart from any other bad paste", () => {
    expect(() => requireBundle({ ok: false, reason: "newerVersion" })).toThrow(
      "This copy comes from a newer pls,fix. Update the add-in.",
    );
  });

  it("says a foreign, malformed, truncated or plain-text-only paste is not a pls,fix copy", () => {
    expect(() => requireBundle({ ok: false, reason: "notBundle" })).toThrow(
      "That is not a pls,fix copy from Excel.",
    );
  });

  it("hands back the bundle unwrapped when the paste was good", () => {
    const bundle = { links: [], inbox: [] };
    expect(requireBundle({ ok: true, bundle })).toBe(bundle);
  });
});

describe("summarizePaste: the remaining full-stop branches", () => {
  it("says just 'Pasted.' for a bundle that carried nothing and did nothing", () => {
    expect(
      summarizePaste({
        links: 0,
        updated: 0,
        current: 0,
        waiting: 0,
        failed: 0,
        failures: [],
        notes: [],
      }),
    ).toBe("Pasted.");
  });

  it("says 'Pasted N link(s).' with no colon when a link came in but nothing happened", () => {
    expect(
      summarizePaste({
        links: 1,
        updated: 0,
        current: 0,
        waiting: 0,
        failed: 0,
        failures: [],
        notes: [],
      }),
    ).toBe("Pasted 1 link.");
    expect(
      summarizePaste({
        links: 2,
        updated: 0,
        current: 0,
        waiting: 0,
        failed: 0,
        failures: [],
        notes: [],
      }),
    ).toBe("Pasted 2 links.");
  });

  it("leads with 'failed' alone, no stray leading comma, when nothing else happened", () => {
    expect(
      summarizePaste({
        links: 1,
        updated: 0,
        current: 0,
        waiting: 0,
        failed: 1,
        failures: ["x"],
        notes: [],
      }),
    ).toBe("Pasted 1 link: 1 failed.");
  });
});
