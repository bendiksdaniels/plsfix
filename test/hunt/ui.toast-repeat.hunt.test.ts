// @vitest-environment jsdom
// A toast shown twice in a row (named in the brief, not in toast.test.ts):
// show() unconditionally clears the pending hide timer and starts a fresh
// one every call, so a repeat - the same sentence, or a different one -
// must extend the visible window rather than hiding on the FIRST call's
// original schedule, and must never leave two timers racing to hide it.

import { describe, expect, it, vi } from "vitest";
import { createToast } from "../../src/ui/toast";

describe("a toast shown twice in a row", () => {
  it("the identical message again resets the hide timer, not stacks it", () => {
    vi.useFakeTimers();
    const root = document.createElement("div");
    const toast = createToast(root, 100);

    toast.show("Saved");
    vi.advanceTimersByTime(50);
    toast.show("Saved"); // repeat, well inside the first window

    // The first call's own 100 ms deadline (t=100) must NOT hide it: the
    // second call reset the clock to its own t=150.
    vi.advanceTimersByTime(50);
    expect(root.className).toContain("visible");
    expect(root.textContent).toContain("Saved");

    vi.advanceTimersByTime(50);
    expect(root.className).toBe("toast");
    vi.useRealTimers();
  });

  it("a different message right after the first also resets the timer", () => {
    vi.useFakeTimers();
    const root = document.createElement("div");
    const toast = createToast(root, 100);

    toast.show("Saved");
    vi.advanceTimersByTime(90);
    toast.show("Undone"); // a different sentence, just before the first hides

    vi.advanceTimersByTime(90);
    expect(root.className).toContain("visible");
    expect(root.textContent).toContain("Undone");

    vi.advanceTimersByTime(10);
    expect(root.className).toBe("toast");
  });

  it("three in a row leave exactly one visible text node, not a pile-up", () => {
    const root = document.createElement("div");
    const toast = createToast(root, 100);

    toast.show("One");
    toast.show("Two");
    toast.show("Three");

    expect(root.querySelectorAll(".toast-text")).toHaveLength(1);
    expect(root.textContent).toContain("Three");
    expect(root.textContent).not.toContain("One");
    expect(root.textContent).not.toContain("Two");
  });
});
