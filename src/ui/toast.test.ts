// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import { createToast } from "./toast";

describe("toast", () => {
  it("shows the message, then hides", () => {
    vi.useFakeTimers();
    const root = document.createElement("div");
    const toast = createToast(root, 100);
    toast.show("Saved");
    expect(root.className).toContain("visible");
    expect(root.textContent).toContain("Saved");
    vi.advanceTimersByTime(101);
    expect(root.className).toBe("toast");
  });
  it("offers a copy button only when details exist", async () => {
    const root = document.createElement("div");
    const write = vi.fn(() => Promise.resolve());
    Object.assign(navigator, { clipboard: { writeText: write } });
    createToast(root, 100).show("Failed", "error", "stack...");
    const button = root.querySelector("button")!;
    button.click();
    expect(write).toHaveBeenCalledWith("stack...");
    createToast(root, 100).show("Plain");
    expect(root.querySelector("button")).toBeNull();
  });
});
