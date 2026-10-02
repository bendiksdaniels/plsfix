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

  it("reports 'Details copied.' once the copy actually lands", async () => {
    const root = document.createElement("div");
    const write = vi.fn(() => Promise.resolve());
    Object.assign(navigator, { clipboard: { writeText: write } });
    createToast(root, 100).show("Failed", "error", "stack...");

    root.querySelector("button")!.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(root.textContent).toContain("Details copied.");
    expect(root.className).toContain("success");
  });

  it("reports the copy failed when neither clipboard path works", async () => {
    const root = document.createElement("div");
    Object.assign(navigator, { clipboard: undefined });
    Object.assign(document, { execCommand: vi.fn(() => false) });
    createToast(root, 100).show("Failed", "error", "stack...");

    root.querySelector("button")!.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(root.textContent).toContain(
      "Copy failed: select the text and copy it by hand.",
    );
    expect(root.className).toContain("error");
  });
});
