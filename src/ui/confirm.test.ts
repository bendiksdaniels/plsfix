// @vitest-environment jsdom
// The shared two-click confirm every destructive or one-way button in both
// panes runs through (styles-delete and delete-names first, slice W's eight
// after them). Exercised against bare buttons built by hand here; each real
// button's own wiring is proven in its module's own tests and in the
// click-through suites.
import { describe, expect, it, vi } from "vitest";
import { armConfirm, CONFIRM_MS } from "./confirm";

function danger(label: string): HTMLButtonElement {
  const button = document.createElement("button");
  button.className = "danger";
  button.textContent = label;
  return button;
}

describe("armConfirm", () => {
  it("arms on the first press without running, runs on the second", () => {
    const button = danger("Remove link");
    const run = vi.fn();
    const confirm = armConfirm(button, run);

    confirm.handleClick();
    expect(confirm.isArmed()).toBe(true);
    expect(run).not.toHaveBeenCalled();
    expect(button.textContent).toBe("Click again to confirm");
    expect(button.classList.contains("armed")).toBe(true);

    confirm.handleClick();
    expect(run).toHaveBeenCalledTimes(1);
    expect(confirm.isArmed()).toBe(false);
    expect(button.textContent).toBe("Remove link");
    expect(button.classList.contains("armed")).toBe(false);
  });

  it("restores a frozen snapshot of the button's own markup by default", () => {
    // The rich action-list shape (clean-past-data): an icon plus a <strong>
    // title and a <small> description, none of it plain text.
    const button = document.createElement("button");
    button.innerHTML =
      '<span class="action-icon toc">⌫</span>' +
      "<span><strong>Clean past the data</strong>" +
      "<small>Drops the empty rows and columns</small></span>";
    const confirm = armConfirm(button, vi.fn());

    confirm.arm();
    // Only the <strong> changes; the icon and the description survive.
    expect(button.querySelector("strong")?.textContent).toBe(
      "Click again to confirm",
    );
    expect(button.querySelector(".action-icon")).not.toBeNull();
    expect(button.querySelector("small")?.textContent).toBe(
      "Drops the empty rows and columns",
    );

    confirm.disarm();
    expect(button.querySelector("strong")?.textContent).toBe(
      "Clean past the data",
    );
    expect(button.querySelector("small")?.textContent).toBe(
      "Drops the empty rows and columns",
    );
    expect(button.querySelector(".action-icon")).not.toBeNull();
  });

  it("reads the resting label fresh from a callback rather than a frozen snapshot", () => {
    // styles-delete's own shape: the resting text carries a live count that
    // can change between an arm and its disarm (a rescan in between).
    const button = danger("Delete 1 unused style");
    let count = 1;
    const confirm = armConfirm(button, vi.fn(), {
      label: () =>
        `Delete ${String(count)} unused ${count === 1 ? "style" : "styles"}`,
    });

    confirm.arm();
    expect(button.textContent).toBe("Click again to confirm");
    count = 3;
    confirm.disarm();
    expect(button.textContent).toBe("Delete 3 unused styles");
  });

  it("keeps an icon-only button's glyph, relabelling its title and aria-label instead", () => {
    // reset-brand's own shape: a fixed-size icon square, no room for a
    // sentence, so only the accessible name and the tooltip change.
    const button = document.createElement("button");
    button.className = "icon-button";
    button.setAttribute("aria-label", "Reset to pls,fix defaults");
    button.setAttribute("title", "Reset to defaults");
    button.textContent = "⟲";
    const confirm = armConfirm(button, vi.fn());

    confirm.arm();
    expect(button.textContent).toBe("⟲");
    expect(button.getAttribute("aria-label")).toBe("Click again to confirm");
    expect(button.getAttribute("title")).toBe("Click again to confirm");
    expect(button.classList.contains("armed")).toBe(true);

    confirm.disarm();
    expect(button.textContent).toBe("⟲");
    expect(button.getAttribute("aria-label")).toBe("Reset to pls,fix defaults");
    expect(button.getAttribute("title")).toBe("Reset to defaults");
  });

  it("lapses on its own after CONFIRM_MS and a press only re-arms", () => {
    vi.useFakeTimers();
    try {
      const button = danger("Forget key");
      const run = vi.fn();
      const confirm = armConfirm(button, run);

      confirm.handleClick();
      expect(confirm.isArmed()).toBe(true);

      vi.advanceTimersByTime(CONFIRM_MS);
      expect(confirm.isArmed()).toBe(false);
      expect(button.textContent).toBe("Forget key");
      expect(run).not.toHaveBeenCalled();

      // The window has lapsed: a press now only re-arms, it does not run.
      confirm.handleClick();
      expect(confirm.isArmed()).toBe(true);
      expect(run).not.toHaveBeenCalled();
    } finally {
      vi.useRealTimers();
    }
  });

  it("a second arm resets the lapse timer rather than stacking two", () => {
    vi.useFakeTimers();
    try {
      const button = danger("Break link");
      const run = vi.fn();
      const confirm = armConfirm(button, run);

      confirm.handleClick();
      vi.advanceTimersByTime(CONFIRM_MS - 1);
      // Disarmed and re-armed by hand (a rerender elsewhere in the pane
      // mid-window): only ONE lapse timer should be pending afterwards.
      confirm.disarm();
      confirm.arm();
      vi.advanceTimersByTime(CONFIRM_MS - 1);
      expect(confirm.isArmed()).toBe(true);
      vi.advanceTimersByTime(1);
      expect(confirm.isArmed()).toBe(false);
    } finally {
      vi.useRealTimers();
    }
  });

  it("when false skips the confirm window and runs at once", () => {
    // generate-key's own shape: a first key needs no confirming.
    const button = danger("Generate");
    const run = vi.fn();
    const confirm = armConfirm(button, run, { when: () => false });

    confirm.handleClick();
    expect(run).toHaveBeenCalledTimes(1);
    expect(confirm.isArmed()).toBe(false);
    expect(button.classList.contains("armed")).toBe(false);
  });

  it("when true still requires the two presses", () => {
    const button = danger("Generate");
    const run = vi.fn();
    const confirm = armConfirm(button, run, { when: () => true });

    confirm.handleClick();
    expect(run).not.toHaveBeenCalled();
    expect(confirm.isArmed()).toBe(true);

    confirm.handleClick();
    expect(run).toHaveBeenCalledTimes(1);
  });

  it("disarm is a safe no-op when nothing is armed", () => {
    const button = danger("Remove link");
    const confirm = armConfirm(button, vi.fn());

    expect(() => confirm.disarm()).not.toThrow();
    expect(confirm.isArmed()).toBe(false);
    expect(button.textContent).toBe("Remove link");
  });

  // The busy latch disables every button while an action runs (setBusy in
  // src/pane/shared.ts and src/ppt/main.ts); disabling is a DOM attribute
  // this module never reads, so an armed button that a caller disarms while
  // disabled (as a rerender elsewhere in the pane would) comes back clean
  // with no stale "Click again to confirm" label once it is re-enabled.
  it("disarms cleanly while the button is disabled", () => {
    const button = danger("Reset all");
    const confirm = armConfirm(button, vi.fn());

    confirm.arm();
    button.disabled = true;
    confirm.disarm();
    button.disabled = false;

    expect(button.textContent).toBe("Reset all");
    expect(button.classList.contains("armed")).toBe(false);
    expect(confirm.isArmed()).toBe(false);
  });
});
