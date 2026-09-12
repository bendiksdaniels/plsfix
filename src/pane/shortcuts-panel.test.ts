// @vitest-environment jsdom
// The shortcut manager's section: the rows it renders off the shipped
// public/shortcuts.json, and the keys it reads back from Office whenever the
// Brand tab becomes visible. Apply is next door in shortcuts-panel.apply.test.ts.
import { beforeEach, describe, expect, it, vi } from "vitest";

import {
  box,
  installOffice,
  installPanel,
  makeActions,
  note,
  paneRoot,
  refusal,
  rowCount,
  serveNothing,
  serveShortcuts,
  showBrandTab,
  UNSUPPORTED,
  type ShortcutMap,
} from "../../test/shortcuts.support";

beforeEach(() => {
  vi.resetModules();
  paneRoot();
  serveShortcuts();
});

describe("the shortcut manager: what it renders", () => {
  it("shows one row per shortcut, its default key and an empty box", async () => {
    installOffice(makeActions());
    await installPanel();

    expect(rowCount()).toBe(42);
    expect(box("PLSFIX_AUTOCOLOR").value).toBe("");
    expect(box("PLSFIX_AUTOCOLOR").disabled).toBe(false);
    expect(box("PLSFIX_AUTOCOLOR").placeholder).toBe("Ctrl+Shift+K");
    expect(note()).toContain("Apply");
  });

  // The docs map Cmd to Ctrl on Windows and Alt to Option on Mac, never Ctrl to
  // Cmd on Mac, so the pane prints what Office takes and relabels nothing.
  it("prints every key exactly as Office takes it", async () => {
    installOffice(makeActions());
    await installPanel();

    const list = document.getElementById("shortcuts-list")?.textContent ?? "";
    expect(list).toContain("Ctrl+Alt+R");
    expect(list).toContain("Ctrl+Shift+Alt+8");
    expect(list).not.toContain("Cmd");
    expect(list).not.toContain("Option");
  });

  it("says what a host without the requirement set can do, and disables the boxes", async () => {
    installOffice(makeActions(), { supported: false });
    await installPanel();

    expect(rowCount()).toBe(42);
    expect(note()).toBe(UNSUPPORTED);
    expect(box("PLSFIX_AUTOCOLOR").disabled).toBe(true);
  });

  it("renders and refuses politely with no Office host at all", async () => {
    Object.assign(globalThis, { Office: undefined });
    const { applyShortcuts, resetShortcuts } = await installPanel();

    expect(rowCount()).toBe(42);
    expect(note()).toBe(UNSUPPORTED);
    expect(await refusal(applyShortcuts)).toBe(UNSUPPORTED);
    expect(await refusal(resetShortcuts)).toBe(UNSUPPORTED);
  });

  it("says so, and refuses, when the list cannot be fetched", async () => {
    installOffice(makeActions());
    serveNothing();
    const { applyShortcuts } = await installPanel();

    expect(rowCount()).toBe(0);
    expect(note()).toContain("could not be read");
    expect(await refusal(applyShortcuts)).toContain("could not be read");
  });

  it("runs the caller's own printable-card handler", async () => {
    installOffice(makeActions());
    const onCard = vi.fn();
    await installPanel(onCard);

    document
      .getElementById("shortcuts-card")
      ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(onCard).toHaveBeenCalledTimes(1);
  });
});

describe("the shortcut manager: reading the user's own keys back", () => {
  it("fills a box only where the host's key differs from the shipped one", async () => {
    const actions = makeActions();
    actions.getShortcuts.mockResolvedValue({
      PLSFIX_AUTOCOLOR: "ctrl+shift+j",
      PLSFIX_FILLRIGHT: "Ctrl+Alt+R",
      PLSFIX_FILLDOWN: null,
    });
    installOffice(actions);
    await installPanel();

    // Canonical, not the raw string the host happened to store.
    expect(box("PLSFIX_AUTOCOLOR").value).toBe("Ctrl+Shift+J");
    // The shipped default and an overridden action are both no customisation.
    expect(box("PLSFIX_FILLRIGHT").value).toBe("");
    expect(box("PLSFIX_FILLDOWN").value).toBe("");
  });

  // src/ui/tabs.ts shows a panel the same way for a click, an arrow key and
  // tabs.activate() from the tool search: the panel may only watch for that.
  it("reads them again whenever the Brand tab becomes visible, however it was reached", async () => {
    const actions = makeActions();
    installOffice(actions);
    await installPanel();
    expect(actions.getShortcuts).toHaveBeenCalledTimes(1);

    actions.getShortcuts.mockResolvedValue({ PLSFIX_UNDO: "Ctrl+Alt+Z" });
    showBrandTab();
    await vi.waitFor(() => {
      expect(box("PLSFIX_UNDO").value).toBe("Ctrl+Alt+Z");
    });
  });

  it("does not overwrite a box typed but not yet applied", async () => {
    const actions = makeActions();
    installOffice(actions);
    await installPanel();

    box("PLSFIX_UNDO").value = "Ctrl+Alt+Q";
    actions.getShortcuts.mockResolvedValue({ PLSFIX_AUTOCOLOR: "Ctrl+Alt+Z" });
    showBrandTab(false);
    showBrandTab();
    await vi.waitFor(() => {
      expect(box("PLSFIX_AUTOCOLOR").value).toBe("Ctrl+Alt+Z");
    });
    expect(box("PLSFIX_UNDO").value).toBe("Ctrl+Alt+Q");
  });

  it("leaves the boxes alone when the host refuses to hand its keys over", async () => {
    const actions = makeActions();
    actions.getShortcuts.mockRejectedValue(new Error("no"));
    installOffice(actions);
    await installPanel();

    expect(box("PLSFIX_UNDO").value).toBe("");
    expect(note()).toContain("Apply");
  });
});

describe("the shortcut manager: Reset all", () => {
  it("sends null for every action and empties every box", async () => {
    const actions = makeActions();
    actions.getShortcuts.mockResolvedValue({ PLSFIX_UNDO: "Ctrl+Alt+Z" });
    installOffice(actions);
    const { resetShortcuts } = await installPanel();
    expect(box("PLSFIX_UNDO").value).toBe("Ctrl+Alt+Z");

    expect(await resetShortcuts()).toBe(
      "Shortcuts are back to the pls,fix defaults.",
    );
    const sent = actions.replaceShortcuts.mock.calls[0]![0] as ShortcutMap;
    expect(Object.keys(sent).length).toBe(42);
    expect(Object.values(sent).every((key) => key === null)).toBe(true);
    expect(box("PLSFIX_UNDO").value).toBe("");
    expect(actions.areShortcutsInUse).not.toHaveBeenCalled();
  });

  // Reset is the one intent that does not need the host's current keys: the
  // whole point is that everything goes back to the shipped defaults.
  it("still resets when the host never handed its keys over", async () => {
    const actions = makeActions();
    actions.getShortcuts.mockRejectedValue(new Error("no"));
    installOffice(actions);
    const { resetShortcuts } = await installPanel();

    await resetShortcuts();
    expect(actions.replaceShortcuts).toHaveBeenCalledTimes(1);
  });
});
