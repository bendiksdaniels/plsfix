// @vitest-environment jsdom
// The shortcut manager's Apply: the exact map it hands Office.actions, what it
// refuses before writing, and the guard that stops an unread panel sending 42
// nulls over a user's keys. Rendering and Reset all are in shortcuts-panel.test.ts.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { describeError } from "../ui/report";
import {
  box,
  installOffice,
  installPanel,
  makeActions,
  paneRoot,
  refusal,
  rejects,
  serveShortcuts,
  showBrandTab,
  type ActionsStub,
  type ShortcutMap,
} from "../../test/shortcuts.support";

const APPLIED = "Shortcuts updated for your account.";
const UNSUPPORTED =
  "Custom shortcuts need Microsoft 365 with a signed-in account.";

function sentMap(actions: ActionsStub): ShortcutMap {
  return actions.replaceShortcuts.mock.calls[0]![0] as ShortcutMap;
}

beforeEach(() => {
  vi.resetModules();
  paneRoot();
  serveShortcuts();
});

afterEach(() => {
  vi.useRealTimers();
});

describe("Apply: the map it sends", () => {
  it("sends null for every untouched action and the combination for the rest", async () => {
    const actions = makeActions();
    installOffice(actions);
    const { applyShortcuts } = await installPanel();

    box("PLSFIX_AUTOCOLOR").value = "alt+ctrl+j";
    expect(await applyShortcuts()).toBe(APPLIED);

    expect(actions.areShortcutsInUse).toHaveBeenCalledWith(["Ctrl+Alt+J"]);
    const sent = sentMap(actions);
    expect(Object.keys(sent).length).toBe(42);
    expect(sent.PLSFIX_AUTOCOLOR).toBe("Ctrl+Alt+J");
    expect(sent.PLSFIX_UNDO).toBeNull();
    // The conflict question comes before the write, never after.
    expect(actions.areShortcutsInUse.mock.invocationCallOrder[0]!).toBeLessThan(
      actions.replaceShortcuts.mock.invocationCallOrder[0]!,
    );
  });

  it("asks nothing when every box is empty and still sends the reverting map", async () => {
    const actions = makeActions();
    installOffice(actions);
    const { applyShortcuts } = await installPanel();

    await applyShortcuts();
    expect(actions.areShortcutsInUse).not.toHaveBeenCalled();
    expect(Object.values(sentMap(actions)).every((key) => key === null)).toBe(
      true,
    );
  });

  it("names the combinations Office says are already in use", async () => {
    const actions = makeActions();
    actions.areShortcutsInUse.mockResolvedValue([
      { shortcut: "Ctrl+Alt+J", inUse: true },
    ]);
    installOffice(actions);
    const { applyShortcuts } = await installPanel();

    box("PLSFIX_AUTOCOLOR").value = "Ctrl+Alt+J";
    expect(await applyShortcuts()).toBe(
      `${APPLIED} Already used elsewhere: Ctrl+Alt+J.`,
    );
    expect(actions.replaceShortcuts).toHaveBeenCalledTimes(1);
  });

  it("goes ahead when the host will not answer the conflict question", async () => {
    const actions = makeActions();
    actions.areShortcutsInUse.mockRejectedValue(new Error("no"));
    installOffice(actions);
    const { applyShortcuts } = await installPanel();

    box("PLSFIX_AUTOCOLOR").value = "Ctrl+Alt+J";
    expect(await applyShortcuts()).toBe(APPLIED);
    expect(actions.replaceShortcuts).toHaveBeenCalledTimes(1);
  });

  it("shows the canonical keys back after a write, typed shorthand and all", async () => {
    const actions = makeActions();
    installOffice(actions);
    const { applyShortcuts } = await installPanel();

    box("PLSFIX_AUTOCOLOR").value = "alt+ctrl+j";
    actions.getShortcuts.mockResolvedValue({ PLSFIX_AUTOCOLOR: "Ctrl+Alt+J" });
    await applyShortcuts();
    expect(box("PLSFIX_AUTOCOLOR").value).toBe("Ctrl+Alt+J");
  });
});

describe("Apply: what it refuses before writing", () => {
  it("refuses a box that is not a key combination, by name", async () => {
    const actions = makeActions();
    installOffice(actions);
    const { applyShortcuts } = await installPanel();

    box("PLSFIX_AUTOCOLOR").value = "Shift+J";
    expect(await refusal(applyShortcuts)).toBe(
      "Autocolor selection: Use Ctrl, Alt or Shift plus one key, for example Ctrl+Shift+A.",
    );
    expect(actions.replaceShortcuts).not.toHaveBeenCalled();
    expect(document.activeElement).toBe(box("PLSFIX_AUTOCOLOR"));
  });

  it("refuses one combination typed into two boxes, naming both actions", async () => {
    const actions = makeActions();
    installOffice(actions);
    const { applyShortcuts } = await installPanel();

    box("PLSFIX_AUTOCOLOR").value = "Ctrl+Alt+J";
    box("PLSFIX_UNDO").value = "ctrl+alt+j";
    expect(await refusal(applyShortcuts)).toBe(
      "Autocolor selection and Undo last pls,fix action would both use Ctrl+Alt+J.",
    );
    expect(actions.replaceShortcuts).not.toHaveBeenCalled();
  });

  // The clash Office cannot report: both keys are ours, so areShortcutsInUse
  // says false while two pls,fix actions end up on one combination.
  it("refuses a custom key that lands on another action's shipped default", async () => {
    const actions = makeActions();
    installOffice(actions);
    const { applyShortcuts } = await installPanel();

    box("PLSFIX_AUTOCOLOR").value = "Ctrl+Shift+M";
    expect(await refusal(applyShortcuts)).toBe(
      "Open pls,fix and Autocolor selection would both use Ctrl+Shift+M.",
    );
    expect(actions.replaceShortcuts).not.toHaveBeenCalled();
  });
});

// The panel's boxes are only the user's intent once they hold what Office
// currently has. Until then an Apply would read 42 empty boxes as "revert
// everything" and wipe a whole custom key map behind a success toast.
describe("Apply: never wipes a key map it has not read", () => {
  async function unreadPanel(): Promise<{
    actions: ActionsStub;
    applyShortcuts: () => Promise<string>;
  }> {
    const actions = makeActions();
    actions.getShortcuts.mockRejectedValueOnce(new Error("not yet"));
    installOffice(actions);
    const { applyShortcuts } = await installPanel();
    expect(box("PLSFIX_AUTOCOLOR").value).toBe("");
    return { actions, applyShortcuts };
  }

  it("reads the keys first when the tab was never opened (tool search runs Apply)", async () => {
    const { actions, applyShortcuts } = await unreadPanel();
    actions.getShortcuts.mockResolvedValue({ PLSFIX_AUTOCOLOR: "Ctrl+Alt+Z" });

    expect(await applyShortcuts()).toBe(APPLIED);
    expect(sentMap(actions).PLSFIX_AUTOCOLOR).toBe("Ctrl+Alt+Z");
    expect(box("PLSFIX_AUTOCOLOR").value).toBe("Ctrl+Alt+Z");
  });

  it("keeps the keys a tab opened by keyboard has already filled in", async () => {
    const { actions, applyShortcuts } = await unreadPanel();
    actions.getShortcuts.mockResolvedValue({ PLSFIX_AUTOCOLOR: "Ctrl+Alt+Z" });
    showBrandTab();
    await vi.waitFor(() => {
      expect(box("PLSFIX_AUTOCOLOR").value).toBe("Ctrl+Alt+Z");
    });

    await applyShortcuts();
    expect(sentMap(actions).PLSFIX_AUTOCOLOR).toBe("Ctrl+Alt+Z");
  });

  it("refuses outright while the keys still cannot be read", async () => {
    const { actions, applyShortcuts } = await unreadPanel();
    actions.getShortcuts.mockRejectedValue(new Error("still no"));

    expect(await refusal(applyShortcuts)).toBe(
      "Your current shortcuts could not be read, so nothing was changed. Reopen the Brand tab and try again.",
    );
    expect(actions.replaceShortcuts).not.toHaveBeenCalled();
  });
});

describe("Apply: what it says when Office refuses", () => {
  it("answers the signed-in sentence and keeps Office's own code for the report", async () => {
    const actions = makeActions();
    actions.replaceShortcuts.mockRejectedValue(
      Object.assign(new Error("InvalidOperation"), {
        code: "InvalidOperation",
        debugInfo: { errorLocation: "Actions.replaceShortcuts" },
      }),
    );
    installOffice(actions);
    const { applyShortcuts } = await installPanel();

    box("PLSFIX_AUTOCOLOR").value = "Ctrl+Alt+J";
    const error = await rejects(applyShortcuts);
    expect(error.message).toBe(UNSUPPORTED);

    // src/ui/report.ts reads code and debugInfo off the error it is handed and
    // never walks .cause, so both have to travel on the thrown error itself.
    const { details } = describeError(error, {
      host: "Excel",
      version: "v0.0.000",
    });
    expect(details).toContain("code: InvalidOperation");
    expect(details).toContain("Actions.replaceShortcuts");
  });

  it("says the write is slow, not that the user is signed out, when it times out", async () => {
    const actions = makeActions();
    actions.replaceShortcuts.mockImplementation(
      () => new Promise(() => undefined),
    );
    installOffice(actions);
    const { applyShortcuts } = await installPanel();
    box("PLSFIX_AUTOCOLOR").value = "Ctrl+Alt+J";

    vi.useFakeTimers();
    const pending = rejects(applyShortcuts);
    await vi.advanceTimersByTimeAsync(20_000);
    const error = await pending;
    expect(error.message).toContain("Office did not answer");
    expect(error.message).toContain("may still have been saved");
    expect(error.message).not.toBe(UNSUPPORTED);
  });
});
