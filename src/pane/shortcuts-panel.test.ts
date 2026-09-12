// @vitest-environment jsdom
// The shortcut manager as the modeller drives it: the rows it renders off the
// shipped public/shortcuts.json, and the exact maps Apply and Reset all hand
// to Office.actions. Office is a stub and fetch is a stub, so no host is
// needed; the markup under test is taskpane.html itself.
import { readFileSync } from "node:fs";
import { join } from "node:path";

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

type ShortcutMap = Record<string, string | null>;

interface ActionsStub {
  getShortcuts: ReturnType<typeof vi.fn>;
  replaceShortcuts: ReturnType<typeof vi.fn>;
  areShortcutsInUse: ReturnType<typeof vi.fn>;
}

const SHORTCUTS_JSON = readFileSync(
  join(process.cwd(), "public/shortcuts.json"),
  "utf8",
);

const UNSUPPORTED =
  "Custom shortcuts need Microsoft 365 with a signed-in account.";

function paneRoot(): void {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
}

function makeActions(): ActionsStub {
  return {
    getShortcuts: vi.fn(async () => ({})),
    replaceShortcuts: vi.fn(async () => undefined),
    areShortcutsInUse: vi.fn(async (keys: string[]) =>
      keys.map((shortcut) => ({ shortcut, inUse: false })),
    ),
  };
}

function installOffice(
  actions: ActionsStub | null,
  options: { supported?: boolean; platform?: string } = {},
): void {
  const supported = options.supported ?? actions !== null;
  Object.assign(globalThis, {
    Office: {
      context: {
        platform: options.platform ?? "PC",
        requirements: { isSetSupported: () => supported },
      },
      actions,
    },
  });
}

function serveShortcuts(body = SHORTCUTS_JSON): void {
  Object.assign(globalThis, {
    fetch: vi.fn(async () => ({
      ok: true,
      status: 200,
      json: async () => JSON.parse(body) as unknown,
    })),
  });
}

async function install(onCard: () => void = () => undefined): Promise<{
  applyShortcuts: () => Promise<string>;
  resetShortcuts: () => Promise<string>;
}> {
  const module = await import("./shortcuts-panel");
  await module.installShortcutsPanel(document, onCard);
  return module;
}

function box(id: string): HTMLInputElement {
  const input = document.getElementById(`shortcut-${id}`);
  if (!(input instanceof HTMLInputElement)) throw new Error(`no box ${id}`);
  return input;
}

function note(): string {
  return (document.getElementById("shortcuts-note")?.textContent ?? "").trim();
}

function rowCount(): number {
  return document.querySelectorAll("#shortcuts-list label").length;
}

function openBrandTab(): void {
  document
    .getElementById("tab-brand")
    ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
}

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

beforeEach(() => {
  vi.resetModules();
  paneRoot();
  serveShortcuts();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("the shortcut manager: what it renders", () => {
  it("shows one row per shortcut, its default key and an empty box", async () => {
    installOffice(makeActions());
    await install();

    expect(rowCount()).toBe(42);
    expect(box("PLSFIX_AUTOCOLOR").value).toBe("");
    expect(box("PLSFIX_AUTOCOLOR").disabled).toBe(false);
    expect(box("PLSFIX_AUTOCOLOR").placeholder).toBe("Ctrl+Shift+K");
    expect(note()).toContain("Apply");
  });

  it("shows the Mac names on a Mac while the box still takes Ctrl", async () => {
    const actions = makeActions();
    installOffice(actions, { platform: "Mac" });
    await install();

    expect(document.getElementById("shortcuts-list")?.textContent).toContain(
      "Cmd+Option+R",
    );
    box("PLSFIX_AUTOCOLOR").value = "Cmd+Shift+J";
    const { applyShortcuts } = await import("./shortcuts-panel");
    await applyShortcuts();
    const sent = actions.replaceShortcuts.mock.calls[0]![0] as ShortcutMap;
    expect(sent.PLSFIX_AUTOCOLOR).toBe("Ctrl+Shift+J");
  });

  it("says what a host without the requirement set can do, and disables the boxes", async () => {
    installOffice(makeActions(), { supported: false });
    await install();

    expect(rowCount()).toBe(42);
    expect(note()).toBe(UNSUPPORTED);
    expect(box("PLSFIX_AUTOCOLOR").disabled).toBe(true);
  });

  it("renders and refuses politely with no Office host at all", async () => {
    Object.assign(globalThis, { Office: undefined });
    const { applyShortcuts, resetShortcuts } = await install();

    expect(rowCount()).toBe(42);
    expect(note()).toBe(UNSUPPORTED);
    expect(await rejects(applyShortcuts)).toBe(UNSUPPORTED);
    expect(await rejects(resetShortcuts)).toBe(UNSUPPORTED);
  });

  it("says so, and refuses, when the list cannot be fetched", async () => {
    installOffice(makeActions());
    Object.assign(globalThis, {
      fetch: vi.fn(async () => {
        throw new Error("offline");
      }),
    });
    const { applyShortcuts } = await install();

    expect(rowCount()).toBe(0);
    expect(note()).toContain("could not be read");
    expect(await rejects(applyShortcuts)).toContain("could not be read");
  });

  it("runs the caller's own printable-card handler", async () => {
    installOffice(makeActions());
    const onCard = vi.fn();
    await install(onCard);

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
      PLSFIX_AUTOCOLOR: "Ctrl+Shift+J",
      PLSFIX_FILLRIGHT: "Ctrl+Alt+R",
      PLSFIX_FILLDOWN: null,
    });
    installOffice(actions);
    await install();

    expect(box("PLSFIX_AUTOCOLOR").value).toBe("Ctrl+Shift+J");
    // The shipped default and an overridden action are both no customisation.
    expect(box("PLSFIX_FILLRIGHT").value).toBe("");
    expect(box("PLSFIX_FILLDOWN").value).toBe("");
  });

  it("reads them again every time the Brand tab is opened", async () => {
    const actions = makeActions();
    installOffice(actions);
    await install();
    expect(actions.getShortcuts).toHaveBeenCalledTimes(1);

    actions.getShortcuts.mockResolvedValue({ PLSFIX_UNDO: "Ctrl+Alt+Z" });
    openBrandTab();
    await vi.waitFor(() => {
      expect(box("PLSFIX_UNDO").value).toBe("Ctrl+Alt+Z");
    });
  });

  it("leaves the boxes alone when the host refuses to hand its keys over", async () => {
    const actions = makeActions();
    actions.getShortcuts.mockRejectedValue(new Error("no"));
    installOffice(actions);
    await install();

    expect(box("PLSFIX_UNDO").value).toBe("");
    expect(note()).toContain("Apply");
  });
});

describe("the shortcut manager: Apply", () => {
  it("sends null for every untouched action and the combination for the rest", async () => {
    const actions = makeActions();
    installOffice(actions);
    const { applyShortcuts } = await install();

    box("PLSFIX_AUTOCOLOR").value = "alt+ctrl+j";
    expect(await applyShortcuts()).toBe("Shortcuts updated for your account.");

    expect(actions.areShortcutsInUse).toHaveBeenCalledWith(["Ctrl+Alt+J"]);
    const sent = actions.replaceShortcuts.mock.calls[0]![0] as ShortcutMap;
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
    const { applyShortcuts } = await install();

    await applyShortcuts();
    expect(actions.areShortcutsInUse).not.toHaveBeenCalled();
    const sent = actions.replaceShortcuts.mock.calls[0]![0] as ShortcutMap;
    expect(Object.values(sent).every((key) => key === null)).toBe(true);
  });

  it("names the combinations Office says are already in use", async () => {
    const actions = makeActions();
    actions.areShortcutsInUse.mockResolvedValue([
      { shortcut: "Ctrl+Alt+J", inUse: true },
    ]);
    installOffice(actions);
    const { applyShortcuts } = await install();

    box("PLSFIX_AUTOCOLOR").value = "Ctrl+Alt+J";
    expect(await applyShortcuts()).toBe(
      "Shortcuts updated for your account. Already used elsewhere: Ctrl+Alt+J.",
    );
    expect(actions.replaceShortcuts).toHaveBeenCalledTimes(1);
  });

  it("goes ahead when the host will not answer the conflict question", async () => {
    const actions = makeActions();
    actions.areShortcutsInUse.mockRejectedValue(new Error("no"));
    installOffice(actions);
    const { applyShortcuts } = await install();

    box("PLSFIX_AUTOCOLOR").value = "Ctrl+Alt+J";
    expect(await applyShortcuts()).toBe("Shortcuts updated for your account.");
    expect(actions.replaceShortcuts).toHaveBeenCalledTimes(1);
  });

  it("refuses a box that is not a key combination, by name, before writing", async () => {
    const actions = makeActions();
    installOffice(actions);
    const { applyShortcuts } = await install();

    box("PLSFIX_AUTOCOLOR").value = "Shift+J";
    expect(await rejects(applyShortcuts)).toBe(
      "Autocolor selection: Use Ctrl, Alt or Shift plus one key, for example Ctrl+Shift+A.",
    );
    expect(actions.replaceShortcuts).not.toHaveBeenCalled();
    expect(document.activeElement).toBe(box("PLSFIX_AUTOCOLOR"));
  });

  it("refuses one combination typed into two boxes", async () => {
    const actions = makeActions();
    installOffice(actions);
    const { applyShortcuts } = await install();

    box("PLSFIX_AUTOCOLOR").value = "Ctrl+Alt+J";
    box("PLSFIX_UNDO").value = "ctrl+alt+j";
    expect(await rejects(applyShortcuts)).toBe(
      "Ctrl+Alt+J is set on two actions.",
    );
    expect(actions.replaceShortcuts).not.toHaveBeenCalled();
  });

  it("answers the signed-in sentence when Office refuses the write", async () => {
    const actions = makeActions();
    actions.replaceShortcuts.mockRejectedValue(
      Object.assign(new Error("InvalidOperation"), {
        code: "InvalidOperation",
      }),
    );
    installOffice(actions);
    const { applyShortcuts } = await install();

    box("PLSFIX_AUTOCOLOR").value = "Ctrl+Alt+J";
    expect(await rejects(applyShortcuts)).toBe(UNSUPPORTED);
  });
});

describe("the shortcut manager: Reset all", () => {
  it("sends null for every action and empties every box", async () => {
    const actions = makeActions();
    actions.getShortcuts.mockResolvedValue({ PLSFIX_UNDO: "Ctrl+Alt+Z" });
    installOffice(actions);
    const { resetShortcuts } = await install();
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
});
