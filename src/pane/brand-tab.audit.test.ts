// @vitest-environment jsdom
// The Brand tab driven through the real pane markup: what the pickers, the
// selects and the JSON import write, and what the tab does on a workbook that
// refuses its settings with localStorage refusing too. The Excel adapter is
// mocked, so no Office host is needed.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CONFIRM_MS } from "../ui/confirm";

vi.mock("../excel", () => ({
  // What ./shared reaches for at import time.
  copySourceLabel: vi.fn(() => null),
  inspectSelection: vi.fn(async () => ({
    address: "A1",
    cells: 1,
    formulas: 0,
    errors: 0,
    blanks: 0,
  })),
  lastUndoSkipped: vi.fn(() => false),
  undoTarget: vi.fn(() => null),
  // What the Brand tab itself calls.
  readWorkbookBrand: vi.fn(async () => null),
  writeWorkbookBrand: vi.fn(async () => undefined),
  setAutocolorOnEdit: vi.fn(async () => undefined),
}));

import {
  readWorkbookBrand,
  setAutocolorOnEdit,
  writeWorkbookBrand,
} from "../excel";

// This Node/jsdom combination ships a global `localStorage` whose methods are
// all undefined, so the tab is driven against a real, controllable one.
let store: Map<string, string>;

function workingStorage(): Pick<Storage, "getItem" | "setItem"> {
  return {
    getItem: (name) => store.get(name) ?? null,
    setItem: (name, value) => {
      store.set(name, value);
    },
  };
}

function throwingStorage(): Pick<Storage, "getItem" | "setItem"> {
  return {
    getItem: () => {
      throw new Error("storage blocked");
    },
    setItem: () => {
      throw new Error("storage blocked");
    },
  };
}

function paneRoot(): void {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
}

// The tab and ./shared both read the DOM at import time, so the markup is in
// place before either module is loaded, and the registry is reset per test.
async function load() {
  vi.resetModules();
  paneRoot();
  const shared = await import("./shared");
  const tab = await import("./brand-tab");
  const settings = await import("../settings");
  return { shared, tab, settings };
}

function hex(slot: string): HTMLInputElement {
  return document.querySelector<HTMLInputElement>(
    `[data-slot-hex="${slot}"]`,
  ) as HTMLInputElement;
}

function toastText(): string {
  return document.getElementById("toast")?.textContent ?? "";
}

function toastKind(): string {
  return document.getElementById("toast")?.className ?? "";
}

// A file input cannot be filled from script in jsdom; the handler only ever
// reads files[0], which is what a real pick hands it.
function pick(id: string, file: File): void {
  const input = document.getElementById(id) as HTMLInputElement;
  Object.defineProperty(input, "files", { value: [file], configurable: true });
  input.dispatchEvent(new Event("change"));
}

// jsdom's Blob has no text(); the pane only ever calls that one method.
function jsonFile(text: string): File {
  const file = new File([text], "palette.json", { type: "application/json" });
  Object.defineProperty(file, "text", {
    value: () => Promise.resolve(text),
    configurable: true,
  });
  return file;
}

async function settle(): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, 0));
}

describe("brand tab", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    store = new Map();
    vi.stubGlobal("localStorage", workingStorage());
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("renders the stored palette into every picker, field and select", async () => {
    const { tab, settings } = await load();
    tab.wireBrand();
    tab.renderBrand();

    const active = settings.getActiveSettings();
    expect(hex("primary").value).toBe(active.primary);
    expect(
      document.querySelector<HTMLInputElement>('[data-slot-color="accent"]')
        ?.value,
    ).toBe(active.accent.toLowerCase());
    expect(
      (document.getElementById("setting-font") as HTMLSelectElement).value,
    ).toBe(active.font);
    expect(
      document.getElementById("currency-format-button")?.textContent,
    ).toContain("1 234");
  });

  it("applies a hex field to the palette and saves it both ways", async () => {
    const { tab, shared, settings } = await load();
    shared.setExcelReady(true);
    tab.wireBrand();
    tab.renderBrand();

    hex("accent").value = "#123456";
    hex("accent").dispatchEvent(new Event("change"));
    await settle();

    expect(settings.getActiveSettings().accent).toBe("#123456");
    expect(store.get("plsfix.brand.v1")).toContain("#123456");
    expect(writeWorkbookBrand).toHaveBeenCalledTimes(1);
  });

  it("says so and repaints the field when the hex is not a colour", async () => {
    const { tab, settings } = await load();
    tab.wireBrand();
    tab.renderBrand();
    const before = settings.getActiveSettings().accent;

    hex("accent").value = "not a colour";
    hex("accent").dispatchEvent(new Event("change"));

    expect(toastText()).toContain("is not a hex color");
    expect(settings.getActiveSettings().accent).toBe(before);
    expect(hex("accent").value).toBe(before);
  });

  it("imports a palette file and leaves a malformed one alone", async () => {
    const { tab, settings } = await load();
    tab.wireBrand();
    tab.renderBrand();

    pick("import-file", jsonFile('{"primary":"#0A0B0C"}'));
    await settle();
    expect(settings.getActiveSettings().primary).toBe("#0A0B0C");
    expect(toastText()).toBe("Palette imported");

    pick("import-file", jsonFile("{not json"));
    await settle();
    expect(settings.getActiveSettings().primary).toBe("#0A0B0C");
    expect(toastText()).toBe("That file is not a valid palette JSON.");
  });

  // A file that parses but carries none of the palette's own keys is somebody
  // else's JSON. Applying it would silently reset every colour to the shipped
  // defaults under a "Palette imported" toast.
  it("refuses a JSON file that carries no palette keys at all", async () => {
    const { tab, settings } = await load();
    tab.wireBrand();
    tab.renderBrand();

    hex("primary").value = "#0A0B0C";
    hex("primary").dispatchEvent(new Event("change"));
    await settle();

    pick("import-file", jsonFile('{"name":"my-model","version":"1.0.0"}'));
    await settle();

    expect(settings.getActiveSettings().primary).toBe("#0A0B0C");
    expect(toastText()).toBe("That file is not a valid palette JSON.");
  });

  it("refuses an empty file and a JSON array the same way", async () => {
    const { tab, settings } = await load();
    tab.wireBrand();
    tab.renderBrand();
    const before = settings.getActiveSettings().primary;

    pick("import-file", jsonFile(""));
    await settle();
    expect(toastText()).toBe("That file is not a valid palette JSON.");

    pick("import-file", jsonFile("[]"));
    await settle();
    expect(toastText()).toBe("That file is not a valid palette JSON.");
    expect(settings.getActiveSettings().primary).toBe(before);
  });

  it("reads a file that will not open as one sentence, not a crash", async () => {
    const { tab } = await load();
    tab.wireBrand();
    tab.renderBrand();

    const broken = jsonFile("{}");
    Object.defineProperty(broken, "text", {
      value: () => Promise.reject(new Error("the file was moved")),
    });
    pick("import-file", broken);
    await settle();

    expect(toastText()).toBe("That file could not be read.");
  });
});

describe("reset-brand: two-click confirm", () => {
  beforeEach(() => {
    store = new Map();
    vi.stubGlobal("localStorage", workingStorage());
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("arms on the first press without resetting anything, resets on the second", async () => {
    const { tab, settings } = await load();
    tab.wireBrand();
    tab.renderBrand();
    const button = document.getElementById("reset-brand") as HTMLButtonElement;
    const before = settings.getActiveSettings().accent;
    hex("accent").value = "#123456";
    hex("accent").dispatchEvent(new Event("change"));
    await settle();
    expect(settings.getActiveSettings().accent).toBe("#123456");

    button.click();
    // An icon-only button (src/ui/confirm.ts): the glyph stays, the
    // accessible name carries the confirm instead.
    expect(button.textContent?.trim()).toBe("⟲");
    expect(button.getAttribute("aria-label")).toBe("Click again to confirm");
    expect(button.classList.contains("armed")).toBe(true);
    expect(settings.getActiveSettings().accent).toBe("#123456");

    button.click();
    expect(settings.getActiveSettings().accent).toBe(before);
    expect(button.classList.contains("armed")).toBe(false);
    expect(button.getAttribute("aria-label")).toBe("Reset brand settings");
  });

  it("lapses on its own after five seconds and a press only re-arms", async () => {
    const { tab, settings } = await load();
    tab.wireBrand();
    tab.renderBrand();
    const button = document.getElementById("reset-brand") as HTMLButtonElement;
    hex("accent").value = "#123456";
    hex("accent").dispatchEvent(new Event("change"));
    await settle();

    vi.useFakeTimers();
    try {
      button.click();
      expect(button.classList.contains("armed")).toBe(true);
      vi.advanceTimersByTime(CONFIRM_MS);
      expect(button.classList.contains("armed")).toBe(false);
    } finally {
      vi.useRealTimers();
    }
    expect(settings.getActiveSettings().accent).toBe("#123456");
  });
});

// The Map's precedence rule, from the winning side: a palette saved in the
// file replaces the machine default the pane booted with.
describe("the workbook's own palette", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    store = new Map();
    vi.stubGlobal("localStorage", workingStorage());
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("replaces the machine palette and repaints every field", async () => {
    const { tab, shared, settings } = await load();
    shared.setExcelReady(true);
    tab.loadSettings();
    tab.wireBrand();
    tab.renderBrand();
    hex("accent").value = "#111111";
    hex("accent").dispatchEvent(new Event("change"));
    await settle();

    vi.mocked(readWorkbookBrand).mockResolvedValue(
      JSON.stringify({ accent: "#00FF00", font: "Georgia" }),
    );
    await tab.adoptWorkbookBrand();

    expect(settings.getActiveSettings().accent).toBe("#00FF00");
    expect(hex("accent").value).toBe("#00FF00");
    expect(
      (document.getElementById("setting-font") as HTMLSelectElement).value,
    ).toBe("Georgia");
    expect(setAutocolorOnEdit).toHaveBeenCalled();
  });

  // A workbook that never carried a brand keeps the machine's, not the
  // shipped colours.
  it("keeps the machine palette when the workbook carries none", async () => {
    vi.mocked(readWorkbookBrand).mockResolvedValue(null);
    const { tab, settings } = await load();
    tab.wireBrand();
    tab.renderBrand();
    hex("accent").value = "#111111";
    hex("accent").dispatchEvent(new Event("change"));
    await settle();

    await tab.adoptWorkbookBrand();

    expect(settings.getActiveSettings().accent).toBe("#111111");
  });
});

// The Map's rule: the workbook wins over localStorage, which is only the
// default a new workbook starts from. Both refusing is a webview in a private
// session on a read-only file, and the tab still has to work.
describe("brand tab with no store at all", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    store = new Map();
    vi.stubGlobal("localStorage", workingStorage());
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("renders the shipped palette when localStorage refuses to be read", async () => {
    vi.stubGlobal("localStorage", throwingStorage());
    const { tab, settings } = await load();

    tab.loadSettings();
    tab.wireBrand();
    tab.renderBrand();

    expect(settings.getActiveSettings()).toEqual(settings.DEFAULT_SETTINGS);
    expect(hex("primary").value).toBe(settings.DEFAULT_SETTINGS.primary);
  });

  it("still applies a colour when neither store will take it", async () => {
    vi.stubGlobal("localStorage", throwingStorage());
    vi.mocked(writeWorkbookBrand).mockRejectedValue(
      new Error("Settings are not available on this workbook."),
    );
    const { tab, shared, settings } = await load();
    shared.setExcelReady(true);
    tab.loadSettings();
    tab.wireBrand();
    tab.renderBrand();

    hex("primary").value = "#123456";
    hex("primary").dispatchEvent(new Event("change"));
    await settle();

    // The palette the sheet paints with is the one the modeller just chose.
    expect(settings.getActiveSettings().primary).toBe("#123456");
    expect(settings.activeTheme().titleFill).toBe("#123456");
    expect(hex("primary").value).toBe("#123456");
    // One sentence about the save, not an exception out of the handler.
    expect(toastText()).toContain("Settings are not available");
    expect(toastKind()).toContain("error");
  });

  // A workbook read that fails is the same story from the other side: the boot
  // adopt step must not take the machine's palette down with it.
  it("keeps the machine palette when the workbook read is refused", async () => {
    vi.mocked(readWorkbookBrand).mockRejectedValue(new Error("no settings"));
    const { tab, settings } = await load();
    tab.loadSettings();
    tab.wireBrand();
    tab.renderBrand();
    hex("accent").value = "#654321";
    hex("accent").dispatchEvent(new Event("change"));
    await settle();

    await tab.adoptWorkbookBrand();

    expect(settings.getActiveSettings().accent).toBe("#654321");
    expect(toastText()).toContain("no settings");
  });

  // Autocolor on edit is an Office.js handler: a pane that has not connected
  // yet has no workbook to register it on, so a palette change must not turn
  // into a red toast about a host nobody asked for.
  it("does not call the host from a palette change before Excel is ready", async () => {
    const { tab } = await load();
    tab.wireBrand();
    tab.renderBrand();

    hex("formula").value = "#222222";
    hex("formula").dispatchEvent(new Event("change"));
    await settle();

    expect(setAutocolorOnEdit).not.toHaveBeenCalled();
    expect(toastKind()).not.toContain("error");
  });
});
