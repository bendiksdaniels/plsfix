// @vitest-environment jsdom
// The palette as a file: what the Copy button puts on the clipboard when the
// webview allows it and when it does not, and which picked files the Brand tab
// will apply as a palette. Driven through the real pane markup, so the two
// controls that reach ./brand-io are the ones under test.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../excel", () => ({
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
  readWorkbookBrand: vi.fn(async () => null),
  writeWorkbookBrand: vi.fn(async () => undefined),
  setAutocolorOnEdit: vi.fn(async () => undefined),
}));

let store: Map<string, string>;

function workingStorage(): Pick<Storage, "getItem" | "setItem"> {
  return {
    getItem: (name) => store.get(name) ?? null,
    setItem: (name, value) => {
      store.set(name, value);
    },
  };
}

async function load() {
  vi.resetModules();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  await import("./shared");
  const tab = await import("./brand-tab");
  const settings = await import("../settings");
  tab.wireBrand();
  tab.renderBrand();
  return { tab, settings };
}

function toastText(): string {
  return document.getElementById("toast")?.textContent ?? "";
}

function deniedClipboard(): void {
  vi.stubGlobal("navigator", {
    ...navigator,
    clipboard: { writeText: () => Promise.reject(new Error("denied")) },
  });
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

function pick(file: File | null): void {
  const input = document.getElementById("import-file") as HTMLInputElement;
  Object.defineProperty(input, "files", {
    value: file ? [file] : [],
    configurable: true,
  });
  input.dispatchEvent(new Event("change"));
}

async function settle(): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, 0));
}

beforeEach(() => {
  vi.clearAllMocks();
  store = new Map();
  vi.stubGlobal("localStorage", workingStorage());
});

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("palette export", () => {
  it("copies the palette as the JSON the import reads back", async () => {
    const written: string[] = [];
    vi.stubGlobal("navigator", {
      ...navigator,
      clipboard: {
        writeText: async (value: string) => {
          written.push(value);
        },
      },
    });
    const { settings } = await load();

    (document.getElementById("export-brand") as HTMLButtonElement).click();
    await settle();

    expect(toastText()).toBe("Palette JSON copied");
    expect(JSON.parse(written[0] ?? "{}")).toEqual(
      settings.getActiveSettings(),
    );
  });

  // Office webviews can deny the Clipboard API; the hidden textarea behind it
  // is the fallback, and a webview without execCommand either must still leave
  // one sentence rather than an unhandled rejection.
  it("falls back to the textarea when the clipboard is denied", async () => {
    deniedClipboard();
    const copy = vi.fn(() => true);
    Object.defineProperty(document, "execCommand", {
      value: copy,
      configurable: true,
    });
    await load();

    (document.getElementById("export-brand") as HTMLButtonElement).click();
    await settle();

    expect(copy).toHaveBeenCalledWith("copy");
    expect(toastText()).toBe("Palette JSON copied");
    expect(document.querySelector("textarea:not(#copy-manual)")).toBeNull();
  });

  it("says the copy failed rather than throwing out of the handler", async () => {
    deniedClipboard();
    Object.defineProperty(document, "execCommand", {
      value: () => {
        throw new Error("not implemented in this webview");
      },
      configurable: true,
    });
    await load();

    (document.getElementById("export-brand") as HTMLButtonElement).click();
    await settle();

    expect(toastText()).toBe("Copy failed");
    expect(document.querySelector("textarea:not(#copy-manual)")).toBeNull();
  });
});

describe("what counts as a palette file", () => {
  it("does nothing at all when the picker was cancelled", async () => {
    const { settings } = await load();
    const before = settings.getActiveSettings().primary;

    pick(null);
    await settle();

    expect(toastText()).toBe("");
    expect(settings.getActiveSettings().primary).toBe(before);
  });

  it("refuses JSON that is not an object at all", async () => {
    const { settings } = await load();
    const before = settings.getActiveSettings().primary;

    pick(jsonFile("null"));
    await settle();
    expect(toastText()).toBe("That file is not a valid palette JSON.");

    pick(jsonFile('"#123456"'));
    await settle();
    expect(toastText()).toBe("That file is not a valid palette JSON.");
    expect(settings.getActiveSettings().primary).toBe(before);
  });
});
