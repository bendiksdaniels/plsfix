// @vitest-environment jsdom
// The rest of the Brand tab: the logo colour picker, the three Output-defaults
// selects, the reset, the palette export and the separators note. Driven
// through the real pane markup with the Excel adapter mocked.
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

import { setAutocolorOnEdit } from "../excel";

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
  const shared = await import("./shared");
  const tab = await import("./brand-tab");
  const settings = await import("../settings");
  tab.wireBrand();
  tab.renderBrand();
  return { shared, tab, settings };
}

function toastText(): string {
  return document.getElementById("toast")?.textContent ?? "";
}

function select(id: string, value: string): void {
  const el = document.getElementById(id) as HTMLSelectElement;
  el.value = value;
  el.dispatchEvent(new Event("change"));
}

async function settle(): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, 0));
}

// ---------------------------------------------------------------------------
// The logo picker needs three browser pieces jsdom does not ship: an object
// URL, an <img> that decodes, and a 2d canvas. Each is stubbed to the shape
// the pane actually uses, so the module's own arithmetic is what runs.
// ---------------------------------------------------------------------------

type ImageOutcome = "load" | "error";

function stubImage(outcome: ImageOutcome, width = 200, height = 100): void {
  class FakeImage {
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    width = width;
    height = height;
    set src(_value: string) {
      setTimeout(() => {
        if (outcome === "load") this.onload?.();
        else this.onerror?.();
      }, 0);
    }
  }
  vi.stubGlobal("Image", FakeImage);
}

// Records the canvas the pane drew on, so the scaling step can be read back.
const drawnOn: HTMLCanvasElement[] = [];

function stubCanvas(pixels: number[] | null): void {
  const context =
    pixels === null
      ? null
      : {
          drawImage: () => undefined,
          getImageData: () => ({ data: new Uint8ClampedArray(pixels) }),
        };
  vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockImplementation(
    function (this: HTMLCanvasElement) {
      drawnOn.push(this);
      return context as unknown as CanvasRenderingContext2D;
    },
  );
}

// Two solid colours on white, opaque: the first bucket is the commoner one.
function twoColourLogo(): number[] {
  const navy = [20, 33, 61, 255];
  const mint = [46, 196, 182, 255];
  const white = [255, 255, 255, 255];
  return [...navy, ...navy, ...navy, ...mint, ...mint, ...white];
}

function pickLogo(): void {
  const input = document.getElementById("logo-file") as HTMLInputElement;
  const file = new File([""], "logo.png", { type: "image/png" });
  Object.defineProperty(input, "files", { value: [file], configurable: true });
  input.dispatchEvent(new Event("change"));
}

beforeEach(() => {
  vi.clearAllMocks();
  store = new Map();
  vi.stubGlobal("localStorage", workingStorage());
  URL.createObjectURL = vi.fn(() => "blob:logo");
  URL.revokeObjectURL = vi.fn();
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("logo colour picker", () => {
  it("offers the colours it read and alternates primary then accent", async () => {
    stubImage("load");
    stubCanvas(twoColourLogo());
    const { settings } = await load();

    pickLogo();
    await settle();

    const strip = document.getElementById("logo-swatches") as HTMLDivElement;
    expect(strip.hidden).toBe(false);
    expect(document.getElementById("logo-hint")?.hidden).toBe(false);
    const swatches = Array.from(strip.querySelectorAll("button"));
    expect(swatches).toHaveLength(2);

    swatches[0]?.click();
    await settle();
    expect(toastText()).toContain("Primary set to");
    const first = settings.getActiveSettings().primary;

    swatches[1]?.click();
    await settle();
    expect(toastText()).toContain("Accent set to");
    expect(settings.getActiveSettings().primary).toBe(first);
  });

  it("says so when the file is not a readable image", async () => {
    stubImage("error");
    stubCanvas(twoColourLogo());
    await load();

    pickLogo();
    await settle();

    expect(toastText()).toBe("That file is not a readable image.");
    expect(document.getElementById("logo-swatches")?.hidden).toBe(true);
  });

  it("says so when the webview will not give a canvas", async () => {
    stubImage("load");
    stubCanvas(null);
    await load();

    pickLogo();
    await settle();

    expect(toastText()).toBe("Could not read that image.");
  });

  // A photo of a white wall, or a logo that is all background.
  it("says so when the image holds no usable colour", async () => {
    stubImage("load");
    stubCanvas([255, 255, 255, 255, 250, 250, 250, 255]);
    await load();

    pickLogo();
    await settle();

    expect(toastText()).toBe("No usable colors found in that image.");
    expect(document.getElementById("logo-swatches")?.hidden).toBe(true);
  });

  // A 4000px logo is scaled to 64 before a single pixel is read; the canvas
  // never grows with the file.
  it("scales a huge image down before reading it", async () => {
    stubImage("load", 4000, 2000);
    stubCanvas(twoColourLogo());
    await load();

    pickLogo();
    await settle();

    // 64 on the long side, the aspect kept, whatever the file measured.
    expect(drawnOn.at(-1)?.width).toBe(64);
    expect(drawnOn.at(-1)?.height).toBe(32);
    expect(document.getElementById("logo-swatches")?.hidden).toBe(false);
  });

  it("clears the swatches when the palette is reset", async () => {
    stubImage("load");
    stubCanvas(twoColourLogo());
    const { settings } = await load();
    pickLogo();
    await settle();

    (document.getElementById("reset-brand") as HTMLButtonElement).click();
    await settle();

    expect(settings.getActiveSettings()).toEqual(settings.DEFAULT_SETTINGS);
    expect(document.getElementById("logo-swatches")?.hidden).toBe(true);
    expect(document.getElementById("logo-hint")?.hidden).toBe(true);
    expect(toastText()).toBe("Palette reset to pls,fix defaults");
  });
});

describe("output defaults", () => {
  it("writes the font, the currency and the language through", async () => {
    const { settings } = await load();

    select("setting-font", "Calibri");
    select("setting-currency", "$");
    select("setting-language", "en");

    const active = settings.getActiveSettings();
    expect(active.font).toBe("Calibri");
    expect(active.currency).toBe("$");
    expect(active.language).toBe("en");
    // The currency button is the same amount in the language's own style.
    expect(document.getElementById("currency-format-button")?.textContent).toBe(
      "$ 1,234",
    );
  });

  it("turns the edit handler on and off with one sentence each", async () => {
    const { shared } = await load();
    shared.setExcelReady(true);
    const box = document.getElementById(
      "setting-autocolor-edit",
    ) as HTMLInputElement;

    box.checked = true;
    box.dispatchEvent(new Event("change"));
    await settle();
    expect(toastText()).toBe("Autocolor runs on every edit");
    expect(setAutocolorOnEdit).toHaveBeenLastCalledWith(true);

    box.checked = false;
    box.dispatchEvent(new Event("change"));
    await settle();
    expect(toastText()).toBe("Autocolor on edit is off");
    expect(setAutocolorOnEdit).toHaveBeenLastCalledWith(false);
  });
});

describe("the separators note", () => {
  it("stays hidden while Excel has not answered, or cannot", async () => {
    const { tab } = await load();
    const note = document.getElementById("separators-note") as HTMLElement;
    expect(note.hidden).toBe(true);

    // Below ExcelApi 1.11 readSeparators answers null and boot passes it on.
    tab.applyExcelSeparators(null);
    expect(note.hidden).toBe(true);
  });

  it("says the house style is already set when Excel agrees", async () => {
    const { tab } = await load();

    tab.applyExcelSeparators({ decimal: ".", thousands: " " });

    const note = document.getElementById("separators-note") as HTMLElement;
    expect(note.hidden).toBe(false);
    expect(note.textContent).toContain("the house style");
  });

  it("names the setting to change when Excel disagrees", async () => {
    const { tab } = await load();

    tab.applyExcelSeparators({ decimal: ",", thousands: "." });

    const note = document.getElementById("separators-note") as HTMLElement;
    expect(note.textContent).toContain("Excel shows 1.094.417,5");
    expect(note.textContent).toContain("the Latviešu style is 1 094 417.5");
    expect(note.textContent).toContain("File > Options > Advanced on Windows");
  });

  it("redraws itself when the language changes under it", async () => {
    const { tab } = await load();
    tab.applyExcelSeparators({ decimal: ".", thousands: " " });
    expect(document.getElementById("separators-note")?.textContent).toContain(
      "the house style",
    );

    select("setting-language", "en");

    expect(document.getElementById("separators-note")?.textContent).toContain(
      "the English style is 1,094,417.5",
    );
  });
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
});
