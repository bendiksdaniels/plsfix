// @vitest-environment jsdom
// The Links tab as the modeller drives it: the list on install, the actions
// behind each button and the link-key panel - masked by default, and refusing
// to generate over a key it could not read. The Excel adapter is mocked, so no
// Office host is needed; the table renderer is covered in links-list.test.ts.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  exportActiveChart,
  exportSelection,
  exportSelectionAsTable,
  exportSelectionAsText,
  goToSource,
  listWorkbookLinks,
  pushLinks,
  removeLink,
  type WorkbookLinkRow,
  listActiveSheetCharts,
} from "../excel";
import type { RegistryEntry } from "../link/model";
import type { RelayApi } from "../link/relay";
import type { KeyStore } from "../link/workspace";
import { WORKSPACE_STORAGE_KEY } from "../link/workspace";
import type { Guard } from "../ui/guard";
import type { Toast, ToastKind } from "../ui/toast";
import { installLinksTab } from "./links-tab";

// The adapter needs a real Excel host, so the tab is wired against mocks; the
// markup under test is taskpane.html itself, which keeps ids from drifting.
vi.mock("../excel", () => ({
  listActiveSheetCharts: vi.fn(async () => []),
  watchActiveSheet: vi.fn(),
  watchWorksheetEdits: vi.fn(),
  exportActiveChart: vi.fn(),
  exportSelection: vi.fn(),
  exportSelectionAsTable: vi.fn(),
  exportSelectionAsText: vi.fn(),
  goToSource: vi.fn(),
  listWorkbookLinks: vi.fn(),
  pushLinks: vi.fn(),
  removeLink: vi.fn(),
  touchWorkbookLinks: vi.fn(async () => 0),
}));

const ID_A = "a".repeat(32);
const ID_B = "b".repeat(32);

function row(
  id: string,
  entry: Partial<RegistryEntry> = {},
  source: "ok" | "missing" = "ok",
): WorkbookLinkRow {
  return {
    entry: {
      id,
      kind: "range",
      anchor: `PLSFIX_LINK_${id.slice(0, 8)}`,
      label: "Model!B4:F12",
      token: "token",
      createdAt: "2026-08-29T11:00:00.000Z",
      lastPushedAt: "2026-08-29T11:58:00.000Z",
      rev: 3,
      ...entry,
    },
    source,
  };
}

interface Harness {
  guard: Guard;
  toast: Toast;
  relay: RelayApi;
  keyStore: KeyStore;
  messages: string[];
  errors: string[];
  toasts: { message: string; kind?: ToastKind; details?: string }[];
  stored: Map<string, string>;
  pending: Promise<void>[];
  // Flipped on to model storage that is momentarily unavailable: a private
  // webview, or a cold pane racing OfficeRuntime.storage.
  failRead: { on: boolean };
}

function harness(): Harness {
  const messages: string[] = [];
  const errors: string[] = [];
  const toasts: { message: string; kind?: ToastKind; details?: string }[] = [];
  const stored = new Map<string, string>();
  const pending: Promise<void>[] = [];
  const failRead = { on: false };
  return {
    messages,
    errors,
    toasts,
    stored,
    pending,
    failRead,
    relay: {} as RelayApi,
    guard: (run) => {
      const done = (async () => {
        try {
          messages.push(await run());
        } catch (error) {
          errors.push(error instanceof Error ? error.message : String(error));
        }
      })();
      pending.push(done);
      return done;
    },
    toast: {
      show: (message, kind, details) => {
        toasts.push({ message, kind, details });
      },
    },
    keyStore: {
      get: async (key) => {
        if (failRead.on) throw new Error("storage unavailable");
        return stored.get(key) ?? null;
      },
      set: async (key, value) => {
        stored.set(key, value);
      },
      remove: async (key) => {
        stored.delete(key);
      },
    },
  };
}

// The real pane markup, so a renamed id fails here instead of in Excel.
// import.meta.url is not a file URL under the jsdom environment, so the pane
// is read from the project root vitest already runs in.
function paneRoot(): Document {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  return document;
}

function click(id: string): void {
  document.querySelector<HTMLButtonElement>(`#${id}`)?.click();
}

function keyDisplay(): string {
  return document.getElementById("workspace-key-display")?.textContent ?? "";
}

function button(id: string): HTMLButtonElement {
  return document.getElementById(id) as HTMLButtonElement;
}

// A click starts an action nobody awaits, and real WebCrypto needs more than
// one turn, so the fake guard hands every in-flight run back to be waited on.
async function settle(h: Harness): Promise<void> {
  while (h.pending.length > 0) await Promise.all(h.pending.splice(0));
  await new Promise((resolve) => setTimeout(resolve, 0));
}

function install(h: Harness): ReturnType<typeof installLinksTab> {
  return installLinksTab({
    guard: h.guard,
    toast: h.toast,
    relay: h.relay,
    keyStore: h.keyStore,
    root: paneRoot(),
  });
}

describe("installLinksTab", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(listWorkbookLinks).mockResolvedValue([]);
  });

  it("lists the workbook's links on install", async () => {
    vi.mocked(listWorkbookLinks).mockResolvedValue([row(ID_A), row(ID_B)]);
    const h = harness();
    install(h);
    await settle(h);

    expect(listWorkbookLinks).toHaveBeenCalled();
    const body = document.getElementById("workbook-links");
    expect(body?.querySelectorAll("tr[data-link-id]")).toHaveLength(2);
  });

  it("refuses to export before a link key exists", async () => {
    const h = harness();
    install(h);
    await settle(h);

    click("export-selection");
    await settle(h);

    expect(h.errors).toEqual(["Generate a link key first (Links > Settings)."]);
    expect(exportSelection).not.toHaveBeenCalled();
  });

  // The second button on the same selection: a different adapter call, the
  // same key, the same report line.
  it("sends the selection as a table from its own button", async () => {
    vi.mocked(exportSelectionAsTable).mockResolvedValue({
      id: ID_B,
      label: "Model!B4:F12 table",
    });
    const h = harness();
    install(h);
    await settle(h);
    click("generate-key");
    await settle(h);

    click("export-table");
    await settle(h);

    expect(exportSelectionAsTable).toHaveBeenCalledTimes(1);
    expect(exportSelection).not.toHaveBeenCalled();
    expect(h.messages).toContain("Sent to PowerPoint: Model!B4:F12 table");
  });

  // The third button on the same selection: one cell, its own adapter call.
  it("sends the selection as text from its own button", async () => {
    vi.mocked(exportSelectionAsText).mockResolvedValue({
      id: ID_B,
      label: "Model!B4 text",
    });
    const h = harness();
    install(h);
    await settle(h);
    click("generate-key");
    await settle(h);

    click("export-text");
    await settle(h);

    expect(exportSelectionAsText).toHaveBeenCalledTimes(1);
    expect(exportSelection).not.toHaveBeenCalled();
    expect(h.messages).toContain("Sent to PowerPoint: Model!B4 text");
  });

  // The key is the secret itself, and the pane is screen-shared on deal calls:
  // it is masked until asked for, and Copy is the route to the whole value.
  it("masks a generated key, reveals it on demand and exports with it", async () => {
    vi.mocked(exportSelection).mockResolvedValue({
      id: ID_A,
      label: "Model!B4:F12",
    });
    const writeText = vi.fn<(text: string) => Promise<void>>(() =>
      Promise.resolve(),
    );
    Object.defineProperty(navigator, "clipboard", {
      value: { writeText },
      configurable: true,
    });
    const h = harness();
    install(h);
    await settle(h);

    click("generate-key");
    await settle(h);
    const key = h.stored.get(WORKSPACE_STORAGE_KEY) ?? "";
    expect(key).toMatch(/^[A-Za-z0-9_-]{43}$/);
    expect(keyDisplay()).toBe(`${key.slice(0, 4)}…${key.slice(-4)}`);
    expect(keyDisplay()).not.toContain(key.slice(4, 20));

    click("copy-key");
    await settle(h);
    expect(writeText).toHaveBeenCalledWith(key);

    click("reveal-key");
    expect(keyDisplay()).toBe(key);
    expect(button("reveal-key").textContent).toBe("Hide");
    click("reveal-key");
    expect(keyDisplay()).not.toBe(key);

    click("export-selection");
    await settle(h);
    expect(exportSelection).toHaveBeenCalledTimes(1);
    expect(h.messages).toContain("Sent to PowerPoint: Model!B4:F12");

    click("forget-key");
    await settle(h);
    expect(h.stored.size).toBe(0);
    expect(keyDisplay()).toBe("No link key yet.");
    expect(button("reveal-key").disabled).toBe(true);
  });

  // "No link key yet." over a read that failed would invite a new key, and a
  // new key unpairs every deck holding the old one.
  it("reports a key that cannot be read and refuses to generate over it", async () => {
    const h = harness();
    h.failRead.on = true;
    install(h);
    await settle(h);

    const unreadable = "Could not read the link key on this computer.";
    expect(keyDisplay()).toBe(unreadable);
    expect(h.toasts).toEqual([
      { message: unreadable, kind: "error", details: "storage unavailable" },
    ]);
    expect(button("generate-key").disabled).toBe(true);

    click("generate-key");
    await settle(h);
    expect(h.stored.size).toBe(0);

    click("export-selection");
    await settle(h);
    expect(h.errors).toEqual([unreadable]);
    expect(exportSelection).not.toHaveBeenCalled();

    // Coming back to the tab reads the key again. The reload runs outside the
    // guard, so it is waited for rather than settled.
    h.failRead.on = false;
    h.stored.set(WORKSPACE_STORAGE_KEY, "k".repeat(43));
    click("tab-links");
    await vi.waitFor(() => {
      expect(button("generate-key").disabled).toBe(false);
    });
    expect(keyDisplay()).toBe("kkkk…kkkk");
  });

  it("reports a push summary and hands the failures to the toast", async () => {
    vi.mocked(listWorkbookLinks).mockResolvedValue([row(ID_A)]);
    vi.mocked(pushLinks).mockResolvedValue({
      pushed: 1,
      missing: 1,
      failed: 1,
      failures: ["Model!B4:F12: relay PUT /api/links: 500"],
    });
    const h = harness();
    install(h);
    await settle(h);

    click("push-all");
    await settle(h);

    expect(pushLinks).toHaveBeenCalledWith("all", h.relay);
    expect(h.messages).toContain("1 pushed, 1 missing, 1 failed");
    expect(h.toasts).toEqual([
      {
        message: "1 pushed, 1 missing, 1 failed",
        kind: "error",
        details: "Model!B4:F12: relay PUT /api/links: 500",
      },
    ]);
  });

  it("pushes only the ticked rows and asks for one first", async () => {
    vi.mocked(listWorkbookLinks).mockResolvedValue([row(ID_A), row(ID_B)]);
    vi.mocked(pushLinks).mockResolvedValue({
      pushed: 1,
      missing: 0,
      failed: 0,
      failures: [],
    });
    const h = harness();
    install(h);
    await settle(h);

    click("push-selected");
    await settle(h);
    expect(h.errors).toEqual(["Select a link in the list first."]);
    expect(pushLinks).not.toHaveBeenCalled();

    document
      .querySelector<HTMLInputElement>("#workbook-links input[type=checkbox]")
      ?.click();
    click("push-selected");
    await settle(h);
    expect(pushLinks).toHaveBeenCalledWith([ID_A], h.relay);
    expect(h.toasts).toEqual([]);
  });

  it("jumps to and removes the ticked link", async () => {
    vi.mocked(listWorkbookLinks).mockResolvedValue([row(ID_A)]);
    const h = harness();
    install(h);
    await settle(h);

    document
      .querySelector<HTMLInputElement>("#workbook-links input[type=checkbox]")
      ?.click();

    click("go-to-source");
    await settle(h);
    expect(goToSource).toHaveBeenCalledWith(ID_A);

    click("remove-link");
    await settle(h);
    expect(removeLink).toHaveBeenCalledWith(ID_A, h.relay);
    expect(h.messages).toContain("Removed 1 link");
  });
});

describe("chart list", () => {
  it("re-reads the sheet's charts when the selection moves", async () => {
    vi.mocked(listActiveSheetCharts).mockResolvedValue(["Revenue chart"]);
    const h = harness();
    const tab = install(h);
    await settle(h);
    const before = vi.mocked(listActiveSheetCharts).mock.calls.length;
    await tab.sheetChanged();
    expect(vi.mocked(listActiveSheetCharts).mock.calls.length).toBe(before + 1);
    const pick =
      document.querySelector<HTMLSelectElement>("#export-chart-pick")!;
    expect(pick.hidden).toBe(false);
    expect(Array.from(pick.options).map((option) => option.value)).toEqual([
      "",
      "Revenue chart",
    ]);
  });
});

// A chart Excel could not describe still leaves as a picture; the toast says
// so in the sentence PowerPoint will repeat beside the inserted picture.
describe("chart export", () => {
  it("passes on why the chart will arrive as a picture", async () => {
    vi.mocked(exportActiveChart).mockResolvedValue({
      id: ID_B,
      label: "Model!Revenue chart",
      note: "as a picture: 7 series; shapes draw up to 6",
    });
    const h = harness();
    install(h);
    await settle(h);
    click("generate-key");
    await settle(h);

    click("export-chart");
    await settle(h);

    expect(exportActiveChart).toHaveBeenCalledTimes(1);
    expect(h.messages).toContain(
      "Sent to PowerPoint: Model!Revenue chart (as a picture: 7 series; shapes draw up to 6)",
    );
  });
});
