// @vitest-environment jsdom
// E1 audit: the Links tab's boot order and its unhappy panels - a relay that
// never answers the boot touch, a link list the workbook refuses, a clipboard
// the webview does not offer. The Excel adapter is mocked, as in the tab's
// other pane tests.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  listActiveSheetCharts,
  listWorkbookLinks,
  pushLinks,
  removeLink,
  restoreAutoPush,
  restoreLinkHighlight,
  touchWorkbookLinks,
  watchActiveSheet,
  watchWorksheetEdits,
  type WorkbookLinkRow,
} from "../excel";
import type { RelayApi } from "../link/relay";
import type { KeyStore } from "../link/workspace";
import type { Guard } from "../ui/guard";
import type { Toast, ToastKind } from "../ui/toast";
import { installLinksTab } from "./links-tab";

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
  readProjectState: vi.fn(async () => ({ names: [], active: undefined })),
  setActiveProject: vi.fn(async () => undefined),
  moveLinksToProject: vi.fn(async () => undefined),
  restoreAutoPush: vi.fn(),
  setAutoPush: vi.fn(),
  restoreLinkHighlight: vi.fn(),
  toggleLinkHighlight: vi.fn(),
}));

interface Harness {
  guard: Guard;
  toast: Toast;
  relay: RelayApi;
  keyStore: KeyStore;
  messages: string[];
  errors: string[];
  toasts: { message: string; kind?: ToastKind }[];
  pending: Promise<void>[];
}

function harness(): Harness {
  const messages: string[] = [];
  const errors: string[] = [];
  const toasts: { message: string; kind?: ToastKind }[] = [];
  const pending: Promise<void>[] = [];
  const stored = new Map<string, string>();
  return {
    messages,
    errors,
    toasts,
    pending,
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
      show: (message, kind) => {
        toasts.push({ message, kind });
      },
    },
    keyStore: {
      get: async (key) => stored.get(key) ?? null,
      set: async (key, value) => {
        stored.set(key, value);
      },
      remove: async (key) => {
        stored.delete(key);
      },
    },
  };
}

function install(h: Harness): void {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  installLinksTab({
    guard: h.guard,
    toast: h.toast,
    relay: h.relay,
    keyStore: h.keyStore,
    root: document,
  });
}

function click(id: string): void {
  document.querySelector<HTMLButtonElement>(`#${id}`)?.click();
}

async function settle(h: Harness): Promise<void> {
  while (h.pending.length > 0) await Promise.all(h.pending.splice(0));
  await new Promise((resolve) => setTimeout(resolve, 0));
}

beforeEach(() => {
  vi.clearAllMocks();
  vi.mocked(listWorkbookLinks).mockResolvedValue([]);
  vi.mocked(restoreAutoPush).mockResolvedValue(false);
  vi.mocked(restoreLinkHighlight).mockResolvedValue(false);
  vi.mocked(touchWorkbookLinks).mockResolvedValue(0);
  vi.mocked(listActiveSheetCharts).mockResolvedValue([]);
  vi.mocked(pushLinks).mockResolvedValue({
    pushed: 1,
    missing: 0,
    failed: 0,
    failures: [],
  });
});

// The TTL touch is one HTTP call with no timeout of its own: a relay behind a
// dead route leaves that promise open for as long as the webview's own socket
// timeout, which is a minute or more. Nothing the tab does after it may wait
// on it.
describe("boot with a relay that never answers", () => {
  it("still arms the watchers and reads the two tick boxes", async () => {
    vi.mocked(touchWorkbookLinks).mockReturnValue(new Promise(() => undefined));
    vi.mocked(restoreAutoPush).mockResolvedValue(true);
    const h = harness();
    install(h);
    await settle(h);

    expect(vi.mocked(listWorkbookLinks)).toHaveBeenCalled();
    expect(vi.mocked(watchWorksheetEdits)).toHaveBeenCalled();
    expect(vi.mocked(watchActiveSheet)).toHaveBeenCalled();
    expect(vi.mocked(restoreAutoPush)).toHaveBeenCalled();
    expect(
      document.querySelector<HTMLInputElement>("#links-autopush")!.checked,
    ).toBe(true);
  });

  it("boots the same way when the touch rejects", async () => {
    vi.mocked(touchWorkbookLinks).mockRejectedValue(
      new Error("relay unreachable"),
    );
    const h = harness();
    install(h);
    await settle(h);

    expect(vi.mocked(watchWorksheetEdits)).toHaveBeenCalled();
    // Never a toast on boot: a failed touch changes nothing the user can see.
    expect(h.toasts).toEqual([]);
  });
});

describe("a link list the workbook refuses", () => {
  it("says so in the table and clears the ticks", async () => {
    vi.mocked(listWorkbookLinks).mockRejectedValue(
      new Error("registry PLSFIX_LINKS: unreadable, not overwriting"),
    );
    const h = harness();
    install(h);
    await settle(h);

    const body = document.getElementById("workbook-links")!;
    expect(body.textContent).toBe(
      "This workbook's links could not be read: registry PLSFIX_LINKS: unreadable, not overwriting",
    );
    expect(body.querySelectorAll("tr[data-link-id]")).toHaveLength(0);

    // Nothing is ticked, so the row actions still ask for a selection first.
    click("push-selected");
    await settle(h);
    expect(h.errors).toEqual(["Select a link in the list first."]);
  });

  it("names the failure even when the host threw a bare string", async () => {
    vi.mocked(listWorkbookLinks).mockRejectedValue("InvalidOperation");
    const h = harness();
    install(h);
    await settle(h);

    expect(document.getElementById("workbook-links")!.textContent).toBe(
      "This workbook's links could not be read: unknown error",
    );
  });
});

function row(id: string, source: "ok" | "missing" = "ok"): WorkbookLinkRow {
  return {
    entry: {
      id,
      kind: "range",
      anchor: `PLSFIX_LINK_${id.slice(0, 8)}`,
      label: `Model!B${id.slice(0, 2)}`,
      token: "token",
      createdAt: "2026-08-29T11:00:00.000Z",
      lastPushedAt: "2026-08-29T11:58:00.000Z",
      rev: 3,
    },
    source,
  };
}

function ticks(): HTMLInputElement[] {
  return [
    ...document.querySelectorAll<HTMLInputElement>(
      "#workbook-links input[type=checkbox]",
    ),
  ];
}

// The list is re-read on every settled edit and after every action, and the
// ticks live beside it: a tick left over from a link that is gone must never
// reach the adapter, and one whose row is still there must survive the paint.
describe("the link table across a refresh", () => {
  it("drops a tick whose link is gone and keeps the rest", async () => {
    const ids = ["a", "b", "c"].map((letter) => letter.repeat(32));
    vi.mocked(listWorkbookLinks).mockResolvedValue(ids.map((id) => row(id)));
    const h = harness();
    install(h);
    await settle(h);

    ticks()[0]!.click();
    ticks()[2]!.click();
    vi.mocked(listWorkbookLinks).mockResolvedValue([
      row(ids[1]!),
      row(ids[2]!),
    ]);
    click("tab-links");
    await vi.waitFor(() => {
      expect(ticks()).toHaveLength(2);
    });

    expect(ticks().map((box) => box.checked)).toEqual([false, true]);
    click("push-selected");
    await settle(h);
    expect(vi.mocked(pushLinks)).toHaveBeenCalledWith([ids[2]], h.relay);
  });

  it("forgets a row the modeller unticked and removes the two that are left", async () => {
    const ids = ["a", "b", "c"].map((letter) => letter.repeat(32));
    vi.mocked(listWorkbookLinks).mockResolvedValue(ids.map((id) => row(id)));
    const h = harness();
    install(h);
    await settle(h);

    for (const box of ticks()) box.click();
    ticks()[0]!.click();

    click("remove-link");
    await settle(h);
    expect(vi.mocked(removeLink).mock.calls.map((call) => call[0])).toEqual([
      ids[1],
      ids[2],
    ]);
    expect(h.messages).toContain("Removed 2 links");
  });

  it("draws two hundred links, badges the broken ones and empties again", async () => {
    const many = Array.from({ length: 200 }, (_unused, index) =>
      row(String(index).padStart(32, "0"), index % 20 === 0 ? "missing" : "ok"),
    );
    vi.mocked(listWorkbookLinks).mockResolvedValue(many);
    const h = harness();
    install(h);
    await settle(h);

    const body = document.getElementById("workbook-links")!;
    expect(body.querySelectorAll("tr[data-link-id]")).toHaveLength(200);
    expect(body.querySelectorAll(".wl-badge")).toHaveLength(10);

    vi.mocked(listWorkbookLinks).mockResolvedValue([]);
    click("tab-links");
    await vi.waitFor(() => {
      expect(body.textContent).toBe("No linked objects in this workbook yet.");
    });
    expect(body.querySelectorAll("tr[data-link-id]")).toHaveLength(0);
  });
});

// The picker only exists to name a chart when none is selected. A host that
// refuses the read is the same as a sheet with no charts: it hides, and the
// export button says what it says.
describe("the chart picker when the host refuses the read", () => {
  it("hides rather than reporting, and the rest of boot still happens", async () => {
    vi.mocked(listActiveSheetCharts).mockRejectedValue(
      new Error("charts unavailable"),
    );
    const h = harness();
    install(h);
    await settle(h);

    const pick =
      document.querySelector<HTMLSelectElement>("#export-chart-pick")!;
    expect(pick.hidden).toBe(true);
    expect(h.toasts).toEqual([]);
    expect(h.errors).toEqual([]);
    expect(vi.mocked(watchWorksheetEdits)).toHaveBeenCalled();
  });

  it("shows the sheet's charts and hides again when the next sheet has none", async () => {
    vi.mocked(listActiveSheetCharts).mockResolvedValue(["Revenue bridge"]);
    const h = harness();
    install(h);
    await settle(h);

    const pick =
      document.querySelector<HTMLSelectElement>("#export-chart-pick")!;
    expect(pick.hidden).toBe(false);
    expect([...pick.options].map((option) => option.value)).toEqual([
      "",
      "Revenue bridge",
    ]);

    vi.mocked(listActiveSheetCharts).mockResolvedValue([]);
    click("tab-links");
    await vi.waitFor(() => {
      expect(pick.hidden).toBe(true);
    });
    expect([...pick.options].map((option) => option.value)).toEqual([""]);
  });
});

// installLinksTab is the only thing main.ts adds for the tab: a renamed id has
// to fail loudly at boot rather than leave a dead button in the pane.
describe("markup the tab cannot work with", () => {
  it("names the element that is missing", () => {
    document.body.innerHTML = "<div></div>";
    const h = harness();
    expect(() =>
      installLinksTab({
        guard: h.guard,
        toast: h.toast,
        relay: h.relay,
        keyStore: h.keyStore,
        root: document,
      }),
    ).toThrow("Missing element #export-chart-pick");
  });
});

// Office webviews deny the async Clipboard API; copyText falls back to a
// hidden textarea, and jsdom implements neither. The button must still report.
describe("the link key without a clipboard", () => {
  it("reports the copy failed instead of a false success", async () => {
    Object.defineProperty(navigator, "clipboard", {
      value: undefined,
      configurable: true,
    });
    const h = harness();
    install(h);
    await settle(h);

    click("generate-key");
    await settle(h);
    click("copy-key");
    await settle(h);

    // jsdom's execCommand("copy") fallback does not really land text on a
    // clipboard, so this exercises the same "neither path worked" case a
    // webview with both APIs denied would hit.
    expect(h.errors).toContain(
      "Copy failed: select the text and copy it by hand.",
    );
    expect(h.messages).not.toContain("Link key copied.");
  });
});
