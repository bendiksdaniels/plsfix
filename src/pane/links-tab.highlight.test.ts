// @vitest-environment jsdom
// The "Highlight linked cells" checkbox in the Links tab: last session's fills
// put back on boot, the toggle reported through the guard, and the box put back
// when the adapter refuses. The Excel adapter is mocked, so no Office host is
// needed.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  listWorkbookLinks,
  restoreLinkHighlight,
  toggleLinkHighlight,
} from "../excel";
import type { RelayApi } from "../link/relay";
import type { KeyStore } from "../link/workspace";
import type { Guard } from "../ui/guard";
import type { Toast, ToastKind } from "../ui/toast";
import { installLinksTab } from "./links-tab";

vi.mock("../excel", () => ({
  listActiveSheetCharts: vi.fn(async () => []),
  watchActiveSheet: vi.fn(),
  exportActiveChart: vi.fn(),
  exportSelection: vi.fn(),
  goToSource: vi.fn(),
  listWorkbookLinks: vi.fn(),
  pushLinks: vi.fn(),
  removeLink: vi.fn(),
  restoreAutoPush: vi.fn(),
  restoreLinkHighlight: vi.fn(),
  setAutoPush: vi.fn(),
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

// The real pane markup, so a renamed id fails here instead of in Excel.
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

function box(): HTMLInputElement {
  return document.getElementById("links-highlight") as HTMLInputElement;
}

function toggle(on: boolean): void {
  const input = box();
  input.checked = on;
  input.dispatchEvent(new Event("change"));
}

async function settle(h: Harness): Promise<void> {
  while (h.pending.length > 0) await Promise.all(h.pending.splice(0));
  await new Promise((resolve) => setTimeout(resolve, 0));
}

describe("highlight linked cells", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(listWorkbookLinks).mockResolvedValue([]);
    vi.mocked(restoreLinkHighlight).mockResolvedValue(false);
    vi.mocked(toggleLinkHighlight).mockResolvedValue(true);
  });

  it("boots clear in a workbook that was never highlighted", async () => {
    const h = harness();
    install(h);
    await settle(h);

    expect(restoreLinkHighlight).toHaveBeenCalled();
    expect(box().checked).toBe(false);
    expect(h.toasts).toEqual([]);
  });

  // The tint is saved with the file: the box boots clear because the fills
  // under it have just been handed back.
  it("says so when last session's fills were put back", async () => {
    vi.mocked(restoreLinkHighlight).mockResolvedValue(true);
    const h = harness();
    install(h);
    await settle(h);

    await vi.waitFor(() => {
      expect(h.toasts).toEqual([
        {
          message: "Linked-cell highlight from the last session was cleared.",
          kind: undefined,
        },
      ]);
    });
    expect(box().checked).toBe(false);
  });

  it("stays clear when the workbook cannot be read", async () => {
    vi.mocked(restoreLinkHighlight).mockRejectedValue(new Error("no host"));
    const h = harness();
    install(h);
    await settle(h);

    expect(box().checked).toBe(false);
    expect(h.toasts).toEqual([]);
    expect(h.errors).toEqual([]);
  });

  it("paints and clears from the box", async () => {
    const h = harness();
    install(h);
    await settle(h);

    toggle(true);
    await settle(h);
    expect(toggleLinkHighlight).toHaveBeenCalled();
    expect(box().checked).toBe(true);
    expect(h.messages).toContain("Linked cells highlighted.");

    vi.mocked(toggleLinkHighlight).mockResolvedValue(false);
    toggle(false);
    await settle(h);
    expect(box().checked).toBe(false);
    expect(h.messages).toContain("Highlight off: the original fills are back.");
  });

  // The box has to show what the workbook is really doing: a refusal - the
  // audit overlay owns the fills - puts it back.
  it("puts the box back when the adapter refuses", async () => {
    vi.mocked(toggleLinkHighlight).mockRejectedValue(
      new Error("highlight: turn the audit overlay off first"),
    );
    const h = harness();
    install(h);
    await settle(h);

    toggle(true);
    await settle(h);

    expect(box().checked).toBe(false);
    expect(h.errors).toEqual(["highlight: turn the audit overlay off first"]);
  });
});
