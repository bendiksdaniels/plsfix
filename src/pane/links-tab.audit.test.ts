// @vitest-environment jsdom
// E1 audit: the Links tab's boot order and its unhappy panels - a relay that
// never answers the boot touch, a link list the workbook refuses, a clipboard
// the webview does not offer. The Excel adapter is mocked, as in the tab's
// other pane tests.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  listWorkbookLinks,
  restoreAutoPush,
  restoreLinkHighlight,
  touchWorkbookLinks,
  watchActiveSheet,
  watchWorksheetEdits,
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

// Office webviews deny the async Clipboard API; copyText falls back to a
// hidden textarea, and jsdom implements neither. The button must still report.
describe("the link key without a clipboard", () => {
  it("still reports the copy", async () => {
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

    expect(h.errors).toEqual([]);
    expect(h.messages).toContain("Link key copied.");
  });
});
