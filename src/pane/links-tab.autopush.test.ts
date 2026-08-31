// @vitest-environment jsdom
// The auto-push checkbox in the Links tab: armed again from the workbook on
// boot, reported through the guard when it is toggled, put back when the toggle
// fails, and the one place the watcher's own messages reach the toast. The
// Excel adapter is mocked, so no Office host is needed.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { listWorkbookLinks, restoreAutoPush, setAutoPush } from "../excel";
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
  goToSource: vi.fn(),
  listWorkbookLinks: vi.fn(),
  pushLinks: vi.fn(),
  removeLink: vi.fn(),
  restoreAutoPush: vi.fn(),
  setAutoPush: vi.fn(),
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
  return document.getElementById("links-autopush") as HTMLInputElement;
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

// The watcher pushes without anyone clicking, so its messages go straight to
// the toast; this is the callback the tab hands it.
function watcherNotify(): (message: string) => void {
  const call = vi.mocked(restoreAutoPush).mock.calls[0];
  if (!call) throw new Error("restoreAutoPush was never called");
  return call[1];
}

describe("auto-push on edit", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(listWorkbookLinks).mockResolvedValue([]);
    vi.mocked(restoreAutoPush).mockResolvedValue(false);
    vi.mocked(setAutoPush).mockResolvedValue(undefined);
  });

  it("leaves the box clear in a workbook that never asked for it", async () => {
    const h = harness();
    install(h);
    await settle(h);

    expect(restoreAutoPush).toHaveBeenCalledWith(h.relay, expect.any(Function));
    expect(box().checked).toBe(false);
    expect(h.toasts).toEqual([]);
  });

  it("ticks the box when the workbook had it on", async () => {
    vi.mocked(restoreAutoPush).mockResolvedValue(true);
    const h = harness();
    install(h);
    await settle(h);

    await vi.waitFor(() => {
      expect(box().checked).toBe(true);
    });
  });

  // A pane that cannot reach Excel yet has nothing to report at boot: the box
  // stays clear rather than raising an error nobody can act on.
  it("stays clear when the workbook cannot be read", async () => {
    vi.mocked(restoreAutoPush).mockRejectedValue(new Error("no host"));
    const h = harness();
    install(h);
    await settle(h);

    expect(box().checked).toBe(false);
    expect(h.toasts).toEqual([]);
    expect(h.errors).toEqual([]);
  });

  it("arms and disarms the watcher from the box", async () => {
    const h = harness();
    install(h);
    await settle(h);

    toggle(true);
    await settle(h);
    expect(setAutoPush).toHaveBeenLastCalledWith(
      true,
      h.relay,
      expect.any(Function),
    );
    expect(h.messages).toContain(
      "Auto-push on: linked pictures follow your edits.",
    );

    toggle(false);
    await settle(h);
    expect(setAutoPush).toHaveBeenLastCalledWith(
      false,
      h.relay,
      expect.any(Function),
    );
    expect(h.messages).toContain("Auto-push off.");
  });

  // The box has to show what the workbook is really doing.
  it("puts the box back when arming fails", async () => {
    vi.mocked(setAutoPush).mockRejectedValue(new Error("handler refused"));
    const h = harness();
    install(h);
    await settle(h);

    toggle(true);
    await settle(h);

    expect(box().checked).toBe(false);
    expect(h.errors).toEqual(["handler refused"]);
  });

  it("shows what the watcher says", async () => {
    const h = harness();
    install(h);
    await settle(h);

    watcherNotify()("Pushed 2 links");
    expect(h.toasts).toEqual([{ message: "Pushed 2 links", kind: undefined }]);
  });
});
