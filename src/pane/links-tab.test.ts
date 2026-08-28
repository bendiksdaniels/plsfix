// @vitest-environment jsdom
// Render tests for the Links list: one row per registry entry, the missing
// badge, the push time and the checkbox callback. installLinksTab is covered
// with the Excel adapter mocked, so no Office host is needed here.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  exportSelection,
  goToSource,
  listWorkbookLinks,
  pushLinks,
  removeLink,
  type WorkbookLinkRow,
} from "../excel";
import type { RegistryEntry } from "../link/model";
import type { RelayApi } from "../link/relay";
import type { KeyStore } from "../link/workspace";
import { WORKSPACE_STORAGE_KEY } from "../link/workspace";
import type { Guard } from "../ui/guard";
import type { Toast, ToastKind } from "../ui/toast";
import { installLinksTab, renderWorkbookLinks } from "./links-tab";

// The adapter needs a real Excel host, so the tab is wired against mocks; the
// markup under test is taskpane.html itself, which keeps ids from drifting.
vi.mock("../excel", () => ({
  exportActiveChart: vi.fn(),
  exportSelection: vi.fn(),
  goToSource: vi.fn(),
  listWorkbookLinks: vi.fn(),
  pushLinks: vi.fn(),
  removeLink: vi.fn(),
}));

const ID_A = "a".repeat(32);
const ID_B = "b".repeat(32);
const NOW = "2026-08-29T12:00:00.000Z";

function tbody(): HTMLTableSectionElement {
  document.body.innerHTML = "<table><tbody id='rows'></tbody></table>";
  return document.getElementById("rows") as HTMLTableSectionElement;
}

function row(
  id: string,
  entry: Partial<RegistryEntry> = {},
  source: "ok" | "missing" = "ok",
): WorkbookLinkRow {
  return {
    entry: {
      id,
      kind: "range",
      anchor: `SMT_LINK_${id.slice(0, 8)}`,
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

function boxes(body: HTMLTableSectionElement): HTMLInputElement[] {
  return Array.from(
    body.querySelectorAll<HTMLInputElement>("input[type=checkbox]"),
  );
}

describe("renderWorkbookLinks", () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date(NOW));
  });
  afterEach(() => {
    vi.useRealTimers();
  });

  it("renders one row per entry with its label and anchor", () => {
    const body = tbody();
    renderWorkbookLinks(
      body,
      [row(ID_A), row(ID_B, { label: "Chart 1", kind: "chart" })],
      new Set(),
      () => undefined,
    );

    const rows = body.querySelectorAll("tr[data-link-id]");
    expect(rows).toHaveLength(2);
    expect(rows[0]?.textContent).toContain("Model!B4:F12");
    expect(rows[0]?.textContent).toContain("SMT_LINK_aaaaaaaa");
    expect(rows[1]?.textContent).toContain("Chart 1");
  });

  it("spells the push time in minutes, hours and days", () => {
    const body = tbody();
    renderWorkbookLinks(
      body,
      [
        row(ID_A, { lastPushedAt: "2026-08-29T11:58:00.000Z" }),
        row(ID_B, { lastPushedAt: "2026-08-29T07:00:00.000Z" }),
        row("c".repeat(32), { lastPushedAt: "2026-08-26T12:00:00.000Z" }),
      ],
      new Set(),
      () => undefined,
    );

    const rows = body.querySelectorAll("tr[data-link-id]");
    expect(rows[0]?.textContent).toContain("Pushed 2 min ago");
    expect(rows[1]?.textContent).toContain("Pushed 5 h ago");
    expect(rows[2]?.textContent).toContain("Pushed 3 days ago");
  });

  it("badges a missing source and says never for a link never pushed", () => {
    const body = tbody();
    renderWorkbookLinks(
      body,
      [row(ID_A, { lastPushedAt: null }, "missing"), row(ID_B)],
      new Set(),
      () => undefined,
    );

    const rows = body.querySelectorAll("tr[data-link-id]");
    expect(rows[0]?.textContent).toContain("Source missing");
    expect(rows[0]?.textContent).toContain("never");
    expect(rows[1]?.textContent).not.toContain("Source missing");
  });

  it("checks the selected rows and reports every toggle", () => {
    const body = tbody();
    const onToggle = vi.fn();
    renderWorkbookLinks(
      body,
      [row(ID_A), row(ID_B)],
      new Set([ID_A]),
      onToggle,
    );

    const [first, second] = boxes(body);
    expect(first?.checked).toBe(true);
    expect(second?.checked).toBe(false);

    second?.click();
    expect(onToggle).toHaveBeenCalledWith(ID_B, true);
    first?.click();
    expect(onToggle).toHaveBeenCalledWith(ID_A, false);
  });

  it("clears stale rows and shows an empty state", () => {
    const body = tbody();
    renderWorkbookLinks(body, [row(ID_A)], new Set(), () => undefined);
    renderWorkbookLinks(body, [], new Set(), () => undefined);

    expect(body.querySelectorAll("tr[data-link-id]")).toHaveLength(0);
    expect(body.textContent).toContain("No linked objects");
  });
});

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
}

function harness(): Harness {
  const messages: string[] = [];
  const errors: string[] = [];
  const toasts: { message: string; kind?: ToastKind; details?: string }[] = [];
  const stored = new Map<string, string>();
  const pending: Promise<void>[] = [];
  return {
    messages,
    errors,
    toasts,
    stored,
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
      show: (message, kind, details) => {
        toasts.push({ message, kind, details });
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

// A click starts an action nobody awaits, and real WebCrypto needs more than
// one turn, so the fake guard hands every in-flight run back to be waited on.
async function settle(h: Harness): Promise<void> {
  while (h.pending.length > 0) await Promise.all(h.pending.splice(0));
  await new Promise((resolve) => setTimeout(resolve, 0));
}

function install(h: Harness): { refresh(): Promise<void> } {
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

  it("shows a generated key and then exports with it", async () => {
    vi.mocked(exportSelection).mockResolvedValue({
      id: ID_A,
      label: "Model!B4:F12",
    });
    const h = harness();
    install(h);
    await settle(h);

    click("generate-key");
    await settle(h);
    const shown = document.getElementById("workspace-key-display")?.textContent;
    expect(shown).toMatch(/^[A-Za-z0-9_-]{43}$/);
    expect(h.stored.get(WORKSPACE_STORAGE_KEY)).toBe(shown);

    click("export-selection");
    await settle(h);
    expect(exportSelection).toHaveBeenCalledTimes(1);
    expect(h.messages).toContain("Sent to PowerPoint: Model!B4:F12");

    click("forget-key");
    await settle(h);
    expect(h.stored.size).toBe(0);
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
