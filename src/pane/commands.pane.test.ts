// @vitest-environment jsdom
// The shortcut card's pane-only actions: Super Find and the style scrubber
// both put their answer in the Workbook tab and nowhere else, so pressing
// their keys with the pane shut has to open it first - the way PLSFIX_SHOWPANE
// does. Every other id on the card writes to the workbook, which is visible
// whether the pane is open or not, and this suite holds that line: a new
// pane-only action either joins NEEDS_PANE or leaves the card.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const order: string[] = [];

const excel = vi.hoisted(() => {
  const names = [
    "addCagrLabel",
    "applyBorderCycle",
    "applyColumnWidthCycle",
    "applyDecimalStep",
    "applyFillCycle",
    "applyFontColorCycle",
    "applyNumberCycle",
    "applyRowHeightCycle",
    "applyRowStyleCycle",
    "applySignFlip",
    "autocolorSelection",
    "fastFillAuto",
    "formatSelectedChart",
    "insertCagr",
    "insertConsistentRounding",
    "insertTornado",
    "insertWaterfall",
    "markCopySource",
    "pasteSpecial",
    "pastePreserveFormulas",
    "scaleSelection",
    "toggleIfErrorGuard",
    "undoLastAction",
    "unpivotSelection",
  ];
  const stub: Record<string, unknown> = {
    lastUndoSkipped: vi.fn(() => undefined),
  };
  for (const name of names) stub[name] = vi.fn(async () => undefined);
  return stub;
});
vi.mock("../excel", () => excel);
vi.mock("./find-panel", () => ({
  focusFind: vi.fn(async () => {
    order.push("find");
  }),
}));
vi.mock("./styles-panel", () => ({
  focusStyles: vi.fn(async () => {
    order.push("styles");
  }),
}));
vi.mock("./paint-slots", () => ({
  capturePaintSlot: vi.fn(async () => undefined),
  applyPaintSlot: vi.fn(async () => undefined),
}));
vi.mock("./share-panel", () => ({
  prepareShare: vi.fn(async () => undefined),
}));
vi.mock("./trace-panel", () => ({
  startTrace: vi.fn(async () => undefined),
  toggleAudit: vi.fn(async () => undefined),
}));
vi.mock("./workbook-tab", () => ({
  insertTocSheet: vi.fn(async () => undefined),
}));
vi.mock("./shared", () => ({
  APP_VERSION: "v9.9.999",
  refreshSelection: vi.fn(async () => undefined),
  renderActionState: vi.fn(() => undefined),
  toast: { show: vi.fn() },
}));

import { autocolorSelection } from "../excel";
import { toast } from "./shared";
import { registerCommands } from "./commands";

type Handler = (event?: { completed: () => void }) => void;

function stubOffice(): Map<string, Handler> {
  const associated = new Map<string, Handler>();
  (globalThis as { Office?: unknown }).Office = {
    actions: {
      associate: (id: string, handler: Handler) => associated.set(id, handler),
    },
    addin: {
      showAsTaskpane: vi.fn(async () => {
        order.push("pane");
      }),
    },
  };
  return associated;
}

const settle = () => new Promise((resolve) => setTimeout(resolve, 0));

async function fire(id: string): Promise<() => void> {
  const associated = stubOffice();
  registerCommands();
  const completed = vi.fn();
  associated.get(id)!({ completed });
  await settle();
  return completed;
}

// The two the ruling covers, and the card that has to carry them.
const PANE_ONLY = ["PLSFIX_FIND", "PLSFIX_STYLES_SCAN"];

function shortcutCardIds(): string[] {
  const json = JSON.parse(
    readFileSync(join(process.cwd(), "public/shortcuts.json"), "utf8"),
  ) as { actions: { id: string }[] };
  return json.actions.map((action) => action.id);
}

beforeEach(() => {
  vi.clearAllMocks();
  order.length = 0;
});
afterEach(() => {
  delete (globalThis as { Office?: unknown }).Office;
});

describe("a shortcut whose whole answer is in the pane", () => {
  it("opens the pane before Super Find runs", async () => {
    const completed = await fire("PLSFIX_FIND");
    expect(order).toEqual(["pane", "find"]);
    expect(completed).toHaveBeenCalledOnce();
    expect(toast.show).not.toHaveBeenCalled();
  });

  it("opens the pane before the style scan runs", async () => {
    const completed = await fire("PLSFIX_STYLES_SCAN");
    expect(order).toEqual(["pane", "styles"]);
    expect(completed).toHaveBeenCalledOnce();
  });

  it("still runs the action when the host refuses to show the pane", async () => {
    const associated = stubOffice();
    const office = (
      globalThis as { Office: { addin: { showAsTaskpane: () => void } } }
    ).Office;
    vi.mocked(office.addin.showAsTaskpane).mockRejectedValueOnce(
      new Error("no host"),
    );
    registerCommands();
    const completed = vi.fn();
    associated.get("PLSFIX_FIND")!({ completed });
    await settle();
    expect(order).toEqual(["find"]);
    expect(completed).toHaveBeenCalledOnce();
    expect(toast.show).not.toHaveBeenCalled();
  });

  it("is on the shortcut card, which is what makes the rule worth having", () => {
    const card = new Set(shortcutCardIds());
    for (const id of PANE_ONLY) expect(card.has(id)).toBe(true);
  });
});

describe("a shortcut the workbook itself answers", () => {
  it("leaves the pane shut: the cells show what happened", async () => {
    const completed = await fire("PLSFIX_AUTOCOLOR");
    expect(order).toEqual([]);
    expect(autocolorSelection).toHaveBeenCalledOnce();
    expect(completed).toHaveBeenCalledOnce();
  });

  it("opens nothing for the trace pair either: they jump the selection", async () => {
    await fire("PLSFIX_TRACE_PRE");
    await fire("PLSFIX_TRACE_DEP");
    expect(order).toEqual([]);
  });
});

describe("PLSFIX_SHOWPANE", () => {
  it("opens the pane through the one helper and completes", async () => {
    const completed = await fire("PLSFIX_SHOWPANE");
    expect(order).toEqual(["pane"]);
    expect(completed).toHaveBeenCalledOnce();
  });
});
