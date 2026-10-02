import { afterEach, describe, expect, it, vi } from "vitest";
import { commandTable, registerCommands, type CommandDeps } from "./commands";

const PRESENTATION_FUNCTIONS = [
  "PLSFIX_PPT_ALIGN_BOTTOM",
  "PLSFIX_PPT_ALIGN_CENTER",
  "PLSFIX_PPT_ALIGN_LEFT",
  "PLSFIX_PPT_ALIGN_MIDDLE",
  "PLSFIX_PPT_ALIGN_RIGHT",
  "PLSFIX_PPT_ALIGN_TOP",
  "PLSFIX_PPT_CAPTURE",
  "PLSFIX_PPT_DIST_ACROSS",
  "PLSFIX_PPT_DIST_DOWN",
  "PLSFIX_PPT_MATCH",
  "PLSFIX_PPT_PAINT",
  "PLSFIX_PPT_SIMILAR",
  "PLSFIX_PPT_SWAP",
  "PLSFIX_PPT_TOOLS",
];

function deps(
  showTools: CommandDeps["showTools"] = vi.fn(async () => undefined),
): CommandDeps {
  return {
    notify: vi.fn(),
    context: { host: "PowerPoint", version: "v9.9.999" },
    showTools,
  };
}

type Handler = (event?: { completed: () => void }) => void;

function stubOffice(): Map<string, Handler> {
  const associated = new Map<string, Handler>();
  (globalThis as { Office?: unknown }).Office = {
    actions: {
      associate: (id: string, handler: Handler) => associated.set(id, handler),
    },
  };
  return associated;
}

const settle = () => new Promise((resolve) => setTimeout(resolve, 0));

describe("PowerPoint ribbon commands", () => {
  afterEach(() => {
    delete (globalThis as { Office?: unknown }).Office;
  });

  it("maps every FunctionName of the Presentation host, and nothing else", () => {
    expect(Object.keys(commandTable(deps())).sort()).toEqual(
      PRESENTATION_FUNCTIONS,
    );
  });

  it("Object tools brings the pane up on its Tools tab and says nothing", async () => {
    const d = deps();
    expect(await commandTable(d)["PLSFIX_PPT_TOOLS"]!()).toBe("");
    expect(d.showTools).toHaveBeenCalledOnce();
  });

  it("registers each command once, toasts its result and completes the event", async () => {
    const associated = stubOffice();
    const d = deps();
    registerCommands(d);
    expect([...associated.keys()].sort()).toEqual(PRESENTATION_FUNCTIONS);

    const completed = vi.fn();
    associated.get("PLSFIX_PPT_TOOLS")!({ completed });
    await settle();
    expect(d.showTools).toHaveBeenCalledOnce();
    expect(d.notify).not.toHaveBeenCalled();
    expect(completed).toHaveBeenCalledOnce();
  });

  it("reports a failed command through the toast, named by its FunctionName", async () => {
    const associated = stubOffice();
    const d = deps(vi.fn(async () => Promise.reject(new Error("no pane"))));
    registerCommands(d);
    const completed = vi.fn();
    associated.get("PLSFIX_PPT_TOOLS")!({ completed });
    await settle();
    expect(d.notify).toHaveBeenCalledWith(
      expect.stringContaining("no pane"),
      "error",
      expect.stringContaining("PLSFIX_PPT_TOOLS"),
    );
    expect(completed).toHaveBeenCalledOnce();
  });

  it("does nothing on a host without Office.actions", () => {
    (globalThis as { Office?: unknown }).Office = {};
    expect(() => registerCommands(deps())).not.toThrow();
  });
});
