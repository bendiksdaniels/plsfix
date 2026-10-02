// Hunt pass 2, target 1: failure injection at every sync step of Break
// link and every object tool, refused or hung. Invariant: a plain
// sentence and a settled promise; object tools are not checked for
// position (the fake tracks only adds as pending, a write can still land).

import { afterEach, describe, expect, it, vi } from "vitest";
import { fakePng } from "../fakepng";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
} from "../fakeppt";
import { bootPpt, memoryStore, seedLink } from "../ppt.support";
import { createWorkspace, type Workspace } from "../../src/link/workspace";
import * as host from "../../src/ppt/host";
import type * as ObjectToolsModule from "../../src/ppt/object-tools";
import { settleHungSync } from "../hung-sync";

enableStrictLoadSemantics();

function expectPlainSentence(message: string): void {
  expect(message).not.toMatch(/InvalidArgument|ItemNotFound|GeneralException/);
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
});

// ---------------------------------------------------------------------------
// Break link
// ---------------------------------------------------------------------------

async function breakScenario(): Promise<{
  helpers: FakePptHelpers;
  presentation: Awaited<ReturnType<typeof bootPpt>>["presentation"];
  found: host.FoundLink;
}> {
  const { helpers, presentation, relay } = await bootPpt();
  const links = await import("../../src/ppt/links");
  const ws: Workspace = await createWorkspace(memoryStore());
  const item = await seedLink(fakePng(200, 100));
  await links.insertFromInbox(item, ws, relay);
  const row = (await links.listLinks(relay))[0]!;
  return { helpers, presentation, found: row.found };
}

const BREAK_TOTAL_SYNCS = await (async () => {
  const scenario = await breakScenario();
  const before = scenario.helpers.syncCount();
  await host.breakLink(scenario.found);
  const total = scenario.helpers.syncCount() - before;
  uninstallFakePpt();
  return total;
})();

describe(`Break link: every sync step (${String(BREAK_TOTAL_SYNCS)} sync${BREAK_TOTAL_SYNCS === 1 ? "" : "s"} when it succeeds)`, () => {
  for (let n = 0; n < BREAK_TOTAL_SYNCS; n += 1) {
    it(`refusing sync ${String(n)} rejects cleanly, never leaves exactly one of the two tags`, async () => {
      const { helpers, presentation, found } = await breakScenario();
      helpers.refuseNextSync(new Error("the host answered with an error"), n);

      await expect(host.breakLink(found)).rejects.toThrow();

      const shape = presentation.slides[0]!.shapes[0]!;
      const hasLink = shape.tags.has("PLSFIX_LINK");
      const hasKey = shape.tags.has("PLSFIX_KEY");
      expect(hasLink).toBe(hasKey);
    });

    it(`a hung sync ${String(n)} settles, never stuck`, async () => {
      const { helpers, found } = await breakScenario();
      helpers.hangNextSync(n);
      vi.useFakeTimers();

      let settled = false;
      try {
        await settleHungSync(host.breakLink(found));
        settled = true;
      } catch {
        settled = true;
      }

      expect(settled).toBe(true);
    });
  }
});

// ---------------------------------------------------------------------------
// Object tools
// ---------------------------------------------------------------------------

interface ToolCase {
  name: string;
  shapeCount: number;
  run: (tools: typeof ObjectToolsModule) => Promise<string>;
  // applyObjectStyle refuses outright with no captured style at all
  // ("Capture an object style first.") - primed with one clean, unarmed
  // capture first so its OWN syncs are what the sweep below actually
  // exercises. object-tools.ts's `painter` is module state, which is
  // exactly why this runs inside the SAME import as `run`, before any
  // failure is armed.
  prime?: (tools: typeof ObjectToolsModule) => Promise<void>;
}

const TOOL_CASES: ToolCase[] = [
  { name: "align", shapeCount: 2, run: (t) => t.alignSelected("left") },
  {
    name: "distribute",
    shapeCount: 3,
    run: (t) => t.distributeSelected("horizontal"),
  },
  { name: "matchSize", shapeCount: 2, run: (t) => t.matchSelectedSize() },
  { name: "swap", shapeCount: 2, run: (t) => t.swapSelected() },
  { name: "selectSimilar", shapeCount: 1, run: (t) => t.selectSimilar() },
  {
    name: "captureObjectStyle",
    shapeCount: 1,
    run: (t) => t.captureObjectStyle(),
  },
  {
    name: "applyObjectStyle",
    shapeCount: 2,
    run: (t) => t.applyObjectStyle(),
    prime: async (t) => {
      await t.captureObjectStyle();
    },
  },
];

async function toolScenario(
  toolCase: ToolCase,
): Promise<{ helpers: FakePptHelpers; tools: typeof ObjectToolsModule }> {
  const { helpers, presentation } = installFakePpt({ slides: 1 });
  const slide = presentation.slides[0]!;
  const ids: string[] = [];
  for (let i = 0; i < toolCase.shapeCount; i += 1) {
    const shape = presentation.addShape(slide, {
      left: i * 100,
      top: 0,
      width: 50,
      height: 50,
    });
    ids.push(shape.id);
  }
  const tools = await import("../../src/ppt/object-tools");
  if (toolCase.prime) {
    // captureObjectStyle (the one prime this file uses) requires EXACTLY
    // one shape selected, whatever the case's own shapeCount is - select
    // just the first for the prime, then restore the real selection.
    helpers.selectShapes([ids[0]!]);
    await toolCase.prime(tools);
  }
  helpers.selectShapes(ids);
  return { helpers, tools };
}

async function measureToolSyncs(toolCase: ToolCase): Promise<number> {
  const { helpers, tools } = await toolScenario(toolCase);
  const before = helpers.syncCount();
  await toolCase.run(tools);
  const total = helpers.syncCount() - before;
  uninstallFakePpt();
  return total;
}

const TOOL_TOTALS = new Map<string, number>();
for (const toolCase of TOOL_CASES) {
  TOOL_TOTALS.set(toolCase.name, await measureToolSyncs(toolCase));
}

describe("Object tools: every sync step refused or hung, one case per step", () => {
  for (const toolCase of TOOL_CASES) {
    const total = TOOL_TOTALS.get(toolCase.name)!;
    describe(`${toolCase.name} (${String(total)} syncs when it succeeds)`, () => {
      for (let n = 0; n < total; n += 1) {
        it(`refusing sync ${String(n)} rejects with a plain sentence`, async () => {
          const { helpers, tools } = await toolScenario(toolCase);
          helpers.refuseNextSync(
            new Error("the host answered with an error"),
            n,
          );

          let threw: unknown;
          try {
            await toolCase.run(tools);
          } catch (error) {
            threw = error;
          }

          expect(threw).toBeDefined();
          expectPlainSentence(messageOf(threw));
        });

        it(`a hung sync ${String(n)} settles, never stuck`, async () => {
          const { helpers, tools } = await toolScenario(toolCase);
          helpers.hangNextSync(n);
          vi.useFakeTimers();

          let settled = false;
          try {
            await settleHungSync(toolCase.run(tools));
            settled = true;
          } catch {
            settled = true;
          }

          expect(settled).toBe(true);
        });
      }
    });
  }
});
