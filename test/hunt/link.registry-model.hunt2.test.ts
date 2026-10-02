// Pass-2 model-based test: a tiny reference model of the link registry,
// driven by random export/push/remove/move/reopen sequences against the
// REAL links.ts + link-projects.ts. Checks: revisions only climb per
// transport space, no duplicate ids, and reopen round-trips every link.

import fc from "fast-check";
import { describe, it, vi } from "vitest";
import type * as LinksModule from "../../src/excel/links";
import { moveLinksToProject } from "../../src/excel/link-projects";
import { isLocalRev } from "../../src/link/local";
import { LocalCollector } from "../../src/link/local-collector";
import type { RelayApi } from "../../src/link/relay";
import {
  createWorkspace,
  type KeyStore,
  type Workspace,
} from "../../src/link/workspace";
import { FakeRelay } from "../fakerelay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "../fakehost";

enableStrictLoadSemantics();

const SEED = 20260927;
const RUNS = 500;
const MAX_COMMANDS = 40;

type Kind = "text" | "table" | "picture" | "chart";
const SLOT_KINDS: readonly Kind[] = ["text", "table", "picture", "chart"];
const SLOT_COUNT = SLOT_KINDS.length;
const CELL: readonly string[] = ["Model!B2", "Model!B4:D6", "Model!B9", ""];
const PROJECTS: readonly (string | undefined)[] = ["Alpha", "Beta", undefined];

interface SlotModel {
  id: string | null;
  removed: boolean;
  project: string | undefined;
  // The last rev sent in each space, so "climbs" is checked per space.
  lastLocalRev: number | null;
  lastRelayRev: number | null;
  // The space THIS command just sent in, for verifySlot's rev check; set
  // and cleared by verifyAfterSend around the one call.
  lastSpaceHint?: "local" | "relay";
}

interface Model {
  slots: SlotModel[];
  transport: "local" | "relay";
  // newEntry stamps a fresh export with the registry's activeProject, and
  // moveLinksToProject sets it too (withActiveProject): a NEW export joins
  // whichever project was last moved to, not only this slot's own move.
  activeProject: string | undefined;
}

interface Sut {
  helpers: FakeHelpers;
  relay: FakeRelay;
  ws: Workspace;
  links: typeof LinksModule;
  reopen(): Promise<void>;
}

function memoryStore(): KeyStore {
  const map = new Map<string, string>();
  return {
    get: async (k) => map.get(k) ?? null,
    set: async (k, v) => {
      map.set(k, v);
    },
    remove: async (k) => {
      map.delete(k);
    },
  };
}

function emptyModel(): Model {
  return {
    slots: SLOT_KINDS.map(() => ({
      id: null,
      removed: false,
      project: undefined,
      lastLocalRev: null,
      lastRelayRev: null,
    })),
    transport: "relay",
    activeProject: undefined,
  };
}

async function setup(): Promise<{ model: Model; real: Sut }> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"] });
  const helpers = host.helpers;
  helpers.seed("Model!B2", [["a short cell of text"]]);
  helpers.seed("Model!B4:D6", [
    ["h1", "h2", "h3"],
    ["1", "2", "3"],
    ["4", "5", "6"],
  ]);
  helpers.seed("Model!B9", [[42]]);
  const chart = helpers.addChart("Model", {
    name: "Chart1",
    chartType: "ColumnClustered",
    title: "T",
    series: [{ name: "S", categories: ["a", "b"], values: [1, 2] }],
  });
  helpers.setActiveChart(chart);

  const relay = new FakeRelay();
  const ws = await createWorkspace(memoryStore());
  const links = await import("../../src/excel/links");

  const real: Sut = {
    helpers,
    relay,
    ws,
    links,
    async reopen() {
      vi.resetModules();
      real.links = await import("../../src/excel/links");
    },
  };
  return { model: emptyModel(), real };
}

function relayFor(model: Model, real: Sut): RelayApi {
  return model.transport === "relay" ? real.relay : new LocalCollector();
}

function slotLabel(slot: number): string {
  return `slot ${String(slot)} (${SLOT_KINDS[slot]})`;
}

// Shorter at every call site below than a bare throw; never returns.
function fail(message: string): never {
  throw new Error(message);
}

// The precondition every command but Export and Reopen shares.
function isLive(m: Readonly<Model>, slot: number): boolean {
  const s = m.slots[slot]!;
  return s.id !== null && !s.removed;
}

// Tags the one call with which space this send just ran in, for verifySlot's
// rev check, then clears it so a later command never inherits it.
async function verifyAfterSend(
  real: Sut,
  slot: number,
  s: SlotModel,
  space: "local" | "relay",
): Promise<void> {
  s.lastSpaceHint = space;
  await verifySlot(real, slot, s);
  s.lastSpaceHint = undefined;
}

// Every command ends here: no duplicate ids; the list's own rev/project for
// this slot match the model, and, once the send's space is known, its rev
// sits in the right space and climbed past this slot's last one there.
async function verifySlot(
  real: Sut,
  slot: number,
  s: SlotModel,
): Promise<void> {
  const rows = await real.links.listWorkbookLinks();
  const ids = rows.map((row) => row.entry.id);
  if (new Set(ids).size !== ids.length) {
    fail(`${slotLabel(slot)}: dup id ${ids.join(",")}`);
  }
  if (s.id === null) return;
  const row = rows.find((one) => one.entry.id === s.id);
  if (s.removed) {
    if (row) fail(`${slotLabel(slot)}: still listed after remove`);
    return;
  }
  if (!row) fail(`${slotLabel(slot)}: missing from the workbook list`);
  if (row.entry.project !== s.project) {
    throw new Error(
      `${slotLabel(slot)}: project mismatch, model=${String(s.project)} list=${String(row.entry.project)}`,
    );
  }
  const space = s.lastSpaceHint;
  if (space === undefined) return;
  const rev = row.entry.rev;
  const local = isLocalRev(rev);
  if (local !== (space === "local")) {
    fail(`${slotLabel(slot)}: ${space} rev in wrong space (${String(rev)})`);
  }
  const last = space === "local" ? s.lastLocalRev : s.lastRelayRev;
  if (last !== null && rev <= last) {
    fail(`${slotLabel(slot)}: ${space} flat ${String(last)}->${String(rev)}`);
  }
  if (space === "local") s.lastLocalRev = rev;
  else s.lastRelayRev = rev;
}

async function doExport(
  real: Sut,
  space: "local" | "relay",
  relay: RelayApi,
  kind: Kind,
  slot: number,
): Promise<string> {
  if (kind === "chart") {
    return (await real.links.exportActiveChart(real.ws, relay)).id;
  }
  real.helpers.select(CELL[slot]!);
  if (kind === "text")
    return (await real.links.exportSelectionAsText(real.ws, relay)).id;
  if (kind === "table")
    return (await real.links.exportSelectionAsTable(real.ws, relay)).id;
  return (await real.links.exportSelection(real.ws, relay)).id;
}

type Cmd = fc.AsyncCommand<Model, Sut>;

function exportCmd(slot: number): Cmd {
  return {
    check: (m) => m.slots[slot]!.id === null && !m.slots[slot]!.removed,
    async run(m, real) {
      const s = m.slots[slot]!;
      const space = m.transport;
      s.id = await doExport(
        real,
        space,
        relayFor(m, real),
        SLOT_KINDS[slot]!,
        slot,
      );
      s.project = m.activeProject;
      await verifyAfterSend(real, slot, s, space);
    },
    toString: () => `Export(${slotLabel(slot)})`,
  };
}

// An explicit Push (pushLinks) or auto-push (pushRegistry, what
// link-watch.ts's flush() calls) - the debounce timing that decides WHEN
// auto-push fires is a separately hunted concern; this checks either keeps
// the registry exactly as consistent as the other.
function pushOneCmd(slot: number, via: "push" | "auto"): Cmd {
  return {
    check: (m) => isLive(m, slot),
    async run(m, real) {
      const s = m.slots[slot]!;
      const space = m.transport;
      const relay = relayFor(m, real);
      const summary =
        via === "push"
          ? await real.links.pushLinks([s.id!], relay)
          : await real.links.pushRegistry([s.id!], relay);
      if (
        summary.pushed !== 1 ||
        summary.failed !== 0 ||
        summary.missing !== 0
      ) {
        fail(`${slotLabel(slot)}: ${via} not ok: ${JSON.stringify(summary)}`);
      }
      await verifyAfterSend(real, slot, s, space);
    },
    toString: () =>
      `${via === "push" ? "Push" : "AutoPush"}(${slotLabel(slot)})`,
  };
}

function pushAllCmd(): Cmd {
  return {
    check: (m) => m.slots.some((s) => s.id !== null && !s.removed),
    async run(m, real) {
      const space = m.transport;
      const targets = m.slots.filter((s) => s.id !== null && !s.removed);
      const summary = await real.links.pushLinks("all", relayFor(m, real));
      if (
        summary.pushed !== targets.length ||
        summary.failed !== 0 ||
        summary.missing !== 0
      ) {
        fail(`push all: ${String(targets.length)},${JSON.stringify(summary)}`);
      }
      for (const [slot, s] of m.slots.entries()) {
        if (s.id !== null && !s.removed)
          await verifyAfterSend(real, slot, s, space);
      }
    },
    toString: () => "PushAll",
  };
}

function removeCmd(slot: number): Cmd {
  return {
    check: (m) => isLive(m, slot),
    async run(m, real) {
      const s = m.slots[slot]!;
      await real.links.removeLink(s.id!, relayFor(m, real));
      s.removed = true;
      await verifySlot(real, slot, s);
    },
    toString: () => `Remove(${slotLabel(slot)})`,
  };
}

// moveLinksToProject also calls withActiveProject (link-projects.ts): the
// project a link just moved to becomes the workbook's active one, so the
// NEXT export joins it too (exportCmd reads m.activeProject).
function moveToProjectCmd(slot: number, project: string | undefined): Cmd {
  return {
    check: (m) => isLive(m, slot),
    async run(m, real) {
      const s = m.slots[slot]!;
      await moveLinksToProject([s.id!], project);
      s.project = project;
      m.activeProject = project;
      await verifySlot(real, slot, s);
    },
    toString: () => `MoveToProject(${slotLabel(slot)}, ${String(project)})`,
  };
}

function transportSwitchCmd(to: "local" | "relay"): Cmd {
  return {
    check: (m) => m.transport !== to,
    run: async (m) => {
      m.transport = to;
    },
    toString: () => `TransportSwitch(${to})`,
  };
}

function sortById<T extends { entry: { id: string } }>(rows: T[]): T[] {
  return [...rows].sort((a, b) => a.entry.id.localeCompare(b.entry.id));
}

// Settings re-read fresh and link-record.ts's per-session trustedThisSession
// memory reset (test/hunt/link.local-rev-floor.hunt2's own mechanism), but
// nothing already sent should be lost or altered.
function reopenCmd(): Cmd {
  return {
    check: () => true,
    async run(_m, real) {
      const before = sortById(await real.links.listWorkbookLinks());
      await real.reopen();
      const after = sortById(await real.links.listWorkbookLinks());
      if (before.length !== after.length) {
        fail(`reopen: count ${String(before.length)}->${String(after.length)}`);
      }
      for (let i = 0; i < before.length; i += 1) {
        const b = before[i]!.entry;
        const a = after[i]!.entry;
        if (a.id !== b.id || a.rev !== b.rev || a.project !== b.project) {
          fail(`reopen: ${JSON.stringify(b)}->${JSON.stringify(a)}`);
        }
      }
    },
    toString: () => "Reopen",
  };
}

function slotArb(): fc.Arbitrary<number> {
  return fc.integer({ min: 0, max: SLOT_COUNT - 1 });
}

const commandsArb = [
  slotArb().map(exportCmd),
  slotArb().map((slot) => pushOneCmd(slot, "push")),
  slotArb().map((slot) => pushOneCmd(slot, "auto")),
  fc.constant(pushAllCmd()),
  slotArb().map(removeCmd),
  fc
    .tuple(slotArb(), fc.constantFrom(...PROJECTS))
    .map(([slot, project]) => moveToProjectCmd(slot, project)),
  fc.constant(transportSwitchCmd("local")),
  fc.constant(transportSwitchCmd("relay")),
  fc.constant(reopenCmd()),
];

describe("registry model: a random sequence of link-lifecycle operations", () => {
  it("keeps revisions climbing per transport space, the list free of duplicates, and reopen exact", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.commands(commandsArb, { maxCommands: MAX_COMMANDS }),
        async (cmds) => {
          await fc.asyncModelRun(setup, cmds);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});
