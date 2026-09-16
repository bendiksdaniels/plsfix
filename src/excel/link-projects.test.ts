// The project adapter over the strict fake Excel host: readProjectState's
// merge of the registry's own project list with what the links carry, and the
// two writers running under link-lock's shared queue so a concurrent export or
// push cannot interleave with them.

import { afterEach, describe, expect, it } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "../../test/fakehost";
import { decodeRegistry, REGISTRY_SETTING } from "../link/model";
import { exclusive, linkQueueStage } from "./link-lock";
import {
  moveLinksToProject,
  readProjectState,
  setActiveProject,
} from "./link-projects";

enableStrictLoadSemantics();
afterEach(() => uninstallFakeHost());

// Long enough for anything that was going to run without waiting on the queue.
async function settle(): Promise<void> {
  for (let turn = 0; turn < 5; turn += 1) {
    await new Promise((done) => setTimeout(done, 0));
  }
}

function boot(): FakeHelpers {
  return installFakeHost({ sheets: ["Model"] }).helpers;
}

/** A minimal RegistryEntry, as JSON: only `id` and `project` ever vary here. */
function rawEntry(id: string, project?: string): Record<string, unknown> {
  return {
    id,
    kind: "range",
    anchor: `PLSFIX_LINK_${id}`,
    label: `Model!A1 (${id})`,
    token: "token",
    createdAt: "2026-01-01T00:00:00.000Z",
    lastPushedAt: null,
    rev: 0,
    ...(project !== undefined ? { project } : {}),
  };
}

function seedRegistry(
  helpers: FakeHelpers,
  registry: Record<string, unknown>,
): void {
  helpers.setSetting(REGISTRY_SETTING, JSON.stringify(registry));
}

function stored(helpers: FakeHelpers) {
  return decodeRegistry(helpers.setting(REGISTRY_SETTING));
}

describe("readProjectState", () => {
  it("merges the registry's own list with the links' names, cleaned and deduped, and reports the active project", async () => {
    const helpers = boot();
    seedRegistry(helpers, {
      v: 1,
      links: [
        rawEntry("a1", "Balcia"),
        rawEntry("a2", "Amasty"),
        rawEntry("a3"),
      ],
      projects: ["Amasty", " Zeta "],
      activeProject: "Amasty",
    });

    const state = await readProjectState();

    expect(state.names).toEqual(["Amasty", "Zeta", "Balcia"]);
    expect(state.active).toBe("Amasty");
  });

  it("reports an empty list and no active project on a workbook with no links", async () => {
    boot();

    expect(await readProjectState()).toEqual({ names: [], active: undefined });
  });
});

describe("setActiveProject", () => {
  it("writes activeProject and adds it to the project list", async () => {
    const helpers = boot();
    seedRegistry(helpers, { v: 1, links: [rawEntry("a1")] });

    await setActiveProject("Amasty");

    const registry = stored(helpers);
    expect(registry.activeProject).toBe("Amasty");
    expect(registry.projects).toContain("Amasty");
  });

  it("deletes activeProject when set to undefined", async () => {
    const helpers = boot();
    seedRegistry(helpers, {
      v: 1,
      links: [rawEntry("a1")],
      activeProject: "Amasty",
    });

    await setActiveProject(undefined);

    expect(stored(helpers).activeProject).toBeUndefined();
  });
});

describe("moveLinksToProject", () => {
  it("sets the project on the named entry only", async () => {
    const helpers = boot();
    seedRegistry(helpers, { v: 1, links: [rawEntry("a1"), rawEntry("a2")] });

    await moveLinksToProject(["a1"], "Balcia");

    const { links } = stored(helpers);
    expect(links.find((entry) => entry.id === "a1")?.project).toBe("Balcia");
    expect(links.find((entry) => entry.id === "a2")?.project).toBeUndefined();
  });

  it("deletes the project off the named entry when moved to undefined", async () => {
    const helpers = boot();
    seedRegistry(helpers, {
      v: 1,
      links: [rawEntry("a1", "Balcia")],
    });

    await moveLinksToProject(["a1"], undefined);

    const { links } = stored(helpers);
    expect(links.find((entry) => entry.id === "a1")?.project).toBeUndefined();
  });
});

describe("the project writers share link-lock's queue", () => {
  it("writes nothing until a held queue resolves, then lands the move", async () => {
    const helpers = boot();
    seedRegistry(helpers, { v: 1, links: [rawEntry("a1")] });

    let release: () => void = () => undefined;
    const deferred = new Promise<void>((resolve) => {
      release = resolve;
    });
    const holding = exclusive("push", () => deferred);
    const moving = moveLinksToProject(["a1"], "Balcia");
    await settle();
    expect(linkQueueStage()).toBe("push");
    // Queued behind the held stage: the registry read never happened yet.
    expect(
      stored(helpers).links.find((entry) => entry.id === "a1")?.project,
    ).toBeUndefined();

    release();
    await holding;
    await moving;

    expect(
      stored(helpers).links.find((entry) => entry.id === "a1")?.project,
    ).toBe("Balcia");
    expect(linkQueueStage()).toBeNull();
  });
});
