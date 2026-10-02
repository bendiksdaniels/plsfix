// test/hunt/links-tab-project.harness.ts
// Shared rig for the Links tab's project flows: a stateful mock of
// readProjectState/setActiveProject/moveLinksToProject that round-trips
// like link-projects.ts, plus the DOM helpers the split test files drive
// it through. vi.mock hoists per file, so each importer mocks
// "../../src/excel" itself before this loads.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { vi } from "vitest";
import {
  moveLinksToProject,
  pushLinks,
  readProjectState,
  setActiveProject,
  type WorkbookLinkRow,
} from "../../src/excel";
import type { RegistryEntry } from "../../src/link/model";
import type { RelayApi } from "../../src/link/relay";
import { TRANSPORT_STORAGE_KEY } from "../../src/link/transport-setting";
import type { KeyStore } from "../../src/link/workspace";
import type { Guard } from "../../src/ui/guard";
import type { Toast, ToastKind } from "../../src/ui/toast";
import { installLinksTab } from "../../src/pane/links-tab";

export const ID_A = "a".repeat(32);
export const ID_B = "b".repeat(32);

export function row(id: string, project?: string): WorkbookLinkRow {
  const entry: RegistryEntry = {
    id,
    kind: "range",
    anchor: `PLSFIX_LINK_${id.slice(0, 8)}`,
    label: "Model!B4:F12",
    token: "token",
    createdAt: "2026-08-29T11:00:00.000Z",
    lastPushedAt: "2026-08-29T11:58:00.000Z",
    rev: 3,
    ...(project !== undefined ? { project } : {}),
  };
  return { entry, source: "ok" };
}

export interface Harness {
  guard: Guard;
  toast: Toast;
  relay: RelayApi;
  keyStore: KeyStore;
  messages: string[];
  errors: string[];
  toasts: { message: string; kind?: ToastKind; details?: string }[];
  pending: Promise<void>[];
  // The real link-projects.ts state this mocked module round-trips through,
  // so chooseProject/moveSelected are driven exactly as the adapter would
  // drive the true one.
  state: { names: string[]; active: string | undefined };
}

export function harness(names: string[] = ["Amasty", "Balcia"]): Harness {
  const messages: string[] = [];
  const errors: string[] = [];
  const toasts: { message: string; kind?: ToastKind; details?: string }[] = [];
  const pending: Promise<void>[] = [];
  const state = { names, active: undefined as string | undefined };
  const stored = new Map<string, string>();
  stored.set(TRANSPORT_STORAGE_KEY, "relay");

  // Mirrors src/link/project.ts's real withActiveProject: setting a new
  // active project also joins the project list (cleaned/trimmed), so a
  // freshly created project's name is there for fillProjectSelect to find
  // an <option> for right after createProject sets the select's value.
  function activate(name: string | undefined): void {
    state.active = name;
    if (name === undefined) return;
    const trimmed = name.trim();
    if (trimmed !== "" && !state.names.includes(trimmed)) {
      state.names = [...state.names, trimmed];
    }
  }

  vi.mocked(pushLinks).mockResolvedValue({
    pushed: 0,
    missing: 0,
    failed: 0,
    failures: [],
  });
  vi.mocked(readProjectState).mockImplementation(async () => ({ ...state }));
  vi.mocked(setActiveProject).mockImplementation(async (name) => {
    activate(name);
  });
  vi.mocked(moveLinksToProject).mockImplementation(async (_ids, name) => {
    // Mirrors src/excel/link-projects.ts's real moveLinksToProject: it also
    // calls withActiveProject(..., project), so a move sets the active
    // project to match, clearing it for a move to "No project".
    activate(name);
  });

  return {
    messages,
    errors,
    toasts,
    pending,
    state,
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

export function paneRoot(): Document {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  return document;
}

export function click(id: string): void {
  document.querySelector<HTMLButtonElement>(`#${id}`)?.click();
}

export function projectSelect(): HTMLSelectElement {
  return document.getElementById("link-project") as HTMLSelectElement;
}

export function chooseProject(value: string): void {
  const select = projectSelect();
  select.value = value;
  select.dispatchEvent(new Event("change"));
}

// linkRow() (src/pane/links-list.ts) puts the tick box inside a plain <tr
// data-link-id> with no id of its own on the box itself.
export function tickRow(id: string): void {
  const rowEl = document.querySelector(`tr[data-link-id="${id}"]`);
  const checkbox = rowEl?.querySelector('input[type="checkbox"]');
  if (!(checkbox instanceof HTMLInputElement)) {
    throw new Error(`no tick box found for row ${id}`);
  }
  checkbox.checked = true;
  checkbox.dispatchEvent(new Event("change"));
}

export async function settle(h: Harness): Promise<void> {
  while (h.pending.length > 0) await Promise.all(h.pending.splice(0));
  await new Promise((resolve) => setTimeout(resolve, 0));
}

export function install(h: Harness): ReturnType<typeof installLinksTab> {
  return installLinksTab({
    guard: h.guard,
    toast: h.toast,
    relay: h.relay,
    keyStore: h.keyStore,
    root: paneRoot(),
  });
}
