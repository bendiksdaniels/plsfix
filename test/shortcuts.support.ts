// Shared jsdom scaffolding for the two shortcut-manager suites: the real
// taskpane markup, a served public/shortcuts.json and a stubbed Office.actions.
// Owns no assertions. Invariant: every helper leaves globalThis ready for a
// fresh dynamic import of src/pane/shortcuts-panel.

import { readFileSync } from "node:fs";
import { join } from "node:path";

import { vi } from "vitest";

import type * as ShortcutsPanel from "../src/pane/shortcuts-panel";

export interface ActionsStub {
  getShortcuts: ReturnType<typeof vi.fn>;
  replaceShortcuts: ReturnType<typeof vi.fn>;
  areShortcutsInUse: ReturnType<typeof vi.fn>;
}

/** The map replaceShortcuts is handed: null means "back to the shipped key". */
export type ShortcutMap = Record<string, string | null>;

export const SHORTCUTS_JSON = readFileSync(
  join(process.cwd(), "public/shortcuts.json"),
  "utf8",
);

export const UNSUPPORTED =
  "Custom shortcuts need Microsoft 365 with a signed-in account.";

export function paneRoot(): void {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
}

export function makeActions(): ActionsStub {
  return {
    getShortcuts: vi.fn(async () => ({})),
    replaceShortcuts: vi.fn(async () => undefined),
    areShortcutsInUse: vi.fn(async (keys: string[]) =>
      keys.map((shortcut) => ({ shortcut, inUse: false })),
    ),
  };
}

export function installOffice(
  actions: ActionsStub | null,
  options: { supported?: boolean } = {},
): void {
  const supported = options.supported ?? actions !== null;
  Object.assign(globalThis, {
    Office: {
      context: { requirements: { isSetSupported: () => supported } },
      actions,
    },
  });
}

export function serveShortcuts(body = SHORTCUTS_JSON): void {
  Object.assign(globalThis, {
    fetch: vi.fn(async () => ({
      ok: true,
      status: 200,
      json: async () => JSON.parse(body) as unknown,
    })),
  });
}

export function serveNothing(): void {
  Object.assign(globalThis, {
    fetch: vi.fn(async () => {
      throw new Error("offline");
    }),
  });
}

/** Loads the panel fresh (module state) and installs it over the real markup. */
export async function installPanel(
  onCard: () => void = () => undefined,
): Promise<typeof ShortcutsPanel> {
  const module = await import("../src/pane/shortcuts-panel");
  await module.installShortcutsPanel(document, onCard);
  return module;
}

export function box(id: string): HTMLInputElement {
  const input = document.getElementById(`shortcut-${id}`);
  if (!(input instanceof HTMLInputElement)) throw new Error(`no box ${id}`);
  return input;
}

export function note(): string {
  return (document.getElementById("shortcuts-note")?.textContent ?? "").trim();
}

export function rowCount(): number {
  return document.querySelectorAll("#shortcuts-list label").length;
}

/**
 * What src/ui/tabs.ts does to a panel however the tab was reached: a click, an
 * arrow key, or tabs.activate() from the tool search. The panel may only watch
 * the panel's own visibility, never the tab button's click.
 */
export function showBrandTab(shown = true): void {
  const panel = document.getElementById("view-brand");
  if (panel !== null) panel.hidden = !shown;
}

export async function rejects(run: () => Promise<unknown>): Promise<Error> {
  try {
    await run();
  } catch (error) {
    return error as Error;
  }
  throw new Error("expected a rejection");
}

export async function refusal(run: () => Promise<unknown>): Promise<string> {
  return (await rejects(run)).message;
}
