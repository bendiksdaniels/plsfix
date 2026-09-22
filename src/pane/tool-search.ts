// The "Find a tool" search box: catalogues every dispatchable button across
// all four tabs into ToolEntry[] (src/tool-search.ts ranks them), then wires
// the static #tool-search input and #tool-search-results list that
// taskpane.html already carries - typing, arrow keys, Enter/click to run a
// row, Escape, and the "/" shortcut. DOM only: no Office.js, and nothing
// here reaches back into ../main.

import { HELP, helpSections } from "../help/copy";
import { rankTools, type ToolEntry } from "../tool-search";
import { getElement } from "../ui/dom";
import { tabs } from "./shared";

const RESULT_LIMIT = 8;
const SENTENCE_CLIP = 90;

function textOf(element: Element): string {
  return (element.textContent ?? "").replace(/\s+/g, " ").trim();
}

// label = the button's aria-label when its face is a glyph ("↻", "⌨", "⟲",
// "←"), else its <strong> line, else its own text - the same priority
// src/ui/help.ts's labelOf uses for a card's dt, so a glyph-only button is
// found by what it does, not by the character a search box cannot type.
function labelOf(button: HTMLElement): string {
  const ariaLabel = (button.getAttribute("aria-label") ?? "")
    .replace(/\s+/g, " ")
    .trim();
  if (ariaLabel) return ariaLabel;
  const strong = button.querySelector("strong");
  return textOf(strong ?? button);
}

// Section key -> button -> sentence, read straight off the live DOM the same
// way src/ui/help.ts does, so a button key that happens to repeat across
// sections never borrows the wrong section's wording.
function sentencesOf(root: ParentNode): Map<HTMLElement, string> {
  const found = new Map<HTMLElement, string>();
  for (const section of helpSections(root)) {
    const copy = HELP[section.key];
    if (copy === undefined) continue;
    for (const button of section.buttons) {
      const sentence = copy.buttons[button.key];
      if (sentence !== undefined) found.set(button.element, sentence);
    }
  }
  return found;
}

interface TabInfo {
  id: string;
  name: string;
}

// The role=tab whose aria-controls names the tabpanel a button sits in - the
// exact pairing src/ui/tabs.ts reads in the other direction to show a panel.
function tabOf(button: HTMLElement, root: ParentNode): TabInfo | null {
  const panel = button.closest<HTMLElement>('[role="tabpanel"]');
  if (panel === null || panel.id === "") return null;
  const tab = root.querySelector<HTMLElement>(
    `[role="tab"][aria-controls="${panel.id}"]`,
  );
  return tab ? { id: tab.id, name: textOf(tab) } : null;
}

/**
 * Every button a search should find: one per data-action, plus any button
 * keyed by id that the help copy already covers (src/help/copy.ts reads a
 * button's key the same way: data-action, else id). A button on another tab
 * is still catalogued - only one hidden inside its own panel, permanently,
 * is left out.
 */
export function toolEntries(root: ParentNode): ToolEntry[] {
  const sentences = sentencesOf(root);
  const entries: ToolEntry[] = [];
  for (const button of root.querySelectorAll<HTMLButtonElement>("button")) {
    const hasAction = button.dataset.action !== undefined;
    const sentence = sentences.get(button);
    if (!hasAction && sentence === undefined) continue;
    if (button.hidden) continue;
    entries.push({
      action: button.dataset.action ?? button.id,
      label: labelOf(button),
      tab: tabOf(button, root)?.name ?? "",
      sentence: sentence ?? "",
    });
  }
  return entries;
}

function clip(text: string, max: number): string {
  return text.length > max ? `${text.slice(0, max - 1)}…` : text;
}

function findButton(root: Document, action: string): HTMLButtonElement | null {
  const byAction = root.querySelector<HTMLButtonElement>(
    `button[data-action="${action}"]`,
  );
  if (byAction) return byAction;
  const byId = root.getElementById(action);
  return byId instanceof HTMLButtonElement ? byId : null;
}

// Whether a press on this button right now would do anything: not disabled
// (Copy report before a check has run, Back before there is anywhere to go),
// and not sitting inside something hidden for a reason other than its tab
// not being the active one - a tab panel is expected to be hidden and
// runEntry switches to it first, but the #project-prompt Create/Cancel pair
// hides for its own reason and stays dead until "New project" opens it.
function canAct(button: HTMLButtonElement): boolean {
  if (button.disabled) return false;
  for (
    let ancestor = button.parentElement;
    ancestor !== null;
    ancestor = ancestor.parentElement
  ) {
    if (ancestor.hidden && ancestor.getAttribute("role") !== "tabpanel") {
      return false;
    }
  }
  return true;
}

function buildRow(entry: ToolEntry, index: number): HTMLLIElement {
  const li = document.createElement("li");
  li.id = `tool-search-option-${String(index)}`;
  li.setAttribute("role", "option");
  li.setAttribute("aria-selected", "false");
  li.dataset.toolAction = entry.action;

  const label = document.createElement("strong");
  label.textContent = entry.label;
  const badge = document.createElement("span");
  badge.className = "tool-search-tab";
  badge.textContent = entry.tab;
  const sentence = document.createElement("small");
  sentence.textContent = clip(entry.sentence, SENTENCE_CLIP);

  li.append(label, badge, sentence);
  return li;
}

interface SearchState {
  root: Document;
  input: HTMLInputElement;
  list: HTMLUListElement;
  hint: HTMLElement;
  tools: ToolEntry[];
  matches: ToolEntry[];
  activeIndex: number;
}

// Only while the box is both empty and focused - typing hides it, blurring
// hides it, a result list showing has nothing left for it to explain.
function updateHint(state: SearchState): void {
  state.hint.hidden = !(
    state.root.activeElement === state.input && state.input.value.trim() === ""
  );
}

function setActive(state: SearchState, index: number): void {
  const rows = Array.from(state.list.children) as HTMLElement[];
  for (const [i, row] of rows.entries()) {
    row.setAttribute("aria-selected", String(i === index));
  }
  state.activeIndex = index;
  const activeId = rows[index]?.id;
  if (activeId) state.input.setAttribute("aria-activedescendant", activeId);
  else state.input.removeAttribute("aria-activedescendant");
}

function hideResults(state: SearchState): void {
  state.list.hidden = true;
  state.list.textContent = "";
  state.matches = [];
  state.activeIndex = -1;
  state.input.removeAttribute("aria-activedescendant");
}

function renderResults(state: SearchState): void {
  // Ranked fresh off the live DOM every keystroke, not off the catalogue's
  // boot-time snapshot: a button's disabled or hidden-ancestor state moves
  // as the pane is used (Copy report, Back, New project's Create/Cancel),
  // and a row must never outlive what it names.
  const candidates = state.tools.filter((entry) => {
    const button = findButton(state.root, entry.action);
    return button !== null && canAct(button);
  });
  state.matches = rankTools(state.input.value, candidates, RESULT_LIMIT);
  state.list.textContent = "";
  state.matches.forEach((entry, index) =>
    state.list.append(buildRow(entry, index)),
  );
  state.list.hidden = state.matches.length === 0;
  if (state.matches.length > 0) setActive(state, 0);
  else state.activeIndex = -1;
  updateHint(state);
}

// Switches to the row's tab through tabs.activate (never a simulated click
// on the DOM tab), then clicks the real button so every guard, busy state
// and toast applies exactly as a modeller's own click would.
function runEntry(state: SearchState, entry: ToolEntry): void {
  const button = findButton(state.root, entry.action);
  // Re-checked rather than trusted from render time: nothing stops a slow
  // click landing after the button's own state moved on.
  if (button === null || !canAct(button)) return;
  const tab = tabOf(button, state.root);
  if (tab !== null) tabs.activate(tab.id);
  button.click();
  state.input.value = "";
  hideResults(state);
  updateHint(state);
}

function handleKeydown(state: SearchState, event: KeyboardEvent): void {
  if (event.key === "ArrowDown" && state.matches.length > 0) {
    event.preventDefault();
    setActive(state, Math.min(state.activeIndex + 1, state.matches.length - 1));
  } else if (event.key === "ArrowUp" && state.matches.length > 0) {
    event.preventDefault();
    setActive(state, Math.max(state.activeIndex - 1, 0));
  } else if (event.key === "Enter") {
    const entry = state.matches[state.activeIndex] ?? state.matches[0];
    if (entry) {
      event.preventDefault();
      runEntry(state, entry);
    }
  } else if (event.key === "Escape") {
    event.preventDefault();
    state.input.value = "";
    hideResults(state);
    updateHint(state);
  }
}

function handleListClick(state: SearchState, event: MouseEvent): void {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;
  const row = target.closest<HTMLElement>("li[data-tool-action]");
  if (row === null) return;
  const entry = state.matches.find((m) => m.action === row.dataset.toolAction);
  if (entry) runEntry(state, entry);
}

// "/" jumps to the box from anywhere in the pane, unless the hands are
// already in a field - a select's own typeahead, a text input, and the box
// itself all get to keep their "/" character instead of losing it to focus.
function handleSlash(state: SearchState, event: KeyboardEvent): void {
  if (event.key !== "/") return;
  const active = state.root.activeElement;
  const typing =
    active instanceof HTMLInputElement ||
    active instanceof HTMLTextAreaElement ||
    active instanceof HTMLSelectElement ||
    (active instanceof HTMLElement && active.isContentEditable);
  if (typing) return;
  event.preventDefault();
  state.input.focus();
}

/**
 * Wires the search box. Called once boot() has the final DOM in place: the
 * catalogue is built here, not recomputed on every keystroke.
 */
export function installToolSearch(root: Document): void {
  const state: SearchState = {
    root,
    input: getElement<HTMLInputElement>("tool-search"),
    list: getElement<HTMLUListElement>("tool-search-results"),
    hint: getElement<HTMLElement>("tool-search-hint"),
    tools: toolEntries(root),
    matches: [],
    activeIndex: -1,
  };

  state.input.addEventListener("input", () => renderResults(state));
  state.input.addEventListener("focus", () => updateHint(state));
  state.input.addEventListener("blur", () => updateHint(state));
  state.input.addEventListener("keydown", (event) =>
    handleKeydown(state, event),
  );
  state.list.addEventListener("click", (event) =>
    handleListClick(state, event),
  );
  root.addEventListener("keydown", (event) => handleSlash(state, event));
}
