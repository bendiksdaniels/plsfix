// @vitest-environment jsdom
// installToolSearch pulls in ./shared, whose top-level installTabs/createToast
// calls need #tab-bar and #toast to already exist - and by the time any
// static top-level import runs, a beforeEach has not fired yet. So every
// test loads the real markup first and only then dynamically imports the
// module under test, with the module registry reset in between so each test
// sees a fresh catalogue built off its own fresh DOM.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";

function paneRoot(): Document {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  return document;
}

// No explicit return type: annotating it would need an `import()` type,
// which the lint config forbids in favour of a top-level type import - and
// none of this module's exported types are needed here.
async function load() {
  vi.resetModules();
  paneRoot();
  return import("./tool-search");
}

function input(): HTMLInputElement {
  return document.getElementById("tool-search") as HTMLInputElement;
}

function results(): HTMLElement {
  return document.getElementById("tool-search-results") as HTMLElement;
}

function type(value: string): void {
  const el = input();
  el.value = value;
  el.dispatchEvent(new Event("input"));
}

function press(key: string): void {
  input().dispatchEvent(new KeyboardEvent("keydown", { key }));
}

describe("toolEntries", () => {
  it("finds a known button on its own tab", async () => {
    const { toolEntries } = await load();
    const entries = toolEntries(document);
    const entry = entries.find((e) => e.action === "export-selection");
    expect(entry).toBeDefined();
    expect(entry?.label).toBe("Export selection");
    expect(entry?.tab).toBe("Links");
    expect(entry?.sentence.length).toBeGreaterThan(0);
  });

  it("leaves out a button that is hidden inside its own panel", async () => {
    const { toolEntries } = await load();
    const entries = toolEntries(document);
    expect(entries.some((e) => e.action === "styles-delete")).toBe(false);
  });

  it("still catalogues a button on a tab that is not the active one", async () => {
    const { toolEntries } = await load();
    // Workbook, Links and Brand all start hidden; only Tools is active.
    document.getElementById("view-links")!.hidden = true;
    const entries = toolEntries(document);
    expect(entries.some((e) => e.action === "export-selection")).toBe(true);
  });

  it("catalogues the hygiene cycles and the sheet tools on their own tabs", async () => {
    const { toolEntries } = await load();
    const entries = toolEntries(document);
    const found = (action: string) =>
      entries.find((entry) => entry.action === action);

    expect(found("cycle-indent")).toMatchObject({
      label: "Indent",
      tab: "Tools",
    });
    expect(found("cycle-align")?.tab).toBe("Tools");
    expect(found("cycle-underline")?.tab).toBe("Tools");
    expect(found("sheets-unhide-all")).toMatchObject({
      label: "Unhide all",
      tab: "Workbook",
    });
    expect(found("sheets-move-end")?.tab).toBe("Workbook");
    expect(found("clean-past-data")).toMatchObject({
      label: "Clean past the data",
      tab: "Workbook",
    });
    // Every one of them carries its help sentence into the search row.
    for (const action of [
      "cycle-indent",
      "cycle-align",
      "cycle-underline",
      "sheets-unhide-all",
      "sheets-show-only",
      "sheets-bury",
      "sheets-move-up",
      "sheets-move-down",
      "sheets-move-end",
      "clean-past-data",
    ]) {
      expect([action, found(action)?.sentence.length ?? 0]).not.toEqual([
        action,
        0,
      ]);
    }
  });
});

describe("installToolSearch", () => {
  it("lists a button when its name is typed", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    type("export");

    expect(results().hidden).toBe(false);
    expect(results().textContent).toContain("Export selection");
  });

  it("Enter switches to the button's tab and clicks it", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    const button = document.getElementById(
      "export-selection",
    ) as HTMLButtonElement;
    const onClick = vi.fn();
    button.addEventListener("click", onClick);

    type("export");
    press("Enter");

    expect(onClick).toHaveBeenCalledTimes(1);
    expect(document.getElementById("view-links")!.hidden).toBe(false);
    expect(document.getElementById("view-tools")!.hidden).toBe(true);
    expect(
      document.getElementById("tab-links")!.getAttribute("aria-selected"),
    ).toBe("true");
    expect(input().value).toBe("");
    expect(results().hidden).toBe(true);
  });

  it("a click on a result row runs it the same way Enter does", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    const button = document.querySelector(
      '[data-action="undo"]',
    ) as HTMLButtonElement;
    const onClick = vi.fn();
    button.addEventListener("click", onClick);

    type("undo");
    const row = results().querySelector("li")!;
    row.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    expect(onClick).toHaveBeenCalledTimes(1);
  });

  it("Escape clears the box and hides the list", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    type("export");
    expect(results().hidden).toBe(false);

    press("Escape");

    expect(input().value).toBe("");
    expect(results().hidden).toBe(true);
  });

  it("shows the hint only while empty and focused", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);
    const hint = document.getElementById("tool-search-hint")!;

    input().focus();
    expect(hint.hidden).toBe(false);

    type("export");
    expect(hint.hidden).toBe(true);

    type("");
    expect(hint.hidden).toBe(false);

    input().blur();
    expect(hint.hidden).toBe(true);
  });

  it("ArrowDown/ArrowUp move the active row, clamped at both ends", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    // "waterfall" matches the Waterfall button by label and the EBITDA
    // bridge template by its help sentence, giving more than one row.
    type("waterfall");
    const rows = () => Array.from(results().querySelectorAll("li"));
    expect(rows().length).toBeGreaterThan(1);
    expect(rows()[0]!.getAttribute("aria-selected")).toBe("true");

    press("ArrowDown");
    expect(rows()[0]!.getAttribute("aria-selected")).toBe("false");
    expect(rows()[1]!.getAttribute("aria-selected")).toBe("true");
    expect(input().getAttribute("aria-activedescendant")).toBe(rows()[1]!.id);

    press("ArrowUp");
    expect(rows()[0]!.getAttribute("aria-selected")).toBe("true");

    // Clamped, not wrapped: one more ArrowUp at the top stays put.
    press("ArrowUp");
    expect(rows()[0]!.getAttribute("aria-selected")).toBe("true");

    for (let i = 0; i < rows().length + 2; i += 1) press("ArrowDown");
    const last = rows().length - 1;
    expect(rows()[last]!.getAttribute("aria-selected")).toBe("true");
  });

  it("Enter with an arrow-selected row runs that row, not the first match", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    type("waterfall");
    const targetAction =
      results().querySelectorAll("li")[1]!.dataset.toolAction!;
    const targetButton = document.querySelector<HTMLButtonElement>(
      `[data-action="${targetAction}"]`,
    )!;
    const firstAction =
      results().querySelectorAll("li")[0]!.dataset.toolAction!;
    const firstButton = document.querySelector<HTMLButtonElement>(
      `[data-action="${firstAction}"]`,
    )!;
    const onTargetClick = vi.fn();
    const onFirstClick = vi.fn();
    targetButton.addEventListener("click", onTargetClick);
    firstButton.addEventListener("click", onFirstClick);

    press("ArrowDown");
    press("Enter");

    expect(onTargetClick).toHaveBeenCalledTimes(1);
    expect(onFirstClick).not.toHaveBeenCalled();
  });

  it('"/" focuses the box from elsewhere in the pane, but not from a field', async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    const other = document.getElementById(
      "refresh-selection",
    ) as HTMLButtonElement;
    other.focus();
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "/" }));
    expect(document.activeElement).toBe(input());

    // Typing "/" while already inside a text field must not be hijacked.
    const findQuery = document.getElementById("find-query") as HTMLInputElement;
    findQuery.focus();
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "/" }));
    expect(document.activeElement).toBe(findQuery);
  });
});
