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
