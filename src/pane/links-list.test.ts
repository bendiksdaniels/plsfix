// @vitest-environment jsdom
// Render tests for the Links list: one row per registry entry, the missing
// badge, the push time and the checkbox callback. The tab that drives it is
// covered in links-tab.test.ts.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { WorkbookLinkRow } from "../excel";
import type { RegistryEntry } from "../link/model";
import { renderWorkbookLinks } from "./links-list";

const ID_A = "a".repeat(32);
const ID_B = "b".repeat(32);
const NOW = "2026-08-29T12:00:00.000Z";

function tbody(): HTMLTableSectionElement {
  document.body.innerHTML = "<table><tbody id='rows'></tbody></table>";
  return document.getElementById("rows") as HTMLTableSectionElement;
}

function row(
  id: string,
  entry: Partial<RegistryEntry> = {},
  source: "ok" | "missing" = "ok",
): WorkbookLinkRow {
  return {
    entry: {
      id,
      kind: "range",
      anchor: `PLSFIX_LINK_${id.slice(0, 8)}`,
      label: "Model!B4:F12",
      token: "token",
      createdAt: "2026-08-29T11:00:00.000Z",
      lastPushedAt: "2026-08-29T11:58:00.000Z",
      rev: 3,
      ...entry,
    },
    source,
  };
}

function boxes(body: HTMLTableSectionElement): HTMLInputElement[] {
  return Array.from(
    body.querySelectorAll<HTMLInputElement>("input[type=checkbox]"),
  );
}

describe("renderWorkbookLinks", () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date(NOW));
  });
  afterEach(() => {
    vi.useRealTimers();
  });

  it("groups rows under their project name", () => {
    const body = tbody();
    renderWorkbookLinks(
      body,
      [
        row(ID_A, { project: "Amasty" }),
        row(ID_B, { project: "Balcia" }),
        row("c".repeat(32)),
      ],
      new Set(),
      () => undefined,
    );
    const headers = Array.from(body.querySelectorAll(".wl-project")).map(
      (cell) => cell.textContent,
    );
    expect(headers).toEqual(["Amasty", "Balcia", "No project"]);
  });

  it("renders one row per entry with its label and kind", () => {
    const body = tbody();
    renderWorkbookLinks(
      body,
      [row(ID_A), row(ID_B, { label: "Chart 1", kind: "chart" })],
      new Set(),
      () => undefined,
    );

    const rows = body.querySelectorAll("tr[data-link-id]");
    expect(rows).toHaveLength(2);
    expect(rows[0]?.textContent).toContain("Model!B4:F12");
    expect(rows[0]?.textContent).toContain("Picture · Model!B4:F12");
    expect(rows[1]?.textContent).toContain("Chart · Chart 1");
  });

  it("moves the internal anchor id to the row's title, out of the visible text", () => {
    const body = tbody();
    renderWorkbookLinks(body, [row(ID_A)], new Set(), () => undefined);

    const tr = body.querySelector("tr[data-link-id]");
    expect(tr?.getAttribute("title")).toBe(`PLSFIX_LINK_${ID_A.slice(0, 8)}`);
    expect(tr?.textContent).not.toContain("PLSFIX_LINK");
  });

  it("shows the kind and where it points on the second line, per kind", () => {
    const body = tbody();
    const ID_C = "c".repeat(32);
    const ID_D = "d".repeat(32);
    renderWorkbookLinks(
      body,
      [
        row(ID_A, { kind: "table", label: "P&L!A10:H24 table" }),
        row(ID_B, { kind: "text", label: "P&L!A8 text" }),
        row(ID_C, { kind: "chart", label: "Bridge: Revenue chart" }),
        row(ID_D, { kind: "range", label: "Data!A10:G15" }),
      ],
      new Set(),
      () => undefined,
    );

    const rows = body.querySelectorAll("tr[data-link-id]");
    expect(rows[0]?.textContent).toContain("Table · P&L!A10:H24");
    expect(rows[1]?.textContent).toContain("Text · P&L!A8");
    expect(rows[2]?.textContent).toContain("Chart · Bridge: Revenue chart");
    expect(rows[3]?.textContent).toContain("Picture · Data!A10:G15");
  });

  it("spells the push time in minutes, hours and days", () => {
    const body = tbody();
    renderWorkbookLinks(
      body,
      [
        row(ID_A, { lastPushedAt: "2026-08-29T11:58:00.000Z" }),
        row(ID_B, { lastPushedAt: "2026-08-29T07:00:00.000Z" }),
        row("c".repeat(32), { lastPushedAt: "2026-08-26T12:00:00.000Z" }),
      ],
      new Set(),
      () => undefined,
    );

    const rows = body.querySelectorAll("tr[data-link-id]");
    expect(rows[0]?.textContent).toContain("Pushed 2 min ago");
    expect(rows[1]?.textContent).toContain("Pushed 5 h ago");
    expect(rows[2]?.textContent).toContain("Pushed 3 days ago");
  });

  it("badges a missing source and says never for a link never pushed", () => {
    const body = tbody();
    renderWorkbookLinks(
      body,
      [row(ID_A, { lastPushedAt: null }, "missing"), row(ID_B)],
      new Set(),
      () => undefined,
    );

    const rows = body.querySelectorAll("tr[data-link-id]");
    expect(rows[0]?.textContent).toContain("Source missing");
    expect(rows[0]?.textContent).toContain("never");
    expect(rows[1]?.textContent).not.toContain("Source missing");
  });

  it("checks the selected rows and reports every toggle", () => {
    const body = tbody();
    const onToggle = vi.fn();
    renderWorkbookLinks(
      body,
      [row(ID_A), row(ID_B)],
      new Set([ID_A]),
      onToggle,
    );

    const [first, second] = boxes(body);
    expect(first?.checked).toBe(true);
    expect(second?.checked).toBe(false);

    second?.click();
    expect(onToggle).toHaveBeenCalledWith(ID_B, true);
    first?.click();
    expect(onToggle).toHaveBeenCalledWith(ID_A, false);
  });

  it("clears stale rows and shows an empty state", () => {
    const body = tbody();
    renderWorkbookLinks(body, [row(ID_A)], new Set(), () => undefined);
    renderWorkbookLinks(body, [], new Set(), () => undefined);

    expect(body.querySelectorAll("tr[data-link-id]")).toHaveLength(0);
    expect(body.textContent).toContain("No linked objects");
  });
});
