// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import { renderInbox, renderLinkRows, statusLabel } from "./views";

describe("renderLinkRows", () => {
  it("renders one row per link with a checkbox and a status badge", () => {
    document.body.innerHTML = "<table><tbody id='rows'></tbody></table>";
    const body = document.getElementById("rows") as HTMLTableSectionElement;
    const onToggle = vi.fn();
    renderLinkRows(
      body,
      [
        {
          key: "s1/sh1",
          slide: 3,
          label: "Model!B4:F12",
          source: "Model_v4.xlsx",
          status: "updateAvailable",
          pushedAt: null,
          selected: false,
        },
      ],
      onToggle,
    );
    expect(body.querySelectorAll("tr")).toHaveLength(1);
    expect(body.textContent).toContain("Update available");
    const badge = body.querySelector(".link-status .badge") as HTMLElement;
    expect(badge.title).toBe("Update available");
    (body.querySelector("input[type=checkbox]") as HTMLInputElement).click();
    expect(onToggle).toHaveBeenCalledWith("s1/sh1", true);
  });
  it("clears stale rows on re-render", () => {
    document.body.innerHTML = "<table><tbody id='rows'></tbody></table>";
    const body = document.getElementById("rows") as HTMLTableSectionElement;
    renderLinkRows(
      body,
      [
        {
          key: "a",
          slide: 1,
          label: "x",
          source: "y",
          status: "current",
          pushedAt: 0,
          selected: true,
        },
      ],
      () => undefined,
    );
    renderLinkRows(body, [], () => undefined);
    expect(body.querySelectorAll("tr")).toHaveLength(0);
  });
});

describe("renderInbox", () => {
  it("lists items with insert buttons", () => {
    const list = document.createElement("div");
    const onInsert = vi.fn();
    const item = {
      id: "a".repeat(32),
      token: "t",
      kind: "range" as const,
      label: "Model!B4:F12",
      src: {
        workbook: "Model_v4.xlsx",
        sheet: "Model",
        ref: "B4:F12",
        anchor: "SMT_LINK_aaaaaaaa",
      },
      createdAt: new Date().toISOString(),
    };
    renderInbox(list, [item], onInsert);
    list.querySelector("button")!.click();
    expect(onInsert).toHaveBeenCalledWith(item);
    renderInbox(list, [], onInsert);
    expect(list.textContent).toContain("Nothing waiting");
  });
});

describe("labels", () => {
  it("names statuses", () => {
    expect(statusLabel("wrongKey")).toBe("Wrong link key");
  });
});
