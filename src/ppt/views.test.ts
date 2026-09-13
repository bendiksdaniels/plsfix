// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import type { InboxItem } from "../link/model";
import {
  renderCandidates,
  renderInbox,
  renderLinkRows,
  renderSlideOptions,
  statusLabel,
} from "./views";

function waiting(workbook: string, label = "Model!B4:F12"): InboxItem {
  return {
    id: workbook,
    token: "t",
    kind: "range",
    label,
    src: {
      workbook,
      sheet: "Model",
      ref: "B4:F12",
      anchor: "PLSFIX_LINK_aaaaaaaa",
    },
    createdAt: new Date().toISOString(),
  };
}

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
          kind: "range",
          status: "updateAvailable",
          pushedAt: null,
          selected: false,
        },
      ],
      onToggle,
    );
    expect(body.querySelectorAll("tr")).toHaveLength(1);
    expect(body.textContent).toContain("Update available");
    // The kind rides in the meta column: a table link and a picture of the
    // same cells carry the same label.
    expect(body.querySelector(".link-source")!.textContent).toBe(
      "Model_v4.xlsx · range",
    );
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
          kind: "table",
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
        anchor: "PLSFIX_LINK_aaaaaaaa",
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

describe("renderCandidates", () => {
  it("lists each export by label, workbook, kind and age, in the order given", () => {
    const select = document.createElement("select");

    renderCandidates(select, [waiting("Model_v5.xlsx"), waiting("Old.xlsx")]);

    const options = [...select.options];
    expect(options.map((option) => option.value)).toEqual([
      "Model_v5.xlsx",
      "Old.xlsx",
    ]);
    expect(options[0]!.textContent).toBe(
      "Model!B4:F12 · Model_v5.xlsx · range · just now",
    );
    expect(options[0]!.title).toBe(options[0]!.textContent);
  });

  it("clears stale options on re-render", () => {
    const select = document.createElement("select");
    renderCandidates(select, [waiting("Model_v5.xlsx")]);
    renderCandidates(select, []);
    expect(select.options).toHaveLength(0);
  });
});

describe("labels", () => {
  it("names statuses", () => {
    expect(statusLabel("wrongKey")).toBe("Wrong link key");
  });
});

describe("renderSlideOptions", () => {
  it("lists This slide plus one option per slide, in order", () => {
    const select = document.createElement("select");

    renderSlideOptions(select, 3);

    const options = [...select.options];
    expect(options.map((option) => option.value)).toEqual(["", "1", "2", "3"]);
    expect(options.map((option) => option.textContent)).toEqual([
      "This slide",
      "Slide 1",
      "Slide 2",
      "Slide 3",
    ]);
    expect(select.value).toBe("");
  });

  it("keeps the picked slide across a re-render, and falls back once it is gone", () => {
    const select = document.createElement("select");
    renderSlideOptions(select, 3);
    select.value = "2";

    renderSlideOptions(select, 3);
    expect(select.value).toBe("2");

    renderSlideOptions(select, 1);
    expect(select.value).toBe("");
  });
});
