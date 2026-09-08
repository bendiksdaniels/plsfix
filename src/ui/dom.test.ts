// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { getElement } from "./dom";

describe("getElement", () => {
  it("returns the element by id", () => {
    document.body.innerHTML = '<div id="thing"></div>';
    expect(getElement("thing")).toBe(document.getElementById("thing"));
  });

  it("throws by name instead of handing back null", () => {
    document.body.innerHTML = "";
    expect(() => getElement("missing")).toThrow("Missing element #missing");
  });
});
