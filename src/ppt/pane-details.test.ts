import { describe, expect, it } from "vitest";
import { createPaneDetails } from "./pane-details";

const CONTEXT = { host: "PowerPoint", version: "v9.9.999" };

describe("createPaneDetails", () => {
  it("joins staged lines in the order they were added", () => {
    const details = createPaneDetails(CONTEXT);
    expect(details.value).toBeUndefined();

    details.add("Model!B4:F12: missing");
    details.add("Source changed: a.xlsx -> b.xlsx");

    expect(details.value).toBe(
      "Model!B4:F12: missing\nSource changed: a.xlsx -> b.xlsx",
    );
  });

  it("takes a whole block from an action, and gives it up again", () => {
    const details = createPaneDetails(CONTEXT);
    details.set("1 failed");
    expect(details.value).toBe("1 failed");

    details.set(undefined);
    expect(details.value).toBeUndefined();
  });

  it("names the step and the error's own message on a failure", () => {
    const details = createPaneDetails(CONTEXT);
    details.addFailure(
      new Error("network error"),
      "The list was not refreshed",
    );
    expect(details.value).toBe("The list was not refreshed: network error");
  });

  it("describes something that is not an Error without throwing", () => {
    const details = createPaneDetails(CONTEXT);
    details.addFailure("relay exploded", "The inbox was not read");
    expect(details.value).toBe(
      "The inbox was not read: The add-in could not complete that action.",
    );
  });
});
