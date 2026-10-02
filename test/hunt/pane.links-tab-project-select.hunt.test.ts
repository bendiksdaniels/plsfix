// @vitest-environment jsdom
// Attacks the Links tab's project select and the "new project" prompt
// (src/pane/links-tab.ts's chooseProject/createProject): the "" (No
// project) value and a typed name, both dead in coverage before this.
// Shares its rig with the move file via links-tab-project.harness.ts.

import { beforeEach, describe, expect, it, vi } from "vitest";
import { listWorkbookLinks, setActiveProject } from "../../src/excel";
import {
  chooseProject,
  click,
  harness,
  install,
  projectSelect,
  settle,
} from "./links-tab-project.harness";

vi.mock("../../src/excel", () => ({
  listActiveSheetCharts: vi.fn(async () => []),
  watchActiveSheet: vi.fn(),
  watchWorksheetEdits: vi.fn(),
  exportActiveChart: vi.fn(),
  exportSelection: vi.fn(),
  exportSelectionAsTable: vi.fn(),
  exportSelectionAsText: vi.fn(),
  goToSource: vi.fn(),
  listWorkbookLinks: vi.fn(),
  pushLinks: vi.fn(),
  removeLink: vi.fn(),
  touchWorkbookLinks: vi.fn(async () => 0),
  readProjectState: vi.fn(async () => ({ names: [], active: undefined })),
  setActiveProject: vi.fn(async () => undefined),
  moveLinksToProject: vi.fn(async () => undefined),
  restoreAutoPush: vi.fn(async () => false),
  setAutoPush: vi.fn(async () => undefined),
  restoreLinkHighlight: vi.fn(async () => false),
  toggleLinkHighlight: vi.fn(async () => false),
}));

describe("choosing a project in the select", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(listWorkbookLinks).mockResolvedValue([]);
  });

  it('sets no active project and says so when "No project" is chosen', async () => {
    const h = harness();
    install(h);
    await settle(h);

    chooseProject("");
    await settle(h);

    expect(setActiveProject).toHaveBeenCalledWith(undefined);
    expect(h.messages).toContain("No project for new exports.");
  });

  it("sets the named active project and says which one", async () => {
    const h = harness();
    install(h);
    await settle(h);

    chooseProject("Amasty");
    await settle(h);

    expect(setActiveProject).toHaveBeenCalledWith("Amasty");
    expect(h.messages).toContain("Project: Amasty");
  });

  it('"All projects" is the same as "No project" for what new exports join', async () => {
    const h = harness();
    install(h);
    await settle(h);

    chooseProject("all");
    await settle(h);

    expect(setActiveProject).toHaveBeenCalledWith(undefined);
    expect(h.messages).toContain("No project for new exports.");
  });

  it("cycles cleanly across three repeated choices: named, no project, named again", async () => {
    const h = harness();
    install(h);
    await settle(h);

    chooseProject("Amasty");
    await settle(h);
    chooseProject("");
    await settle(h);
    chooseProject("Balcia");
    await settle(h);

    expect(vi.mocked(setActiveProject).mock.calls).toEqual([
      ["Amasty"],
      [undefined],
      ["Balcia"],
    ]);
    expect(h.messages.slice(-3)).toEqual([
      "Project: Amasty",
      "No project for new exports.",
      "Project: Balcia",
    ]);
  });
});

describe("creating a project from the prompt", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(listWorkbookLinks).mockResolvedValue([]);
  });

  function projectNameInput(): HTMLInputElement {
    return document.getElementById("project-name") as HTMLInputElement;
  }

  it("refuses a whitespace-only name without calling setActiveProject", async () => {
    const h = harness();
    install(h);
    await settle(h);

    click("new-project");
    await settle(h);
    projectNameInput().value = "   ";
    click("project-ok");
    await settle(h);

    expect(h.errors).toContain("Type a project name.");
    expect(setActiveProject).not.toHaveBeenCalled();
  });

  it("trims the name, sets it active and reports it", async () => {
    const h = harness();
    install(h);
    await settle(h);

    click("new-project");
    await settle(h);
    projectNameInput().value = "  Delfin31  ";
    click("project-ok");
    await settle(h);

    expect(setActiveProject).toHaveBeenCalledWith("  Delfin31  ");
    expect(h.messages).toContain("Project: Delfin31");
    expect(projectSelect().value).toBe("Delfin31");
  });

  it("cancel hides the prompt and clears the typed name without calling setActiveProject", async () => {
    const h = harness();
    install(h);
    await settle(h);

    click("new-project");
    await settle(h);
    projectNameInput().value = "Half-typed";
    click("project-cancel");
    await settle(h);

    expect(setActiveProject).not.toHaveBeenCalled();
    expect(document.getElementById("project-prompt")).toHaveProperty(
      "hidden",
      true,
    );
    expect(projectNameInput().value).toBe("");
  });
});
