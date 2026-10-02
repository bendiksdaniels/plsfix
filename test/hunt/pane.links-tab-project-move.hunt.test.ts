// @vitest-environment jsdom
// Attacks the Links tab's "Move to project" and Push all's project filter
// (src/pane/links-tab.ts's moveSelected/shownPushIds): moveLinksToProject
// asserted with a real project name, and Push all narrowed to the shown
// project rather than every link. Shares its rig via links-tab-project.harness.ts.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  listWorkbookLinks,
  moveLinksToProject,
  pushLinks,
} from "../../src/excel";
import {
  ID_A,
  ID_B,
  chooseProject,
  click,
  harness,
  install,
  row,
  settle,
  tickRow,
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

describe("Move to project", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('refuses with an id ticked, before touching the adapter, when "All projects" is picked', async () => {
    const h = harness();
    vi.mocked(listWorkbookLinks).mockResolvedValue([row(ID_A)]);
    install(h);
    await settle(h);
    tickRow(ID_A);
    chooseProject("all");
    await settle(h);

    click("move-to-project");
    await settle(h);

    expect(h.errors).toContain("Pick a project first, or create one.");
    expect(moveLinksToProject).not.toHaveBeenCalled();
  });

  it('moves the ticked link to "No project" and reports it, clearing the active project too', async () => {
    const h = harness();
    vi.mocked(listWorkbookLinks).mockResolvedValue([row(ID_A, "Amasty")]);
    install(h);
    await settle(h);
    chooseProject("Amasty"); // sets the active project, as picking it in Excel would
    await settle(h);
    tickRow(ID_A);
    chooseProject(""); // now aim the move at "No project"
    await settle(h);

    click("move-to-project");
    await settle(h);

    expect(moveLinksToProject).toHaveBeenCalledWith([ID_A], undefined);
    expect(h.messages).toContain("Moved 1 to No project.");
    expect(h.state.active).toBeUndefined();
  });

  it("moves the ticked link to a named project and reports it", async () => {
    const h = harness();
    vi.mocked(listWorkbookLinks).mockResolvedValue([row(ID_A)]);
    install(h);
    await settle(h);
    tickRow(ID_A);
    chooseProject("Balcia");
    await settle(h);

    click("move-to-project");
    await settle(h);

    expect(moveLinksToProject).toHaveBeenCalledWith([ID_A], "Balcia");
    expect(h.messages).toContain("Moved 1 to Balcia.");
  });

  it("moves every ticked link on a two-link selection, not just the first", async () => {
    const h = harness();
    vi.mocked(listWorkbookLinks).mockResolvedValue([row(ID_A), row(ID_B)]);
    install(h);
    await settle(h);
    tickRow(ID_A);
    tickRow(ID_B);
    chooseProject("Amasty");
    await settle(h);

    click("move-to-project");
    await settle(h);

    expect(moveLinksToProject).toHaveBeenCalledWith([ID_A, ID_B], "Amasty");
    expect(h.messages).toContain("Moved 2 to Amasty.");
  });
});

describe("Push all honours the shown project (shownPushIds)", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(pushLinks).mockResolvedValue({
      pushed: 0,
      missing: 0,
      failed: 0,
      failures: [],
    });
  });

  it("pushes only the links in the chosen project, not every link in the workbook", async () => {
    const h = harness();
    vi.mocked(listWorkbookLinks).mockResolvedValue([
      row(ID_A, "Amasty"),
      row(ID_B, "Balcia"),
    ]);
    install(h);
    await settle(h);
    chooseProject("Amasty");
    await settle(h);

    click("push-all");
    await settle(h);

    expect(pushLinks).toHaveBeenCalledWith([ID_A], h.relay);
  });

  it('pushes only the links with no project when "No project" is shown', async () => {
    const h = harness();
    vi.mocked(listWorkbookLinks).mockResolvedValue([
      row(ID_A),
      row(ID_B, "Amasty"),
    ]);
    install(h);
    await settle(h);
    chooseProject("");
    await settle(h);

    click("push-all");
    await settle(h);

    expect(pushLinks).toHaveBeenCalledWith([ID_A], h.relay);
  });

  it('passes the "all" sentinel through unfiltered when "All projects" is shown', async () => {
    // shownPushIds only narrows to an explicit id list for a NAMED project
    // or "No project"; for "all" it hands the literal sentinel to
    // pushLinks, which does its own "every link" pass server-side of the
    // pane (src/excel/links.ts's pushRegistry: `ids === "all" ? null : ...`).
    const h = harness();
    vi.mocked(listWorkbookLinks).mockResolvedValue([
      row(ID_A, "Amasty"),
      row(ID_B),
    ]);
    install(h);
    await settle(h);
    chooseProject("all");
    await settle(h);

    click("push-all");
    await settle(h);

    expect(pushLinks).toHaveBeenCalledWith("all", h.relay);
  });
});
