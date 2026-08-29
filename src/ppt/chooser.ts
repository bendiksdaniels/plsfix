// The "Change source" picker in the PowerPoint pane's Links view: the button
// that opens it, the inline list of candidate exports, and its Confirm /
// Cancel. Nothing is remembered while the picker is open - Confirm reads the
// tick and the inbox again - so a rescan underneath it can never re-point a
// picture the user is no longer looking at. The pane hands in the state to
// read and what to do afterwards; no Office.js and no relay rules live here.

import { sourceLabel, type InboxItem } from "../link/model";
import type { RelayApi } from "../link/relay";
import type { Workspace } from "../link/workspace";
import { getElement } from "../ui/dom";
import {
  candidatesFor,
  changeSource,
  kindWarning,
  pickCandidate,
  requireCandidates,
  requireOneRow,
} from "./change-source";
import type { LinkRow } from "./links";
import { renderCandidates } from "./views";

export interface ChangeSourceDeps {
  // The pane's guarded runner: busy state, toast, error reporting.
  act(run: () => Promise<string>, action: string): void;
  relay: RelayApi;
  // The ticked rows and the waiting exports, read fresh on every click.
  rows(): LinkRow[];
  inbox(): InboxItem[];
  workspace(): Workspace;
  // A line for the toast's details, beside the summary the action returns.
  note(line: string): void;
  // Re-read the inbox and rescan the deck once a source has changed.
  after(): Promise<void>;
}

// Installs the three buttons and returns the one thing the pane must call
// itself: the enabled-state rule, re-applied whenever the ticks or the busy
// state move.
export function installChangeSource(deps: ChangeSourceDeps): () => void {
  const button = getElement<HTMLButtonElement>("change-source");
  const panel = getElement("change-source-chooser");
  const list = getElement<HTMLSelectElement>("change-source-list");

  function open(): Promise<string> {
    deps.workspace();
    const row = requireOneRow(deps.rows());
    renderCandidates(list, requireCandidates(candidatesFor(row, deps.inbox())));
    panel.hidden = false;
    const label = sourceLabel(row.found.tag.src, row.found.tag.kind);
    return Promise.resolve(`Choose the export to point ${label} at.`);
  }

  async function confirm(): Promise<string> {
    const ws = deps.workspace();
    const row = requireOneRow(deps.rows());
    const item = pickCandidate(candidatesFor(row, deps.inbox()), list.value);
    const warning = kindWarning(row, item);
    if (warning !== undefined) deps.note(warning);
    const summary = await changeSource(row, item, ws, deps.relay);
    panel.hidden = true;
    await deps.after();
    return summary;
  }

  function cancel(): Promise<string> {
    panel.hidden = true;
    return Promise.resolve("Change source cancelled. Nothing was re-pointed.");
  }

  const handlers = {
    "change-source": open,
    "change-source-confirm": confirm,
    "change-source-cancel": cancel,
  };
  for (const [id, run] of Object.entries(handlers)) {
    getElement<HTMLButtonElement>(id).addEventListener("click", () => {
      deps.act(run, id);
    });
  }

  return () => {
    button.disabled = deps.rows().length !== 1;
    // A picker left standing over a row that is no longer the only tick would
    // confirm against another link, so it closes with the button.
    if (button.disabled) panel.hidden = true;
  };
}
