// The Links tab's two workbook-wide tick boxes: auto-push on edit and the
// linked-cell highlight. Both are told by the workbook when the pane boots, and
// both show what the adapter reports rather than what the click asked for, so a
// refusal puts the box back before the guard says why. The tab owns the guard
// and the busy state; Office.js only reaches here through src/excel.

import {
  restoreAutoPush,
  restoreLinkHighlight,
  setAutoPush,
  toggleLinkHighlight,
} from "../excel";
import type { RelayApi } from "../link/relay";
import type { Toast } from "../ui/toast";

export interface Toggles {
  autopush: HTMLInputElement;
  highlight: HTMLInputElement;
  relay: RelayApi;
  toast: Toast;
}

const HIGHLIGHT_CLEARED =
  "Linked-cell highlight from the last session was cleared.";

export async function restoreToggles(toggles: Toggles): Promise<void> {
  await restoreWatch(toggles);
  await restoreHighlight(toggles);
}

export function setTogglesBusy(toggles: Toggles, busy: boolean): void {
  toggles.autopush.disabled = busy;
  toggles.highlight.disabled = busy;
}

// Re-arming reads a workbook setting, which a pane that cannot reach Excel yet
// has no way to do: the box then simply stays clear, and ticking it reports
// through the guard like every other action.
async function restoreWatch(toggles: Toggles): Promise<void> {
  try {
    toggles.autopush.checked = await restoreAutoPush(
      toggles.relay,
      notify(toggles),
    );
  } catch {
    toggles.autopush.checked = false;
  }
}

// The tint is saved with the file while the snapshot that undoes it dies with
// the pane, so last session's fills go back before they can be mistaken for
// model formatting: the box always boots clear.
async function restoreHighlight(toggles: Toggles): Promise<void> {
  toggles.highlight.checked = false;
  try {
    if (await restoreLinkHighlight()) toggles.toast.show(HIGHLIGHT_CLEARED);
  } catch {
    return;
  }
}

// Auto-push runs on an edit, not on a click, so its own messages go straight to
// the toast rather than through the guard.
export function notify(toggles: Toggles): (message: string) => void {
  return (message) => {
    toggles.toast.show(message);
  };
}

export async function toggleAutoPush(toggles: Toggles): Promise<string> {
  const on = toggles.autopush.checked;
  try {
    await setAutoPush(on, toggles.relay, notify(toggles));
  } catch (error) {
    toggles.autopush.checked = !on;
    throw error;
  }
  return on
    ? "Auto-push on: linked pictures follow your edits."
    : "Auto-push off.";
}

// The adapter is the toggle itself, so the box takes the state it hands back; a
// refusal - the audit overlay owns the fills, or the linked ranges are past the
// cap - leaves the box where it was.
export async function toggleHighlight(toggles: Toggles): Promise<string> {
  const wanted = toggles.highlight.checked;
  try {
    toggles.highlight.checked = await toggleLinkHighlight();
  } catch (error) {
    toggles.highlight.checked = !wanted;
    throw error;
  }
  return toggles.highlight.checked
    ? "Linked cells highlighted."
    : "Highlight off: the original fills are back.";
}
