// The "New here?" card shown once per machine, dismissed for good with "Got
// it". Host-agnostic like help.ts: no Office.js, so both panes install it the
// same way. Every storage read and write is wrapped in try/catch - a
// locked-down webview can refuse localStorage entirely, and the only safe
// fallback is for the card to simply show again next time.

const DONE = "done";

function hasFinished(key: string): boolean {
  try {
    return localStorage.getItem(key) === DONE;
  } catch {
    return false;
  }
}

function markFinished(key: string): void {
  try {
    localStorage.setItem(key, DONE);
  } catch {
    // Storage can be unavailable in a locked-down webview; the card simply
    // shows again next time, which is the documented fallback, not a bug.
  }
}

/**
 * Shows or hides `#${sectionId}` from `localStorage[key]`, and wires its own
 * "Got it" button (`#${sectionId}-dismiss`). When `onShortcuts` is given, the
 * card's "Shortcut card" button (`#${sectionId}-shortcuts`, if present) calls
 * it instead of being left dead.
 */
export function installFirstRun(
  root: ParentNode,
  key: string,
  sectionId: string,
  onShortcuts?: () => void,
): void {
  const section = root.querySelector<HTMLElement>(`#${sectionId}`);
  if (section === null) return;

  section.hidden = hasFinished(key);

  section
    .querySelector<HTMLButtonElement>(`#${sectionId}-dismiss`)
    ?.addEventListener("click", () => {
      markFinished(key);
      section.hidden = true;
    });

  if (onShortcuts === undefined) return;
  section
    .querySelector<HTMLButtonElement>(`#${sectionId}-shortcuts`)
    ?.addEventListener("click", onShortcuts);
}
