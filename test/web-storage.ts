// web-storage.ts: a fresh machine's web storage for a simulated pane boot. Under Node 22 (CI) the
// jsdom `localStorage` works and keeps what an earlier boot wrote, such as the first-run card's
// dismissal; Node 25's own global `localStorage` has no working methods without a storage file.
// Clearing at every boot makes both start empty. Invariant: never throws.
export function freshWebStorage(): void {
  try {
    globalThis.localStorage.clear();
  } catch {
    // A storage without working methods holds nothing to clear.
  }
}
