// Guard: the one place every pane action funnels through - busy state on,
// run the action, refresh, decorate and notify on success, describe and
// notify on failure, then always clear busy. Host-agnostic: Office.js only
// reaches this module through the callbacks the caller supplies.
// Invariant: with `busyMessage` set this is also the pane's re-entrancy latch
// - one action at a time, the second answered rather than interleaved.

import type { ToastKind } from "./toast";

export interface GuardDeps {
  setBusy(busy: boolean): void;
  notify(message: string, kind?: ToastKind, details?: string): void;
  describe(
    error: unknown,
    action: string | undefined,
  ): { message: string; details: string };
  after?(): Promise<void>;
  decorate?(message: string): string;
  finally?(): void;
  // The sentence a second action gets while the first is still in the host.
  // Absent (the Excel pane) leaves the latch off and nothing changes: every
  // caller runs the moment it is called, as it always did.
  busyMessage?: string;
}

export type Guard = (
  run: () => Promise<string>,
  action?: string,
) => Promise<void>;

export function makeGuard(deps: GuardDeps): Guard {
  const busyMessage = deps.busyMessage;
  let running = false;
  return async (run, action) => {
    if (busyMessage !== undefined && running) {
      deps.notify(busyMessage, "error");
      return;
    }
    running = true;
    deps.setBusy(true);
    try {
      const message = await run();
      await deps.after?.();
      deps.notify(deps.decorate ? deps.decorate(message) : message);
    } catch (error) {
      const { message, details } = deps.describe(error, action);
      deps.notify(message, "error", details);
    } finally {
      deps.finally?.();
      deps.setBusy(false);
      running = false;
    }
  };
}
