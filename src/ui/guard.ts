// Guard: the one place every pane action funnels through - busy state on,
// run the action, refresh, decorate and notify on success, describe and
// notify on failure, then always clear busy. Host-agnostic: Office.js only
// reaches this module through the callbacks the caller supplies.

import type { ToastKind } from "./toast";

export interface GuardDeps {
  setBusy(busy: boolean): void;
  notify(message: string, kind?: ToastKind, details?: string): void;
  describe(error: unknown, action: string | undefined): { message: string; details: string };
  after?(): Promise<void>;
  decorate?(message: string): string;
  finally?(): void;
}

export type Guard = (run: () => Promise<string>, action?: string) => Promise<void>;

export function makeGuard(deps: GuardDeps): Guard {
  return async (run, action) => {
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
    }
  };
}
