// A virtual Clock for the debouncer: time only moves when a test says so, and
// the callbacks that fall due run synchronously inside advance(). Keeps the
// three-second auto-push window out of the suite's wall clock.

import type { Clock, Timer } from "../src/link/debounce";

export interface VirtualClock extends Clock {
  advance(ms: number): void;
  waiting(): number;
}

export function virtualClock(): VirtualClock {
  let now = 0;
  let scheduled: { at: number; run: () => void }[] = [];

  return {
    now: () => now,
    after(ms: number, run: () => void): Timer {
      const entry = { at: now + ms, run };
      scheduled.push(entry);
      return {
        cancel: () => {
          scheduled = scheduled.filter((other) => other !== entry);
        },
      };
    },
    advance(ms: number): void {
      now += ms;
      const due = scheduled.filter((entry) => entry.at <= now);
      scheduled = scheduled.filter((entry) => entry.at > now);
      for (const entry of due) entry.run();
    },
    waiting: () => scheduled.length,
  };
}
