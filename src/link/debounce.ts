// A debouncer over string keys with an injectable clock: every touch pushes a
// key's deadline out by the delay, and the keys that fall due together are
// handed to one flush call. Auto-push keys it by worksheet, so a burst of
// typing on two sheets still ends in a single push. The clock is the only I/O
// this module has, which is what makes it testable without waiting.

export interface Timer {
  cancel(): void;
}

export interface Clock {
  now(): number;
  after(ms: number, run: () => void): Timer;
}

export function systemClock(): Clock {
  return {
    now: () => Date.now(),
    after: (ms, run) => {
      const handle = setTimeout(run, ms);
      return {
        cancel: () => {
          clearTimeout(handle);
        },
      };
    },
  };
}

export class Debouncer {
  private readonly delay: number;
  private readonly flush: (keys: string[]) => void;
  private readonly clock: Clock;
  private readonly deadlines = new Map<string, number>();
  private timer: Timer | null = null;

  constructor(
    delay: number,
    flush: (keys: string[]) => void,
    clock: Clock = systemClock(),
  ) {
    this.delay = delay;
    this.flush = flush;
    this.clock = clock;
  }

  touch(keys: Iterable<string>): void {
    const due = this.clock.now() + this.delay;
    for (const key of keys) this.deadlines.set(key, due);
    this.arm();
  }

  // Everything waiting is dropped: switching auto-push off must not fire a push
  // three seconds later.
  cancel(): void {
    this.deadlines.clear();
    this.timer?.cancel();
    this.timer = null;
  }

  pending(): string[] {
    return [...this.deadlines.keys()];
  }

  // One timer, set to the nearest deadline. A clock that fires early leaves
  // nothing due, and the re-arm below simply waits out the remainder.
  private arm(): void {
    this.timer?.cancel();
    this.timer = null;
    let earliest: number | null = null;
    for (const deadline of this.deadlines.values()) {
      if (earliest === null || deadline < earliest) earliest = deadline;
    }
    if (earliest === null) return;
    const wait = Math.max(0, earliest - this.clock.now());
    this.timer = this.clock.after(wait, () => {
      this.fire();
    });
  }

  private fire(): void {
    this.timer = null;
    const now = this.clock.now();
    const due: string[] = [];
    for (const [key, deadline] of this.deadlines) {
      if (deadline <= now) {
        due.push(key);
        this.deadlines.delete(key);
      }
    }
    this.arm();
    if (due.length > 0) this.flush(due);
  }
}
