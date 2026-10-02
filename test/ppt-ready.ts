// Waits for PowerPoint pane test work to finish instead of guessing a fixed
// number of setTimeout(0) turns. Owns trackPptBoot (follows Office.onReady's
// own callback in src/ppt/main.ts to its actual settlement) and settlePpt
// (drains real event-loop turns, then keeps draining while the busy latch
// setBusy() holds). Invariant: a PowerPoint pane test waits for the work,
// never for a count of turns.

const BUSY_POLL_LIMIT = 2000;

interface ReadyInfo {
  host: string;
}
type ReadyCallback = (info: ReadyInfo) => unknown;
interface FakeOffice {
  onReady: (callback?: ReadyCallback) => unknown;
}

function fakeOffice(): FakeOffice {
  return (globalThis as unknown as { Office: FakeOffice }).Office;
}

// Call once installFakePpt(...) has put Office on globalThis and before the
// pane module is imported - importing it is what actually calls
// Office.onReady. Wraps that call so the promise this returns settles once
// the pane's own onReady callback (main.ts's async boot sequence, which
// awaits WebCrypto key derivations whose latency grows with the suite's
// load) has resolved or rejected, rather than guessing how many turns it
// needs.
export function trackPptBoot(): Promise<void> {
  const office = fakeOffice();
  const original = office.onReady;
  return new Promise<void>((resolve) => {
    const done = (): void => resolve();
    office.onReady = (callback?: ReadyCallback) =>
      original((info) => {
        const result = callback?.(info);
        // .then(done, done), never a bare await here: a boot that rejects
        // must still let the test proceed, and the rejection must stay
        // handled rather than becoming an unhandled rejection.
        void Promise.resolve(result).then(done, done);
        return result;
      });
  });
}

async function turn(): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, 0));
}

// True exactly when src/ppt/main.ts's setBusy(true) is holding the pane: it
// disables every ".app-shell button", the tab strip included, and nothing
// else ever disables a tab. No tabs painted yet reads as not busy, not as
// vacuously busy.
function pptBusy(): boolean {
  const tabs = document.querySelectorAll<HTMLButtonElement>("[role=tab]");
  return tabs.length > 0 && [...tabs].every((tab) => tab.disabled);
}

// The fixed turns most actions finish inside, then kept turning while the
// busy latch still holds - bounded, so a latch a test deliberately leaves
// stuck (a relay call mocked to hang forever) still returns instead of
// hanging the run.
export async function settlePpt(rounds = 12): Promise<void> {
  for (let round = 0; round < rounds; round += 1) await turn();
  for (let extra = 0; extra < BUSY_POLL_LIMIT && pptBusy(); extra += 1) {
    await turn();
  }
}
