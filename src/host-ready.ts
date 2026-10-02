// When the pane may talk to its host. Office.onReady is the signal; on Excel
// for the web it never settles when the custom-functions runtime fails to
// initialise, although the Excel API already answers. So once onReady has had
// its head start, a host that answers a real API probe counts as ready too,
// flagged as degraded so the pane can say what it noticed. Pure: timing and
// the probe are injected.

export interface HostReadiness<H> {
  host: H;
  /** True when the host answered the probe while Office.onReady never did. */
  degraded: boolean;
}

export interface HostReadyOptions<H> {
  /** The host as Office.onReady names it. */
  onReady: Promise<H>;
  /** The host that answered a real API call, or null when none did. */
  probe: () => Promise<H | null>;
  /** onReady's head start before the first probe. */
  headStartMs: number;
  probeEveryMs: number;
  /** Past this the probing stops and only onReady can end the wait. */
  giveUpAfterMs: number;
  sleep: (ms: number) => Promise<void>;
}

export async function hostReady<H>(
  options: HostReadyOptions<H>,
): Promise<HostReadiness<H>> {
  let settled = false;
  const ready = options.onReady.then((host) => {
    settled = true;
    return { host, degraded: false };
  });
  const probed = probeUntil(options, () => settled).then((host) =>
    host === null ? ready : { host, degraded: true },
  );
  return Promise.race([ready, probed]);
}

async function probeUntil<H>(
  options: HostReadyOptions<H>,
  done: () => boolean,
): Promise<H | null> {
  await options.sleep(options.headStartMs);
  for (
    let waited = options.headStartMs;
    !done() && waited <= options.giveUpAfterMs;
    waited += options.probeEveryMs
  ) {
    const host = await options.probe().catch(() => null);
    if (host !== null) return host;
    await options.sleep(options.probeEveryMs);
  }
  return null;
}
