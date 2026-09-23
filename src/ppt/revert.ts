// "Revert last update" for the PowerPoint pane: PowerPoint cannot undo what an
// add-in wrote, and the relay keeps two revisions, so a deck that dislikes an
// update goes back exactly one step. Each row is repainted from the revision
// below the one its tag holds, and that rev is written back into the tag, so
// the newer one shows up as available again on the next scan. No Office.js -
// the host adapter and the relay are both injectable.

import { deriveLinkKeys, open } from "../link/crypto";
import { previousRevOf } from "../link/local";
import { decodePayload, sourceLabel, type Payload } from "../link/model";
import { isRelayError, type RelayApi } from "../link/relay";
import type { FoundLink, RefreshRequest } from "./host";
import * as realHost from "./host";
import { applyBatch, failureLine, type LinkRow, type PptHost } from "./links";

export interface RevertSummary {
  reverted: number;
  noPrevious: number;
  failed: number;
  failures: string[];
}

// The relay (or this computer's store) holds the tag's revision and the one
// before it, so a link that was never updated - or one already reverted - has
// nowhere left to go. Local revs stop at the first one of their own space.
function previousRev(found: FoundLink): number | null {
  return previousRevOf(found.tag.rev);
}

async function fetchRevision(
  found: FoundLink,
  rev: number,
  relay: RelayApi,
): Promise<Payload> {
  const keys = await deriveLinkKeys(found.token);
  const result = await relay.getLinkRev(found.tag.id, keys.auth, rev);
  return decodePayload(await open(keys.enc, found.tag.id, result.blob));
}

// One row's failure never stops the rest, and the repaints the fetches earn
// all travel together - the update path's rules, because a revert is the same
// repaint with an older picture.
export async function revertLinks(
  rows: LinkRow[],
  relay: RelayApi,
  host: PptHost = realHost,
): Promise<RevertSummary> {
  const summary: RevertSummary = {
    reverted: 0,
    noPrevious: 0,
    failed: 0,
    failures: [],
  };
  const batch: RefreshRequest[] = [];
  for (const row of rows) {
    const rev = previousRev(row.found);
    if (rev === null) {
      summary.noPrevious += 1;
      continue;
    }
    try {
      const payload = await fetchRevision(row.found, rev, relay);
      batch.push({ found: row.found, payload, rev });
    } catch (error) {
      countFailure(summary, row.found, rev, error);
    }
  }
  summary.reverted += await applyBatch(batch, host, (found, error) => {
    countFailure(summary, found, found.tag.rev - 1, error);
  });
  return summary;
}

// A revision the relay dropped (a third push, or the seven-day TTL) is the one
// failure worth its own sentence: nothing is wrong with the deck or the key,
// the picture is simply gone, and the row still shows the update it holds.
function countFailure(
  summary: RevertSummary,
  found: FoundLink,
  rev: number,
  error: unknown,
): void {
  summary.failed += 1;
  const gone = isRelayError(error) && error.kind === "missing";
  summary.failures.push(
    gone
      ? `revert ${sourceLabel(found.tag.src, found.tag.kind)}: the previous version is no longer available.`
      : failureLine(found, error),
  );
}

export function summarizeRevert(summary: RevertSummary): string {
  const parts = [
    [summary.reverted, "reverted"],
    [summary.noPrevious, "without a previous version"],
    [summary.failed, "failed"],
  ] as const;
  const text = parts
    .filter(([count]) => count > 0)
    .map(([count, label]) => `${String(count)} ${label}`)
    .join(", ");
  return text || "Nothing to revert";
}
