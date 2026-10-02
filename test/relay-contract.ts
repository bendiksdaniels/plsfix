// test/relay-contract.ts
// The read rules every RelayApi the PowerPoint pane runs against must keep:
// FakeRelay (the server's rules) and LocalStore (local mode). A rule kept by
// one and not the other is a transport the deck cannot trust.
import { describe, expect, it } from "vitest";
import { isRelayError, type RelayApi } from "../src/link/relay";
import { deriveStatus } from "../src/link/status";

export interface ContractTransport {
  relay: RelayApi;
  auth: string;
  // Stores the blobs as consecutive revisions of one link, oldest first, and
  // answers the revs it gave them.
  seed(id: string, blobs: Uint8Array[]): Promise<number[]>;
  unknown: "missing" | "notPasted";
}

const ID = "0123456789abcdef0123456789abcdef";
const OTHER = "fedcba9876543210fedcba9876543210";
const [ONE, TWO, THREE] = [1, 2, 3].map((n) => new Uint8Array([n]));

export function relayReadContract(
  name: string,
  make: () => Promise<ContractTransport>,
): void {
  describe(`${name}: the read rules the deck relies on`, () => {
    it("answers the newest revision in status and fetch", async () => {
      const t = await make();
      const [, rev2] = await t.seed(ID, [ONE!, TWO!]);
      const [status] = await t.relay.status([{ id: ID, auth: t.auth }]);
      expect(status?.rev).toBe(rev2);
      const fetched = await t.relay.fetchLinks([{ id: ID, auth: t.auth }]);
      expect(fetched.items).toEqual([{ id: ID, rev: rev2, blob: TWO }]);
    });

    it("omits a link whose newest revision the deck already holds", async () => {
      const t = await make();
      const [rev1] = await t.seed(ID, [ONE!]);
      const fetched = await t.relay.fetchLinks([
        { id: ID, auth: t.auth, knownRev: rev1 },
      ]);
      expect(fetched).toEqual({
        items: [],
        omitted: [{ id: ID, reason: "unchanged" }],
      });
    });

    it("reaches the revision before the newest and nothing older", async () => {
      const t = await make();
      const [rev1, rev2] = await t.seed(ID, [ONE!, TWO!, THREE!]);
      await expect(t.relay.getLinkRev(ID, t.auth, rev2!)).resolves.toEqual({
        rev: rev2,
        blob: TWO,
      });
      const gone = await t.relay
        .getLinkRev(ID, t.auth, rev1!)
        .catch((error: unknown) => error);
      expect(isRelayError(gone) && gone.kind === "missing").toBe(true);
    });

    it("says why an unknown link has no picture", async () => {
      const t = await make();
      const fetched = await t.relay.fetchLinks([{ id: OTHER, auth: t.auth }]);
      expect(fetched.omitted).toEqual([{ id: OTHER, reason: t.unknown }]);
      const [status] = await t.relay.status([{ id: OTHER, auth: t.auth }]);
      expect(deriveStatus(3, status)).toBe(t.unknown);
    });
  });
}
