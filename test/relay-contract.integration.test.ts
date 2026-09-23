// test/relay-contract.integration.test.ts
// The same read rules, run on the fake relay and on this computer's store.
import { LOCAL_REV_BASE } from "../src/link/local";
import { memoryPersistence } from "../src/link/local-persist";
import { LocalStore } from "../src/link/local-store";
import { FakeRelay } from "./fakerelay";
import { relayReadContract } from "./relay-contract";

relayReadContract("FakeRelay", async () => {
  const relay = new FakeRelay();
  return {
    relay,
    auth: "auth-1",
    unknown: "missing",
    seed: async (id, blobs) => {
      const revs: number[] = [];
      for (const blob of blobs)
        revs.push((await relay.putLink(id, "auth-1", blob)).rev);
      return revs;
    },
  };
});

relayReadContract("LocalStore", async () => {
  const store = await LocalStore.open(memoryPersistence());
  return {
    relay: store,
    auth: "any",
    unknown: "notPasted",
    seed: async (id, blobs) => {
      const revs = blobs.map((_, index) => LOCAL_REV_BASE + 1 + index);
      for (const [index, blob] of blobs.entries()) {
        await store.ingest(
          {
            links: [{ id, rev: revs[index]!, sentAt: index, blob }],
            inbox: [],
          },
          new Set(),
        );
      }
      return revs;
    },
  };
});
