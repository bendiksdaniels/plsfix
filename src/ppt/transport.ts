// src/ppt/transport.ts
// Which RelayApi and workspace the PowerPoint pane's flows run against: the
// relay client always, or this computer's LocalStore once local mode has
// been used. Owns the per-device mode and the lazily opened, kept-open local
// store. No Office.js, no DOM: main.ts wires it to the Settings select.
// Invariant: relay() and workspace() never await; local mode's store is open
// by the time either is asked to answer for it.

import { localWorkspace } from "../link/local";
import {
  openLocalPersistence,
  type LocalPersistence,
} from "../link/local-persist";
import { LocalStore } from "../link/local-store";
import type { RelayApi } from "../link/relay";
import {
  loadTransport,
  saveTransport,
  type LinkTransport,
} from "../link/transport-setting";
import type { KeyStore, Workspace } from "../link/workspace";

export interface PptTransport {
  mode(): LinkTransport;
  // The RelayClient in relay mode, or this computer's LocalStore in local
  // mode (open by the time this can be asked, per the invariant above).
  relay(): RelayApi;
  // The paired key in relay mode (null before pairing), or the fixed local
  // workspace, which is always there.
  workspace(): Workspace | null;
  // The local store once opened, whichever mode is current now; null until
  // local mode has been used at least once this session.
  store(): LocalStore | null;
  setMode(mode: LinkTransport): Promise<void>;
}

export interface PptTransportDeps {
  keyStore: KeyStore;
  remote: RelayApi;
  // The relay pairing main.ts already tracks (loadWorkspace/saveKey/
  // forgetKey); read fresh, never cached, so a pair or a forget after boot
  // is seen immediately.
  paired: () => Workspace | null;
  persistence?: () => Promise<LocalPersistence>;
}

// What main.ts paints with before Office.onReady has opened the real one:
// relay mode, unpaired, no store. Never reached by a button (ready is false).
export function preBootTransport(remote: RelayApi): PptTransport {
  return {
    mode: () => "relay",
    relay: () => remote,
    workspace: () => null,
    store: () => null,
    setMode: async () => undefined,
  };
}

export async function openPptTransport(
  deps: PptTransportDeps,
): Promise<PptTransport> {
  const openPersistence = deps.persistence ?? openLocalPersistence;
  let mode = await loadTransport(deps.keyStore);
  let store: LocalStore | null = null;
  let localWs: Workspace | null = null;

  // Kept once open: a later switch back to local reuses it rather than
  // reopening IndexedDB and losing nothing that was pasted meanwhile.
  async function openLocal(): Promise<void> {
    if (store !== null) return;
    localWs = await localWorkspace();
    store = await LocalStore.open(await openPersistence());
  }

  if (mode === "local") await openLocal();

  function relay(): RelayApi {
    if (mode !== "local") return deps.remote;
    if (store === null) {
      // Cannot happen: mode only ever becomes "local" through the two paths
      // above, both of which await openLocal() first.
      throw new Error("PptTransport: local mode is not open yet");
    }
    return store;
  }

  return {
    mode: () => mode,
    relay,
    workspace: () => (mode === "local" ? localWs : deps.paired()),
    store: () => store,
    async setMode(next) {
      if (next === "local") await openLocal();
      mode = next;
      await saveTransport(deps.keyStore, next);
    },
  };
}
