// src/link/transport-setting.ts
// Which way links travel on this device: "local" (one copy, one paste) or
// "relay". Stored per device beside the link key. Invariant: nothing stored
// means relay only where a key is already here, so existing users keep what
// works and new installs start local.

import { WORKSPACE_STORAGE_KEY, type KeyStore } from "./workspace";

export type LinkTransport = "local" | "relay";

export const TRANSPORT_STORAGE_KEY = "plsfix.link.transport.v1";

async function read(store: KeyStore, key: string): Promise<string | null> {
  try {
    return await store.get(key);
  } catch {
    return null;
  }
}

export async function loadTransport(store: KeyStore): Promise<LinkTransport> {
  const saved = await read(store, TRANSPORT_STORAGE_KEY);
  if (saved === "local" || saved === "relay") return saved;
  return (await read(store, WORKSPACE_STORAGE_KEY)) === null
    ? "local"
    : "relay";
}

export async function saveTransport(
  store: KeyStore,
  transport: LinkTransport,
): Promise<void> {
  await store.set(TRANSPORT_STORAGE_KEY, transport);
}
