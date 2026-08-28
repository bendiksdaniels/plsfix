// Workspace pairing key: the 32 random bytes an Excel pane and a PowerPoint pane
// share so an inbox item sealed on one host opens on the other. Owns the derived
// {id, enc, auth} triple, the export string a user copies between hosts, and the
// host key/value store it is remembered in. Unreadable storage reads as unpaired.
import { fromBase64Url, hkdf, randomBytes, toBase64Url } from "./crypto";

export interface KeyStore {
  get(key: string): Promise<string | null>;
  set(key: string, value: string): Promise<void>;
  remove(key: string): Promise<void>;
}

export const WORKSPACE_STORAGE_KEY = "smt.link.workspace.v1";
const KEY_BYTES = 32;

// The one process-wide fallback map, so two calls in a host without storage
// still see the same pairing for the life of the pane.
const fallback = new Map<string, string>();

export function officeKeyStore(): KeyStore {
  const runtime =
    typeof OfficeRuntime === "undefined" ? undefined : OfficeRuntime;
  if (runtime) {
    const storage = runtime.storage;
    return {
      get: (key) => storage.getItem(key),
      set: (key, value) => storage.setItem(key, value),
      remove: (key) => storage.removeItem(key),
    };
  }
  const web = webStorage();
  if (web) {
    return {
      get: async (key) => web.getItem(key),
      set: async (key, value) => {
        web.setItem(key, value);
      },
      remove: async (key) => {
        web.removeItem(key);
      },
    };
  }
  return {
    get: async (key) => fallback.get(key) ?? null,
    set: async (key, value) => {
      fallback.set(key, value);
    },
    remove: async (key) => {
      fallback.delete(key);
    },
  };
}

export interface Workspace {
  id: string;
  enc: Uint8Array;
  auth: string;
  exportKey: string;
}

// id and auth are public (a relay path and a bearer token); enc never leaves
// the host, and exportKey is the secret itself, shown only when a user pairs.
export async function deriveWorkspace(secret: Uint8Array): Promise<Workspace> {
  return {
    id: toBase64Url(await hkdf(secret, "smt-ws-id")),
    enc: await hkdf(secret, "smt-ws-enc"),
    auth: toBase64Url(await hkdf(secret, "smt-ws-auth")),
    exportKey: toBase64Url(secret),
  };
}

export async function loadWorkspace(
  store: KeyStore,
): Promise<Workspace | null> {
  const stored = await store.get(WORKSPACE_STORAGE_KEY);
  if (stored === null) return null;
  const secret = readKey(stored);
  return secret === null ? null : await deriveWorkspace(secret);
}

export async function createWorkspace(store: KeyStore): Promise<Workspace> {
  const workspace = await deriveWorkspace(randomBytes(KEY_BYTES));
  await store.set(WORKSPACE_STORAGE_KEY, workspace.exportKey);
  return workspace;
}

export async function importWorkspace(
  store: KeyStore,
  exportKey: string,
): Promise<Workspace> {
  const secret = readKey(exportKey);
  if (secret === null) throw new Error("importWorkspace: invalid link key");
  const workspace = await deriveWorkspace(secret);
  await store.set(WORKSPACE_STORAGE_KEY, workspace.exportKey);
  return workspace;
}

export async function forgetWorkspace(store: KeyStore): Promise<void> {
  await store.remove(WORKSPACE_STORAGE_KEY);
}

// Null rather than a throw: a truncated or foreign value means "not paired
// here", and the pasted text is never echoed back because it is the key.
function readKey(text: string): Uint8Array | null {
  try {
    const bytes = fromBase64Url(text.trim());
    return bytes.length === KEY_BYTES ? bytes : null;
  } catch {
    return null;
  }
}

// localStorage exists but is unusable in several hosts (a Node test runner, a
// browser with site data blocked), so it is probed by use, not by typeof.
function webStorage(): Storage | null {
  const probe = `${WORKSPACE_STORAGE_KEY}.probe`;
  try {
    const storage = globalThis.localStorage;
    storage.setItem(probe, "1");
    storage.removeItem(probe);
    return storage;
  } catch {
    return null;
  }
}
