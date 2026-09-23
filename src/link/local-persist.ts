// src/link/local-persist.ts
// Where pasted links live between sessions: the persistence port LocalStore
// writes through, its memory stand-in, and (Task 6) IndexedDB. Invariant: no
// method throws into a paste; a failed write only turns `durable` false.

export interface StoredLocalLink {
  id: string;
  rev: number;
  sentAt: number;
  blob: Uint8Array;
  pastedAt: number;
  previous?: { rev: number; sentAt: number; blob: Uint8Array };
}

export interface StoredLocalInbox {
  id: string;
  createdAt: number;
  blob: Uint8Array;
  pastedAt: number;
}

export interface LocalPersistence {
  readonly durable: boolean;
  load(): Promise<{ links: StoredLocalLink[]; inbox: StoredLocalInbox[] }>;
  putLinks(links: StoredLocalLink[]): Promise<void>;
  putInbox(rows: StoredLocalInbox[]): Promise<void>;
  deleteLinks(ids: string[]): Promise<void>;
  deleteInbox(ids: string[]): Promise<void>;
  clear(): Promise<void>;
}

// Nothing survives the pane: LocalStore's own maps are the memory.
export function memoryPersistence(): LocalPersistence {
  const nothing = async (): Promise<void> => undefined;
  return {
    durable: false,
    load: async () => ({ links: [], inbox: [] }),
    putLinks: nothing,
    putInbox: nothing,
    deleteLinks: nothing,
    deleteInbox: nothing,
    clear: nothing,
  };
}
