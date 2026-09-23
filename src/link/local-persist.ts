// src/link/local-persist.ts
// Where pasted links live between sessions: the persistence port LocalStore
// writes through, its memory stand-in, and IndexedDB. Invariant: no method
// throws into a paste; a failed write only turns `durable` false.

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

const DB_NAME = "plsfix-links";
const DB_VERSION = 1;
const LINKS = "links";
const INBOX = "inbox";

export async function openLocalPersistence(
  factory: IDBFactory | undefined = globalThis.indexedDB,
): Promise<LocalPersistence> {
  if (factory === undefined) return memoryPersistence();
  try {
    return new IdbPersistence(await openDb(factory));
  } catch {
    return memoryPersistence();
  }
}

function openDb(factory: IDBFactory): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = factory.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(LINKS))
        db.createObjectStore(LINKS, { keyPath: "id" });
      if (!db.objectStoreNames.contains(INBOX))
        db.createObjectStore(INBOX, { keyPath: "id" });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () =>
      reject(request.error ?? new Error("indexedDB: open failed"));
    request.onblocked = () => reject(new Error("indexedDB: open blocked"));
  });
}

function readAll<T>(db: IDBDatabase, name: string): Promise<T[]> {
  return new Promise((resolve, reject) => {
    const request = db.transaction(name, "readonly").objectStore(name).getAll();
    request.onsuccess = () => resolve(request.result as T[]);
    request.onerror = () =>
      reject(request.error ?? new Error(`indexedDB: read ${name}`));
  });
}

class IdbPersistence implements LocalPersistence {
  durable = true;

  constructor(private readonly db: IDBDatabase) {}

  async load(): Promise<{
    links: StoredLocalLink[];
    inbox: StoredLocalInbox[];
  }> {
    try {
      const [links, inbox] = await Promise.all([
        readAll<StoredLocalLink>(this.db, LINKS),
        readAll<StoredLocalInbox>(this.db, INBOX),
      ]);
      return { links, inbox };
    } catch {
      this.durable = false;
      return { links: [], inbox: [] };
    }
  }

  putLinks(links: StoredLocalLink[]): Promise<void> {
    return this.write([LINKS], (tx) =>
      links.forEach((link) => tx.objectStore(LINKS).put(link)),
    );
  }
  putInbox(rows: StoredLocalInbox[]): Promise<void> {
    return this.write([INBOX], (tx) =>
      rows.forEach((row) => tx.objectStore(INBOX).put(row)),
    );
  }
  deleteLinks(ids: string[]): Promise<void> {
    return this.write([LINKS], (tx) =>
      ids.forEach((id) => tx.objectStore(LINKS).delete(id)),
    );
  }
  deleteInbox(ids: string[]): Promise<void> {
    return this.write([INBOX], (tx) =>
      ids.forEach((id) => tx.objectStore(INBOX).delete(id)),
    );
  }
  clear(): Promise<void> {
    return this.write([LINKS, INBOX], (tx) => {
      tx.objectStore(LINKS).clear();
      tx.objectStore(INBOX).clear();
    });
  }

  // Resolves either way: a full or refused store must not fail the paste that
  // is already in memory; it only stops claiming to be durable.
  private write(
    names: string[],
    fill: (tx: IDBTransaction) => void,
  ): Promise<void> {
    return new Promise((resolve) => {
      try {
        const tx = this.db.transaction(names, "readwrite");
        tx.oncomplete = () => resolve();
        tx.onerror = tx.onabort = () => {
          this.durable = false;
          resolve();
        };
        fill(tx);
      } catch {
        this.durable = false;
        resolve();
      }
    });
  }
}
