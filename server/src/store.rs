//! sqlite store behind the relay: sealed link revisions plus the schema both
//! tables share. Owns ownership by `sha256(authKey)`, the two-revision
//! retention rule, TTL expiry (30 d links, 7 d inbox) and the row and byte
//! counts the storage ceiling and `/version` are read from; the inbox rows
//! themselves live in `store_inbox`.
//! Invariant: blobs are opaque ciphertext - the server holds no key.

use std::{path::Path, sync::Mutex, sync::MutexGuard};

use rusqlite::{params, Connection};
use sha2::{Digest, Sha256};

/// A link lives thirty days from its last push or touch: a deck reopened
/// after a month of holidays still repaints.
pub const LINK_TTL: i64 = 30 * 24 * 3600;
/// An inbox item lives a week - long enough to reach the deck, not to linger.
pub const INBOX_TTL: i64 = 7 * 24 * 3600;

// The inbox key carries the writer's hash: the server cannot tell that a
// bearer belongs to a workspace (the two are independent HKDF branches of a
// secret it never sees), so any key may POST to any `ws` path. With the hash
// in the primary key a foreign writer gets its own row, which nobody else can
// list or delete, instead of squatting the real pane's `(ws, id)` slot.
// inbox_v2 replaces the `PRIMARY KEY (ws, id)` table; the inbox is a
// short-lived buffer, so the old rows are dropped rather than migrated.
const SCHEMA: &str = "
CREATE TABLE IF NOT EXISTS links (id TEXT NOT NULL, rev INTEGER NOT NULL, auth_hash BLOB NOT NULL, pushed_at INTEGER NOT NULL, expires_at INTEGER NOT NULL, blob BLOB NOT NULL, PRIMARY KEY (id, rev));
CREATE TABLE IF NOT EXISTS inbox_v2 (ws TEXT NOT NULL, id TEXT NOT NULL, auth_hash BLOB NOT NULL, created_at INTEGER NOT NULL, expires_at INTEGER NOT NULL, blob BLOB NOT NULL, PRIMARY KEY (ws, id, auth_hash));
DROP TABLE IF EXISTS inbox;
";

const INSERT_REV: &str = "INSERT INTO links (id, rev, auth_hash, pushed_at, expires_at, blob) VALUES (?1, ?2, ?3, ?4, ?5, ?6)";

/// Sealed blobs keyed by link id and workspace, guarded by one connection.
pub struct Store {
    conn: Mutex<Connection>,
}

/// Outcome of a push: a new link, a new revision, or a foreign owner.
#[derive(Debug)]
pub enum Put {
    Created(i64),
    Updated(i64),
    Forbidden,
}

/// The newest live revision of a link.
#[derive(Debug)]
pub struct Found {
    pub rev: i64,
    pub pushed_at: i64,
    pub blob: Vec<u8>,
}

#[derive(Debug)]
pub enum Get {
    Found(Found),
    Forbidden,
    Missing,
}

#[derive(Debug)]
pub enum Delete {
    Deleted,
    Forbidden,
    Missing,
}

/// What the store holds right now, for `/version`: distinct links, the
/// revisions behind them, live-or-not inbox rows and the blob bytes of both.
#[derive(Debug)]
pub struct Counts {
    pub links: i64,
    pub revisions: i64,
    pub inbox: i64,
    pub bytes: i64,
}

/// One row of a status batch: `rev` is `None` when unknown, expired or foreign.
#[derive(Debug)]
pub struct StatusRow {
    pub id: String,
    pub rev: Option<i64>,
    pub pushed_at: Option<i64>,
    pub auth_error: bool,
}

/// What the store keeps of a bearer key: never the key itself.
pub fn auth_hash(bearer: &str) -> [u8; 32] {
    let mut hasher = Sha256::new();
    hasher.update(bearer.as_bytes());
    hasher.finalize().into()
}

/// The newest unexpired revision of a link, without its blob.
struct Head {
    rev: i64,
    pushed_at: i64,
    auth: Vec<u8>,
}

/// "No such row" is an outcome here, not a failure: every lookup below asks
/// for a row that may legitimately be gone.
fn optional<T>(result: rusqlite::Result<T>) -> rusqlite::Result<Option<T>> {
    match result {
        Ok(value) => Ok(Some(value)),
        Err(rusqlite::Error::QueryReturnedNoRows) => Ok(None),
        Err(other) => Err(other),
    }
}

fn head(conn: &Connection, id: &str, now: i64) -> rusqlite::Result<Option<Head>> {
    optional(conn.query_row(
        "SELECT rev, pushed_at, auth_hash FROM links WHERE id = ?1 AND expires_at > ?2 ORDER BY rev DESC LIMIT 1",
        params![id, now],
        |row| {
            Ok(Head {
                rev: row.get(0)?,
                pushed_at: row.get(1)?,
                auth: row.get(2)?,
            })
        },
    ))
}

/// Expiry is enforced on read and on write, so an unswept row is still dead.
pub(crate) fn sweep_locked(conn: &Connection, now: i64) -> rusqlite::Result<usize> {
    let links = conn.execute("DELETE FROM links WHERE expires_at <= ?1", params![now])?;
    let inbox = conn.execute("DELETE FROM inbox_v2 WHERE expires_at <= ?1", params![now])?;
    Ok(links + inbox)
}

fn cannot_open(path: &Path, error: &std::io::Error) -> rusqlite::Error {
    rusqlite::Error::SqliteFailure(
        rusqlite::ffi::Error::new(rusqlite::ffi::SQLITE_CANTOPEN),
        Some(format!("store open {}: {error}", path.display())),
    )
}

impl Store {
    /// Opens (creating the directory and file) a WAL database at `path`.
    pub fn open(path: &Path) -> rusqlite::Result<Store> {
        if let Some(parent) = path.parent().filter(|dir| !dir.as_os_str().is_empty()) {
            std::fs::create_dir_all(parent).map_err(|error| cannot_open(parent, &error))?;
        }
        let conn = Connection::open(path)?;
        // journal_mode answers with the mode it set, so query instead of execute.
        let _: String = conn.query_row("PRAGMA journal_mode=WAL", [], |row| row.get(0))?;
        conn.pragma_update(None, "synchronous", "NORMAL")?;
        Store::with_connection(conn)
    }

    /// A throwaway store for tests: same schema, nothing on disk.
    pub fn in_memory() -> rusqlite::Result<Store> {
        Store::with_connection(Connection::open_in_memory()?)
    }

    fn with_connection(conn: Connection) -> rusqlite::Result<Store> {
        conn.execute_batch(SCHEMA)?;
        Ok(Store {
            conn: Mutex::new(conn),
        })
    }

    pub(crate) fn conn(&self) -> MutexGuard<'_, Connection> {
        self.conn.lock().expect("store connection mutex poisoned")
    }

    /// First push creates rev 1 and fixes the owner; later pushes from the same
    /// owner add a revision, keep the last two and refresh the TTL.
    pub fn put_link(
        &self,
        id: &str,
        auth_hash: &[u8; 32],
        blob: &[u8],
        now: i64,
    ) -> rusqlite::Result<Put> {
        let conn = self.conn();
        sweep_locked(&conn, now)?;
        let expires = now + LINK_TTL;
        let tx = conn.unchecked_transaction()?;
        let outcome = match head(&tx, id, now)? {
            None => {
                tx.execute(
                    INSERT_REV,
                    params![id, 1, auth_hash.as_slice(), now, expires, blob],
                )?;
                Put::Created(1)
            }
            Some(head) if head.auth != auth_hash.as_slice() => Put::Forbidden,
            Some(head) => {
                let rev = head.rev + 1;
                tx.execute(
                    INSERT_REV,
                    params![id, rev, auth_hash.as_slice(), now, expires, blob],
                )?;
                tx.execute(
                    "DELETE FROM links WHERE id = ?1 AND rev < ?2",
                    params![id, rev - 1],
                )?;
                tx.execute(
                    "UPDATE links SET expires_at = ?2 WHERE id = ?1",
                    params![id, expires],
                )?;
                Put::Updated(rev)
            }
        };
        tx.commit()?;
        Ok(outcome)
    }

    /// The newest live revision, for the owner only.
    pub fn get_link(&self, id: &str, auth_hash: &[u8; 32], now: i64) -> rusqlite::Result<Get> {
        let conn = self.conn();
        let Some(head) = head(&conn, id, now)? else {
            return Ok(Get::Missing);
        };
        if head.auth != auth_hash.as_slice() {
            return Ok(Get::Forbidden);
        }
        let blob: Vec<u8> = conn.query_row(
            "SELECT blob FROM links WHERE id = ?1 AND rev = ?2",
            params![id, head.rev],
            |row| row.get(0),
        )?;
        Ok(Get::Found(Found {
            rev: head.rev,
            pushed_at: head.pushed_at,
            blob,
        }))
    }

    /// One revision by name, for the owner only: what "Revert last update"
    /// asks for. Ownership is the link's, not the row's, so a foreign key is
    /// refused whether or not the revision it names is still held; a rev the
    /// two-revision rule has dropped, or one that never existed, is missing.
    pub fn get_link_rev(
        &self,
        id: &str,
        auth_hash: &[u8; 32],
        rev: i64,
        now: i64,
    ) -> rusqlite::Result<Get> {
        let conn = self.conn();
        let Some(head) = head(&conn, id, now)? else {
            return Ok(Get::Missing);
        };
        if head.auth != auth_hash.as_slice() {
            return Ok(Get::Forbidden);
        }
        let found = optional(conn.query_row(
            "SELECT pushed_at, blob FROM links WHERE id = ?1 AND rev = ?2 AND expires_at > ?3",
            params![id, rev, now],
            |row| {
                Ok(Found {
                    rev,
                    pushed_at: row.get(0)?,
                    blob: row.get(1)?,
                })
            },
        ))?;
        Ok(found.map_or(Get::Missing, Get::Found))
    }

    /// Removes every revision of a link.
    pub fn delete_link(
        &self,
        id: &str,
        auth_hash: &[u8; 32],
        now: i64,
    ) -> rusqlite::Result<Delete> {
        let conn = self.conn();
        sweep_locked(&conn, now)?;
        let Some(head) = head(&conn, id, now)? else {
            return Ok(Delete::Missing);
        };
        if head.auth != auth_hash.as_slice() {
            return Ok(Delete::Forbidden);
        }
        conn.execute("DELETE FROM links WHERE id = ?1", params![id])?;
        Ok(Delete::Deleted)
    }

    /// One batch poll: unknown links report no rev, foreign ones an auth error.
    pub fn status(
        &self,
        items: &[(String, [u8; 32])],
        now: i64,
    ) -> rusqlite::Result<Vec<StatusRow>> {
        let conn = self.conn();
        let mut rows = Vec::with_capacity(items.len());
        for (id, auth) in items {
            rows.push(match head(&conn, id, now)? {
                Some(head) if head.auth == auth.as_slice() => StatusRow {
                    id: id.clone(),
                    rev: Some(head.rev),
                    pushed_at: Some(head.pushed_at),
                    auth_error: false,
                },
                Some(_) => StatusRow {
                    id: id.clone(),
                    rev: None,
                    pushed_at: None,
                    auth_error: true,
                },
                None => StatusRow {
                    id: id.clone(),
                    rev: None,
                    pushed_at: None,
                    auth_error: false,
                },
            });
        }
        Ok(rows)
    }

    /// Pushes the TTL of every revision of each link the caller still owns
    /// back to a full `LINK_TTL`, and answers with how many links moved. A
    /// link that is unknown, already expired or owned by another key is
    /// skipped in silence: the status batch already tells a matching key that
    /// its link exists, so this reveals nothing new to a wrong one.
    pub fn touch_links(&self, items: &[(String, [u8; 32])], now: i64) -> rusqlite::Result<usize> {
        let conn = self.conn();
        let expires = now + LINK_TTL;
        let mut touched = 0;
        // One transaction for the whole batch: a workbook touches up to 200
        // links on boot, and that is one commit, not two hundred.
        let tx = conn.unchecked_transaction()?;
        for (id, auth) in items {
            let Some(head) = head(&tx, id, now)? else {
                continue;
            };
            if head.auth != auth.as_slice() {
                continue;
            }
            // Every revision, not just the head: a revert reaches the one
            // below it, and a half-expired link would break that.
            tx.execute(
                "UPDATE links SET expires_at = ?2 WHERE id = ?1",
                params![id, expires],
            )?;
            touched += 1;
        }
        tx.commit()?;
        Ok(touched)
    }

    /// Blob bytes both tables hold together: what the storage ceiling counts.
    /// Expired rows count until they are swept, which is why a write that
    /// finds itself over the ceiling sweeps before it refuses.
    pub fn total_bytes(&self) -> rusqlite::Result<i64> {
        self.conn().query_row(
            "SELECT (SELECT COALESCE(SUM(LENGTH(blob)), 0) FROM links)
                  + (SELECT COALESCE(SUM(LENGTH(blob)), 0) FROM inbox_v2)",
            [],
            |row| row.get(0),
        )
    }

    /// What `/version` reports: rows held, links behind them, bytes on disk.
    pub fn counts(&self) -> rusqlite::Result<Counts> {
        let conn = self.conn();
        let one =
            |sql: &str| -> rusqlite::Result<i64> { conn.query_row(sql, [], |row| row.get(0)) };
        let links = one("SELECT COUNT(DISTINCT id) FROM links")?;
        let revisions = one("SELECT COUNT(*) FROM links")?;
        let inbox = one("SELECT COUNT(*) FROM inbox_v2")?;
        drop(conn);
        Ok(Counts {
            links,
            revisions,
            inbox,
            bytes: self.total_bytes()?,
        })
    }

    /// Deletes expired rows from both tables; runs on every write and hourly.
    pub fn sweep(&self, now: i64) -> rusqlite::Result<usize> {
        let conn = self.conn();
        sweep_locked(&conn, now)
    }

    /// Test support: revisions held for `id`, expired ones included.
    pub fn rev_count(&self, id: &str) -> usize {
        let count: i64 = self
            .conn()
            .query_row(
                "SELECT COUNT(*) FROM links WHERE id = ?1",
                params![id],
                |row| row.get(0),
            )
            .unwrap_or(0);
        usize::try_from(count).unwrap_or(0)
    }
}
