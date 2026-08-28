//! sqlite store behind the relay: sealed link revisions and workspace inbox
//! items. Owns the schema, ownership by `sha256(authKey)`, the two-revision
//! retention rule and TTL expiry (7 d links, 24 h inbox).
//! Invariant: blobs are opaque ciphertext - the server holds no key.

use std::{path::Path, sync::Mutex, sync::MutexGuard};

use rusqlite::{params, Connection};
use sha2::{Digest, Sha256};

/// A link lives seven days from its last push.
pub const LINK_TTL: i64 = 7 * 24 * 3600;
/// An inbox item lives a day - long enough to reach the deck, not to linger.
pub const INBOX_TTL: i64 = 24 * 3600;

const SCHEMA: &str = "
CREATE TABLE IF NOT EXISTS links (id TEXT NOT NULL, rev INTEGER NOT NULL, auth_hash BLOB NOT NULL, pushed_at INTEGER NOT NULL, expires_at INTEGER NOT NULL, blob BLOB NOT NULL, PRIMARY KEY (id, rev));
CREATE TABLE IF NOT EXISTS inbox (ws TEXT NOT NULL, id TEXT NOT NULL, auth_hash BLOB NOT NULL, created_at INTEGER NOT NULL, expires_at INTEGER NOT NULL, blob BLOB NOT NULL, PRIMARY KEY (ws, id));
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

/// One row of a status batch: `rev` is `None` when unknown, expired or foreign.
#[derive(Debug)]
pub struct StatusRow {
    pub id: String,
    pub rev: Option<i64>,
    pub pushed_at: Option<i64>,
    pub auth_error: bool,
}

#[derive(Debug)]
pub struct InboxRow {
    pub id: String,
    pub created_at: i64,
    pub blob: Vec<u8>,
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

fn head(conn: &Connection, id: &str, now: i64) -> rusqlite::Result<Option<Head>> {
    conn.query_row(
        "SELECT rev, pushed_at, auth_hash FROM links WHERE id = ?1 AND expires_at > ?2 ORDER BY rev DESC LIMIT 1",
        params![id, now],
        |row| {
            Ok(Head {
                rev: row.get(0)?,
                pushed_at: row.get(1)?,
                auth: row.get(2)?,
            })
        },
    )
    .map(Some)
    .or_else(|error| match error {
        rusqlite::Error::QueryReturnedNoRows => Ok(None),
        other => Err(other),
    })
}

/// Expiry is enforced on read and on write, so an unswept row is still dead.
fn sweep_locked(conn: &Connection, now: i64) -> rusqlite::Result<usize> {
    let links = conn.execute("DELETE FROM links WHERE expires_at <= ?1", params![now])?;
    let inbox = conn.execute("DELETE FROM inbox WHERE expires_at <= ?1", params![now])?;
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

    fn conn(&self) -> MutexGuard<'_, Connection> {
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

    /// Removes every revision of a link.
    pub fn delete_link(&self, id: &str, auth_hash: &[u8; 32], now: i64) -> rusqlite::Result<Delete> {
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
    pub fn status(&self, items: &[(String, [u8; 32])], now: i64) -> rusqlite::Result<Vec<StatusRow>> {
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

    /// Drops a sealed item into a workspace inbox; a re-export replaces the row.
    pub fn post_inbox(
        &self,
        ws: &str,
        auth_hash: &[u8; 32],
        id: &str,
        blob: &[u8],
        now: i64,
    ) -> rusqlite::Result<()> {
        let conn = self.conn();
        sweep_locked(&conn, now)?;
        conn.execute(
            "INSERT INTO inbox (ws, id, auth_hash, created_at, expires_at, blob) VALUES (?1, ?2, ?3, ?4, ?5, ?6)
             ON CONFLICT (ws, id) DO UPDATE SET auth_hash = excluded.auth_hash, created_at = excluded.created_at, expires_at = excluded.expires_at, blob = excluded.blob",
            params![ws, id, auth_hash.as_slice(), now, now + INBOX_TTL, blob],
        )?;
        Ok(())
    }

    /// Live items of a workspace, newest first; a foreign key simply sees none.
    pub fn list_inbox(
        &self,
        ws: &str,
        auth_hash: &[u8; 32],
        now: i64,
    ) -> rusqlite::Result<Vec<InboxRow>> {
        let conn = self.conn();
        let mut statement = conn.prepare(
            "SELECT id, created_at, blob FROM inbox WHERE ws = ?1 AND auth_hash = ?2 AND expires_at > ?3 ORDER BY created_at DESC, id DESC",
        )?;
        let rows = statement.query_map(params![ws, auth_hash.as_slice(), now], |row| {
            Ok(InboxRow {
                id: row.get(0)?,
                created_at: row.get(1)?,
                blob: row.get(2)?,
            })
        })?;
        rows.collect()
    }

    /// Removes one inbox item once its deck has taken it.
    pub fn delete_inbox(
        &self,
        ws: &str,
        auth_hash: &[u8; 32],
        id: &str,
        now: i64,
    ) -> rusqlite::Result<Delete> {
        let conn = self.conn();
        sweep_locked(&conn, now)?;
        let owner: Option<Vec<u8>> = conn
            .query_row(
                "SELECT auth_hash FROM inbox WHERE ws = ?1 AND id = ?2 AND expires_at > ?3",
                params![ws, id, now],
                |row| row.get(0),
            )
            .map(Some)
            .or_else(|error| match error {
                rusqlite::Error::QueryReturnedNoRows => Ok(None),
                other => Err(other),
            })?;
        let Some(owner) = owner else {
            return Ok(Delete::Missing);
        };
        if owner != auth_hash.as_slice() {
            return Ok(Delete::Forbidden);
        }
        conn.execute(
            "DELETE FROM inbox WHERE ws = ?1 AND id = ?2",
            params![ws, id],
        )?;
        Ok(Delete::Deleted)
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
