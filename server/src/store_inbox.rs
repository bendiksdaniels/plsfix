//! The workspace-inbox half of the store: `inbox_v2` rows keyed by
//! `(ws, id, sha256(authKey))`, their TTL and the count the per-workspace row
//! cap is measured against. The schema, the connection and the links table
//! live in `store`.
//! Invariant: a row is only ever seen by the key that wrote it.

use rusqlite::params;

use crate::store::{sweep_locked, Store, INBOX_TTL};
use crate::store_room::{inbox_room_locked, room_locked};

#[derive(Debug)]
pub struct InboxRow {
    pub id: String,
    pub created_at: i64,
    pub blob: Vec<u8>,
}

/// Outcome of an inbox write: stored, or no room for it - the store's byte
/// ceiling or the workspace's row cap.
#[derive(Debug)]
pub enum Posted {
    Stored,
    Full,
}

impl Store {
    /// Drops a sealed item into a workspace inbox; a re-export from the same
    /// key replaces its own row. A foreign key writes a row of its own, which
    /// only that key can list or delete, so it can neither block nor shadow
    /// the pane's item. Both ceilings are read after the sweep and inside the
    /// same lock as the insert, so concurrent posts cannot all pass one count.
    pub fn post_inbox(
        &self,
        ws: &str,
        auth_hash: &[u8; 32],
        id: &str,
        blob: &[u8],
        now: i64,
    ) -> rusqlite::Result<Posted> {
        let conn = self.conn();
        sweep_locked(&conn, now)?;
        let caps = self.caps();
        if !room_locked(&conn, blob.len(), caps.max_bytes)?
            || !inbox_room_locked(&conn, ws, now, caps.inbox_rows)?
        {
            return Ok(Posted::Full);
        }
        conn.execute(
            "INSERT INTO inbox_v2 (ws, id, auth_hash, created_at, expires_at, blob) VALUES (?1, ?2, ?3, ?4, ?5, ?6)
             ON CONFLICT (ws, id, auth_hash) DO UPDATE SET created_at = excluded.created_at, expires_at = excluded.expires_at, blob = excluded.blob",
            params![ws, id, auth_hash.as_slice(), now, now + INBOX_TTL, blob],
        )?;
        Ok(Posted::Stored)
    }

    /// Live items of a workspace, newest first. `created_at` is whole seconds,
    /// so two exports in the same second tie; the old tiebreak, `id DESC`,
    /// sorted on the link's random hex id and so picked either one with even
    /// odds. `rowid` only ever grows (this table keeps its default rowid), so
    /// a tie now reads as "the row inserted last of the two" - for two
    /// distinct links, the one exported last, which is what "Paste latest
    /// linked" promises; a re-export upserts its existing row rather than
    /// reinserting it, so it keeps its original rowid and does not itself
    /// move up a tie. A foreign key simply sees no rows.
    pub fn list_inbox(
        &self,
        ws: &str,
        auth_hash: &[u8; 32],
        now: i64,
    ) -> rusqlite::Result<Vec<InboxRow>> {
        let conn = self.conn();
        let mut statement = conn.prepare(
            "SELECT id, created_at, blob FROM inbox_v2 WHERE ws = ?1 AND auth_hash = ?2 AND expires_at > ?3 ORDER BY created_at DESC, rowid DESC",
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

    /// Removes one inbox item once its deck has taken it. Only the key that
    /// wrote the row can see it, so a foreign key deletes nothing and is told
    /// the item is missing - never that someone else owns it.
    pub fn delete_inbox(
        &self,
        ws: &str,
        auth_hash: &[u8; 32],
        id: &str,
        now: i64,
    ) -> rusqlite::Result<bool> {
        let conn = self.conn();
        sweep_locked(&conn, now)?;
        let removed = conn.execute(
            "DELETE FROM inbox_v2 WHERE ws = ?1 AND id = ?2 AND auth_hash = ?3",
            params![ws, id, auth_hash.as_slice()],
        )?;
        Ok(removed > 0)
    }

    /// Live rows one workspace holds, over every key that wrote into it: what
    /// the per-workspace cap is measured against. The cap itself is applied
    /// inside `post_inbox`; this is the same count for a caller that only asks.
    pub fn inbox_count(&self, ws: &str, now: i64) -> rusqlite::Result<i64> {
        self.conn().query_row(
            "SELECT COUNT(*) FROM inbox_v2 WHERE ws = ?1 AND expires_at > ?2",
            params![ws, now],
            |row| row.get(0),
        )
    }
}
