//! How full the store is, and what a write must fit under: the blob-byte sum
//! over both tables, the counters `/version` reads (cached, because counting
//! scans both tables under the one connection mutex every push also needs),
//! and the two ceilings a write is checked against.
//! Invariant: a ceiling is read inside the same lock as the insert it guards,
//! so two concurrent writes can never both pass the same pre-write total.

use std::sync::{Mutex, PoisonError};

use rusqlite::{params, Connection};

use crate::store::{Counts, Store};

/// How long `/version` may reuse its counters. The route is anonymous and sits
/// on the Access-bypassed path, so without this a flood of `/version` would
/// hold the connection mutex against every pane's push.
pub const VERSION_COUNTS_TTL: i64 = 30;

/// The ceilings a write is checked against: blob bytes over both tables, and
/// live inbox rows in one workspace. A store built without them holds whatever
/// it is given, which is what the seeding half of the tests wants.
#[derive(Clone, Copy, Debug)]
pub struct Caps {
    pub max_bytes: i64,
    pub inbox_rows: i64,
}

impl Caps {
    /// No ceiling at all: a store nobody has configured.
    pub const OPEN: Caps = Caps {
        max_bytes: i64::MAX,
        inbox_rows: i64::MAX,
    };
}

/// The last counted answer and the second it was counted at.
#[derive(Default)]
pub struct CountsCache {
    last: Mutex<Option<(i64, Counts)>>,
}

impl CountsCache {
    /// The counters if they were taken within the TTL, else nothing and the
    /// caller counts again. A poisoned cache is a stale number, never a panic.
    pub fn get(&self, now: i64) -> Option<Counts> {
        let last = self.last.lock().unwrap_or_else(PoisonError::into_inner);
        last.as_ref()
            .filter(|(at, _)| now - *at < VERSION_COUNTS_TTL && now >= *at)
            .map(|(_, counts)| counts.clone())
    }

    pub fn set(&self, now: i64, counts: &Counts) {
        let mut last = self.last.lock().unwrap_or_else(PoisonError::into_inner);
        *last = Some((now, counts.clone()));
    }

    /// Forgets the held answer: the sweeper calls this after dropping dead
    /// rows, so `/version` never reports a store the sweep has just emptied.
    pub fn clear(&self) {
        let mut last = self.last.lock().unwrap_or_else(PoisonError::into_inner);
        *last = None;
    }
}

/// Blob bytes both tables hold together: what the storage ceiling counts.
/// Expired rows count until they are swept, which is why every write sweeps
/// before it measures.
pub(crate) fn bytes_locked(conn: &Connection) -> rusqlite::Result<i64> {
    conn.query_row(
        "SELECT (SELECT COALESCE(SUM(LENGTH(blob)), 0) FROM links)
              + (SELECT COALESCE(SUM(LENGTH(blob)), 0) FROM inbox_v2)",
        [],
        |row| row.get(0),
    )
}

/// Whether `adding` more bytes still fit. Called from inside the write's own
/// lock, after the sweep every write starts with, so nothing can be written
/// between the answer and the insert it authorises.
pub(crate) fn room_locked(
    conn: &Connection,
    adding: usize,
    max_bytes: i64,
) -> rusqlite::Result<bool> {
    let adding = i64::try_from(adding).unwrap_or(i64::MAX);
    Ok(bytes_locked(conn)?.saturating_add(adding) <= max_bytes)
}

/// Whether one workspace has room for another inbox row. Counted over every
/// key that wrote into the workspace, because the cap protects the store, not
/// one pane's view of it.
pub(crate) fn inbox_room_locked(
    conn: &Connection,
    ws: &str,
    now: i64,
    max_rows: i64,
) -> rusqlite::Result<bool> {
    let rows: i64 = conn.query_row(
        "SELECT COUNT(*) FROM inbox_v2 WHERE ws = ?1 AND expires_at > ?2",
        params![ws, now],
        |row| row.get(0),
    )?;
    Ok(rows < max_rows)
}

impl Store {
    /// Blob bytes both tables hold together.
    pub fn total_bytes(&self) -> rusqlite::Result<i64> {
        bytes_locked(&self.conn())
    }

    /// What `/version` reports: rows held, links behind them, bytes on disk.
    /// Uncached - `CountsCache` is what keeps the route off this.
    pub fn counts(&self) -> rusqlite::Result<Counts> {
        let conn = self.conn();
        let one =
            |sql: &str| -> rusqlite::Result<i64> { conn.query_row(sql, [], |row| row.get(0)) };
        Ok(Counts {
            links: one("SELECT COUNT(DISTINCT id) FROM links")?,
            revisions: one("SELECT COUNT(*) FROM links")?,
            inbox: one("SELECT COUNT(*) FROM inbox_v2")?,
            bytes: bytes_locked(&conn)?,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// M3 of the N1 security review: one panic while the connection was locked
    /// used to turn every later request into a panic of its own.
    #[test]
    fn a_panic_under_the_lock_does_not_wedge_the_store() {
        let store = std::sync::Arc::new(Store::in_memory().unwrap());
        let poisoner = std::sync::Arc::clone(&store);
        let panicked = std::thread::spawn(move || {
            let _held = poisoner.conn();
            panic!("a handler died holding the connection");
        })
        .join();
        assert!(panicked.is_err());
        assert_eq!(store.total_bytes().unwrap(), 0);
    }

    /// The counters are reused for the TTL and counted again after it.
    #[test]
    fn the_counts_cache_answers_for_its_ttl_only() {
        let cache = CountsCache::default();
        let counts = Counts {
            links: 1,
            revisions: 2,
            inbox: 3,
            bytes: 4,
        };
        cache.set(1_000, &counts);
        assert_eq!(cache.get(1_000).map(|held| held.bytes), Some(4));
        assert_eq!(
            cache.get(1_000 + VERSION_COUNTS_TTL - 1).map(|c| c.links),
            Some(1)
        );
        assert!(cache.get(1_000 + VERSION_COUNTS_TTL).is_none());
        // A clock that stepped back is not a fresh answer either.
        assert!(cache.get(900).is_none());
    }
}
