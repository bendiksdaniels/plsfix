//! How full the store is, and what a write must fit under: the blob-byte sum
//! over both tables, the counters `/version` reads (cached, because counting
//! scans both tables under the one connection mutex every push also needs),
//! and the two ceilings a write is checked against.
//! Invariant: a ceiling is read inside the same lock as the insert it guards,
//! so two concurrent writes can never both pass the same pre-write total.

use std::future::Future;

use rusqlite::{params, Connection};
use tokio::sync::Mutex;

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
    /// The counters, recounted at most once every `VERSION_COUNTS_TTL` AND at
    /// most once at a time: the lock is held across the recount, so a second
    /// reader that arrived on the same miss waits for the first one's answer
    /// instead of scanning the store again.
    pub async fn counted<F>(&self, now: i64, recount: F) -> Option<Counts>
    where
        F: Future<Output = Option<Counts>>,
    {
        let mut held = self.last.lock().await;
        if let Some(fresh) = fresh(&held, now) {
            return Some(fresh);
        }
        let counted = recount.await?;
        *held = Some((now, counted.clone()));
        Some(counted)
    }

    /// Forgets the held answer: the sweeper calls this after dropping dead
    /// rows, so `/version` never reports a store the sweep has just emptied.
    pub async fn clear(&self) {
        *self.last.lock().await = None;
    }
}

/// The held answer if it was taken within the TTL. A clock that stepped back
/// is not a fresh answer either.
fn fresh(held: &Option<(i64, Counts)>, now: i64) -> Option<Counts> {
    held.as_ref()
        .filter(|(at, _)| now - *at < VERSION_COUNTS_TTL && now >= *at)
        .map(|(_, counts)| counts.clone())
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
    use std::{
        sync::{
            atomic::{AtomicUsize, Ordering},
            Arc,
        },
        time::Duration,
    };

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

    fn counts(links: i64) -> Counts {
        Counts {
            links,
            revisions: 2,
            inbox: 3,
            bytes: 4,
        }
    }

    /// The counters are reused for the TTL and counted again after it.
    #[tokio::test]
    async fn the_counts_cache_answers_for_its_ttl_only() {
        let cache = CountsCache::default();
        let scans = AtomicUsize::new(0);
        let count = || async {
            scans.fetch_add(1, Ordering::SeqCst);
            Some(counts(1))
        };

        assert_eq!(
            cache.counted(1_000, count()).await.map(|c| c.bytes),
            Some(4)
        );
        let within = 1_000 + VERSION_COUNTS_TTL - 1;
        assert_eq!(
            cache.counted(within, count()).await.map(|c| c.links),
            Some(1)
        );
        assert_eq!(scans.load(Ordering::SeqCst), 1);

        cache.counted(1_000 + VERSION_COUNTS_TTL, count()).await;
        // A clock that stepped back is not a fresh answer either.
        cache.counted(900, count()).await;
        assert_eq!(scans.load(Ordering::SeqCst), 3);

        cache.clear().await;
        cache.counted(900, count()).await;
        assert_eq!(scans.load(Ordering::SeqCst), 4);
    }

    /// Two readers that miss together must not both scan the store: the second
    /// waits for the first one's answer.
    #[tokio::test(flavor = "multi_thread", worker_threads = 2)]
    async fn two_readers_that_miss_together_count_once() {
        let cache = Arc::new(CountsCache::default());
        let scans = Arc::new(AtomicUsize::new(0));
        let slow = |scans: Arc<AtomicUsize>| async move {
            scans.fetch_add(1, Ordering::SeqCst);
            tokio::time::sleep(Duration::from_millis(80)).await;
            Some(counts(7))
        };

        let mut readers = Vec::new();
        for _ in 0..2 {
            let (cache, scans) = (Arc::clone(&cache), Arc::clone(&scans));
            readers.push(tokio::spawn(async move {
                cache.counted(5_000, slow(Arc::clone(&scans))).await
            }));
        }
        for reader in readers {
            assert_eq!(reader.await.unwrap().map(|c| c.links), Some(7));
        }
        assert_eq!(scans.load(Ordering::SeqCst), 1);
    }
}
