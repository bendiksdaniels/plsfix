//! Hunt pass 2 shared model: a plain HashMap standing in for the sqlite store
//! (links with at most two revisions, a workspace inbox with its row cap),
//! used by `hunt2_store_model.rs` to compare against the real `Store`.
//! Invariant: every method here mirrors one real `Store`/`store_inbox` method
//! exactly, including its call order (an auth check before an exhausted
//! check, a sweep before a room check), so a mismatch is the real bug.

use std::collections::HashMap;

use plsfix_server::store::{INBOX_TTL, LINK_TTL};
use plsfix_server::store_room::Caps;

pub const N_IDS: usize = 3;
pub const N_WS: usize = 2;
pub const N_AUTHS: usize = 2;

pub fn id_str(id: usize) -> String {
    format!("id-{id}")
}

pub fn ws_str(ws: usize) -> String {
    format!("ws-{ws}")
}

pub fn auths() -> [[u8; 32]; N_AUTHS] {
    [[1u8; 32], [2u8; 32]]
}

/// One link, at most two kept revisions (oldest first); every kept revision
/// of an id shares one `expires_at` in the real store (one UPDATE with no rev
/// filter refreshes them together on every push), so the model keeps one too.
#[derive(Clone)]
pub struct ModelLink {
    pub auth: usize,
    pub revs: Vec<(i64, i64, Vec<u8>)>, // (rev, pushed_at, blob)
    pub expires_at: i64,
}

#[derive(Clone)]
struct ModelInboxRow {
    rowid: i64,
    created_at: i64,
    expires_at: i64,
    blob: Vec<u8>,
}

#[derive(Debug, PartialEq)]
pub enum ModelPut {
    Created(i64),
    Updated(i64),
    Forbidden,
    Full,
    Exhausted,
}

#[derive(Debug, PartialEq)]
pub enum ModelGet {
    Found(i64, i64, Vec<u8>),
    Forbidden,
    Missing,
}

#[derive(Debug, PartialEq)]
pub enum ModelDelete {
    Deleted,
    Forbidden,
    Missing,
}

#[derive(Debug, PartialEq)]
pub enum ModelPosted {
    Stored,
    Full,
}

pub struct Model {
    pub links: HashMap<usize, ModelLink>,
    inbox: HashMap<(usize, usize, usize), ModelInboxRow>, // (ws, id, auth)
    next_rowid: i64,
    pub now: i64,
    pub caps: Caps,
}

impl Model {
    pub fn new(caps: Caps) -> Model {
        Model {
            links: HashMap::new(),
            inbox: HashMap::new(),
            next_rowid: 0,
            now: 0,
            caps,
        }
    }

    /// Blob bytes over both "tables", swept or not - the same total the real
    /// ceiling is read against (`store_room::bytes_locked`).
    pub fn total_bytes(&self) -> i64 {
        let links: i64 = self
            .links
            .values()
            .flat_map(|link| link.revs.iter())
            .map(|(_, _, blob)| blob.len() as i64)
            .sum();
        let inbox: i64 = self.inbox.values().map(|row| row.blob.len() as i64).sum();
        links + inbox
    }

    /// Physically drops expired rows, like `sweep_locked`. Only the write
    /// operations below call this; reads filter by `expires_at` live instead,
    /// which is why an unswept expired row still counts toward the byte cap.
    /// Counts dead REVISION rows, not dead link ids: `sweep_locked` deletes
    /// from `links`, keyed `(id, rev)`, so one link with two revisions past
    /// its TTL costs two rows, the same number `DELETE`'s row count reports.
    pub fn sweep(&mut self) -> usize {
        let now = self.now;
        let dead_revs: usize = self
            .links
            .values()
            .filter(|link| link.expires_at <= now)
            .map(|link| link.revs.len())
            .sum();
        let dead_inbox = self
            .inbox
            .values()
            .filter(|row| row.expires_at <= now)
            .count();
        self.links.retain(|_, link| link.expires_at > now);
        self.inbox.retain(|_, row| row.expires_at > now);
        dead_revs + dead_inbox
    }

    pub fn put_link(&mut self, id: usize, auth: usize, blob: &[u8]) -> ModelPut {
        self.sweep();
        if self.total_bytes() + blob.len() as i64 > self.caps.max_bytes {
            return ModelPut::Full;
        }
        let expires = self.now + LINK_TTL;
        match self.links.get(&id) {
            None => {
                self.links.insert(
                    id,
                    ModelLink {
                        auth,
                        revs: vec![(1, self.now, blob.to_vec())],
                        expires_at: expires,
                    },
                );
                ModelPut::Created(1)
            }
            Some(link) if link.auth != auth => ModelPut::Forbidden,
            Some(link) if link.revs.last().unwrap().0 == i64::MAX => ModelPut::Exhausted,
            Some(_) => {
                let link = self.links.get_mut(&id).unwrap();
                let rev = link.revs.last().unwrap().0 + 1;
                link.revs.push((rev, self.now, blob.to_vec()));
                if link.revs.len() > 2 {
                    link.revs.remove(0);
                }
                link.expires_at = expires;
                ModelPut::Updated(rev)
            }
        }
    }

    pub fn get_link(&self, id: usize, auth: usize) -> ModelGet {
        match self.links.get(&id) {
            None => ModelGet::Missing,
            Some(link) if link.expires_at <= self.now => ModelGet::Missing,
            Some(link) if link.auth != auth => ModelGet::Forbidden,
            Some(link) => {
                let (rev, pushed_at, blob) = link.revs.last().unwrap().clone();
                ModelGet::Found(rev, pushed_at, blob)
            }
        }
    }

    pub fn get_link_rev(&self, id: usize, auth: usize, rev: i64) -> ModelGet {
        match self.links.get(&id) {
            None => ModelGet::Missing,
            Some(link) if link.expires_at <= self.now => ModelGet::Missing,
            Some(link) if link.auth != auth => ModelGet::Forbidden,
            Some(link) => match link.revs.iter().find(|kept| kept.0 == rev) {
                Some((rev, pushed_at, blob)) => ModelGet::Found(*rev, *pushed_at, blob.clone()),
                None => ModelGet::Missing,
            },
        }
    }

    pub fn delete_link(&mut self, id: usize, auth: usize) -> ModelDelete {
        self.sweep();
        match self.links.get(&id) {
            None => ModelDelete::Missing,
            Some(link) if link.auth != auth => ModelDelete::Forbidden,
            Some(_) => {
                self.links.remove(&id);
                ModelDelete::Deleted
            }
        }
    }

    /// `(rev, pushed_at, auth_error)` per item, the same shape `StatusRow`
    /// answers with minus the echoed id.
    pub fn status(&self, items: &[(usize, usize)]) -> Vec<(Option<i64>, Option<i64>, bool)> {
        items
            .iter()
            .map(|&(id, auth)| match self.links.get(&id) {
                None => (None, None, false),
                Some(link) if link.expires_at <= self.now => (None, None, false),
                Some(link) if link.auth != auth => (None, None, true),
                Some(link) => {
                    let (rev, pushed_at, _) = link.revs.last().unwrap();
                    (Some(*rev), Some(*pushed_at), false)
                }
            })
            .collect()
    }

    pub fn touch_links(&mut self, items: &[(usize, usize)]) -> usize {
        let now = self.now;
        let mut touched = 0;
        for &(id, auth) in items {
            let live_and_owned = matches!(
                self.links.get(&id),
                Some(link) if link.expires_at > now && link.auth == auth
            );
            if live_and_owned {
                self.links.get_mut(&id).unwrap().expires_at = now + LINK_TTL;
                touched += 1;
            }
        }
        touched
    }

    pub fn post_inbox(&mut self, ws: usize, auth: usize, id: usize, blob: &[u8]) -> ModelPosted {
        self.sweep();
        if self.total_bytes() + blob.len() as i64 > self.caps.max_bytes {
            return ModelPosted::Full;
        }
        let key = (ws, id, auth);
        let replacing = self.inbox.contains_key(&key);
        if !replacing {
            let now = self.now;
            let live_in_ws = self
                .inbox
                .iter()
                .filter(|(&(row_ws, _, _), row)| row_ws == ws && row.expires_at > now)
                .count() as i64;
            if live_in_ws >= self.caps.inbox_rows {
                return ModelPosted::Full;
            }
        }
        let rowid = match self.inbox.get(&key) {
            Some(row) => row.rowid,
            None => {
                let rowid = self.next_rowid;
                self.next_rowid += 1;
                rowid
            }
        };
        self.inbox.insert(
            key,
            ModelInboxRow {
                rowid,
                created_at: self.now,
                expires_at: self.now + INBOX_TTL,
                blob: blob.to_vec(),
            },
        );
        ModelPosted::Stored
    }

    /// `(id, created_at, blob)`, newest first then highest rowid first - the
    /// same `ORDER BY created_at DESC, rowid DESC` `list_inbox` runs.
    pub fn list_inbox(&self, ws: usize, auth: usize) -> Vec<(usize, i64, Vec<u8>)> {
        let now = self.now;
        let mut rows: Vec<(usize, i64, i64, Vec<u8>)> = self
            .inbox
            .iter()
            .filter(|(&(row_ws, _, row_auth), row)| {
                row_ws == ws && row_auth == auth && row.expires_at > now
            })
            .map(|(&(_, id, _), row)| (id, row.created_at, row.rowid, row.blob.clone()))
            .collect();
        rows.sort_by(|a, b| b.1.cmp(&a.1).then(b.2.cmp(&a.2)));
        rows.into_iter()
            .map(|(id, created_at, _, blob)| (id, created_at, blob))
            .collect()
    }

    pub fn delete_inbox(&mut self, ws: usize, auth: usize, id: usize) -> bool {
        self.sweep();
        self.inbox.remove(&(ws, id, auth)).is_some()
    }

    /// Live rows of one workspace, over every auth - what the per-workspace
    /// cap and `Store::inbox_count` both count.
    pub fn live_inbox_count(&self, ws: usize) -> i64 {
        let now = self.now;
        self.inbox
            .iter()
            .filter(|(&(row_ws, _, _), row)| row_ws == ws && row.expires_at > now)
            .count() as i64
    }
}
