//! `POST /api/links/fetch`: the batched half of "Update all" - every changed
//! picture of a deck in one response instead of one GET per link. Bearer-less
//! like the status batch, because it carries a key per item; capped at the same
//! 200 items and at 4 MiB of blobs, and whatever the cap leaves out is named
//! `deferred` so the client asks for it one link at a time.
//! Invariant: an item is answered from its key's hash, never its key, and a
//! link the key does not own is omitted, never described.

use axum::{body::Bytes, extract::State, response::IntoResponse, Json};
use base64::{engine::general_purpose::URL_SAFE_NO_PAD, Engine as _};
use serde::{Deserialize, Serialize};

use crate::relay::{failed, now, Api, Refused, Reply, MAX_BATCH_ITEMS};
use crate::store::{auth_hash, Found, Get, StatusRow, Store};

/// Blobs one response may carry. A deck of full-slide renders is far past this,
/// so the tail is deferred to individual GETs rather than truncated or refused.
pub(crate) const FETCH_BLOB_CAP: usize = 4 * 1024 * 1024;

/// What the caller holds for one link: its key, and the revision already in
/// the deck. A query without `knownRev` holds nothing and always gets a blob.
#[derive(Deserialize)]
pub(crate) struct FetchQuery {
    id: String,
    auth: String,
    #[serde(rename = "knownRev")]
    known_rev: Option<i64>,
}

/// A changed picture: the revision the blob belongs to, base64url encoded
/// because the batch answers in JSON.
#[derive(Serialize)]
struct FetchItem {
    id: String,
    rev: i64,
    blob: String,
}

/// Why a requested link carries no blob. `unchanged`, `missing` and `auth` are
/// final; `deferred` only means "not in this response".
#[derive(Serialize)]
struct Omitted {
    id: String,
    reason: &'static str,
}

#[derive(Serialize, Default)]
pub(crate) struct FetchOut {
    items: Vec<FetchItem>,
    omitted: Vec<Omitted>,
}

impl FetchOut {
    fn omit(&mut self, id: &str, reason: &'static str) {
        self.omitted.push(Omitted {
            id: id.to_string(),
            reason,
        });
    }

    fn take(&mut self, id: &str, found: Found) {
        self.items.push(FetchItem {
            id: id.to_string(),
            rev: found.rev,
            blob: URL_SAFE_NO_PAD.encode(found.blob),
        });
    }
}

/// The head of every requested link in one pass, so a deck that is already up
/// to date costs no blob reads at all.
fn heads(store: &Store, items: &[FetchQuery], now: i64) -> rusqlite::Result<Vec<StatusRow>> {
    let queries: Vec<(String, [u8; 32])> = items
        .iter()
        .map(|item| (item.id.clone(), auth_hash(&item.auth)))
        .collect();
    store.status(&queries, now)
}

/// One item's blob, if the cap still has room for it. `room` is `None` once an
/// item has been turned away: the rest of the batch is deferred with it, so a
/// client is never left guessing which answer was the whole picture.
fn blob(
    store: &Store,
    item: &FetchQuery,
    now: i64,
    out: &mut FetchOut,
    room: &mut Option<usize>,
) -> rusqlite::Result<()> {
    let Some(left) = *room else {
        out.omit(&item.id, "deferred");
        return Ok(());
    };
    match store.get_link(&item.id, &auth_hash(&item.auth), now)? {
        Get::Found(found) if found.blob.len() <= left => {
            *room = Some(left - found.blob.len());
            out.take(&item.id, found);
        }
        Get::Found(_) => {
            *room = None;
            out.omit(&item.id, "deferred");
        }
        // Only a push that raced this batch gets here: the head above already
        // agreed the key owns a live revision.
        Get::Forbidden => out.omit(&item.id, "auth"),
        Get::Missing => out.omit(&item.id, "missing"),
    }
    Ok(())
}

/// One batch: the heads first, then a blob for every link whose revision the
/// caller does not already hold, in request order and while the cap allows.
pub(crate) fn collect(store: &Store, items: &[FetchQuery], now: i64) -> rusqlite::Result<FetchOut> {
    let mut out = FetchOut::default();
    let mut room = Some(FETCH_BLOB_CAP);
    for (item, head) in items.iter().zip(heads(store, items, now)?) {
        match head.rev {
            _ if head.auth_error => out.omit(&item.id, "auth"),
            None => out.omit(&item.id, "missing"),
            Some(rev) if Some(rev) == item.known_rev => out.omit(&item.id, "unchanged"),
            Some(_) => blob(store, item, now, &mut out, &mut room)?,
        }
    }
    Ok(out)
}

/// Like the status batch, this route carries a key per item and no bearer.
pub(crate) async fn fetch(State(state): Api, body: Bytes) -> Reply {
    let items: Vec<FetchQuery> = serde_json::from_slice(&body).map_err(|_| Refused::BadBody)?;
    if items.len() > MAX_BATCH_ITEMS {
        return Err(Refused::TooManyItems);
    }
    let out =
        collect(&state.store, &items, now()).map_err(|error| failed("fetch", "batch", &error))?;
    Ok(Json(out).into_response())
}
