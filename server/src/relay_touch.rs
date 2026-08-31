//! `POST /api/links/touch`: the Links tab telling the relay which links a
//! workbook still holds, so they keep their full TTL instead of expiring under
//! a deck nobody opened for a month. Bearer-less like the other batches - it
//! carries a key per item - and capped at the same 200 items.
//! Invariant: a link is touched from its key's hash, and one that does not
//! match is skipped in silence, never named.

use axum::{body::Bytes, extract::State, response::IntoResponse, Json};
use serde::Deserialize;

use crate::relay::{failed, now, Api, Refused, Reply, MAX_BATCH_ITEMS};
use crate::store::auth_hash;

/// One link a workbook still holds: which link, and the key that owns it.
#[derive(Deserialize)]
struct TouchQuery {
    id: String,
    auth: String,
}

/// Answers with how many links moved. Unknown, expired and foreign ids are
/// simply not counted: the status batch already tells a matching key that its
/// link exists, so a count adds no oracle a wrong key could read.
pub(crate) async fn touch(State(state): Api, body: Bytes) -> Reply {
    let items: Vec<TouchQuery> = serde_json::from_slice(&body).map_err(|_| Refused::BadBody)?;
    if items.len() > MAX_BATCH_ITEMS {
        return Err(Refused::TooManyItems);
    }
    let queries: Vec<(String, [u8; 32])> = items
        .into_iter()
        .map(|item| (item.id, auth_hash(&item.auth)))
        .collect();
    let touched = state
        .store
        .touch_links(&queries, now())
        .map_err(|error| failed("touch", "batch", &error))?;
    Ok(Json(serde_json::json!({ "touched": touched })).into_response())
}
