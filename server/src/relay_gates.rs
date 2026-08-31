//! What a write passes before it reaches the store: the per-client rate limit
//! (the buckets themselves live in `limits`), the byte ceiling over both
//! tables and the per-workspace inbox row cap. Owns the three refusals and
//! their log lines; the routes own everything else.
//! Invariant: a gate that finds itself full sweeps before it refuses, so a
//! store of dead rows heals itself rather than turning a pane away.

use axum::{
    extract::{Request, State},
    middleware::Next,
    response::{IntoResponse, Response},
};

use crate::limits::{client_key, is_write, Allowed};
use crate::relay::{failed, now, short, Api, AppState, Refused, INBOX_MAX_PER_WS};

/// One token per request, spent before the body is read: the write allowance
/// for a push or a delete, the read one for a GET or a batch. The refusal
/// names the bucket and the wait, never anything the request carried.
pub(crate) async fn rate_limit(State(state): Api, request: Request, next: Next) -> Response {
    let write = is_write(request.method(), request.uri().path());
    let key = client_key(request.headers(), request.extensions());
    let limiter = if write { &state.writes } else { &state.reads };
    match limiter.take(&key, now()) {
        Allowed::Yes => next.run(request).await,
        Allowed::No { retry_after } => {
            let kind = if write { "write" } else { "read" };
            eprintln!("relay rate {kind} {key}: retry after {retry_after}s");
            Refused::RateLimited(retry_after).into_response()
        }
    }
}

/// The byte gate every write passes. A store over its ceiling sweeps first, so
/// a month of dead links heals the write instead of refusing it; only a store
/// that is still full afterwards answers 507.
pub(crate) fn room_for(
    state: &AppState,
    stage: &str,
    id: &str,
    adding: usize,
) -> Result<(), Refused> {
    let adding = i64::try_from(adding).unwrap_or(i64::MAX);
    let fits = |bytes: i64| bytes.saturating_add(adding) <= state.max_bytes;
    let mut bytes = state
        .store
        .total_bytes()
        .map_err(|e| failed(stage, id, &e))?;
    if fits(bytes) {
        return Ok(());
    }
    state
        .store
        .sweep(now())
        .map_err(|e| failed(stage, id, &e))?;
    bytes = state
        .store
        .total_bytes()
        .map_err(|e| failed(stage, id, &e))?;
    if fits(bytes) {
        return Ok(());
    }
    eprintln!(
        "store full {stage} {}: {} of {} bytes",
        short(id),
        bytes.saturating_add(adding),
        state.max_bytes
    );
    Err(Refused::Full)
}

/// The row gate on one workspace's inbox, swept the same way. Counted over
/// every key that wrote into the workspace, because the cap protects the
/// store, not one pane's view of it.
pub(crate) fn inbox_room(state: &AppState, ws: &str, id: &str) -> Result<(), Refused> {
    let count = || {
        state
            .store
            .inbox_count(ws, now())
            .map_err(|e| failed("post_inbox", id, &e))
    };
    if count()? < INBOX_MAX_PER_WS {
        return Ok(());
    }
    state
        .store
        .sweep(now())
        .map_err(|e| failed("post_inbox", id, &e))?;
    let rows = count()?;
    if rows < INBOX_MAX_PER_WS {
        return Ok(());
    }
    eprintln!(
        "store full post_inbox {}: workspace holds {rows} of {INBOX_MAX_PER_WS} rows",
        short(id)
    );
    Err(Refused::Full)
}
