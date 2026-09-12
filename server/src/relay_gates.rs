//! What a request passes before it reaches the store: the per-client rate
//! limit by request and, for a write, a second one by body size (the buckets
//! themselves live in `limits`). Owns the client key a request is counted
//! under, the two refusals and their log lines; the storage ceilings are the
//! store's own, checked inside the write it guards (`store_room`).
//! Invariant: a client is named once, by the middleware, and the bytes of its
//! body are charged to that same bucket.

use std::sync::Arc;

use axum::{
    extract::{Request, State},
    middleware::Next,
    response::{IntoResponse, Response},
};

use crate::limits::{client_key, is_write, token_cost, Allowed};
use crate::relay::{now, short, Api, AppState, Refused};

/// Who a request is counted as, resolved once and carried in the request
/// extensions so a handler charges its body to the bucket the request itself
/// was counted in.
#[derive(Clone, Debug)]
pub(crate) struct ClientKey(pub String);

/// One token per request, spent before the body is read: the write allowance
/// for a push or a delete, the read one for a GET or a batch. The refusal
/// names the bucket and the wait, never anything the request carried.
pub(crate) async fn rate_limit(State(state): Api, mut request: Request, next: Next) -> Response {
    let write = is_write(request.method(), request.uri().path());
    let key = client_key(
        state.trusted_proxy,
        request.headers(),
        request.extensions(),
    );
    let limiter = if write { &state.writes } else { &state.reads };
    match limiter.take(&key, now()) {
        Allowed::Yes => {
            request.extensions_mut().insert(ClientKey(key));
            next.run(request).await
        }
        Allowed::No { retry_after } => {
            refused(write, &key, retry_after);
            Refused::RateLimited(retry_after).into_response()
        }
    }
}

/// The second half of the write allowance: a token per KiB of body, charged
/// once the size is known. A minute of 300 pushes is 300 requests but it is
/// also up to 1.2 GiB, which is the whole store, so the requests alone were
/// never a budget (C1 of the N1 security review).
pub(crate) fn charge_bytes(
    state: &Arc<AppState>,
    client: Option<&ClientKey>,
    bytes: usize,
) -> Result<(), Refused> {
    let key = client.map_or("unknown", |client| client.0.as_str());
    match state.write_bytes.take_tokens(key, now(), token_cost(bytes)) {
        Allowed::Yes => Ok(()),
        Allowed::No { retry_after } => {
            refused(true, key, retry_after);
            Err(Refused::RateLimited(retry_after))
        }
    }
}

fn refused(write: bool, key: &str, retry_after: u64) {
    let kind = if write { "write" } else { "read" };
    eprintln!("relay rate {kind} {key}: retry after {retry_after}s");
}

/// The one log line a full store gets. The store itself decides there is no
/// room, inside the write's own lock; this is only how it reads in the
/// journal.
pub(crate) fn full(stage: &str, id: &str, state: &AppState) -> Refused {
    eprintln!(
        "store full {stage} {}: at the {} byte ceiling",
        short(id),
        state.max_bytes()
    );
    Refused::Full
}
