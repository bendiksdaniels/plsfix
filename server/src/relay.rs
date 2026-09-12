//! `/api` routes for the relay: sealed link revisions (4 MiB), workspace inbox
//! items (64 KiB) and the three bearer-less batches - status, fetch and touch
//! (64 KiB each; those handlers live in `fetch` and `relay_touch`). Owns
//! bearer extraction, id validation, the JSON error shape, the ETag/304
//! contract and the router the write gates in `relay_gates` wrap; the inbox
//! routes live in `relay_inbox`.
//! Invariant: a request is answered from the bearer's hash, never its key.

use std::{convert::Infallible, sync::Arc, time::Duration};

use axum::{
    body::Bytes,
    extract::{DefaultBodyLimit, Extension, Path, Query, State},
    http::{header, HeaderMap, StatusCode},
    middleware,
    response::{IntoResponse, Response},
    routing::{delete, get, post, put},
    Json, Router,
};
use serde::{Deserialize, Serialize};
use tower::{limit::GlobalConcurrencyLimitLayer, ServiceBuilder};
use tower_http::timeout::RequestBodyTimeoutLayer;

use crate::blocking::store_call;
use crate::limits::{RateLimiter, TrustedProxy};
use crate::manifest::ManifestSource;
use crate::relay_gates::{charge_bytes, full, rate_limit, ClientKey};
use crate::store::{auth_hash, Delete, Found, Get, Put, StatusRow, Store};
use crate::store_room::{Caps, CountsCache};
use crate::{fetch, relay_inbox, relay_touch};

/// A sealed picture is the big payload; 4 MiB covers a full-slide render.
const LINK_LIMIT: usize = 4 * 1024 * 1024;
/// Inbox items carry metadata only.
const INBOX_LIMIT: usize = 64 * 1024;
/// The batches are the three routes with no bearer (they carry a key per item)
/// and they sit behind the Access bypass, so they must not inherit the payload
/// limit: a full 200-item batch is about 20 KiB, and anything past this is
/// refused at the socket instead of parsed on the shared connection. The
/// blobs a fetch batch answers with are capped separately, in `fetch`.
const BATCH_LIMIT: usize = 64 * 1024;
/// One poll, fetch or touch covers a deck; a longer batch is a client bug.
pub(crate) const MAX_BATCH_ITEMS: usize = 200;
/// The inbox POST names its link here - the body is the sealed item.
pub(crate) const LINK_ID_HEADER: &str = "x-plsfix-link-id";
/// Blob bytes the store may hold before a write is refused (MODELIS_MAX_BYTES).
pub const DEFAULT_MAX_BYTES: i64 = 1024 * 1024 * 1024;
/// Writes are the expensive half - a push carries a picture - so they get the
/// smaller allowance (MODELIS_RATE_WRITE_PER_MIN). Held below
/// `DEFAULT_MAX_INFLIGHT_WRITES x 60 / DEFAULT_BODY_DEADLINE_SECS`, so a client
/// spending its whole minute on bodies that never arrive still cannot hold
/// every permit of the payload router (`relay_budget.rs` pins the arithmetic).
pub const DEFAULT_WRITE_PER_MIN: u32 = 100;
/// How long a write's body may go without a frame before the request is given
/// up on. It bounds how long one stalled client holds a permit, so it is well
/// inside `REQUEST_TIMEOUT_SECS`; a slow but steady upload never meets it,
/// because the deadline is per frame, not per body.
pub const DEFAULT_BODY_DEADLINE_SECS: u64 = 10;
/// Reads are polls and batches (MODELIS_RATE_READ_PER_MIN).
pub const DEFAULT_READ_PER_MIN: u32 = 1200;
/// A second write budget, charged by the body rather than by the request
/// (MODELIS_RATE_WRITE_KIB_PER_MIN): 64 MiB a minute per client. A working day
/// of exports is far below it - a few hundred pushes of 50-500 KiB and the odd
/// 1 MiB chart picture is single-digit MiB a minute - while a client trying to
/// fill a 1 GiB store now needs a quarter of an hour instead of a minute.
pub const DEFAULT_WRITE_KIB_PER_MIN: u32 = 64 * 1024;
/// Writes the payload router runs at once (MODELIS_MAX_INFLIGHT_WRITES). Each
/// one buffers its body in memory, so this is what bounds the relay's RAM.
pub const DEFAULT_MAX_INFLIGHT_WRITES: usize = 32;
/// Live inbox rows one workspace may hold. A deck takes its items within
/// minutes, so a workspace this deep is a loop, not a busy week.
pub const INBOX_MAX_PER_WS: i64 = 500;

/// Everything the routes share: one store, the storage ceiling, the three rate
/// limiters, which forwarding header names a client, how many writes may run
/// at once, the cached `/version` counters and the manifest source. Built by
/// `AppState::new`, so a new limit cannot be forgotten at one call site;
/// `main` overrides the fields from the environment.
pub struct AppState {
    pub store: Store,
    pub writes: RateLimiter,
    pub write_bytes: RateLimiter,
    pub reads: RateLimiter,
    pub trusted_proxy: TrustedProxy,
    pub max_inflight_writes: usize,
    pub body_deadline: Duration,
    pub counts: CountsCache,
    pub manifest: ManifestSource,
}

impl AppState {
    pub fn new(store: Store) -> AppState {
        let mut store = store;
        store.set_caps(Caps {
            max_bytes: DEFAULT_MAX_BYTES,
            inbox_rows: INBOX_MAX_PER_WS,
        });
        AppState {
            store,
            writes: RateLimiter::new(DEFAULT_WRITE_PER_MIN),
            write_bytes: RateLimiter::new(DEFAULT_WRITE_KIB_PER_MIN),
            reads: RateLimiter::new(DEFAULT_READ_PER_MIN),
            trusted_proxy: TrustedProxy::default(),
            max_inflight_writes: DEFAULT_MAX_INFLIGHT_WRITES,
            body_deadline: Duration::from_secs(DEFAULT_BODY_DEADLINE_SECS),
            counts: CountsCache::default(),
            manifest: ManifestSource::default(),
        }
    }

    /// The storage ceiling this relay runs with: the store enforces it inside
    /// each write, `/version` reports it, and it lives in exactly one place.
    pub fn max_bytes(&self) -> i64 {
        self.store.caps().max_bytes
    }

    /// Writes admitted at once, never zero: a limit of nothing would admit
    /// nothing, which is a silently bricked relay rather than a tight one.
    /// Clamped like the rate allowances beside it.
    pub fn inflight_writes(&self) -> usize {
        self.max_inflight_writes.max(1)
    }
}

pub(crate) type Api = State<Arc<AppState>>;

/// Which refusal a route answers with is decided here; how it reads is
/// `refused.rs`, re-exported so every handler keeps one import path.
pub(crate) use crate::refused::Refused;

/// Handlers answer with a body or with a refusal, never a panic.
pub(crate) type Reply = Result<Response, Refused>;

/// Wall clock in seconds; the store takes `now` so tests can travel in time.
pub fn now() -> i64 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|since| since.as_secs() as i64)
        .unwrap_or(0)
}

/// One router per body limit, so only the payload route carries 4 MiB. The two
/// payload METHODS carry more: each write in flight buffers its whole body in
/// memory, so without a concurrency limit the practical ceiling is RAM rather
/// than `MODELIS_MAX_BYTES` (I2 of the N1 security review), and a body that
/// stops arriving would hold its permit until the 30 s request deadline, so it
/// gets a deadline of its own inside that limit. A link GET carries no body
/// and never queues behind either.
pub fn routes(state: Arc<AppState>) -> Router {
    // Outermost first: the permit is taken, and only then does the body get
    // its deadline, so a stalled body cannot hold a permit past that deadline.
    // GLOBAL, because a `MethodRouter` layers each method separately and a
    // plain `ConcurrencyLimitLayer` would hand PUT and DELETE a pool each.
    let guard = ServiceBuilder::new()
        .layer(GlobalConcurrencyLimitLayer::new(state.inflight_writes()))
        .layer(RequestBodyTimeoutLayer::new(state.body_deadline));
    let writes = put(put_link)
        .delete(delete_link)
        .layer::<_, Infallible>(guard);
    let links = Router::new()
        .route("/:id", get(get_link).merge(writes))
        .layer(DefaultBodyLimit::max(LINK_LIMIT));
    let batch = Router::new()
        .route("/status", post(status))
        .route("/fetch", post(fetch::fetch))
        .route("/touch", post(relay_touch::touch))
        .layer(DefaultBodyLimit::max(BATCH_LIMIT));
    let inbox = Router::new()
        .route(
            "/:ws",
            get(relay_inbox::list_inbox).post(relay_inbox::post_inbox),
        )
        .route("/:ws/:id", delete(relay_inbox::delete_inbox))
        .layer(DefaultBodyLimit::max(INBOX_LIMIT));
    Router::new()
        .nest("/api/links", links.merge(batch))
        .nest("/api/inbox", inbox)
        .layer(middleware::from_fn_with_state(state.clone(), rate_limit))
        .with_state(state)
}

pub(crate) fn ok_json() -> Response {
    Json(serde_json::json!({ "ok": true })).into_response()
}

/// A store failure is the one thing worth a log line: stage, id, cause.
pub(crate) fn failed(stage: &str, id: &str, error: &rusqlite::Error) -> Refused {
    eprintln!("store {stage} {}: {error}", short(id));
    Refused::Store
}

/// Ids are logged short: eight characters name a link in a log line without
/// writing the whole handle of a deck's picture into the journal.
pub(crate) fn short(id: &str) -> String {
    id.chars().take(8).collect()
}

pub(crate) fn is_link_id(id: &str) -> bool {
    id.len() == 32
        && id
            .bytes()
            .all(|byte| matches!(byte, b'0'..=b'9' | b'a'..=b'f'))
}

/// Workspace ids and bearer keys are both 43 base64url characters (32 bytes).
fn is_key(value: &str) -> bool {
    value.len() == 43
        && value
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || byte == b'-' || byte == b'_')
}

/// The bearer's hash, or 401 - the key itself is never stored or logged.
fn bearer(headers: &HeaderMap) -> Result<[u8; 32], Refused> {
    headers
        .get(header::AUTHORIZATION)
        .and_then(|value| value.to_str().ok())
        .and_then(|value| value.strip_prefix("Bearer "))
        .filter(|key| is_key(key))
        .map(auth_hash)
        .ok_or(Refused::Unauthorized)
}

fn link_auth(headers: &HeaderMap, id: &str) -> Result<[u8; 32], Refused> {
    let auth = bearer(headers)?;
    if !is_link_id(id) {
        return Err(Refused::BadId);
    }
    Ok(auth)
}

pub(crate) fn workspace_auth(headers: &HeaderMap, ws: &str) -> Result<[u8; 32], Refused> {
    let auth = bearer(headers)?;
    if !is_key(ws) {
        return Err(Refused::BadId);
    }
    Ok(auth)
}

async fn put_link(
    State(state): Api,
    Path(id): Path<String>,
    client: Option<Extension<ClientKey>>,
    headers: HeaderMap,
    body: Bytes,
) -> Reply {
    let auth = link_auth(&headers, &id)?;
    charge_bytes(&state, client.as_deref(), body.len())?;
    let (key, blob, at) = (id.clone(), body.clone(), now());
    let put = store_call(&state, "put_link", &id, move |store| {
        store.put_link(&key, &auth, &blob, at)
    })
    .await?;
    match put {
        Put::Created(rev) | Put::Updated(rev) => {
            Ok(Json(serde_json::json!({ "rev": rev })).into_response())
        }
        Put::Forbidden => Err(Refused::Forbidden),
        Put::Full => Err(full("put_link", &id, &state)),
    }
}

/// `?rev=<n>` on the GET, which is how "Revert last update" asks for the
/// revision below the head. Absent means "the newest one" as before.
#[derive(Deserialize)]
struct RevQuery {
    rev: Option<String>,
}

// Parsed here rather than by serde, so a rev that is not a positive integer
// answers with the JSON error shape every other refusal uses - and never
// falls back to the head, which would silently repaint what the user is
// trying to undo.
fn wanted_rev(query: &RevQuery) -> Result<Option<i64>, Refused> {
    let Some(text) = query.rev.as_deref() else {
        return Ok(None);
    };
    match text.parse::<i64>() {
        Ok(rev) if rev >= 1 => Ok(Some(rev)),
        _ => Err(Refused::BadRev),
    }
}

async fn get_link(
    State(state): Api,
    Path(id): Path<String>,
    Query(query): Query<RevQuery>,
    headers: HeaderMap,
) -> Reply {
    let auth = link_auth(&headers, &id)?;
    let (key, wanted, at) = (id.clone(), wanted_rev(&query)?, now());
    let found = store_call(&state, "get_link", &id, move |store| match wanted {
        Some(rev) => store.get_link_rev(&key, &auth, rev, at),
        None => store.get_link(&key, &auth, at),
    })
    .await?;
    match found {
        Get::Found(found) => Ok(link_response(&headers, found)),
        Get::Forbidden => Err(Refused::Forbidden),
        Get::Missing => Err(Refused::Missing),
    }
}

/// `ETag: "<rev>"`, so a pane that already holds the newest rev gets a 304.
/// Compared weakly (RFC 7232 §3.2): Cloudflare hands the webview `W/"3"`
/// whenever it compresses the body, and that is the tag the webview sends back.
fn link_response(headers: &HeaderMap, found: Found) -> Response {
    let etag = format!("\"{}\"", found.rev);
    let known = headers
        .get(header::IF_NONE_MATCH)
        .and_then(|value| value.to_str().ok())
        .is_some_and(|value| {
            value
                .split(',')
                .any(|tag| tag.trim().trim_start_matches("W/") == etag)
        });
    if known {
        return (StatusCode::NOT_MODIFIED, [(header::ETAG, etag)]).into_response();
    }
    (
        [
            (header::ETAG, etag),
            (header::CONTENT_TYPE, "application/octet-stream".to_string()),
        ],
        found.blob,
    )
        .into_response()
}

async fn delete_link(State(state): Api, Path(id): Path<String>, headers: HeaderMap) -> Reply {
    let auth = link_auth(&headers, &id)?;
    let (key, at) = (id.clone(), now());
    let deleted = store_call(&state, "delete_link", &id, move |store| {
        store.delete_link(&key, &auth, at)
    })
    .await?;
    match deleted {
        Delete::Deleted => Ok(ok_json()),
        Delete::Forbidden => Err(Refused::Forbidden),
        Delete::Missing => Err(Refused::Missing),
    }
}

#[derive(Deserialize)]
struct StatusQuery {
    id: String,
    auth: String,
}

#[derive(Serialize)]
struct StatusOut {
    id: String,
    rev: Option<i64>,
    #[serde(rename = "pushedAt")]
    pushed_at: Option<i64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    error: Option<&'static str>,
}

/// The batch poll carries a key per item, so it needs no bearer header.
async fn status(State(state): Api, body: Bytes) -> Reply {
    let items: Vec<StatusQuery> = serde_json::from_slice(&body).map_err(|_| Refused::BadBody)?;
    if items.len() > MAX_BATCH_ITEMS {
        return Err(Refused::TooManyItems);
    }
    let queries: Vec<(String, [u8; 32])> = items
        .into_iter()
        .map(|item| (item.id, auth_hash(&item.auth)))
        .collect();
    let at = now();
    let rows = store_call(&state, "status", "batch", move |store| {
        store.status(&queries, at)
    })
    .await?;
    let out: Vec<StatusOut> = rows.into_iter().map(status_out).collect();
    Ok(Json(out).into_response())
}

fn status_out(row: StatusRow) -> StatusOut {
    StatusOut {
        id: row.id,
        rev: row.rev,
        pushed_at: row.pushed_at,
        error: row.auth_error.then_some("auth"),
    }
}
