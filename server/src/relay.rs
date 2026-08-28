//! `/api` routes for the relay: sealed link revisions (4 MiB) and workspace
//! inbox items (64 KiB). Owns bearer extraction, id validation, the JSON error
//! shape and the ETag/304 contract; all state lives in `store`.
//! Invariant: a request is answered from the bearer's hash, never its key.

use std::sync::Arc;

use axum::{
    body::Bytes,
    extract::{DefaultBodyLimit, Path, State},
    http::{header, HeaderMap, StatusCode},
    response::{IntoResponse, Response},
    routing::{delete, get, post},
    Json, Router,
};
use base64::{engine::general_purpose::URL_SAFE_NO_PAD, Engine as _};
use serde::{Deserialize, Serialize};

use crate::store::{auth_hash, Delete, Found, Get, Put, StatusRow, Store};

/// A sealed picture is the big payload; 4 MiB covers a full-slide render.
const LINK_LIMIT: usize = 4 * 1024 * 1024;
/// Inbox items carry metadata only.
const INBOX_LIMIT: usize = 64 * 1024;
/// One poll covers a deck; a longer batch is a client bug.
const MAX_STATUS_ITEMS: usize = 200;
/// The inbox POST names its link here - the body is the sealed item.
const LINK_ID_HEADER: &str = "x-smt-link-id";

/// Everything the routes share: one store, one connection.
pub struct AppState {
    pub store: Store,
}

type Api = State<Arc<AppState>>;

/// Every refusal the routes can produce, mapped to code and message in one
/// place so the JSON error shape cannot drift between handlers.
enum Refused {
    Unauthorized,
    BadId,
    BadBody,
    TooManyItems,
    Forbidden,
    Missing,
    Store,
}

impl Refused {
    fn parts(&self) -> (StatusCode, &'static str) {
        match self {
            Refused::Unauthorized => (StatusCode::UNAUTHORIZED, "bearer required"),
            Refused::BadId => (StatusCode::BAD_REQUEST, "bad id"),
            Refused::BadBody => (StatusCode::BAD_REQUEST, "bad body"),
            Refused::TooManyItems => (StatusCode::BAD_REQUEST, "too many items"),
            Refused::Forbidden => (StatusCode::FORBIDDEN, "another key owns this"),
            Refused::Missing => (StatusCode::NOT_FOUND, "not found"),
            Refused::Store => (StatusCode::INTERNAL_SERVER_ERROR, "store error"),
        }
    }
}

impl IntoResponse for Refused {
    fn into_response(self) -> Response {
        let (status, message) = self.parts();
        (status, Json(serde_json::json!({ "error": message }))).into_response()
    }
}

/// Handlers answer with a body or with a refusal, never a panic.
type Reply = Result<Response, Refused>;

/// Wall clock in seconds; the store takes `now` so tests can travel in time.
pub fn now() -> i64 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|since| since.as_secs() as i64)
        .unwrap_or(0)
}

/// Two nested routers so the inbox keeps its own, much smaller body limit.
pub fn routes(state: Arc<AppState>) -> Router {
    let links = Router::new()
        .route("/:id", get(get_link).put(put_link).delete(delete_link))
        .route("/status", post(status))
        .layer(DefaultBodyLimit::max(LINK_LIMIT));
    let inbox = Router::new()
        .route("/:ws", get(list_inbox).post(post_inbox))
        .route("/:ws/:id", delete(delete_inbox))
        .layer(DefaultBodyLimit::max(INBOX_LIMIT));
    Router::new()
        .nest("/api/links", links)
        .nest("/api/inbox", inbox)
        .with_state(state)
}

fn ok_json() -> Response {
    Json(serde_json::json!({ "ok": true })).into_response()
}

/// A store failure is the one thing worth a log line: stage, id, cause.
fn failed(stage: &str, id: &str, error: &rusqlite::Error) -> Refused {
    let short: String = id.chars().take(8).collect();
    eprintln!("store {stage} {short}: {error}");
    Refused::Store
}

fn is_link_id(id: &str) -> bool {
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

fn workspace_auth(headers: &HeaderMap, ws: &str) -> Result<[u8; 32], Refused> {
    let auth = bearer(headers)?;
    if !is_key(ws) {
        return Err(Refused::BadId);
    }
    Ok(auth)
}

async fn put_link(
    State(state): Api,
    Path(id): Path<String>,
    headers: HeaderMap,
    body: Bytes,
) -> Reply {
    let auth = link_auth(&headers, &id)?;
    let put = state
        .store
        .put_link(&id, &auth, &body, now())
        .map_err(|error| failed("put_link", &id, &error))?;
    match put {
        Put::Created(rev) | Put::Updated(rev) => {
            Ok(Json(serde_json::json!({ "rev": rev })).into_response())
        }
        Put::Forbidden => Err(Refused::Forbidden),
    }
}

async fn get_link(State(state): Api, Path(id): Path<String>, headers: HeaderMap) -> Reply {
    let auth = link_auth(&headers, &id)?;
    let found = state
        .store
        .get_link(&id, &auth, now())
        .map_err(|error| failed("get_link", &id, &error))?;
    match found {
        Get::Found(found) => Ok(link_response(&headers, found)),
        Get::Forbidden => Err(Refused::Forbidden),
        Get::Missing => Err(Refused::Missing),
    }
}

/// `ETag: "<rev>"`, so a pane that already holds the newest rev gets a 304.
fn link_response(headers: &HeaderMap, found: Found) -> Response {
    let etag = format!("\"{}\"", found.rev);
    let known = headers
        .get(header::IF_NONE_MATCH)
        .and_then(|value| value.to_str().ok())
        .is_some_and(|value| value.split(',').any(|tag| tag.trim() == etag));
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
    let deleted = state
        .store
        .delete_link(&id, &auth, now())
        .map_err(|error| failed("delete_link", &id, &error))?;
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
    if items.len() > MAX_STATUS_ITEMS {
        return Err(Refused::TooManyItems);
    }
    let queries: Vec<(String, [u8; 32])> = items
        .into_iter()
        .map(|item| (item.id, auth_hash(&item.auth)))
        .collect();
    let rows = state
        .store
        .status(&queries, now())
        .map_err(|error| failed("status", "batch", &error))?;
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

#[derive(Serialize)]
struct InboxOut {
    id: String,
    #[serde(rename = "createdAt")]
    created_at: i64,
    blob: String,
}

async fn post_inbox(
    State(state): Api,
    Path(ws): Path<String>,
    headers: HeaderMap,
    body: Bytes,
) -> Reply {
    let auth = workspace_auth(&headers, &ws)?;
    let id = headers
        .get(LINK_ID_HEADER)
        .and_then(|value| value.to_str().ok())
        .filter(|id| is_link_id(id))
        .ok_or(Refused::BadId)?;
    let accepted = state
        .store
        .post_inbox(&ws, &auth, id, &body, now())
        .map_err(|error| failed("post_inbox", id, &error))?;
    if !accepted {
        return Err(Refused::Forbidden);
    }
    Ok(ok_json())
}

async fn list_inbox(State(state): Api, Path(ws): Path<String>, headers: HeaderMap) -> Reply {
    let auth = workspace_auth(&headers, &ws)?;
    let rows = state
        .store
        .list_inbox(&ws, &auth, now())
        .map_err(|error| failed("list_inbox", &ws, &error))?;
    let out: Vec<InboxOut> = rows
        .into_iter()
        .map(|row| InboxOut {
            id: row.id,
            created_at: row.created_at,
            blob: URL_SAFE_NO_PAD.encode(row.blob),
        })
        .collect();
    Ok(Json(out).into_response())
}

async fn delete_inbox(
    State(state): Api,
    Path((ws, id)): Path<(String, String)>,
    headers: HeaderMap,
) -> Reply {
    let auth = workspace_auth(&headers, &ws)?;
    if !is_link_id(&id) {
        return Err(Refused::BadId);
    }
    let deleted = state
        .store
        .delete_inbox(&ws, &auth, &id, now())
        .map_err(|error| failed("delete_inbox", &id, &error))?;
    match deleted {
        Delete::Deleted => Ok(ok_json()),
        Delete::Forbidden => Err(Refused::Forbidden),
        Delete::Missing => Err(Refused::Missing),
    }
}
