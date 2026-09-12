//! pls,fix host: the panes' static assets plus the encrypted link
//! relay under `/api`. Owns the router and the cache-control middleware.
//! Invariant: the relay stores ciphertext and `sha256(authKey)` only, so the
//! Cloudflare Access bypass on this path never exposes readable content; the
//! one document served on purpose is `/manifest.xml`, made to be handed out.

mod blocking;
mod fetch;
pub mod limits;
pub mod manifest;
pub mod relay;
mod relay_gates;
mod relay_inbox;
mod relay_touch;
pub mod store;
pub mod store_inbox;
pub mod store_room;

use std::{path::PathBuf, sync::Arc, time::Duration};

use axum::{
    extract::{Request, State},
    http::{header, HeaderValue, StatusCode},
    middleware::{self, Next},
    response::{IntoResponse, Json, Redirect, Response},
    routing::get,
    Router,
};
use tower_http::{services::ServeDir, timeout::TimeoutLayer};

use crate::blocking::store_call;
use crate::relay::{now, AppState};
use crate::relay_gates::rate_limit;
use crate::store::Counts;

/// A stalled connection - a slow client, a wedged upstream - must not hold a
/// worker forever: past this, the layer below answers 408 on its own.
pub const REQUEST_TIMEOUT_SECS: u64 = 30;

/// The gateway parses `version`, so that field never changes shape. `relay` is
/// what the VPS is watched by: rows held, the links behind them and how close
/// the store is to its ceiling. A store that cannot be counted still answers,
/// with nulls, because `/version` is also the health check the suite polls.
async fn version(State(state): State<Arc<AppState>>) -> Json<serde_json::Value> {
    let counts = relay_counts(&state).await;
    Json(serde_json::json!({
        "version": env!("CARGO_PKG_VERSION"),
        "source": option_env!("PLSFIX_VERSION").unwrap_or("dev"),
        "relay": {
            "links": counts.as_ref().map(|counts| counts.links),
            "revisions": counts.as_ref().map(|counts| counts.revisions),
            "inbox": counts.as_ref().map(|counts| counts.inbox),
            "bytes": counts.as_ref().map(|counts| counts.bytes),
            "max_bytes": state.max_bytes(),
        },
    }))
}

/// Counting scans both tables under the connection mutex every push also
/// needs, and this route is anonymous, so the answer is reused for
/// `VERSION_COUNTS_TTL` seconds (I3 of the N1 security review).
async fn relay_counts(state: &Arc<AppState>) -> Option<Counts> {
    let at = now();
    if let Some(cached) = state.counts.get(at) {
        return Some(cached);
    }
    let counted = store_call(state, "counts", "version", |store| store.counts())
        .await
        .ok()?;
    state.counts.set(at, &counted);
    Some(counted)
}

fn healthz() -> Json<serde_json::Value> {
    Json(serde_json::json!({ "ok": true }))
}

/// The manifest to sideload for THIS host: the committed file, or the same
/// file re-pointed at `MODELIS_PUBLIC_URL` with an id of its own (`manifest.rs`).
/// A file that cannot be read is a 404 with the reason on stderr, never a
/// half manifest.
async fn manifest_xml(State(state): State<Arc<AppState>>) -> Response {
    match manifest::render(&state.manifest) {
        Ok(xml) => (
            [(header::CONTENT_TYPE, "application/xml; charset=utf-8")],
            xml,
        )
            .into_response(),
        Err(error) => {
            eprintln!(
                "manifest: cannot read {}: {error}",
                state.manifest.file.display()
            );
            StatusCode::NOT_FOUND.into_response()
        }
    }
}

// Office webviews cache aggressively, so a JS-only redeploy must reach them:
// everything is no-cache except vite's hashed bundles, which never change
// under the same name (only icons share `/assets/`, and they are not .js/.css).
// Sealed blobs are `no-store, private` instead: `no-cache` still permits a
// shared cache to keep the body, and one in front that ignored `Authorization`
// could hand one client's ciphertext to another (M4 of the N1 security
// review). The pane sends `If-None-Match` from its own state, so nothing in
// the 304 path depends on a cache holding the body.
async fn cache_control(request: Request, next: Next) -> Response {
    let path = request.uri().path();
    let hashed = path.starts_with("/assets/") && (path.ends_with(".js") || path.ends_with(".css"));
    let sealed = path.starts_with("/api/");
    let mut response = next.run(request).await;
    let value = match (hashed, sealed) {
        (_, true) => "no-store, private",
        (true, _) => "public, max-age=31536000, immutable",
        _ => "no-cache",
    };
    let headers = response.headers_mut();
    headers.insert(header::CACHE_CONTROL, HeaderValue::from_static(value));
    if sealed {
        headers.insert(header::VARY, HeaderValue::from_static("Authorization"));
    }
    response
}

// No path segment of a built pane starts with a dot, and this path carries the
// Access bypass, so a request for one is a 404 rather than content: ServeDir
// would hand back a `.env` or a `.git` file that ever landed in the deploy.
// The percent-encoded forms are folded first, because ServeDir decodes before
// it opens.
async fn no_dotfiles(request: Request, next: Next) -> Response {
    let path = request
        .uri()
        .path()
        .replace("%2e", ".")
        .replace("%2E", ".")
        .replace("%2f", "/")
        .replace("%2F", "/");
    if path.split('/').any(|segment| segment.starts_with('.')) {
        return StatusCode::NOT_FOUND.into_response();
    }
    next.run(request).await
}

/// Suite endpoints, the manifest, the relay and the built panes, in that
/// order, every request bounded by `timeout`. Split from `app` so a test can
/// pass a short deadline instead of waiting out the real one.
/// The three suite routes carry the read limiter too: they are anonymous, they
/// sit on the Access-bypassed path, and `/version` and `/manifest.xml` each do
/// real work (I3 of the N1 security review).
pub fn app_with_timeout(static_dir: PathBuf, state: Arc<AppState>, timeout: Duration) -> Router {
    let suite = Router::new()
        .route("/healthz", get(|| async { healthz() }))
        .route("/version", get(version))
        .route("/manifest.xml", get(manifest_xml))
        .layer(middleware::from_fn_with_state(state.clone(), rate_limit))
        .with_state(state.clone());
    Router::new()
        .route("/", get(|| async { Redirect::temporary("taskpane.html") }))
        .merge(suite)
        .merge(relay::routes(state))
        .fallback_service(ServeDir::new(static_dir))
        .layer(middleware::from_fn(no_dotfiles))
        .layer(middleware::from_fn(cache_control))
        .layer(TimeoutLayer::new(timeout))
}

/// Suite endpoints, the manifest, the relay and the built panes, in that
/// order, bounded by REQUEST_TIMEOUT_SECS.
pub fn app(static_dir: PathBuf, state: Arc<AppState>) -> Router {
    app_with_timeout(static_dir, state, Duration::from_secs(REQUEST_TIMEOUT_SECS))
}
