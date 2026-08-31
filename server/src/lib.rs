//! pls,fix host: the panes' static assets plus the encrypted link
//! relay under `/api`. Owns the router and the cache-control middleware.
//! Invariant: the relay stores ciphertext and `sha256(authKey)` only, so the
//! Cloudflare Access bypass on this path never exposes readable content.

mod fetch;
pub mod limits;
pub mod relay;
mod relay_gates;
mod relay_inbox;
mod relay_touch;
pub mod store;
mod store_inbox;

use std::{path::PathBuf, sync::Arc};

use axum::{
    extract::{Request, State},
    http::{header, HeaderValue},
    middleware::{self, Next},
    response::{Json, Redirect, Response},
    routing::get,
    Router,
};
use tower_http::services::ServeDir;

use crate::relay::AppState;

/// The gateway parses `version`, so that field never changes shape. `relay` is
/// counted on request and is what the VPS is watched by: rows held, the links
/// behind them and how close the store is to its ceiling. A store that cannot
/// be counted still answers, with nulls, because `/version` is also the health
/// check the suite polls.
async fn version(State(state): State<Arc<AppState>>) -> Json<serde_json::Value> {
    let counts = state.store.counts().ok();
    Json(serde_json::json!({
        "version": env!("CARGO_PKG_VERSION"),
        "source": option_env!("PLSFIX_VERSION").unwrap_or("dev"),
        "relay": {
            "links": counts.as_ref().map(|counts| counts.links),
            "revisions": counts.as_ref().map(|counts| counts.revisions),
            "inbox": counts.as_ref().map(|counts| counts.inbox),
            "bytes": counts.as_ref().map(|counts| counts.bytes),
            "max_bytes": state.max_bytes,
        },
    }))
}

fn healthz() -> Json<serde_json::Value> {
    Json(serde_json::json!({ "ok": true }))
}

// Office webviews cache aggressively, so a JS-only redeploy must reach them:
// everything is no-cache except vite's hashed bundles, which never change
// under the same name (only icons share `/assets/`, and they are not .js/.css).
async fn cache_control(request: Request, next: Next) -> Response {
    let hashed = {
        let path = request.uri().path();
        path.starts_with("/assets/") && (path.ends_with(".js") || path.ends_with(".css"))
    };
    let mut response = next.run(request).await;
    let value = if hashed {
        "public, max-age=31536000, immutable"
    } else {
        "no-cache"
    };
    response
        .headers_mut()
        .insert(header::CACHE_CONTROL, HeaderValue::from_static(value));
    response
}

/// Suite endpoints, the relay and the built panes, in that order.
pub fn app(static_dir: PathBuf, state: Arc<AppState>) -> Router {
    Router::new()
        .route("/healthz", get(|| async { healthz() }))
        .route("/version", get(version))
        .route("/", get(|| async { Redirect::temporary("taskpane.html") }))
        .with_state(state.clone())
        .merge(relay::routes(state))
        .fallback_service(ServeDir::new(static_dir))
        .layer(middleware::from_fn(cache_control))
}
