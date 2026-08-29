//! pls,fix host: the panes' static assets plus the encrypted link
//! relay under `/api`. Owns the router and the cache-control middleware.
//! Invariant: the relay stores ciphertext and `sha256(authKey)` only, so the
//! Cloudflare Access bypass on this path never exposes readable content.

mod fetch;
pub mod relay;
pub mod store;

use std::{path::PathBuf, sync::Arc};

use axum::{
    extract::Request,
    http::{header, HeaderValue},
    middleware::{self, Next},
    response::{Json, Redirect, Response},
    routing::get,
    Router,
};
use tower_http::services::ServeDir;

use crate::relay::AppState;

fn version() -> Json<serde_json::Value> {
    Json(serde_json::json!({
        "version": env!("CARGO_PKG_VERSION"),
        "source": option_env!("PLSFIX_VERSION").unwrap_or("dev"),
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
        .route("/version", get(|| async { version() }))
        .route("/", get(|| async { Redirect::temporary("taskpane.html") }))
        .merge(relay::routes(state))
        .fallback_service(ServeDir::new(static_dir))
        .layer(middleware::from_fn(cache_control))
}
