//! Static host for the Model Tools Excel add-in pane.
//!
//! Serves the built pane (`dist/` locally, `static/` on the server) plus the
//! suite's `/healthz` and `/version`; the gateway strips the public
//! `/modelis/` prefix before proxying here, and a Cloudflare Access bypass
//! keeps the path reachable for Excel's webview (auth walls break panes).

use std::{env, net::SocketAddr, path::PathBuf};

use axum::{
    extract::Request,
    http::{header, HeaderValue},
    middleware::{self, Next},
    response::{Json, Redirect, Response},
    routing::get,
    Router,
};
use tower_http::services::ServeDir;

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
// under the same name.
async fn cache_control(request: Request, next: Next) -> Response {
    let hashed = {
        let path = request.uri().path();
        path.contains("/assets/taskpane-")
            && (path.ends_with(".js") || path.ends_with(".css"))
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

fn app(static_dir: PathBuf) -> Router {
    Router::new()
        .route("/healthz", get(|| async { healthz() }))
        .route("/version", get(|| async { version() }))
        .route("/", get(|| async { Redirect::temporary("taskpane.html") }))
        .fallback_service(ServeDir::new(static_dir))
        .layer(middleware::from_fn(cache_control))
}

#[tokio::main]
async fn main() {
    let port: u16 = env::var("MODELIS_PORT")
        .ok()
        .and_then(|value| value.parse().ok())
        .unwrap_or(8804);
    let static_dir = PathBuf::from(
        env::var("MODELIS_STATIC").unwrap_or_else(|_| "dist".to_string()),
    );
    if !static_dir.join("taskpane.html").is_file() {
        eprintln!(
            "warning: {} has no taskpane.html - build the pane first (npm run build)",
            static_dir.display()
        );
    }

    let addr = SocketAddr::from(([127, 0, 0, 1], port));
    let listener = tokio::net::TcpListener::bind(addr)
        .await
        .unwrap_or_else(|error| panic!("cannot bind {addr}: {error}"));
    println!("plsfix-server serving {} on {addr}", static_dir.display());
    axum::serve(listener, app(static_dir))
        .await
        .expect("server error");
}

#[cfg(test)]
mod tests {
    use super::*;
    use axum::body::Body;
    use axum::http::{Request, StatusCode};
    use http_body_util::BodyExt;
    use tower::ServiceExt;

    fn test_dir() -> PathBuf {
        let dir = std::env::temp_dir().join(format!(
            "modelis-test-{}",
            std::process::id()
        ));
        let assets = dir.join("assets");
        std::fs::create_dir_all(&assets).unwrap();
        std::fs::write(dir.join("taskpane.html"), "<title>Model Tools</title>").unwrap();
        std::fs::write(dir.join("shortcuts.json"), "{\"actions\":[]}").unwrap();
        std::fs::write(assets.join("taskpane-Bfs8s79m.js"), "// bundle").unwrap();
        std::fs::write(assets.join("icon-32.png"), [0x89, 0x50]).unwrap();
        dir
    }

    async fn call(path: &str) -> (StatusCode, Option<String>, String) {
        let response = app(test_dir())
            .oneshot(Request::get(path).body(Body::empty()).unwrap())
            .await
            .unwrap();
        let status = response.status();
        let cache = response
            .headers()
            .get(header::CACHE_CONTROL)
            .map(|value| value.to_str().unwrap().to_string());
        let body = response.into_body().collect().await.unwrap().to_bytes();
        (status, cache, String::from_utf8_lossy(&body).to_string())
    }

    #[tokio::test]
    async fn healthz_is_suite_shaped() {
        let (status, _, body) = call("/healthz").await;
        assert_eq!(status, StatusCode::OK);
        assert!(body.contains("\"ok\":true"));
    }

    #[tokio::test]
    async fn version_names_a_version() {
        let (status, _, body) = call("/version").await;
        assert_eq!(status, StatusCode::OK);
        assert!(body.contains("\"version\":\"1.1.0\""));
    }

    #[tokio::test]
    async fn pane_and_shortcuts_are_served_uncached() {
        let (status, cache, body) = call("/taskpane.html").await;
        assert_eq!(status, StatusCode::OK);
        assert!(body.contains("Model Tools"));
        assert_eq!(cache.as_deref(), Some("no-cache"));

        let (status, cache, _) = call("/shortcuts.json").await;
        assert_eq!(status, StatusCode::OK);
        assert_eq!(cache.as_deref(), Some("no-cache"));
    }

    #[tokio::test]
    async fn hashed_bundles_are_immutable_but_icons_are_not() {
        let (status, cache, _) = call("/assets/taskpane-Bfs8s79m.js").await;
        assert_eq!(status, StatusCode::OK);
        assert_eq!(
            cache.as_deref(),
            Some("public, max-age=31536000, immutable")
        );

        let (_, cache, _) = call("/assets/icon-32.png").await;
        assert_eq!(cache.as_deref(), Some("no-cache"));
    }

    #[tokio::test]
    async fn root_redirects_to_the_pane_relatively() {
        let response = app(test_dir())
            .oneshot(Request::get("/").body(Body::empty()).unwrap())
            .await
            .unwrap();
        assert_eq!(response.status(), StatusCode::TEMPORARY_REDIRECT);
        // Relative, so the redirect survives the gateway's /modelis/ prefix.
        assert_eq!(
            response.headers().get(header::LOCATION).unwrap(),
            "taskpane.html"
        );
    }

    #[tokio::test]
    async fn missing_files_are_404() {
        let (status, _, _) = call("/nope.html").await;
        assert_eq!(status, StatusCode::NOT_FOUND);
    }
}
