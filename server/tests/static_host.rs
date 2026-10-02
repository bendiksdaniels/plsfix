//! Static host tests: the suite endpoints, the pane files and the cache rule
//! that lets an Office webview pick up a redeploy (hashed bundles immutable,
//! everything else no-cache).

use std::{path::PathBuf, sync::Arc, sync::OnceLock};

use plsfix_server::{app, relay::AppState, relay::DEFAULT_MAX_BYTES, store::Store};

mod tests {
    use super::*;
    use axum::body::Body;
    use axum::http::{header, Request, StatusCode};
    use http_body_util::BodyExt;
    use tower::ServiceExt;

    /// One directory per test process, written exactly once: these tests run
    /// on separate threads, and re-writing a file while another thread's
    /// ServeDir was reading it served an empty pane every so often.
    fn test_dir() -> PathBuf {
        static DIR: OnceLock<PathBuf> = OnceLock::new();
        DIR.get_or_init(|| {
            let dir = std::env::temp_dir().join(format!("modelis-test-{}", std::process::id()));
            let assets = dir.join("assets");
            std::fs::create_dir_all(&assets).unwrap();
            std::fs::write(dir.join("taskpane.html"), "<title>pls,fix</title>").unwrap();
            std::fs::write(dir.join("shortcuts.json"), "{\"actions\":[]}").unwrap();
            std::fs::write(dir.join("support.html"), "<title>pls,fix support</title>").unwrap();
            std::fs::write(dir.join("privacy.html"), "<title>pls,fix privacy</title>").unwrap();
            std::fs::write(assets.join("taskpane-Bfs8s79m.js"), "// bundle").unwrap();
            std::fs::write(assets.join("pptpane-abc.js"), "// bundle").unwrap();
            std::fs::write(assets.join("icon-32.png"), [0x89, 0x50]).unwrap();
            dir
        })
        .clone()
    }

    fn test_state() -> Arc<AppState> {
        Arc::new(AppState::new(Store::in_memory().unwrap()))
    }

    async fn call(path: &str) -> (StatusCode, Option<String>, String) {
        let response = app(test_dir(), test_state())
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
        assert!(body.contains(&format!("\"version\":\"{}\"", env!("CARGO_PKG_VERSION"))));
    }

    /// The suite watches the VPS by this object: what the relay holds, and how
    /// close it is to the ceiling it will start refusing writes at.
    #[tokio::test]
    async fn version_counts_what_the_relay_holds() {
        let state = Arc::new(AppState::new(Store::in_memory().unwrap()));
        let auth = plsfix_server::store::auth_hash("key");
        let id = "0123456789abcdef0123456789abcdef";
        let now = plsfix_server::relay::now();
        state.store.put_link(id, &auth, b"12345", now).unwrap();
        state.store.put_link(id, &auth, b"12345", now).unwrap();
        state
            .store
            .post_inbox("ws", &auth, id, b"123", now)
            .unwrap();

        let response = app(test_dir(), state)
            .oneshot(Request::get("/version").body(Body::empty()).unwrap())
            .await
            .unwrap();
        let body = response.into_body().collect().await.unwrap().to_bytes();
        let json: serde_json::Value = serde_json::from_slice(&body).unwrap();
        // The gateway parses this one, so it keeps its shape whatever else moves.
        assert_eq!(json["version"], env!("CARGO_PKG_VERSION"));
        assert_eq!(json["relay"]["links"], 1);
        assert_eq!(json["relay"]["revisions"], 2);
        assert_eq!(json["relay"]["inbox"], 1);
        assert_eq!(json["relay"]["bytes"], 13);
        assert_eq!(json["relay"]["max_bytes"], DEFAULT_MAX_BYTES);
    }

    #[tokio::test]
    async fn pane_and_shortcuts_are_served_uncached() {
        let (status, cache, body) = call("/taskpane.html").await;
        assert_eq!(status, StatusCode::OK);
        assert!(body.contains("pls,fix"));
        assert_eq!(cache.as_deref(), Some("no-cache"));

        let (status, cache, _) = call("/shortcuts.json").await;
        assert_eq!(status, StatusCode::OK);
        assert_eq!(cache.as_deref(), Some("no-cache"));
    }

    #[tokio::test]
    async fn hashed_bundles_are_immutable_but_icons_are_not() {
        for bundle in ["/assets/taskpane-Bfs8s79m.js", "/assets/pptpane-abc.js"] {
            let (status, cache, _) = call(bundle).await;
            assert_eq!(status, StatusCode::OK);
            assert_eq!(
                cache.as_deref(),
                Some("public, max-age=31536000, immutable")
            );
        }

        let (_, cache, _) = call("/assets/icon-32.png").await;
        assert_eq!(cache.as_deref(), Some("no-cache"));
    }

    #[tokio::test]
    async fn assets_prefix_is_matched_not_just_contained() {
        // A path that merely embeds "/assets/" past its start (e.g. a bad
        // gateway prefix) must not be treated as a hashed, immutable bundle.
        let (_, cache, _) = call("/nope/assets/taskpane-x.js").await;
        assert_eq!(cache.as_deref(), Some("no-cache"));

        let (_, cache, _) = call("/assets/pptpane-x.js").await;
        assert_eq!(
            cache.as_deref(),
            Some("public, max-age=31536000, immutable")
        );
    }

    #[tokio::test]
    async fn root_redirects_to_the_pane_relatively() {
        let response = app(test_dir(), test_state())
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

    // public/support.html and public/privacy.html ride into dist/ with every
    // other static file (Vite copies public/ verbatim), so MODELIS_STATIC
    // serving them is the same ServeDir fallback as the pane itself - this
    // only proves the two names are not shadowed by a route or by no_dotfiles.
    #[tokio::test]
    async fn support_and_privacy_pages_are_served_as_html() {
        for path in ["/support.html", "/privacy.html"] {
            let response = app(test_dir(), test_state())
                .oneshot(Request::get(path).body(Body::empty()).unwrap())
                .await
                .unwrap();
            assert_eq!(response.status(), StatusCode::OK, "{path}");
            let content_type = response
                .headers()
                .get(header::CONTENT_TYPE)
                .and_then(|value| value.to_str().ok())
                .unwrap_or("")
                .to_string();
            assert!(
                content_type.starts_with("text/html"),
                "{path} content-type was {content_type}"
            );
        }
    }
}
