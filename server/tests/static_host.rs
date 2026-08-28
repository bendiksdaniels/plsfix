//! Static host tests: the suite endpoints, the pane files and the cache rule
//! that lets an Office webview pick up a redeploy (hashed bundles immutable,
//! everything else no-cache).

use std::{path::PathBuf, sync::Arc};

use plsfix_server::{app, relay::AppState, store::Store};

mod tests {
    use super::*;
    use axum::body::Body;
    use axum::http::{header, Request, StatusCode};
    use http_body_util::BodyExt;
    use tower::ServiceExt;

    fn test_dir() -> PathBuf {
        let dir = std::env::temp_dir().join(format!("modelis-test-{}", std::process::id()));
        let assets = dir.join("assets");
        std::fs::create_dir_all(&assets).unwrap();
        std::fs::write(
            dir.join("taskpane.html"),
            "<title>Model Tools</title>",
        )
        .unwrap();
        std::fs::write(dir.join("shortcuts.json"), "{\"actions\":[]}").unwrap();
        std::fs::write(assets.join("taskpane-Bfs8s79m.js"), "// bundle").unwrap();
        std::fs::write(assets.join("pptpane-abc.js"), "// bundle").unwrap();
        std::fs::write(assets.join("icon-32.png"), [0x89, 0x50]).unwrap();
        dir
    }

    fn test_state() -> Arc<AppState> {
        Arc::new(AppState {
            store: Store::in_memory().unwrap(),
        })
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
}
