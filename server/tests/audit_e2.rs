//! The relay host on its own: it stands up on a data directory that does not
//! exist yet, says which path it could not open when it cannot, refuses every
//! `/api` route without a key, and serves nothing but what MODELIS_STATIC
//! holds. The relay's own contracts are in audit_e2_relay.rs.
//! Invariant: this path carries an Access bypass, so nothing readable may ever
//! be served from it without the key that sealed it.

mod common;

use std::sync::Arc;

use axum::http::{header, StatusCode};
use common::*;
use plsfix_server::{app, relay::AppState, store::Store};
use tower::ServiceExt;

/// A first deploy, or a volume mounted empty: the store creates its directory
/// and both panes come back from the static tree beside it.
#[tokio::test]
async fn an_empty_data_dir_is_created_and_both_panes_are_served() {
    let data = std::env::temp_dir()
        .join(format!("modelis-e2-data-{}", std::process::id()))
        .join("fresh")
        .join("relay");
    let _ = std::fs::remove_dir_all(&data);
    let database = data.join("relay.sqlite");
    let store = Store::open(&database).expect("a missing data dir is created, not a panic");
    assert!(database.is_file());
    // WAL, so a reader and the hourly sweeper do not lock each other out.
    assert!(data.join("relay.sqlite-wal").exists());

    let app = app(static_dir("panes"), Arc::new(AppState::new(store)));
    for (path, expect) in [
        ("/healthz", "\"ok\":true"),
        ("/version", "\"version\":\""),
        ("/taskpane.html", "pls,fix"),
        ("/pptpane.html", "pls,fix deck"),
    ] {
        let (status, body) = get(&app, path).await;
        assert_eq!(status, StatusCode::OK, "{path}");
        assert!(body.contains(expect), "{path}: {body}");
    }
    let _ = std::fs::remove_dir_all(data.parent().unwrap());
}

/// A data directory the unit cannot write to is the one startup failure that
/// must read as itself: every refusal names the path it could not open, so the
/// journal says which directory to fix instead of looping on a bare error.
#[test]
fn a_data_dir_that_cannot_be_written_names_itself() {
    let root = std::env::temp_dir().join(format!("modelis-e2-ro-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&root);
    std::fs::create_dir_all(&root).unwrap();
    let writable = std::fs::metadata(&root).unwrap().permissions();
    let mut locked = writable.clone();
    locked.set_readonly(true);
    std::fs::set_permissions(&root, locked).unwrap();

    // MODELIS_DATA pointing at a file rather than a directory is refused too.
    let file = std::env::temp_dir().join(format!("modelis-e2-file-{}", std::process::id()));
    std::fs::write(&file, "not a directory").unwrap();
    for (path, named) in [
        (root.join("relay.sqlite"), &root),
        (root.join("deep/relay.sqlite"), &root),
        (file.join("relay.sqlite"), &file),
    ] {
        let said = match Store::open(&path) {
            Ok(_) => panic!("{} opened a store it cannot write", path.display()),
            Err(error) => error.to_string(),
        };
        assert!(
            said.contains(named.to_str().unwrap()),
            "the refusal must name the path: {said}"
        );
    }

    std::fs::set_permissions(&root, writable).unwrap();
    let _ = std::fs::remove_dir_all(&root);
    let _ = std::fs::remove_file(&file);
}

/// Every route that can hand back stored bytes wants the key first. The three
/// batch routes carry a key per item instead of a bearer, so they answer - but
/// with nothing a caller without the key could read.
#[tokio::test]
async fn no_api_route_answers_with_content_without_the_key() {
    let app = relay_app();
    push(&app, b"sealed").await;

    for (method, path) in [
        ("GET", format!("/api/links/{ID}")),
        ("PUT", format!("/api/links/{ID}")),
        ("DELETE", format!("/api/links/{ID}")),
        ("GET", format!("/api/inbox/{AUTH}")),
        ("POST", format!("/api/inbox/{AUTH}")),
        ("DELETE", format!("/api/inbox/{AUTH}/{ID}")),
    ] {
        let (status, body) = call(&app, method, &path, None, b"x").await;
        assert_eq!(status, StatusCode::UNAUTHORIZED, "{method} {path}");
        assert_eq!(body, r#"{"error":"bearer required"}"#);
        // A bearer that is not this link's key never reads what is stored: the
        // link routes refuse it, and a foreign key's inbox listing is empty.
        let (status, body) = call(&app, method, &path, Some(OTHER), b"x").await;
        assert!(!body.contains("sealed"), "{method} {path}: {body}");
        assert!(
            status.is_client_error() || body == "[]" || body == r#"{"ok":true}"#,
            "{method} {path} answered {status} {body} to a foreign key"
        );
    }

    // The batch routes answer without a bearer, and say nothing but "not yours".
    for (route, expect) in [
        (
            "status",
            json(r#"[{"id":"","rev":null,"pushedAt":null,"error":"auth"}]"#),
        ),
        (
            "fetch",
            json(r#"{"items":[],"omitted":[{"id":"","reason":"auth"}]}"#),
        ),
        ("touch", json(r#"{"touched":0}"#)),
    ] {
        let path = format!("/api/links/{route}");
        let (status, body) = call(&app, "POST", &path, None, &one(OTHER)).await;
        assert_eq!(status, StatusCode::OK, "{route}");
        assert!(!body.contains("sealed"), "{route} handed back the blob");
        assert_eq!(json(&body.replace(ID, "")), expect, "{route}");
    }
}

/// The static host may only ever hand back what the build put in
/// MODELIS_STATIC: no dotfile, no directory listing, and nothing a `..`
/// (encoded or not) points at outside it.
#[tokio::test]
async fn the_static_host_serves_no_dotfiles_no_listing_and_nothing_above_it() {
    let app = app(
        static_dir("static"),
        Arc::new(AppState::new(Store::in_memory().unwrap())),
    );
    for path in [
        "/.env",
        "/%2Eenv",
        "/%2eenv",
        "/../outside.txt",
        "/%2e%2e/outside.txt",
        "/..%2foutside.txt",
        "/assets/",
        "/assets/../.env",
        "/assets%2f.env",
        "/assets%2F.env",
    ] {
        let (status, body) = get(&app, path).await;
        assert_eq!(status, StatusCode::NOT_FOUND, "{path} answered {status}");
        assert!(!body.contains("MODELIS_SECRET"), "{path} served a dotfile");
        assert!(!body.contains("not the pane"), "{path} escaped the dir");
    }
    // The pane itself still comes back, dots in a hashed bundle name included.
    assert_eq!(get(&app, "/taskpane.html").await.0, StatusCode::OK);
}

/// The cache rule lets an Office webview pick up a redeploy, and it must never
/// mark a relay answer cacheable: only the hashed bundles are immutable.
#[tokio::test]
async fn no_api_answer_is_ever_marked_cacheable() {
    let state = Arc::new(AppState::new(Store::in_memory().unwrap()));
    let auth = plsfix_server::store::auth_hash(AUTH);
    let now = plsfix_server::relay::now();
    state.store.put_link(ID, &auth, b"sealed", now).unwrap();
    let app = app(static_dir("cache"), state);
    for (method, path, body) in [
        ("GET", format!("/api/links/{ID}"), b"".as_slice()),
        ("PUT", format!("/api/links/{ID}"), b"x"),
        ("POST", "/api/links/status".to_string(), b"[]"),
        ("POST", "/api/links/fetch".to_string(), b"[]"),
        ("POST", "/api/links/touch".to_string(), b"[]"),
        ("GET", format!("/api/inbox/{AUTH}"), b""),
    ] {
        let response = app
            .clone()
            .oneshot(req(method, &path, Some(AUTH), body.to_vec()))
            .await
            .unwrap();
        let cache = response.headers().get(header::CACHE_CONTROL);
        assert_eq!(
            cache.map(|value| value.to_str().unwrap()),
            Some("no-cache"),
            "{method} {path}"
        );
    }
}
