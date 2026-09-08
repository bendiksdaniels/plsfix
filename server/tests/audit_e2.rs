//! Independence and contract tests for the relay host: the server stands up on
//! an empty data dir, refuses every `/api` route without a key, serves only
//! what `MODELIS_STATIC` holds, and keeps the ETag, inbox-delete and batch-cap
//! contracts the pane's client is written against.
//! Invariant: nothing readable reaches this path without the key that sealed it.

use std::{path::PathBuf, sync::Arc};

use axum::{
    body::Body,
    http::{header, Request, StatusCode},
    Router,
};
use http_body_util::BodyExt;
use plsfix_server::{app, limits::RateLimiter, relay::*, store::Store};
use tower::ServiceExt;

const ID: &str = "0123456789abcdef0123456789abcdef";
const AUTH: &str = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"; // 43 chars
const OTHER: &str = "BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB";

fn relay_app() -> Router {
    routes(Arc::new(AppState::new(Store::in_memory().unwrap())))
}

fn req(method: &str, path: &str, auth: Option<&str>, body: Vec<u8>) -> Request<Body> {
    let mut builder = Request::builder().method(method).uri(path);
    if let Some(auth) = auth {
        builder = builder.header(header::AUTHORIZATION, format!("Bearer {auth}"));
    }
    if method == "POST" && path.starts_with("/api/inbox/") {
        builder = builder.header("X-PLSFIX-Link-Id", ID);
    }
    builder.body(Body::from(body)).unwrap()
}

async fn send(app: &Router, request: Request<Body>) -> (StatusCode, String) {
    let response = app.clone().oneshot(request).await.unwrap();
    let status = response.status();
    let body = response.into_body().collect().await.unwrap().to_bytes();
    (status, String::from_utf8_lossy(&body).to_string())
}

async fn push(app: &Router, blob: Vec<u8>) {
    let (status, _) = send(
        app,
        req("PUT", &format!("/api/links/{ID}"), Some(AUTH), blob),
    )
    .await;
    assert_eq!(status, StatusCode::OK);
}

/// A static tree of its own per test, so one test's dotfile cannot be another's
/// missing file. The sibling outside it is what a traversal would reach.
fn static_dir(name: &str) -> PathBuf {
    let root = std::env::temp_dir().join(format!("modelis-e2-{}-{name}", std::process::id()));
    let dir = root.join("dist");
    std::fs::create_dir_all(dir.join("assets")).unwrap();
    std::fs::write(dir.join("taskpane.html"), "<title>pls,fix</title>").unwrap();
    std::fs::write(dir.join("pptpane.html"), "<title>pls,fix deck</title>").unwrap();
    std::fs::write(dir.join(".env"), "MODELIS_SECRET=1").unwrap();
    std::fs::write(root.join("outside.txt"), "not the pane").unwrap();
    dir
}

/// The relay comes up on a data directory that does not exist yet - a first
/// deploy, or a volume mounted empty - and serves both panes plus the two
/// suite endpoints from it.
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
        let (status, body) = send(&app, Request::get(path).body(Body::empty()).unwrap()).await;
        assert_eq!(status, StatusCode::OK, "{path}");
        assert!(body.contains(expect), "{path}: {body}");
    }
    let _ = std::fs::remove_dir_all(data.parent().unwrap());
}

/// Every route that can hand back stored bytes wants the key first. The three
/// batch routes carry a key per item instead of a bearer, so they answer - but
/// with nothing a caller without the key could read.
#[tokio::test]
async fn no_api_route_answers_with_content_without_the_key() {
    let app = relay_app();
    push(&app, b"sealed".to_vec()).await;

    for (method, path) in [
        ("GET", format!("/api/links/{ID}")),
        ("PUT", format!("/api/links/{ID}")),
        ("DELETE", format!("/api/links/{ID}")),
        ("GET", format!("/api/inbox/{AUTH}")),
        ("POST", format!("/api/inbox/{AUTH}")),
        ("DELETE", format!("/api/inbox/{AUTH}/{ID}")),
    ] {
        let (status, body) = send(&app, req(method, &path, None, vec![1])).await;
        assert_eq!(status, StatusCode::UNAUTHORIZED, "{method} {path}");
        assert_eq!(body, r#"{"error":"bearer required"}"#);
        // A bearer that is not this link's key never reads what is stored: the
        // link routes refuse it, and a foreign key's inbox listing is empty.
        let (status, body) = send(&app, req(method, &path, Some(OTHER), vec![1])).await;
        assert!(!body.contains("sealed"), "{method} {path}: {body}");
        assert!(
            status.is_client_error() || body == "[]" || body == r#"{"ok":true}"#,
            "{method} {path} answered {status} {body} to a foreign key"
        );
    }

    // The batch routes answer without a bearer, and say nothing but "not yours".
    let wrong = format!(r#"[{{"id":"{ID}","auth":"{OTHER}"}}]"#);
    for (route, expect) in [
        (
            "status",
            r#"[{"id":"0123456789abcdef0123456789abcdef","rev":null,"pushedAt":null,"error":"auth"}]"#,
        ),
        (
            "fetch",
            r#"{"items":[],"omitted":[{"id":"0123456789abcdef0123456789abcdef","reason":"auth"}]}"#,
        ),
        ("touch", r#"{"touched":0}"#),
    ] {
        let (status, body) = send(
            &app,
            req(
                "POST",
                &format!("/api/links/{route}"),
                None,
                wrong.clone().into_bytes(),
            ),
        )
        .await;
        assert_eq!(status, StatusCode::OK, "{route}");
        assert_eq!(body, expect, "{route}");
        assert!(!body.contains("sealed"), "{route} handed back the blob");
    }
}

/// The pane's path carries an Access bypass, so the static host may only ever
/// hand back what the build put in MODELIS_STATIC: no dotfile, no directory
/// listing, and nothing a `..` (encoded or not) points at outside it.
#[tokio::test]
async fn the_static_host_serves_no_dotfiles_no_listing_and_nothing_above_it() {
    let dir = static_dir("static");
    let app = app(dir, Arc::new(AppState::new(Store::in_memory().unwrap())));
    for path in [
        "/.env",
        "/%2Eenv",
        "/%2eenv",
        "/../outside.txt",
        "/%2e%2e/outside.txt",
        "/..%2foutside.txt",
        "/assets/",
        "/assets/../.env",
    ] {
        let (status, body) = send(&app, Request::get(path).body(Body::empty()).unwrap()).await;
        assert_eq!(status, StatusCode::NOT_FOUND, "{path} answered {status}");
        assert!(
            !body.contains("MODELIS_SECRET"),
            "{path} served the dotfile"
        );
        assert!(
            !body.contains("not the pane"),
            "{path} escaped the static dir"
        );
    }
    // The pane itself still comes back, dots in a hashed bundle name included.
    let (status, _) = send(
        &app,
        Request::get("/taskpane.html").body(Body::empty()).unwrap(),
    )
    .await;
    assert_eq!(status, StatusCode::OK);
}

/// RFC 7232 §3.2: If-None-Match compares weakly. Cloudflare hands the webview
/// a weak tag whenever it compresses the body, so the tag that comes back can
/// be `W/"3"` - and a deck that already holds that revision must still get a
/// 304 rather than the whole picture again.
#[tokio::test]
async fn a_weakened_if_none_match_is_still_the_revision_the_deck_holds() {
    let app = relay_app();
    push(&app, b"picture".to_vec()).await;
    for tag in ["\"1\"", "W/\"1\"", "\"0\", W/\"1\""] {
        let mut request = req("GET", &format!("/api/links/{ID}"), Some(AUTH), vec![]);
        request
            .headers_mut()
            .insert(header::IF_NONE_MATCH, tag.parse().unwrap());
        let (status, body) = send(&app, request).await;
        assert_eq!(status, StatusCode::NOT_MODIFIED, "{tag}");
        assert!(body.is_empty(), "{tag}");
    }
    // A revision the deck does not hold still arrives.
    let mut request = req("GET", &format!("/api/links/{ID}"), Some(AUTH), vec![]);
    request
        .headers_mut()
        .insert(header::IF_NONE_MATCH, "W/\"9\"".parse().unwrap());
    let (status, body) = send(&app, request).await;
    assert_eq!(status, StatusCode::OK);
    assert_eq!(body, "picture");
}

/// The inbox delete a deck makes after it has inserted an item: the second one
/// is a 404, and so is one from a key that never wrote the row. Neither is
/// something the deck can act on, which is why the client treats both as done
/// (see src/link/relay.ts deleteInbox).
#[tokio::test]
async fn deleting_an_inbox_item_that_is_already_gone_is_404() {
    let app = relay_app();
    let path = format!("/api/inbox/{AUTH}/{ID}");
    let (status, _) = send(
        &app,
        req(
            "POST",
            &format!("/api/inbox/{AUTH}"),
            Some(AUTH),
            b"item".to_vec(),
        ),
    )
    .await;
    assert_eq!(status, StatusCode::OK);
    let (status, _) = send(&app, req("DELETE", &path, Some(AUTH), vec![])).await;
    assert_eq!(status, StatusCode::OK);
    let (status, body) = send(&app, req("DELETE", &path, Some(AUTH), vec![])).await;
    assert_eq!(status, StatusCode::NOT_FOUND);
    assert_eq!(body, r#"{"error":"not found"}"#);
    // A foreign key sees the same answer as an absent row, never "not yours".
    send(
        &app,
        req(
            "POST",
            &format!("/api/inbox/{AUTH}"),
            Some(AUTH),
            b"item".to_vec(),
        ),
    )
    .await;
    let (status, other) = send(&app, req("DELETE", &path, Some(OTHER), vec![])).await;
    assert_eq!(status, StatusCode::NOT_FOUND);
    assert_eq!(other, body);
}

/// The two ceilings a deck's update-all is cut to: 200 queries per request,
/// and 4 MiB of blobs per answer - the item that does not fit travels alone,
/// and everything behind it is deferred rather than sieved.
#[tokio::test]
async fn a_batch_fetch_holds_at_two_hundred_items_and_four_mebibytes() {
    let app = relay_app();
    let second = "1123456789abcdef0123456789abcdef";
    push(&app, vec![7u8; 4 * 1024 * 1024]).await;
    send(
        &app,
        req("PUT", &format!("/api/links/{second}"), Some(AUTH), vec![1]),
    )
    .await;

    let both = format!(r#"[{{"id":"{ID}","auth":"{AUTH}"}},{{"id":"{second}","auth":"{AUTH}"}}]"#);
    let (status, body) = send(
        &app,
        req("POST", "/api/links/fetch", None, both.into_bytes()),
    )
    .await;
    assert_eq!(status, StatusCode::OK);
    let answer: serde_json::Value = serde_json::from_str(&body).unwrap();
    assert_eq!(answer["items"].as_array().unwrap().len(), 1);
    assert_eq!(answer["items"][0]["id"], ID);
    assert_eq!(answer["omitted"][0]["reason"], "deferred");

    for (count, expect) in [(200usize, StatusCode::OK), (201, StatusCode::BAD_REQUEST)] {
        let items: Vec<String> = (0..count)
            .map(|index| format!(r#"{{"id":"{index:032x}","auth":"{AUTH}"}}"#))
            .collect();
        let body = format!("[{}]", items.join(","));
        for route in ["fetch", "touch", "status"] {
            let (status, _) = send(
                &app,
                req(
                    "POST",
                    &format!("/api/links/{route}"),
                    None,
                    body.clone().into_bytes(),
                ),
            )
            .await;
            assert_eq!(status, expect, "{route} with {count} items");
        }
    }
}

/// The bucket a request is counted under is the forwarding header the gateway
/// passes through, which only Cloudflare may write: nginx listens on
/// 127.0.0.1:8750 and the relay on 127.0.0.1:8804, so the sole way in is the
/// tunnel, and Cloudflare overwrites CF-Connecting-IP on every request it
/// proxies. A peer that could reach the port directly would pick its own
/// bucket - which is what this pins, so exposing either port is a visible
/// change and not a silent one.
#[tokio::test]
async fn the_rate_bucket_follows_the_forwarding_header_the_gateway_passes_on() {
    let state = Arc::new(AppState {
        writes: RateLimiter::new(1),
        ..AppState::new(Store::in_memory().unwrap())
    });
    let app = routes(state);
    let mut codes = Vec::new();
    for client in ["9.9.9.1", "9.9.9.1", "9.9.9.2"] {
        let mut request = req(
            "PUT",
            &format!("/api/links/{ID}"),
            Some(AUTH),
            b"x".to_vec(),
        );
        request
            .headers_mut()
            .insert("cf-connecting-ip", client.parse().unwrap());
        codes.push(send(&app, request).await.0);
    }
    assert_eq!(
        codes,
        vec![
            StatusCode::OK,
            StatusCode::TOO_MANY_REQUESTS,
            StatusCode::OK
        ]
    );
}

/// `/version` is what the suite watches the relay by, and it counts what is
/// held right now: a sweep that drops a month-old link drops it from the
/// counters too.
#[tokio::test]
async fn version_counts_fall_when_the_sweeper_drops_dead_rows() {
    let state = Arc::new(AppState::new(Store::in_memory().unwrap()));
    let auth = plsfix_server::store::auth_hash(AUTH);
    let long_ago = now() - 40 * 24 * 3600;
    state.store.put_link(ID, &auth, b"12345", long_ago).unwrap();
    let app = app(static_dir("version"), state.clone());
    let read = |app: Router| async move {
        let (_, body) = send(&app, Request::get("/version").body(Body::empty()).unwrap()).await;
        serde_json::from_str::<serde_json::Value>(&body).unwrap()
    };
    let before = read(app.clone()).await;
    assert_eq!(before["relay"]["links"], 1);
    assert_eq!(before["relay"]["bytes"], 5);
    assert_eq!(state.store.sweep(now()).unwrap(), 1);
    let after = read(app).await;
    assert_eq!(after["relay"]["links"], 0);
    assert_eq!(after["relay"]["revisions"], 0);
    assert_eq!(after["relay"]["bytes"], 0);
}
