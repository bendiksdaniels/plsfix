//! The relay contracts the pane's client is written against: which revision an
//! answer belongs to, what a revert finds once a sweep has been through, what
//! an inbox delete says the second time, the two batch ceilings, the bucket a
//! request is rated under, and what `/version` counts.
//! Invariant: the revision a body carries is the one its ETag names.

mod common;

use std::sync::Arc;

use axum::http::{header, StatusCode};
use common::*;
use http_body_util::BodyExt;
use plsfix_server::{
    app,
    limits::{RateLimiter, TrustedProxy},
    relay::*,
    store::Store,
    store_room::VERSION_COUNTS_TTL,
};
use tower::ServiceExt;

/// RFC 7232 §3.2: If-None-Match compares weakly. Cloudflare hands the webview
/// a weak tag whenever it compresses the body, so the tag that comes back can
/// be `W/"3"` - and a deck that already holds that revision must still get a
/// 304 rather than the whole picture again.
#[tokio::test]
async fn a_weakened_if_none_match_is_still_the_revision_the_deck_holds() {
    let app = relay_app();
    push(&app, b"picture").await;
    let ask = |tag: &'static str| {
        let mut request = req("GET", &format!("/api/links/{ID}"), Some(AUTH), vec![]);
        request
            .headers_mut()
            .insert(header::IF_NONE_MATCH, tag.parse().unwrap());
        request
    };
    for tag in ["\"1\"", "W/\"1\"", "\"0\", W/\"1\""] {
        let (status, body) = send(&app, ask(tag)).await;
        assert_eq!(status, StatusCode::NOT_MODIFIED, "{tag}");
        assert!(body.is_empty(), "{tag}");
    }
    // A revision the deck does not hold still arrives.
    let (status, body) = send(&app, ask("W/\"9\"")).await;
    assert_eq!(status, StatusCode::OK);
    assert_eq!(body, "picture");
}

/// A push that lands between a deck's poll and its fetch moves the blob and
/// its revision together: the answer is never the new picture under the old
/// rev, which would leave a tag the next update reads as current.
#[tokio::test]
async fn a_push_between_a_poll_and_a_fetch_moves_the_blob_and_its_rev_together() {
    let app = relay_app();
    push(&app, b"one").await;
    let (_, polled) = call(&app, "POST", "/api/links/status", None, &one(AUTH)).await;
    assert_eq!(json(&polled)[0]["rev"], 1);

    // The workbook pushes again before the deck asks for the picture.
    push(&app, b"two").await;
    let (_, fetched) = call(&app, "POST", "/api/links/fetch", None, &one(AUTH)).await;
    assert_eq!(json(&fetched)["items"][0]["rev"], 2);
    assert_eq!(json(&fetched)["items"][0]["blob"], "dHdv"); // "two"

    // The single GET says the same thing in its ETag.
    let response = app
        .clone()
        .oneshot(req("GET", &format!("/api/links/{ID}"), Some(AUTH), vec![]))
        .await
        .unwrap();
    assert_eq!(response.headers().get(header::ETAG).unwrap(), "\"2\"");
    let blob = response.into_body().collect().await.unwrap().to_bytes();
    assert_eq!(&blob[..], b"two");
}

/// "Revert last update" asks for the revision below the deck's tag. Retention
/// drops it after two more pushes, and the TTL sweep takes the whole link:
/// both are a 404, which the client turns into kind "missing".
#[tokio::test]
async fn a_revision_a_sweep_took_is_404_like_one_retention_dropped() {
    let state = Arc::new(AppState::new(Store::in_memory().unwrap()));
    let auth = plsfix_server::store::auth_hash(AUTH);
    let long_ago = now() - 40 * 24 * 3600;
    state.store.put_link(ID, &auth, b"one", long_ago).unwrap();
    state.store.put_link(ID, &auth, b"two", long_ago).unwrap();
    assert_eq!(state.store.rev_count(ID), 2);

    let app = routes(state);
    let named = format!("/api/links/{ID}?rev=1");
    let (status, body) = call(&app, "GET", &named, Some(AUTH), b"").await;
    assert_eq!(status, StatusCode::NOT_FOUND, "expired revisions are gone");
    assert_eq!(body, r#"{"error":"not found"}"#);
    // And so is the head, which is what makes the deck say "missing", not "stale".
    let head = format!("/api/links/{ID}");
    assert_eq!(call(&app, "GET", &head, Some(AUTH), b"").await.0, 404);
}

/// The inbox delete a deck makes after it has inserted an item: the second one
/// is a 404, and so is one from a key that never wrote the row. Neither is
/// something the deck can act on, which is why the client treats both as done
/// (see src/link/relay.ts deleteInbox).
#[tokio::test]
async fn deleting_an_inbox_item_that_is_already_gone_is_404() {
    let app = relay_app();
    let ws = format!("/api/inbox/{AUTH}");
    let row = format!("/api/inbox/{AUTH}/{ID}");
    assert_eq!(call(&app, "POST", &ws, Some(AUTH), b"item").await.0, 200);
    assert_eq!(call(&app, "DELETE", &row, Some(AUTH), b"").await.0, 200);

    let (status, body) = call(&app, "DELETE", &row, Some(AUTH), b"").await;
    assert_eq!(status, StatusCode::NOT_FOUND);
    assert_eq!(body, r#"{"error":"not found"}"#);
    // A foreign key sees the same answer as an absent row, never "not yours".
    call(&app, "POST", &ws, Some(AUTH), b"item").await;
    let (status, other) = call(&app, "DELETE", &row, Some(OTHER), b"").await;
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
    push(&app, &vec![7u8; 4 * 1024 * 1024]).await;
    let other = format!("/api/links/{second}");
    call(&app, "PUT", &other, Some(AUTH), b"small").await;

    let both = format!(r#"[{{"id":"{ID}","auth":"{AUTH}"}},{{"id":"{second}","auth":"{AUTH}"}}]"#);
    let (status, body) = call(&app, "POST", "/api/links/fetch", None, both.as_bytes()).await;
    assert_eq!(status, StatusCode::OK);
    let answer = json(&body);
    assert_eq!(answer["items"].as_array().unwrap().len(), 1);
    assert_eq!(answer["items"][0]["id"], ID);
    assert_eq!(answer["omitted"][0]["reason"], "deferred");

    for (count, expect) in [(200usize, StatusCode::OK), (201, StatusCode::BAD_REQUEST)] {
        let items: Vec<String> = (0..count)
            .map(|index| format!(r#"{{"id":"{index:032x}","auth":"{AUTH}"}}"#))
            .collect();
        let batch = format!("[{}]", items.join(","));
        for route in ["fetch", "touch", "status"] {
            let path = format!("/api/links/{route}");
            let (status, _) = call(&app, "POST", &path, None, batch.as_bytes()).await;
            assert_eq!(status, expect, "{route} with {count} items");
        }
    }
}

/// The bucket a request is counted under follows `MODELIS_TRUSTED_PROXY`, and
/// nothing else. The hosted deployment declares `cloudflare`: nginx listens on
/// 127.0.0.1:8750 and the relay on 127.0.0.1:8804, so the sole way in is the
/// tunnel, and Cloudflare overwrites CF-Connecting-IP on every request it
/// proxies. A deployment that declares nothing - the default, and every
/// self-hoster until they say otherwise - ignores the header, because any
/// client can write one and would then pick its own bucket (I1 of the N1
/// security review).
#[tokio::test]
async fn the_rate_bucket_follows_the_trusted_proxy_and_nothing_else() {
    let cloudflare = Arc::new(AppState {
        writes: RateLimiter::new(1),
        trusted_proxy: TrustedProxy::Cloudflare,
        ..AppState::new(Store::in_memory().unwrap())
    });
    assert_eq!(
        header_buckets(&routes(cloudflare)).await,
        vec![
            StatusCode::OK,
            StatusCode::TOO_MANY_REQUESTS,
            StatusCode::OK
        ]
    );

    // Trusting nothing, the three requests are one client (the route tests
    // carry no peer address either), so the second one is already too many and
    // a new header buys the third nothing.
    let untrusted = Arc::new(AppState {
        writes: RateLimiter::new(1),
        ..AppState::new(Store::in_memory().unwrap())
    });
    assert_eq!(
        header_buckets(&routes(untrusted)).await,
        vec![
            StatusCode::OK,
            StatusCode::TOO_MANY_REQUESTS,
            StatusCode::TOO_MANY_REQUESTS
        ]
    );
}

/// Three pushes naming two different clients in `CF-Connecting-IP`.
async fn header_buckets(app: &axum::Router) -> Vec<StatusCode> {
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
        codes.push(send(app, request).await.0);
    }
    codes
}

/// `/version` is what the suite watches the relay by, and it counts what is
/// held right now: a sweep that drops a month-old link drops it from the
/// counters too, once the sweeper has cleared the cached answer with it.
#[tokio::test]
async fn version_counts_fall_when_the_sweeper_drops_dead_rows() {
    let state = Arc::new(AppState::new(Store::in_memory().unwrap()));
    let auth = plsfix_server::store::auth_hash(AUTH);
    let long_ago = now() - 40 * 24 * 3600;
    state.store.put_link(ID, &auth, b"12345", long_ago).unwrap();
    let app = app(static_dir("version"), state.clone());

    let before = json(&get(&app, "/version").await.1);
    assert_eq!(before["relay"]["links"], 1);
    assert_eq!(before["relay"]["bytes"], 5);
    assert_eq!(state.store.sweep(now()).unwrap(), 1);
    state.counts.clear();
    let after = json(&get(&app, "/version").await.1);
    assert_eq!(after["relay"]["links"], 0);
    assert_eq!(after["relay"]["revisions"], 0);
    assert_eq!(after["relay"]["bytes"], 0);
}

/// Counting scans both tables under the one connection mutex every push needs,
/// and `/version` is anonymous on an Access-bypassed path, so the answer is
/// reused: a link pushed between two reads does not show up in the second
/// (I3 of the N1 security review). The hourly sweeper clears it.
#[tokio::test]
async fn version_reuses_its_counters_instead_of_scanning_per_request() {
    let state = Arc::new(AppState::new(Store::in_memory().unwrap()));
    let auth = plsfix_server::store::auth_hash(AUTH);
    let second = "1123456789abcdef0123456789abcdef";
    state.store.put_link(ID, &auth, b"one", now()).unwrap();
    let app = app(static_dir("version-cache"), state.clone());

    assert_eq!(json(&get(&app, "/version").await.1)["relay"]["links"], 1);
    state.store.put_link(second, &auth, b"two", now()).unwrap();
    assert_eq!(json(&get(&app, "/version").await.1)["relay"]["links"], 1);

    state.counts.clear();
    assert_eq!(json(&get(&app, "/version").await.1)["relay"]["links"], 2);
    assert_eq!(VERSION_COUNTS_TTL, 30);
}
