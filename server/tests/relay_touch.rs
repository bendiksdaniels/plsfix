//! `POST /api/links/touch`: what the Links tab calls on boot so a workbook's
//! links keep their full TTL. Every revision of a matching link moves, a wrong
//! key moves nothing, an expired link is skipped, and the batch ceiling is the
//! same 200 items the other two carry.

use plsfix_server::{
    relay::*,
    store::{auth_hash, Get, Store, LINK_TTL},
};

use axum::{
    body::Body,
    http::{header, Request, StatusCode},
    Router,
};
use http_body_util::BodyExt;
use std::sync::Arc;
use tower::ServiceExt;

const ID: &str = "0123456789abcdef0123456789abcdef";
const OTHER: &str = "fedcba9876543210fedcba9876543210";
const MINE: &str = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"; // 43 chars
const THEIRS: &str = "BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB";

fn state() -> Arc<AppState> {
    Arc::new(AppState::new(Store::in_memory().unwrap()))
}

/// Two revisions that are alive now and die in a minute: the route runs on the
/// wall clock, so a seeded row has to be dated against it.
fn seed_dying_link(state: &AppState) {
    let pushed = now() - LINK_TTL + 60;
    for blob in [b"one", b"two"] {
        state
            .store
            .put_link(ID, &auth_hash(MINE), blob, pushed)
            .unwrap();
    }
    assert_eq!(state.store.rev_count(ID), 2);
}

async fn touch(app: &Router, body: String) -> (StatusCode, String) {
    let request = Request::builder()
        .method("POST")
        .uri("/api/links/touch")
        .header(header::CONTENT_TYPE, "application/json")
        .body(Body::from(body))
        .unwrap();
    let response = app.clone().oneshot(request).await.unwrap();
    let status = response.status();
    let body = response.into_body().collect().await.unwrap().to_bytes();
    (status, String::from_utf8_lossy(&body).to_string())
}

fn one(id: &str, auth: &str) -> String {
    format!(r#"[{{"id":"{id}","auth":"{auth}"}}]"#)
}

/// An hour out: past the seeded expiry, well inside a touched one.
fn later() -> i64 {
    now() + 3600
}

#[tokio::test]
async fn touch_moves_the_expiry_of_every_revision() {
    let state = state();
    let app = routes(state.clone());
    seed_dying_link(&state);
    assert!(matches!(
        state.store.get_link(ID, &auth_hash(MINE), later()).unwrap(),
        Get::Missing
    ));

    let (status, body) = touch(&app, one(ID, MINE)).await;
    assert_eq!(status, StatusCode::OK);
    assert_eq!(body, r#"{"touched":1}"#);

    // The head and the revision below it both live past the old expiry: a
    // revert has to reach rev 1 for as long as rev 2 is there.
    assert!(matches!(
        state.store.get_link(ID, &auth_hash(MINE), later()).unwrap(),
        Get::Found(_)
    ));
    assert!(matches!(
        state
            .store
            .get_link_rev(ID, &auth_hash(MINE), 1, later())
            .unwrap(),
        Get::Found(_)
    ));
}

#[tokio::test]
async fn a_wrong_key_touches_nothing_and_is_told_nothing() {
    let state = state();
    let app = routes(state.clone());
    seed_dying_link(&state);

    let (status, body) = touch(&app, one(ID, THEIRS)).await;
    assert_eq!(status, StatusCode::OK);
    // Same answer an unknown id gets: the count never says "this link exists".
    assert_eq!(body, r#"{"touched":0}"#);
    assert!(matches!(
        state.store.get_link(ID, &auth_hash(MINE), later()).unwrap(),
        Get::Missing
    ));
}

#[tokio::test]
async fn expired_and_unknown_ids_are_skipped() {
    let state = state();
    let app = routes(state.clone());
    // Pushed longer ago than the TTL: dead, and a touch does not raise it.
    let pushed = now() - LINK_TTL - 10;
    state
        .store
        .put_link(ID, &auth_hash(MINE), b"one", pushed)
        .unwrap();

    let (status, body) = touch(&app, one(ID, MINE)).await;
    assert_eq!(status, StatusCode::OK);
    assert_eq!(body, r#"{"touched":0}"#);
    assert!(matches!(
        state.store.get_link(ID, &auth_hash(MINE), now()).unwrap(),
        Get::Missing
    ));

    let (status, body) = touch(&app, one(OTHER, MINE)).await;
    assert_eq!(status, StatusCode::OK);
    assert_eq!(body, r#"{"touched":0}"#);
}

#[tokio::test]
async fn the_batch_ceiling_and_the_body_shape_are_refused_like_the_others() {
    let app = routes(state());
    let items: Vec<String> = (0..201)
        .map(|_| format!(r#"{{"id":"{ID}","auth":"{MINE}"}}"#))
        .collect();
    let (status, body) = touch(&app, format!("[{}]", items.join(","))).await;
    assert_eq!(status, StatusCode::BAD_REQUEST);
    assert_eq!(body, r#"{"error":"too many items"}"#);

    let (status, body) = touch(&app, "not json".to_string()).await;
    assert_eq!(status, StatusCode::BAD_REQUEST);
    assert_eq!(body, r#"{"error":"bad body"}"#);
}
