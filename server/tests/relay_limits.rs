//! The two write gates on the routes: the storage ceiling (507, after a sweep
//! that may heal it), the per-workspace inbox row cap (507) and the per-client
//! rate limit (429 with `Retry-After`). Tiny ceilings and tiny allowances are
//! set on the state, so nothing here writes a gigabyte or waits a minute.

use plsfix_server::{
    limits::RateLimiter,
    relay::*,
    store::{auth_hash, Store, LINK_TTL},
};

use axum::{
    body::Body,
    http::{header, Request, StatusCode},
    response::Response,
    Router,
};
use http_body_util::BodyExt;
use std::sync::Arc;
use tower::ServiceExt;

const ID: &str = "0123456789abcdef0123456789abcdef";
const OTHER: &str = "fedcba9876543210fedcba9876543210";
const WS: &str = "CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"; // 43 chars
const MINE: &str = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";

fn state(build: impl FnOnce(&mut AppState)) -> Arc<AppState> {
    let mut state = AppState::new(Store::in_memory().unwrap());
    build(&mut state);
    Arc::new(state)
}

fn put(id: &str, bytes: usize) -> Request<Body> {
    request("PUT", &format!("/api/links/{id}"), None, vec![b'x'; bytes])
}

fn request(method: &str, path: &str, client: Option<&str>, body: Vec<u8>) -> Request<Body> {
    let mut builder = Request::builder()
        .method(method)
        .uri(path)
        .header(header::AUTHORIZATION, format!("Bearer {MINE}"))
        .header(header::CONTENT_TYPE, "application/octet-stream")
        .header("X-PLSFIX-Link-Id", ID);
    if let Some(client) = client {
        builder = builder.header("CF-Connecting-IP", client);
    }
    builder.body(Body::from(body)).unwrap()
}

async fn send(app: &Router, request: Request<Body>) -> Response {
    app.clone().oneshot(request).await.unwrap()
}

async fn body_of(response: Response) -> String {
    let bytes = response.into_body().collect().await.unwrap().to_bytes();
    String::from_utf8_lossy(&bytes).to_string()
}

#[tokio::test]
async fn a_push_over_the_byte_ceiling_sweeps_first_and_only_then_refuses() {
    let state = state(|state| state.max_bytes = 100);
    let app = routes(state.clone());
    // 100 dead bytes: over the ceiling on paper, gone the moment it matters.
    state
        .store
        .put_link(OTHER, &auth_hash(MINE), &[b'x'; 100], now() - LINK_TTL - 10)
        .unwrap();

    assert_eq!(send(&app, put(ID, 50)).await.status(), StatusCode::OK);

    // 50 live bytes now, so another 60 does not fit and no sweep can help.
    let response = send(&app, put(OTHER, 60)).await;
    assert_eq!(response.status(), StatusCode::INSUFFICIENT_STORAGE);
    assert_eq!(body_of(response).await, r#"{"error":"storage full"}"#);
}

#[tokio::test]
async fn an_inbox_post_over_the_workspace_row_cap_is_refused() {
    let state = state(|_| ());
    let app = routes(state.clone());
    for row in 0..INBOX_MAX_PER_WS {
        let id = format!("{row:032x}");
        state
            .store
            .post_inbox(WS, &auth_hash(MINE), &id, b"item", now())
            .unwrap();
    }

    let response = send(
        &app,
        request("POST", &format!("/api/inbox/{WS}"), None, b"x".to_vec()),
    )
    .await;
    assert_eq!(response.status(), StatusCode::INSUFFICIENT_STORAGE);
    assert_eq!(body_of(response).await, r#"{"error":"storage full"}"#);

    // Another workspace is unaffected: the cap is per workspace, not global.
    let other_ws = "D".repeat(43);
    let response = send(
        &app,
        request(
            "POST",
            &format!("/api/inbox/{other_ws}"),
            None,
            b"x".to_vec(),
        ),
    )
    .await;
    assert_eq!(response.status(), StatusCode::OK);
}

#[tokio::test]
async fn writes_past_the_allowance_get_429_and_a_retry_after() {
    let state = state(|state| state.writes = RateLimiter::new(2));
    let app = routes(state);

    for _ in 0..2 {
        assert_eq!(
            send(
                &app,
                request(
                    "PUT",
                    &format!("/api/links/{ID}"),
                    Some("1.2.3.4"),
                    b"x".to_vec()
                )
            )
            .await
            .status(),
            StatusCode::OK
        );
    }
    let response = send(
        &app,
        request(
            "PUT",
            &format!("/api/links/{ID}"),
            Some("1.2.3.4"),
            b"x".to_vec(),
        ),
    )
    .await;
    assert_eq!(response.status(), StatusCode::TOO_MANY_REQUESTS);
    // Whole seconds: 2 a minute is one every 30.
    assert_eq!(response.headers().get(header::RETRY_AFTER).unwrap(), "30");
    assert_eq!(body_of(response).await, r#"{"error":"too many requests"}"#);

    // The reads have their own allowance, and another client its own bucket.
    assert_eq!(
        send(
            &app,
            request("GET", &format!("/api/links/{ID}"), Some("1.2.3.4"), vec![])
        )
        .await
        .status(),
        StatusCode::OK
    );
    assert_eq!(
        send(
            &app,
            request(
                "PUT",
                &format!("/api/links/{ID}"),
                Some("5.6.7.8"),
                b"x".to_vec()
            )
        )
        .await
        .status(),
        StatusCode::OK
    );
}

#[tokio::test]
async fn an_inbox_post_spends_the_write_allowance_not_the_read_one() {
    let state = state(|state| state.writes = RateLimiter::new(1));
    let app = routes(state);
    let path = format!("/api/inbox/{WS}");

    assert_eq!(
        send(&app, request("POST", &path, Some("1.2.3.4"), b"x".to_vec()))
            .await
            .status(),
        StatusCode::OK
    );
    assert_eq!(
        send(&app, request("POST", &path, Some("1.2.3.4"), b"x".to_vec()))
            .await
            .status(),
        StatusCode::TOO_MANY_REQUESTS
    );
    // The batch routes are reads even though they POST.
    let batch = Request::builder()
        .method("POST")
        .uri("/api/links/touch")
        .header("CF-Connecting-IP", "1.2.3.4")
        .header(header::CONTENT_TYPE, "application/json")
        .body(Body::from("[]"))
        .unwrap();
    assert_eq!(send(&app, batch).await.status(), StatusCode::OK);
}
