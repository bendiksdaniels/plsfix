//! Hunt pass 1 (edges + repeats): empty and content-type-free bodies, the
//! bearer/workspace key length boundary, the one batch route that had no
//! malformed-JSON pin (`/api/links/status`), and a same-item repeat press.

mod common;

use axum::{
    body::Body,
    http::{header, Request, StatusCode},
};
use common::*;
use http_body_util::BodyExt;
use tower::ServiceExt;

/// PUT and the inbox POST read the body as raw `Bytes`, and never require a
/// minimum length: an empty push is a legal, if unusual, ciphertext of zero
/// bytes, and must round-trip exactly - not be silently dropped or turned
/// into a store error.
#[tokio::test]
async fn an_empty_link_body_is_stored_and_served_back_empty() {
    let app = relay_app();
    let path = format!("/api/links/{ID}");
    let (status, body) = call(&app, "PUT", &path, Some(AUTH), b"").await;
    assert_eq!(status, StatusCode::OK, "{body}");
    assert_eq!(body, r#"{"rev":1}"#);

    let response = app
        .clone()
        .oneshot(req("GET", &path, Some(AUTH), vec![]))
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::OK);
    assert_eq!(response.headers().get(header::ETAG).unwrap(), "\"1\"");
    let blob = response.into_body().collect().await.unwrap().to_bytes();
    assert!(blob.is_empty(), "expected an empty blob, got {blob:?}");
}

/// An empty inbox item behaves the same way: stored, listed, its blob an
/// empty string once base64-encoded.
#[tokio::test]
async fn an_empty_inbox_item_is_stored_and_listed_with_an_empty_blob() {
    let app = relay_app();
    let ws = format!("/api/inbox/{AUTH}");
    let (status, _) = call(&app, "POST", &ws, Some(AUTH), b"").await;
    assert_eq!(status, StatusCode::OK);

    let (status, body) = call(&app, "GET", &ws, Some(AUTH), b"").await;
    assert_eq!(status, StatusCode::OK);
    assert_eq!(json(&body)[0]["blob"], "");
}

/// PUT takes raw `Bytes` and `/api/links/status` parses its JSON by hand, not
/// through axum's `Json` extractor, so neither looks at Content-Type: a proxy
/// that strips or rewrites the header must not break a push or a status call.
#[tokio::test]
async fn put_and_status_ignore_content_type() {
    let app = relay_app();
    let put_no_type = Request::builder()
        .method("PUT")
        .uri(format!("/api/links/{ID}"))
        .header(header::AUTHORIZATION, format!("Bearer {AUTH}"))
        .body(Body::from(b"x".to_vec()))
        .unwrap();
    assert_eq!(send(&app, put_no_type).await.0, StatusCode::OK);

    let put_wrong_type = Request::builder()
        .method("PUT")
        .uri(format!("/api/links/{ID}"))
        .header(header::AUTHORIZATION, format!("Bearer {AUTH}"))
        .header(header::CONTENT_TYPE, "text/plain")
        .body(Body::from(b"y".to_vec()))
        .unwrap();
    assert_eq!(send(&app, put_wrong_type).await.0, StatusCode::OK);

    let status_no_type = Request::builder()
        .method("POST")
        .uri("/api/links/status")
        .body(Body::from(one(AUTH)))
        .unwrap();
    let (status, body) = send(&app, status_no_type).await;
    assert_eq!(status, StatusCode::OK, "{body}");
    assert!(body.contains(r#""rev":2"#), "{body}");

    let status_wrong_type = Request::builder()
        .method("POST")
        .uri("/api/links/status")
        .header(header::CONTENT_TYPE, "text/plain")
        .body(Body::from(one(AUTH)))
        .unwrap();
    assert_eq!(send(&app, status_wrong_type).await.0, StatusCode::OK);
}

/// `/api/links/status` was the one batch route with no pin for a body that is
/// not valid JSON at all: `fetch` and `touch` both had one already.
#[tokio::test]
async fn a_malformed_status_batch_is_a_bad_request() {
    let app = relay_app();
    let (status, body) = call(&app, "POST", "/api/links/status", None, b"not json").await;
    assert_eq!(status, StatusCode::BAD_REQUEST);
    assert_eq!(body, r#"{"error":"bad body"}"#);

    // Valid JSON, wrong shape (an object where an array of items belongs).
    let (status, body) = call(&app, "POST", "/api/links/status", None, b"{}").await;
    assert_eq!(status, StatusCode::BAD_REQUEST);
    assert_eq!(body, r#"{"error":"bad body"}"#);

    // An item missing its "auth" field is a shape error too, not a panic.
    let (status, body) = call(
        &app,
        "POST",
        "/api/links/status",
        None,
        format!(r#"[{{"id":"{ID}"}}]"#).as_bytes(),
    )
    .await;
    assert_eq!(status, StatusCode::BAD_REQUEST);
    assert_eq!(body, r#"{"error":"bad body"}"#);
}

/// A bearer or a workspace id one character short or long of the 43-character
/// key length is refused like any other unauthorized or badly shaped request
/// - never treated as almost right.
#[tokio::test]
async fn a_key_one_character_off_the_right_length_is_refused() {
    let app = relay_app();
    let too_short = &AUTH[..AUTH.len() - 1];
    let too_long = format!("{AUTH}A");
    for bad in [too_short, too_long.as_str()] {
        let (status, _) = call(&app, "GET", &format!("/api/links/{ID}"), Some(bad), b"").await;
        assert_eq!(status, StatusCode::UNAUTHORIZED, "bearer len={}", bad.len());
    }

    // A workspace id of the wrong length is a bad id, checked after a VALID
    // bearer is found - so this is BadId (400), not Unauthorized (401).
    let short_ws = &AUTH[..AUTH.len() - 1];
    let (status, body) = call(
        &app,
        "GET",
        &format!("/api/inbox/{short_ws}"),
        Some(AUTH),
        b"",
    )
    .await;
    assert_eq!(status, StatusCode::BAD_REQUEST);
    assert_eq!(body, r#"{"error":"bad id"}"#);
}

/// Posting the same inbox item three times running must leave exactly one
/// row, each press replacing the last rather than stacking - the idempotent
/// case the brief asks every repeated action to prove.
#[tokio::test]
async fn posting_the_same_inbox_item_three_times_leaves_exactly_one_row() {
    let app = relay_app();
    let ws = format!("/api/inbox/{AUTH}");
    for blob in [b"one".as_slice(), b"two".as_slice(), b"three".as_slice()] {
        let (status, _) = call(&app, "POST", &ws, Some(AUTH), blob).await;
        assert_eq!(status, StatusCode::OK);
    }
    let (status, body) = call(&app, "GET", &ws, Some(AUTH), b"").await;
    assert_eq!(status, StatusCode::OK);
    let rows = json(&body);
    assert_eq!(rows.as_array().unwrap().len(), 1, "{body}");
    assert_eq!(rows[0]["blob"], "dGhyZWU"); // base64url of "three"
}
