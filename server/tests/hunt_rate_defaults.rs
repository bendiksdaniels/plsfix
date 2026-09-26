//! Hunt pass 1 (edges + repeats): the production rate defaults, wired end to
//! end through `AppState::new` and the real router. Every existing HTTP test
//! for the write limiter overrides it with a small number (1, 2, 3...) to
//! keep the test fast; none of them exercises the literal
//! `DEFAULT_WRITE_PER_MIN` (100) a real deployment starts with, so a typo
//! that changed the constant without touching its unit test would not be
//! caught at the layer that actually matters: the router.

mod common;

use axum::http::{header, StatusCode};
use common::*;
use plsfix_server::relay::*;
use tower::ServiceExt;

#[tokio::test]
async fn the_hundredth_write_succeeds_and_the_hundred_and_first_waits() {
    let app = relay_app();
    for index in 0..DEFAULT_WRITE_PER_MIN {
        let id = format!("{index:032x}");
        let (status, body) = call(&app, "PUT", &format!("/api/links/{id}"), Some(AUTH), b"x").await;
        assert_eq!(status, StatusCode::OK, "write {index}: {body}");
    }
    let (status, body) = call(
        &app,
        "PUT",
        &format!("/api/links/{:032x}", DEFAULT_WRITE_PER_MIN),
        Some(AUTH),
        b"x",
    )
    .await;
    assert_eq!(status, StatusCode::TOO_MANY_REQUESTS, "{body}");
    assert_eq!(body, r#"{"error":"too many requests"}"#);

    // A read is a different allowance (DEFAULT_READ_PER_MIN, 1200) and is
    // untouched by a client that has just spent its whole write budget.
    let (status, _) = call(
        &app,
        "GET",
        &format!("/api/links/{:032x}", DEFAULT_WRITE_PER_MIN - 1),
        Some(AUTH),
        b"",
    )
    .await;
    assert_eq!(status, StatusCode::OK);
}

/// A refused write at the production allowance still carries a Retry-After a
/// pane could obey: 100 a minute is one every 0.6 s, rounded up to whole
/// seconds like every other refusal.
#[tokio::test]
async fn the_production_write_limit_answers_with_a_retry_after() {
    let app = relay_app();
    for index in 0..DEFAULT_WRITE_PER_MIN {
        let id = format!("{index:032x}");
        call(&app, "PUT", &format!("/api/links/{id}"), Some(AUTH), b"x").await;
    }
    let request = req(
        "PUT",
        &format!("/api/links/{:032x}", DEFAULT_WRITE_PER_MIN),
        Some(AUTH),
        b"x".to_vec(),
    );
    let response = app.oneshot(request).await.unwrap();
    assert_eq!(response.status(), StatusCode::TOO_MANY_REQUESTS);
    let retry_after: u64 = response
        .headers()
        .get(header::RETRY_AFTER)
        .unwrap()
        .to_str()
        .unwrap()
        .parse()
        .unwrap();
    assert!(retry_after >= 1, "must never be zero: {retry_after}");
}
