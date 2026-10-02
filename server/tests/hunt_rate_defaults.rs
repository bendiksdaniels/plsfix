//! Hunt pass 1 (edges + repeats): the production write allowance
//! (`DEFAULT_WRITE_PER_MIN`, 100) through `AppState::new` and the real router.
//! `main.rs` builds its own limiter from the environment (`state_from_env`):
//! this pins the library default, not that wiring.

mod common;

use axum::{
    http::{header, HeaderValue, StatusCode},
    Router,
};
use common::*;
use http_body_util::BodyExt;
use plsfix_server::relay::*;
use tower::ServiceExt;

/// A router that has spent its whole write allowance, and its answer to one
/// write more.
struct Refusal {
    app: Router,
    status: StatusCode,
    retry_after: Option<HeaderValue>,
    body: String,
}

/// One attempt on a fresh router. `None` when the wall clock reached the next
/// second meanwhile: the limiter refills by whole seconds (1.67 tokens each),
/// so such an attempt may have earned a token and proves nothing either way.
async fn spend_all_then_one_more() -> Option<Refusal> {
    let app = relay_app();
    let started = now();
    for index in 0..DEFAULT_WRITE_PER_MIN {
        let path = format!("/api/links/{index:032x}");
        let (status, body) = call(&app, "PUT", &path, Some(AUTH), b"x").await;
        assert_eq!(status, StatusCode::OK, "write {index}: {body}");
    }
    let path = format!("/api/links/{:032x}", DEFAULT_WRITE_PER_MIN);
    let request = req("PUT", &path, Some(AUTH), b"x".to_vec());
    let response = app.clone().oneshot(request).await.unwrap();
    let (status, retry_after) = (
        response.status(),
        response.headers().get(header::RETRY_AFTER).cloned(),
    );
    let bytes = response.into_body().collect().await.unwrap().to_bytes();
    let body = String::from_utf8_lossy(&bytes).to_string();
    (now() == started).then_some(Refusal {
        app,
        status,
        retry_after,
        body,
    })
}

/// Five attempts in a row crossing a second boundary means a machine too slow
/// for this test to mean anything, and the test says so instead of passing.
async fn refused_extra_write() -> Refusal {
    for _ in 0..5 {
        if let Some(refusal) = spend_all_then_one_more().await {
            return refusal;
        }
    }
    panic!("five attempts in a row crossed a second boundary");
}

#[tokio::test]
async fn the_hundredth_write_succeeds_and_the_hundred_and_first_waits() {
    let refusal = refused_extra_write().await;
    assert_eq!(
        refusal.status,
        StatusCode::TOO_MANY_REQUESTS,
        "{}",
        refusal.body
    );
    assert_eq!(refusal.body, r#"{"error":"too many requests"}"#);

    // A read is a different allowance (DEFAULT_READ_PER_MIN, 1200) and is
    // untouched by a client that has just spent its whole write budget.
    let path = format!("/api/links/{:032x}", DEFAULT_WRITE_PER_MIN - 1);
    let (status, _) = call(&refusal.app, "GET", &path, Some(AUTH), b"").await;
    assert_eq!(status, StatusCode::OK);
}

/// A refused write at the production allowance still carries a Retry-After a
/// pane could obey: 100 a minute is one every 0.6 s, rounded up to whole
/// seconds like every other refusal.
#[tokio::test]
async fn the_production_write_limit_answers_with_a_retry_after() {
    let refusal = refused_extra_write().await;
    assert_eq!(refusal.status, StatusCode::TOO_MANY_REQUESTS);
    let retry_after: u64 = refusal
        .retry_after
        .expect("a refusal carries Retry-After")
        .to_str()
        .unwrap()
        .parse()
        .unwrap();
    assert!(retry_after >= 1, "must never be zero: {retry_after}");
}
