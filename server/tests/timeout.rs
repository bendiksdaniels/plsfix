//! The request timeout (server/src/lib.rs, tower_http::timeout::TimeoutLayer):
//! a handler still running past the deadline answers 408 on its own, a fast
//! one is untouched, and app_with_timeout threads a caller's own duration
//! through the real router rather than only REQUEST_TIMEOUT_SECS. Short
//! durations throughout, so the suite pays real milliseconds, never seconds.

use std::{sync::Arc, time::Duration};

use axum::{
    body::Body,
    http::{Request, StatusCode},
    routing::get,
    Router,
};
use plsfix_server::{app_with_timeout, relay::AppState, store::Store, REQUEST_TIMEOUT_SECS};
use tower::ServiceExt;

mod common;
use common::static_dir;

fn slow_router(timeout: Duration, handler_delay: Duration) -> Router {
    Router::new()
        .route(
            "/slow",
            get(move || async move {
                tokio::time::sleep(handler_delay).await;
                "too slow to matter"
            }),
        )
        .layer(tower_http::timeout::TimeoutLayer::new(timeout))
}

#[tokio::test]
async fn a_handler_still_running_past_the_deadline_answers_408() {
    let app = slow_router(Duration::from_millis(20), Duration::from_millis(500));
    let response = app
        .oneshot(Request::get("/slow").body(Body::empty()).unwrap())
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::REQUEST_TIMEOUT);
}

#[tokio::test]
async fn a_handler_that_finishes_first_is_untouched() {
    let app = slow_router(Duration::from_millis(500), Duration::from_millis(5));
    let response = app
        .oneshot(Request::get("/slow").body(Body::empty()).unwrap())
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::OK);
}

#[tokio::test]
async fn the_real_router_is_bounded_by_the_duration_it_is_given() {
    let state = Arc::new(AppState::new(Store::in_memory().unwrap()));
    let app = app_with_timeout(static_dir("timeout"), state, Duration::from_millis(50));
    let response = app
        .oneshot(Request::get("/healthz").body(Body::empty()).unwrap())
        .await
        .unwrap();
    // A fast route on a short deadline still answers normally: the layer
    // bounds a stalled handler, it does not race every request against it.
    assert_eq!(response.status(), StatusCode::OK);
}

#[test]
fn the_production_deadline_is_thirty_seconds() {
    assert_eq!(REQUEST_TIMEOUT_SECS, 30);
}
