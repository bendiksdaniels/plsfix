//! The request timeout (server/src/lib.rs, tower_http::timeout::TimeoutLayer):
//! a handler still running past the deadline answers 408 on its own, a fast
//! one is untouched, and app_with_timeout threads a caller's own duration
//! through the real router rather than only REQUEST_TIMEOUT_SECS. Short
//! durations throughout, so the suite pays real milliseconds, never seconds.

use std::{sync::Arc, time::Duration};

use axum::{
    body::{Body, Bytes},
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

// The two tests above build their own router with their own TimeoutLayer, so
// deleting the layer from the shipped router in lib.rs would leave them
// green - they guard the layer's own behaviour, not that app_with_timeout
// actually installs one on the real routes.
//
// A timing race against a real route turns out not to be reliably
// deterministic here (Duration::ZERO against /healthz was tried first, per
// fix-round-1 review): none of the shipped handlers take a controllable
// delay - /healthz never awaits anything and resolves on its very first poll
// regardless of the deadline (Duration::ZERO measured 0/10 timeouts against
// it), and even the ServeDir file read behind /taskpane.html, though it does
// yield once through a real
// spawn_blocking dispatch, resolves faster than any deadline from 0ms to 10ms
// could reliably beat (0-40% across repeated runs, worse the larger the
// deadline). A request whose BODY never finishes arriving sidesteps the race
// entirely: PUT /api/links/:id reads its body with the `Bytes` extractor
// before put_link's own code (and its auth check) ever runs, so a body that
// never produces a chunk and never signals EOF makes normal completion
// impossible - only the layer's own deadline can ever answer this request,
// so any positive duration is deterministic, not just likely.
#[tokio::test]
async fn the_real_router_answers_408_when_its_own_layer_is_exhausted() {
    let state = Arc::new(AppState::new(Store::in_memory().unwrap()));
    let app = app_with_timeout(static_dir("timeout"), state, Duration::from_millis(10));
    let never_arrives = futures_util::stream::pending::<Result<Bytes, std::io::Error>>();
    let request = Request::put(format!("/api/links/{}", "a".repeat(32)))
        .header(
            "authorization",
            "Bearer AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
        )
        .body(Body::from_stream(never_arrives))
        .unwrap();
    let response = app.oneshot(request).await.unwrap();
    assert_eq!(response.status(), StatusCode::REQUEST_TIMEOUT);
}

#[test]
fn the_production_deadline_is_thirty_seconds() {
    assert_eq!(REQUEST_TIMEOUT_SECS, 30);
}
