//! What C1 and I2/I3 of the N1 security review cost an anonymous client: a
//! write allowance counted in bytes rather than requests, a storage ceiling
//! that cannot be raced by concurrent pushes, and the three suite routes
//! spending the read allowance like everything else.
//! Real bodies here - four mebibytes at a time - because the point is the size
//! of a push, not the count.

use std::sync::Arc;

use axum::{
    body::Body,
    http::{header, Request, StatusCode},
    Router,
};
use plsfix_server::{
    app,
    limits::{RateLimiter, TrustedProxy},
    relay::*,
    store::Store,
    store_room::Caps,
};
use tower::ServiceExt;

mod common;
use common::static_dir;

const AUTH: &str = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"; // 43 chars
const MIB: usize = 1024 * 1024;

fn state(build: impl FnOnce(&mut AppState)) -> Arc<AppState> {
    let mut state = AppState::new(Store::in_memory().unwrap());
    state.trusted_proxy = TrustedProxy::Cloudflare;
    build(&mut state);
    Arc::new(state)
}

/// One push of `bytes` from `client`, to a link of its own.
fn put(index: usize, client: &str, bytes: usize) -> Request<Body> {
    Request::builder()
        .method("PUT")
        .uri(format!("/api/links/{index:032x}"))
        .header(header::AUTHORIZATION, format!("Bearer {AUTH}"))
        .header("CF-Connecting-IP", client)
        .body(Body::from(vec![b'x'; bytes]))
        .unwrap()
}

async fn send(app: &Router, request: Request<Body>) -> StatusCode {
    app.clone().oneshot(request).await.unwrap().status()
}

/// C1: 300 writes a minute was 1.2 GiB a minute, which is the whole store, so
/// the request count was never a budget. Charged by the KiB, one client's
/// minute stops at 64 MiB - sixteen full-size pushes - long before the global
/// ceiling has moved at all.
#[tokio::test]
async fn a_client_minute_of_pushes_is_capped_in_bytes_not_requests() {
    let state = state(|_| ());
    let app = routes(state.clone());
    let mut codes = Vec::new();
    for index in 0..20 {
        codes.push(send(&app, put(index, "1.2.3.4", 4 * MIB)).await);
    }

    // 64 MiB of allowance, 4 MiB a push. The bucket refills at about 1 MiB a
    // second, so the seventeenth would need the first sixteen to have taken
    // nearly four seconds - they take milliseconds.
    assert_eq!(codes[..16], [StatusCode::OK; 16]);
    assert_eq!(codes[16], StatusCode::TOO_MANY_REQUESTS);
    assert!(!codes.contains(&StatusCode::INSUFFICIENT_STORAGE));
    let held = state.store.total_bytes().unwrap();
    assert_eq!(held, 16 * 4 * MIB as i64);
    assert!(held < state.max_bytes() / 4, "{held} bytes of the ceiling");
}

/// The budget is per client, like the request allowance beside it: one client
/// spending its minute leaves the next one whole.
#[tokio::test]
async fn two_clients_have_independent_byte_budgets() {
    // 4 KiB a minute, so two 2 KiB pushes and no more.
    let state = state(|state| state.write_bytes = RateLimiter::new(4));
    let app = routes(state);
    for index in 0..2 {
        assert_eq!(
            send(&app, put(index, "1.2.3.4", 2048)).await,
            StatusCode::OK
        );
    }
    let refused = send(&app, put(2, "1.2.3.4", 2048)).await;
    assert_eq!(refused, StatusCode::TOO_MANY_REQUESTS);
    assert_eq!(send(&app, put(3, "5.6.7.8", 2048)).await, StatusCode::OK);
}

/// A refused push says when to come back, the same contract the request
/// allowance has, so a pane retries instead of guessing.
#[tokio::test]
async fn a_push_over_the_byte_budget_answers_429_with_a_retry_after() {
    let state = state(|state| state.write_bytes = RateLimiter::new(1));
    let app = routes(state);
    assert_eq!(send(&app, put(0, "1.2.3.4", 1024)).await, StatusCode::OK);
    let response = app.clone().oneshot(put(1, "1.2.3.4", 1024)).await.unwrap();
    assert_eq!(response.status(), StatusCode::TOO_MANY_REQUESTS);
    assert_eq!(response.headers().get(header::RETRY_AFTER).unwrap(), "60");
}

/// I2: the ceiling used to be read, the lock dropped, and the row written
/// later, so every push in flight passed the same pre-write total and the
/// overshoot was however many were running. Eight at once against room for
/// four now admit four, and the store ends the test exactly full.
#[tokio::test(flavor = "multi_thread", worker_threads = 4)]
async fn parallel_pushes_at_the_ceiling_admit_exactly_what_fits() {
    let state = state(|state| {
        state.store.set_caps(Caps {
            max_bytes: 4 * MIB as i64,
            inbox_rows: INBOX_MAX_PER_WS,
        });
    });
    let app = routes(state.clone());
    let gate = Arc::new(tokio::sync::Barrier::new(8));

    let mut racing = Vec::new();
    for index in 0..8 {
        let (app, gate) = (app.clone(), Arc::clone(&gate));
        racing.push(tokio::spawn(async move {
            gate.wait().await;
            app.oneshot(put(index, "1.2.3.4", MIB))
                .await
                .unwrap()
                .status()
        }));
    }
    let mut codes = Vec::new();
    for task in racing {
        codes.push(task.await.unwrap());
    }

    let stored = codes.iter().filter(|code| code.is_success()).count();
    let full = codes
        .iter()
        .filter(|code| **code == StatusCode::INSUFFICIENT_STORAGE)
        .count();
    assert_eq!((stored, full), (4, 4), "{codes:?}");
    assert_eq!(state.store.total_bytes().unwrap(), 4 * MIB as i64);
}

/// I3: the suite routes sat outside the limiter, and `/version` scanned both
/// tables under the mutex every push needs. They spend the read allowance now,
/// so an anonymous flood is refused rather than served.
#[tokio::test]
async fn the_suite_routes_spend_the_read_allowance() {
    let state = state(|state| state.reads = RateLimiter::new(2));
    let app = app(static_dir("suite-reads"), state);
    let read = |path: &'static str| {
        let app = app.clone();
        async move {
            app.oneshot(Request::get(path).body(Body::empty()).unwrap())
                .await
                .unwrap()
                .status()
        }
    };
    assert_eq!(read("/healthz").await, StatusCode::OK);
    assert_eq!(read("/version").await, StatusCode::OK);
    // The third read of the minute, whichever route it lands on.
    assert_eq!(read("/manifest.xml").await, StatusCode::TOO_MANY_REQUESTS);
    assert_eq!(read("/healthz").await, StatusCode::TOO_MANY_REQUESTS);
}
