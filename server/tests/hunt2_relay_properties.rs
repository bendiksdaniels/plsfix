//! Hunt pass 2 (properties): the HTTP router fuzzed - bodies, paths, headers,
//! methods - must never panic or answer 500; every refusal body matches a
//! defined shape; a GET never hands the blob to the wrong key; `/version`
//! matches the store it reports on.

mod common;

use std::panic::AssertUnwindSafe;

use axum::{
    body::Body,
    http::{HeaderName, HeaderValue, Method, Request, Response, StatusCode},
    Router,
};
use common::*;
use http_body_util::BodyExt;
use plsfix_server::{app, relay::AppState, store::Store};
use proptest::prelude::*;
use proptest::test_runner::{Config as ProptestConfig, FileFailurePersistence, RngSeed};
use tower::ServiceExt;

fn config() -> ProptestConfig {
    ProptestConfig {
        cases: 150,
        rng_seed: RngSeed::Fixed(0xC0FFEE),
        failure_persistence: Some(Box::new(FileFailurePersistence::Off)),
        ..ProptestConfig::default()
    }
}

/// `relay_app()` has no async runtime of its own, so this drives one request
/// to completion on a fresh throwaway runtime instead.
fn run<F: std::future::Future>(future: F) -> F::Output {
    tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .expect("build a throwaway runtime")
        .block_on(future)
}

/// `Ok` with the response, or `Err` with the panic message - a panic inside a
/// handler that runs on the async task itself (not one caught already by
/// `blocking::store_call`'s `spawn_blocking`) would otherwise abort this test
/// process rather than fail one case.
fn dispatch(app: &Router, request: Request<Body>) -> Result<Response<Body>, String> {
    std::panic::catch_unwind(AssertUnwindSafe(|| run(app.clone().oneshot(request))))
        .map(|infallible| infallible.unwrap())
        // `&*payload`, not `&payload`: `Box<dyn Any + Send>` is itself `Any`
        // (the blanket impl), so a bare `&payload` downcasts against the BOX
        // and always misses; deref once first to reach the real payload.
        .map_err(|payload| panic_message(&*payload))
}

fn panic_message(payload: &(dyn std::any::Any + Send)) -> String {
    if let Some(message) = payload.downcast_ref::<&str>() {
        message.to_string()
    } else if let Some(message) = payload.downcast_ref::<String>() {
        message.clone()
    } else {
        "panic payload was neither &str nor String".to_string()
    }
}

async fn body_bytes(response: Response<Body>) -> Vec<u8> {
    response
        .into_body()
        .collect()
        .await
        .unwrap()
        .to_bytes()
        .to_vec()
}

// ---------------------------------------------------------------------------
// Generators: methods, path segments under /api/, header sets, bodies.
// ---------------------------------------------------------------------------

const METHODS: &[&str] = &["GET", "PUT", "POST", "DELETE", "PATCH", "HEAD", "OPTIONS"];

fn method_strategy() -> impl Strategy<Value = Method> {
    (0..METHODS.len()).prop_map(|i| Method::from_bytes(METHODS[i].as_bytes()).unwrap())
}

// Every raw byte here is a valid, unencoded RFC 3986 pchar, so the generated
// URI always parses: an unparseable string is a fact about the `http` crate,
// not the relay, and would only get in the way of exercising its router.
const SAFE_CHARS: &[char] = &[
    'a', 'b', 'c', 'x', 'y', 'z', 'A', 'Z', '0', '1', '9', '-', '_', '.', '~', '!', '$', '&', '\'',
    '(', ')', '*', '+', ',', ';', '=', ':', '@',
];

/// Mostly short garbage, never 32 hex characters (a link id's shape); mixes
/// in a right-shaped id sometimes so the segment after `links` can pass.
fn garbage_segment_strategy() -> impl Strategy<Value = String> {
    prop_oneof![
        5 => prop::collection::vec(prop::sample::select(SAFE_CHARS), 0..14)
            .prop_map(|chars| chars.into_iter().collect::<String>()),
        2 => prop::collection::vec(any::<u8>(), 16)
            .prop_map(|bytes| bytes.iter().map(|b| format!("{b:02x}")).collect::<String>()),
        1 => Just("%2e%2e".to_string()),
        1 => Just("%00".to_string()),
        1 => Just("..".to_string()),
    ]
}

/// Mostly a real route prefix (so handlers are actually reached, not only the
/// 404 fallback) with 0-3 garbage segments tacked on; sometimes a fully
/// invented head instead.
fn path_tail_strategy() -> impl Strategy<Value = String> {
    let head = prop_oneof![
        3 => Just("links".to_string()),
        3 => Just("inbox".to_string()),
        1 => Just("links/status".to_string()),
        1 => Just("links/fetch".to_string()),
        1 => Just("links/touch".to_string()),
        2 => garbage_segment_strategy(),
    ];
    (
        head,
        prop::collection::vec(garbage_segment_strategy(), 0..3),
    )
        .prop_map(|(head, rest)| {
            let mut parts = vec![head];
            parts.extend(rest);
            parts.join("/")
        })
}

const HEADER_NAMES: &[&str] = &[
    "authorization",
    "content-type",
    "if-none-match",
    "x-plsfix-link-id",
    "x-forwarded-for",
    "cf-connecting-ip",
    "x-hunt2-nonsense",
];

fn headers_strategy() -> impl Strategy<Value = Vec<(String, Vec<u8>)>> {
    let name = prop::sample::select(HEADER_NAMES).prop_map(str::to_string);
    // Values are mostly arbitrary bytes, never `Bearer ` + a real 43-char
    // key (50 bytes); a right-shaped bearer some of the time so it can pass.
    let value = prop_oneof![
        4 => prop::collection::vec(any::<u8>(), 0..40),
        1 => Just(format!("Bearer {AUTH}").into_bytes()),
    ];
    prop::collection::vec((name, value), 0..5)
}

fn body_strategy() -> impl Strategy<Value = Vec<u8>> {
    prop::collection::vec(any::<u8>(), 0..300)
}

/// Only a header whose bytes a real client could actually send at all - the
/// rest is skipped, the way a well-formed HTTP client would never offer them.
fn build_request(
    method: &Method,
    path: &str,
    headers: &[(String, Vec<u8>)],
    body: &[u8],
) -> Request<Body> {
    let mut builder = Request::builder()
        .method(method.clone())
        .uri(format!("/api/{path}"));
    for (name, value) in headers {
        if let (Ok(name), Ok(value)) = (
            HeaderName::from_bytes(name.as_bytes()),
            HeaderValue::from_bytes(value),
        ) {
            builder = builder.header(name, value);
        }
    }
    builder.body(Body::from(body.to_vec())).unwrap()
}

/// The exact (status, message) pairs `refused.rs` can produce. Kept here
/// rather than imported, because the point is to notice the day a refusal's
/// wording or status code drifts from this list, not to import the list it
/// would then trivially agree with.
fn is_a_defined_refusal(status: u16, message: &str) -> bool {
    matches!(
        (status, message),
        (401, "bearer required")
            | (400, "bad id")
            | (400, "bad rev")
            | (400, "bad body")
            | (400, "too many items")
            | (403, "another key owns this")
            | (404, "not found")
            | (507, "storage full")
            | (429, "too many requests")
            | (500, "store error")
    )
}

proptest! {
    #![proptest_config(config())]

    /// Property: whatever a client sends under `/api/`, the relay answers
    /// something - it never panics, and it never falls back to the generic
    /// 500 a caught panic or an unmapped store error would produce, because
    /// nothing in a malformed request should ever reach the store layer at
    /// all before its shape is validated.
    #[test]
    fn an_arbitrary_api_request_never_panics_or_answers_five_hundred(
        method in method_strategy(),
        path in path_tail_strategy(),
        headers in headers_strategy(),
        body in body_strategy(),
    ) {
        let app = relay_app();
        let request = build_request(&method, &path, &headers, &body);
        let outcome = dispatch(&app, request);
        let response = match outcome {
            Ok(response) => response,
            Err(message) => {
                prop_assert!(false, "{} /api/{} panicked: {}", method, path, message);
                unreachable!()
            }
        };
        prop_assert_ne!(
            response.status(),
            StatusCode::INTERNAL_SERVER_ERROR,
            "{} /api/{} answered 500",
            method,
            path
        );
    }

    /// Property: every response whose body is JSON shaped like a refusal -
    /// one object, one key `"error"`, a string value - is one of the
    /// (status, message) pairs `refused.rs` actually defines. The relay never
    /// invents a refusal shape ad hoc, and never puts the right message on
    /// the wrong status code.
    #[test]
    fn every_json_refusal_body_matches_a_defined_shape(
        method in method_strategy(),
        path in path_tail_strategy(),
        headers in headers_strategy(),
        body in body_strategy(),
    ) {
        let app = relay_app();
        let request = build_request(&method, &path, &headers, &body);
        let Ok(response) = dispatch(&app, request) else {
            return Ok(()); // the panic property above already covers this case
        };
        let status = response.status().as_u16();
        let bytes = run(body_bytes(response));
        let Ok(value) = serde_json::from_slice::<serde_json::Value>(&bytes) else {
            return Ok(()); // not a JSON body at all - axum/tower's own rejection
        };
        let Some(object) = value.as_object() else {
            return Ok(());
        };
        if object.len() == 1 {
            if let Some(message) = object.get("error").and_then(|v| v.as_str()) {
                prop_assert!(
                    is_a_defined_refusal(status, message),
                    "{} /api/{} answered {} {{\"error\":\"{}\"}}, not a defined shape",
                    method,
                    path,
                    status,
                    message
                );
            }
        }
    }
}

const KEY_CHARS: &[char] = &['A', 'B', 'C', 'a', 'b', 'c', '0', '1', '9', '-', '_'];

fn bearer_strategy() -> impl Strategy<Value = String> {
    prop_oneof![
        // The right length, wrong content - the case most likely to slip
        // past a check that only looks at shape.
        3 => prop::collection::vec(prop::sample::select(KEY_CHARS), 43)
            .prop_map(|chars| chars.into_iter().collect::<String>()),
        2 => prop::collection::vec(prop::sample::select(KEY_CHARS), 0..60)
            .prop_map(|chars| chars.into_iter().collect::<String>()),
        1 => Just(String::new()),
    ]
}

proptest! {
    #![proptest_config(config())]

    /// Property: once a link is pushed under one key, a GET with any other
    /// bearer - right shape or wrong, empty or long - never comes back with
    /// that blob. A wrong-shaped key is 401 before the store is even asked;
    /// a right-shaped-but-wrong key is 403; neither is ever 200 or 304 (which
    /// leaves the blob's own bytes as the only thing left to check, and 304
    /// carries none).
    #[test]
    fn a_get_never_returns_the_blob_to_the_wrong_key(wrong_bearer in bearer_strategy()) {
        const ID: &str = "0123456789abcdef0123456789abcdef";
        const REAL_AUTH: &str = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";
        prop_assume!(wrong_bearer != REAL_AUTH);

        let app = relay_app();
        let push = Request::builder()
            .method("PUT")
            .uri(format!("/api/links/{ID}"))
            .header("authorization", format!("Bearer {REAL_AUTH}"))
            .body(Body::from(b"the real blob".to_vec()))
            .unwrap();
        let pushed = dispatch(&app, push).expect("the push itself must not panic");
        prop_assert_eq!(pushed.status(), StatusCode::OK);

        let get = Request::builder()
            .method("GET")
            .uri(format!("/api/links/{ID}"))
            .header("authorization", format!("Bearer {wrong_bearer}"))
            .body(Body::empty())
            .unwrap();
        let response = dispatch(&app, get).expect("a wrong-key GET must not panic either");
        let status = response.status();
        prop_assert_ne!(status, StatusCode::OK, "the wrong key read the blob back");
        prop_assert_ne!(status, StatusCode::NOT_MODIFIED, "a 304 implies a matching etag");
    }
}

// ---------------------------------------------------------------------------
// /version: only the full `app()` router serves it, so this half uses that
// instead of `relay_app()` - the bare `/api` router has no such route at all.
// ---------------------------------------------------------------------------

fn version_app() -> (Router, std::sync::Arc<AppState>) {
    let state = std::sync::Arc::new(AppState::new(Store::in_memory().unwrap()));
    let app = app(static_dir("hunt2-relay-version"), state.clone());
    (app, state)
}

fn version_counts(app: &Router) -> (Option<i64>, Option<i64>, Option<i64>, Option<i64>) {
    let response = dispatch(app, Request::get("/version").body(Body::empty()).unwrap())
        .expect("a GET /version must not panic");
    let bytes = run(body_bytes(response));
    let value: serde_json::Value = serde_json::from_slice(&bytes).unwrap();
    let relay = &value["relay"];
    (
        relay["links"].as_i64(),
        relay["revisions"].as_i64(),
        relay["inbox"].as_i64(),
        relay["bytes"].as_i64(),
    )
}

proptest! {
    #![proptest_config(ProptestConfig { cases: 40, ..config() })]

    /// Property: after any sequence of pushes, `/version`'s counters, with
    /// the cache force-cleared each time, equal `Store::counts()` at that
    /// instant - checking the field mapping into `/version`'s JSON. The
    /// force-clear is load-bearing: only the hourly sweeper clears the cache
    /// in production, never a write.
    #[test]
    fn version_counts_always_match_what_the_store_reports(
        pushes in prop::collection::vec((0usize..3, prop::collection::vec(any::<u8>(), 0..40)), 1..6),
    ) {
        const AUTH: &str = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";
        let (app, state) = version_app();
        for (id, blob) in &pushes {
            let request = Request::builder()
                .method("PUT")
                .uri(format!("/api/links/{id:032x}"))
                .header("authorization", format!("Bearer {AUTH}"))
                .body(Body::from(blob.clone()))
                .unwrap();
            dispatch(&app, request).expect("a push must not panic");

            run(state.counts.clear());
            let (links, revisions, inbox, bytes) = version_counts(&app);
            let truth = state.store.counts().unwrap();
            prop_assert_eq!(links, Some(truth.links), "links after pushing id {}", id);
            prop_assert_eq!(revisions, Some(truth.revisions), "revisions after pushing id {}", id);
            prop_assert_eq!(inbox, Some(truth.inbox), "inbox after pushing id {}", id);
            prop_assert_eq!(bytes, Some(truth.bytes), "bytes after pushing id {}", id);
        }
    }
}

/// Pins the mechanism `dispatch()` relies on to report a readable panic
/// message instead of the generic fallback: `Box<dyn Any + Send>` is itself
/// `Any` (the blanket impl), so `payload.downcast_ref()` on a bare reference
/// to the box matches the BOX, never the value inside it, and always misses -
/// found by hand when a real handler panic came back as "neither &str nor
/// String" although the exact same payload downcast fine one line earlier.
#[test]
fn panic_message_recovers_the_real_text_not_a_generic_fallback() {
    let result = std::panic::catch_unwind(AssertUnwindSafe(|| -> Response<Body> {
        panic!("a distinctive message nothing else in this file produces");
    }));
    let Err(payload) = result else {
        panic!("expected the probe closure to panic");
    };
    assert_eq!(
        panic_message(&*payload),
        "a distinctive message nothing else in this file produces"
    );
}
