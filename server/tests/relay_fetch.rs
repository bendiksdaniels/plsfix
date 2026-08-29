//! Route tests for the batched fetch `POST /api/links/fetch`, which is what a
//! PowerPoint "Update all" makes instead of one GET per link: the changed
//! blobs in one response, everything else named in `omitted`, the 4 MiB cap
//! deferring the rest of the batch, and the same 200-item and 64 KiB limits
//! the status batch carries.

use plsfix_server::{relay::*, store::Store};

use axum::{
    body::Body,
    http::{header, Request, StatusCode},
    response::Response,
    Router,
};
use http_body_util::BodyExt;
use tower::ServiceExt;

const CHANGED: &str = "0123456789abcdef0123456789abcdef";
const UNCHANGED: &str = "11111111111111111111111111111111";
const UNKNOWN: &str = "ffffffffffffffffffffffffffffffff";
const FOREIGN: &str = "22222222222222222222222222222222";
const MINE: &str = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"; // 43 chars
const THEIRS: &str = "BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB";

fn app() -> Router {
    routes(std::sync::Arc::new(AppState {
        store: Store::in_memory().unwrap(),
    }))
}

async fn send(app: &Router, request: Request<Body>) -> Response {
    app.clone().oneshot(request).await.unwrap()
}

async fn push(app: &Router, id: &str, auth: &str, blob: Vec<u8>) {
    let request = Request::builder()
        .method("PUT")
        .uri(format!("/api/links/{id}"))
        .header(header::AUTHORIZATION, format!("Bearer {auth}"))
        .header(header::CONTENT_TYPE, "application/octet-stream")
        .body(Body::from(blob))
        .unwrap();
    assert_eq!(send(app, request).await.status(), StatusCode::OK);
}

fn batch(body: Vec<u8>) -> Request<Body> {
    Request::builder()
        .method("POST")
        .uri("/api/links/fetch")
        .header(header::CONTENT_TYPE, "application/json")
        .body(Body::from(body))
        .unwrap()
}

/// `{"id":"..","auth":"..","knownRev":n}`, the deck's side of one link.
fn query(id: &str, auth: &str, known_rev: Option<i64>) -> String {
    let rev = known_rev.map_or("null".to_string(), |rev| rev.to_string());
    format!(r#"{{"id":"{id}","auth":"{auth}","knownRev":{rev}}}"#)
}

async fn fetched(app: &Router, queries: &[String]) -> String {
    let response = send(app, batch(format!("[{}]", queries.join(",")).into_bytes())).await;
    assert_eq!(response.status(), StatusCode::OK);
    let bytes = response.into_body().collect().await.unwrap().to_bytes();
    String::from_utf8(bytes.to_vec()).unwrap()
}

#[tokio::test]
async fn a_mixed_batch_answers_with_the_changed_blobs_only() {
    let app = app();
    push(&app, CHANGED, MINE, b"one".to_vec()).await;
    push(&app, CHANGED, MINE, b"two".to_vec()).await;
    push(&app, UNCHANGED, MINE, b"held".to_vec()).await;
    push(&app, FOREIGN, THEIRS, b"someone else".to_vec()).await;

    let text = fetched(
        &app,
        &[
            query(CHANGED, MINE, Some(1)),
            query(UNCHANGED, MINE, Some(1)),
            query(UNKNOWN, MINE, Some(1)),
            query(FOREIGN, MINE, Some(1)),
        ],
    )
    .await;

    // "two" is the head, base64url encoded; the other three carry no blob and
    // say why, in request order.
    assert_eq!(
        text,
        format!(
            concat!(
                r#"{{"items":[{{"id":"{changed}","rev":2,"blob":"dHdv"}}],"#,
                r#""omitted":[{{"id":"{unchanged}","reason":"unchanged"}},"#,
                r#"{{"id":"{unknown}","reason":"missing"}},"#,
                r#"{{"id":"{foreign}","reason":"auth"}}]}}"#
            ),
            changed = CHANGED,
            unchanged = UNCHANGED,
            unknown = UNKNOWN,
            foreign = FOREIGN
        )
    );
}

// A deck that holds nothing yet asks without a revision, and gets the head.
#[tokio::test]
async fn a_query_without_a_known_rev_always_gets_the_blob() {
    let app = app();
    push(&app, CHANGED, MINE, b"two".to_vec()).await;
    let text = fetched(&app, &[query(CHANGED, MINE, None)]).await;
    assert!(text.contains(r#""rev":1,"blob":"dHdv""#), "{text}");
}

#[tokio::test]
async fn the_blob_cap_defers_the_rest_of_the_batch() {
    let app = app();
    let big = vec![0u8; 2_500_000];
    push(&app, CHANGED, MINE, big.clone()).await;
    push(&app, UNCHANGED, MINE, big).await;
    // Small enough to fit in what the first blob left, and still deferred: the
    // cap stops the batch, it does not sieve it.
    push(&app, FOREIGN, MINE, b"tiny".to_vec()).await;

    let text = fetched(
        &app,
        &[
            query(CHANGED, MINE, None),
            query(UNCHANGED, MINE, None),
            query(FOREIGN, MINE, None),
        ],
    )
    .await;

    assert!(
        text.starts_with(&format!(r#"{{"items":[{{"id":"{CHANGED}","rev":1,"blob":"#)),
        "the first blob is the only one in the response"
    );
    assert!(text.ends_with(&format!(
        concat!(
            r#""omitted":[{{"id":"{unchanged}","reason":"deferred"}},"#,
            r#"{{"id":"{tiny}","reason":"deferred"}}]}}"#
        ),
        unchanged = UNCHANGED,
        tiny = FOREIGN
    )));
}

#[tokio::test]
async fn an_overlong_or_oversize_batch_is_refused() {
    let app = app();
    let one = query(CHANGED, MINE, Some(1));
    let too_many: Vec<String> = std::iter::repeat_n(one, 201).collect();
    let body = format!("[{}]", too_many.join(",")).into_bytes();
    assert!(
        body.len() < 64 * 1024,
        "the item count is what refuses this"
    );
    assert_eq!(
        send(&app, batch(body)).await.status(),
        StatusCode::BAD_REQUEST
    );
    // Past the body limit nothing is parsed at all: refused at the socket.
    assert_eq!(
        send(&app, batch(vec![b'x'; 70 * 1024])).await.status(),
        StatusCode::PAYLOAD_TOO_LARGE
    );
    assert_eq!(
        send(&app, batch(b"not json".to_vec())).await.status(),
        StatusCode::BAD_REQUEST
    );
    // An empty deck is a legal batch, and answers with two empty lists.
    assert_eq!(fetched(&app, &[]).await, r#"{"items":[],"omitted":[]}"#);
}
