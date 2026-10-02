//! Route tests for the revision fetch `GET /api/links/{id}?rev=<n>`, which is
//! what "Revert last update" in the PowerPoint pane reaches for: the exact rev
//! with its own ETag, 404 once retention has dropped it, 403 for a foreign
//! bearer and 400 for a rev that is not a positive integer.

use plsfix_server::{relay::*, store::Store};

use axum::{
    body::Body,
    http::{header, Request, StatusCode},
    response::Response,
    Router,
};
use http_body_util::BodyExt;
use tower::ServiceExt;

const ID: &str = "0123456789abcdef0123456789abcdef";
const MINE: &str = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"; // 43 chars
const THEIRS: &str = "BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB";

fn app() -> Router {
    routes(std::sync::Arc::new(AppState::new(
        Store::in_memory().unwrap(),
    )))
}

fn call(method: &str, path: &str, auth: &str, body: Vec<u8>) -> Request<Body> {
    Request::builder()
        .method(method)
        .uri(path)
        .header(header::AUTHORIZATION, format!("Bearer {auth}"))
        .header(header::CONTENT_TYPE, "application/octet-stream")
        .body(Body::from(body))
        .unwrap()
}

async fn send(app: &Router, request: Request<Body>) -> Response {
    app.clone().oneshot(request).await.unwrap()
}

async fn push(app: &Router, blob: &[u8]) {
    let path = format!("/api/links/{ID}");
    assert_eq!(
        send(app, call("PUT", &path, MINE, blob.to_vec()))
            .await
            .status(),
        StatusCode::OK
    );
}

fn rev_path(rev: &str) -> String {
    format!("/api/links/{ID}?rev={rev}")
}

#[tokio::test]
async fn a_named_rev_answers_with_its_own_blob_and_etag() {
    let app = app();
    push(&app, b"one").await;
    push(&app, b"two").await;

    let response = send(&app, call("GET", &rev_path("1"), MINE, vec![])).await;
    assert_eq!(response.status(), StatusCode::OK);
    assert_eq!(response.headers().get(header::ETAG).unwrap(), "\"1\"");
    let body = response.into_body().collect().await.unwrap().to_bytes();
    assert_eq!(&body[..], b"one");

    // Without the query the route still serves the head, ETag and all.
    let path = format!("/api/links/{ID}");
    let head = send(&app, call("GET", &path, MINE, vec![])).await;
    assert_eq!(head.headers().get(header::ETAG).unwrap(), "\"2\"");
}

#[tokio::test]
async fn a_dropped_or_unknown_rev_is_404_and_a_foreign_bearer_403() {
    let app = app();
    push(&app, b"one").await;
    push(&app, b"two").await;
    // A foreign key is refused before the revision is even looked for.
    assert_eq!(
        send(&app, call("GET", &rev_path("1"), THEIRS, vec![]))
            .await
            .status(),
        StatusCode::FORBIDDEN
    );
    assert_eq!(
        send(&app, call("GET", &rev_path("7"), MINE, vec![]))
            .await
            .status(),
        StatusCode::NOT_FOUND
    );
    // The third push drops rev 1: retention, not expiry, and still a 404.
    push(&app, b"three").await;
    assert_eq!(
        send(&app, call("GET", &rev_path("1"), MINE, vec![]))
            .await
            .status(),
        StatusCode::NOT_FOUND
    );
}

#[tokio::test]
async fn a_rev_that_is_not_a_positive_integer_is_a_bad_request() {
    let app = app();
    push(&app, b"one").await;
    for rev in ["abc", "0", "-1", ""] {
        assert_eq!(
            send(&app, call("GET", &rev_path(rev), MINE, vec![]))
                .await
                .status(),
            StatusCode::BAD_REQUEST,
            "rev={rev}"
        );
    }
}
