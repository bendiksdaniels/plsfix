//! Route tests for the relay: auth and validation codes, ETag round trip,
//! body limits, the status batch and the inbox lifecycle. In-memory store, so
//! each test starts from an empty relay.

use plsfix_server::{relay::*, store::Store};

mod tests {
    use super::*;
    use axum::{
        body::Body,
        http::{header, Request, StatusCode},
    };
    use http_body_util::BodyExt;
    use tower::ServiceExt;

    const ID: &str = "0123456789abcdef0123456789abcdef";
    const AUTH: &str = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"; // 43 chars

    fn app() -> axum::Router {
        routes(std::sync::Arc::new(AppState {
            store: Store::in_memory().unwrap(),
        }))
    }

    fn req(method: &str, path: &str, auth: Option<&str>, body: Vec<u8>) -> Request<Body> {
        let mut builder = Request::builder()
            .method(method)
            .uri(path)
            .header(header::CONTENT_TYPE, "application/octet-stream");
        if let Some(auth) = auth {
            builder = builder.header(header::AUTHORIZATION, format!("Bearer {auth}"));
        }
        // The inbox POST carries the link id in a header; the body is the item.
        if method == "POST" && path.starts_with("/api/inbox/") {
            builder = builder.header("X-PLSFIX-Link-Id", ID);
        }
        builder.body(Body::from(body)).unwrap()
    }

    #[tokio::test]
    async fn put_get_round_trip_with_etag() {
        let app = app();
        let response = app
            .clone()
            .oneshot(req(
                "PUT",
                &format!("/api/links/{ID}"),
                Some(AUTH),
                b"blob".to_vec(),
            ))
            .await
            .unwrap();
        assert_eq!(response.status(), StatusCode::OK);
        let body = response.into_body().collect().await.unwrap().to_bytes();
        assert_eq!(&body[..], br#"{"rev":1}"#);

        let response = app
            .clone()
            .oneshot(req("GET", &format!("/api/links/{ID}"), Some(AUTH), vec![]))
            .await
            .unwrap();
        assert_eq!(response.status(), StatusCode::OK);
        assert_eq!(response.headers().get(header::ETAG).unwrap(), "\"1\"");

        let mut request = req("GET", &format!("/api/links/{ID}"), Some(AUTH), vec![]);
        request
            .headers_mut()
            .insert(header::IF_NONE_MATCH, "\"1\"".parse().unwrap());
        assert_eq!(
            app.oneshot(request).await.unwrap().status(),
            StatusCode::NOT_MODIFIED
        );
    }

    #[tokio::test]
    async fn auth_and_validation_codes() {
        let app = app();
        assert_eq!(
            app.clone()
                .oneshot(req("PUT", &format!("/api/links/{ID}"), None, vec![1]))
                .await
                .unwrap()
                .status(),
            StatusCode::UNAUTHORIZED
        );
        assert_eq!(
            app.clone()
                .oneshot(req("PUT", "/api/links/not-hex", Some(AUTH), vec![1]))
                .await
                .unwrap()
                .status(),
            StatusCode::BAD_REQUEST
        );
        app.clone()
            .oneshot(req("PUT", &format!("/api/links/{ID}"), Some(AUTH), vec![1]))
            .await
            .unwrap();
        let other = "BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB";
        assert_eq!(
            app.clone()
                .oneshot(req("GET", &format!("/api/links/{ID}"), Some(other), vec![]))
                .await
                .unwrap()
                .status(),
            StatusCode::FORBIDDEN
        );
        assert_eq!(
            app.oneshot(req(
                "GET",
                "/api/links/ffffffffffffffffffffffffffffffff",
                Some(AUTH),
                vec![]
            ))
            .await
            .unwrap()
            .status(),
            StatusCode::NOT_FOUND
        );
    }

    #[tokio::test]
    async fn oversize_bodies_are_413() {
        let app = app();
        let big = vec![0u8; 4 * 1024 * 1024 + 1];
        assert_eq!(
            app.oneshot(req("PUT", &format!("/api/links/{ID}"), Some(AUTH), big))
                .await
                .unwrap()
                .status(),
            StatusCode::PAYLOAD_TOO_LARGE
        );
    }

    #[tokio::test]
    async fn status_batch_reports_rev_missing_and_auth() {
        let app = app();
        app.clone()
            .oneshot(req("PUT", &format!("/api/links/{ID}"), Some(AUTH), vec![1]))
            .await
            .unwrap();
        let body = format!(
            r#"[{{"id":"{ID}","auth":"{AUTH}"}},{{"id":"ffffffffffffffffffffffffffffffff","auth":"{AUTH}"}}]"#
        );
        let request = Request::builder()
            .method("POST")
            .uri("/api/links/status")
            .header(header::CONTENT_TYPE, "application/json")
            .body(Body::from(body))
            .unwrap();
        let response = app.oneshot(request).await.unwrap();
        assert_eq!(response.status(), StatusCode::OK);
        let text = String::from_utf8(
            response
                .into_body()
                .collect()
                .await
                .unwrap()
                .to_bytes()
                .to_vec(),
        )
        .unwrap();
        assert!(text.contains(r#""rev":1"#));
        assert!(text.contains(r#""rev":null"#));
    }

    #[tokio::test]
    async fn inbox_post_is_idempotent_and_lists_latest() {
        let app = app();
        let ws = "CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC";
        assert_eq!(
            app.clone()
                .oneshot(req(
                    "POST",
                    &format!("/api/inbox/{ws}"),
                    Some(AUTH),
                    b"item".to_vec()
                ))
                .await
                .unwrap()
                .status(),
            StatusCode::OK
        );
        // the same id is idempotent (re-export replaces the row)
        assert_eq!(
            app.clone()
                .oneshot(req(
                    "POST",
                    &format!("/api/inbox/{ws}"),
                    Some(AUTH),
                    b"item2".to_vec()
                ))
                .await
                .unwrap()
                .status(),
            StatusCode::OK
        );
        let response = app
            .oneshot(req("GET", &format!("/api/inbox/{ws}"), Some(AUTH), vec![]))
            .await
            .unwrap();
        let text = String::from_utf8(
            response
                .into_body()
                .collect()
                .await
                .unwrap()
                .to_bytes()
                .to_vec(),
        )
        .unwrap();
        assert!(text.contains("\"blob\":\"aXRlbTI\""));
    }

    #[tokio::test]
    async fn inbox_delete_removes_the_item() {
        let app = app();
        let ws = "CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC";
        assert_eq!(
            app.clone()
                .oneshot(req(
                    "POST",
                    &format!("/api/inbox/{ws}"),
                    Some(AUTH),
                    b"item".to_vec()
                ))
                .await
                .unwrap()
                .status(),
            StatusCode::OK
        );
        assert_eq!(
            app.oneshot(req(
                "DELETE",
                &format!("/api/inbox/{ws}/{ID}"),
                Some(AUTH),
                vec![]
            ))
            .await
            .unwrap()
            .status(),
            StatusCode::OK
        );
    }
}

// Body limits that are not the payload's, and what a foreign bearer can do to
// a workspace inbox. These build their own requests (several keys, several
// methods), so they sit apart from the shared `req` helper above.
mod limits_and_keys {
    use super::*;
    use axum::body::Body;
    use axum::http::{header, Request, StatusCode};
    use axum::response::Response;
    use axum::Router;
    use http_body_util::BodyExt;
    use tower::ServiceExt;

    const WS: &str = "CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC";
    const LINK: &str = "0123456789abcdef0123456789abcdef";
    const MINE: &str = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";
    const THEIRS: &str = "BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB";

    fn app() -> Router {
        routes(std::sync::Arc::new(AppState {
            store: Store::in_memory().unwrap(),
        }))
    }

    fn post_item(auth: &str, body: Vec<u8>) -> Request<Body> {
        Request::builder()
            .method("POST")
            .uri(format!("/api/inbox/{WS}"))
            .header(header::AUTHORIZATION, format!("Bearer {auth}"))
            .header("x-plsfix-link-id", LINK)
            .header(header::CONTENT_TYPE, "application/octet-stream")
            .body(Body::from(body))
            .unwrap()
    }

    fn call(auth: &str, method: &str, path: String) -> Request<Body> {
        Request::builder()
            .method(method)
            .uri(path)
            .header(header::AUTHORIZATION, format!("Bearer {auth}"))
            .body(Body::empty())
            .unwrap()
    }

    async fn send(app: &Router, request: Request<Body>) -> Response {
        app.clone().oneshot(request).await.unwrap()
    }

    async fn listing(app: &Router, auth: &str) -> String {
        let response = send(app, call(auth, "GET", format!("/api/inbox/{WS}"))).await;
        let bytes = response.into_body().collect().await.unwrap().to_bytes();
        String::from_utf8(bytes.to_vec()).unwrap()
    }

    #[tokio::test]
    async fn inbox_bodies_over_64k_are_413() {
        let app = app();
        assert_eq!(
            send(&app, post_item(MINE, b"item".to_vec())).await.status(),
            StatusCode::OK
        );
        assert_eq!(
            send(&app, post_item(MINE, vec![0u8; 64 * 1024 + 1]))
                .await
                .status(),
            StatusCode::PAYLOAD_TOO_LARGE
        );
    }

    #[tokio::test]
    async fn a_foreign_key_writes_only_its_own_inbox_row() {
        let app = app();
        assert_eq!(
            send(&app, post_item(MINE, b"item".to_vec())).await.status(),
            StatusCode::OK
        );
        // A bearer that is not the workspace's is accepted - the server cannot
        // tell the two apart - but it lands in a row of its own, and the
        // owner's listing is untouched.
        assert_eq!(
            send(&app, post_item(THEIRS, b"squat".to_vec()))
                .await
                .status(),
            StatusCode::OK
        );
        let mine = listing(&app, MINE).await;
        assert!(mine.contains("\"blob\":\"aXRlbQ\""));
        assert!(!mine.contains("c3F1YXQ"));
        // A DELETE from the squatter takes its own row and leaves the owner's.
        let path = format!("/api/inbox/{WS}/{LINK}");
        assert_eq!(
            send(&app, call(THEIRS, "DELETE", path.clone()))
                .await
                .status(),
            StatusCode::OK
        );
        assert!(listing(&app, MINE).await.contains("\"blob\":\"aXRlbQ\""));
        // A key that wrote nothing sees no row to delete: 404, never 403.
        let stranger = "DDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDD";
        assert_eq!(
            send(&app, call(stranger, "DELETE", path.clone()))
                .await
                .status(),
            StatusCode::NOT_FOUND
        );
        assert_eq!(
            send(&app, call(MINE, "DELETE", path)).await.status(),
            StatusCode::OK
        );
    }

    #[tokio::test]
    async fn an_oversize_status_batch_is_refused_at_the_socket() {
        // The status route has no bearer and sits behind the Access bypass: it
        // must not inherit the 4 MiB payload limit and parse an attacker's JSON.
        let app = app();
        let batch = |body: Vec<u8>| {
            Request::builder()
                .method("POST")
                .uri("/api/links/status")
                .header(header::CONTENT_TYPE, "application/json")
                .body(Body::from(body))
                .unwrap()
        };
        assert_eq!(
            send(&app, batch(vec![b'x'; 70 * 1024])).await.status(),
            StatusCode::PAYLOAD_TOO_LARGE
        );
        // A real batch is nowhere near the limit and still answers.
        assert_eq!(
            send(&app, batch(b"[]".to_vec())).await.status(),
            StatusCode::OK
        );
    }
}
