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
            builder = builder.header("X-SMT-Link-Id", ID);
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
    async fn inbox_lifecycle() {
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
            .clone()
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

#[tokio::test]
async fn inbox_bodies_over_64k_are_413_and_foreign_keys_are_403() {
    use axum::body::Body;
    use axum::http::{header, Request, StatusCode};
    use tower::ServiceExt;
    let app = routes(std::sync::Arc::new(AppState {
        store: Store::in_memory().unwrap(),
    }));
    let ws = "CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC";
    let id = "0123456789abcdef0123456789abcdef";
    let build = |auth: &str, size: usize| {
        Request::builder()
            .method("POST")
            .uri(format!("/api/inbox/{ws}"))
            .header(header::AUTHORIZATION, format!("Bearer {auth}"))
            .header("x-smt-link-id", id)
            .header(header::CONTENT_TYPE, "application/octet-stream")
            .body(Body::from(vec![0u8; size]))
            .unwrap()
    };
    let mine = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";
    let theirs = "BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB";
    let status = |r: axum::response::Response| r.status();
    assert_eq!(
        status(app.clone().oneshot(build(mine, 60 * 1024)).await.unwrap()),
        StatusCode::OK
    );
    assert_eq!(
        status(
            app.clone()
                .oneshot(build(mine, 64 * 1024 + 1))
                .await
                .unwrap()
        ),
        StatusCode::PAYLOAD_TOO_LARGE
    );
    assert_eq!(
        status(app.oneshot(build(theirs, 10)).await.unwrap()),
        StatusCode::FORBIDDEN
    );
}
