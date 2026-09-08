//! Shared harness for the E2 audit suites: a router over an in-memory store,
//! one request builder and one send, and a static tree per test. No assertions
//! of its own - each suite owns what it proves.
//! Invariant: every helper here is host-shaped, so a suite never hand-rolls a
//! request that the pane's client would not make.
#![allow(dead_code)] // each suite uses a different half of this file

use std::{path::PathBuf, sync::Arc};

use axum::{
    body::Body,
    http::{header, Request, StatusCode},
    Router,
};
use http_body_util::BodyExt;
use plsfix_server::{relay::*, store::Store};
use tower::ServiceExt;

pub const ID: &str = "0123456789abcdef0123456789abcdef";
pub const AUTH: &str = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"; // 43 chars
pub const OTHER: &str = "BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB";

pub fn relay_app() -> Router {
    routes(Arc::new(AppState::new(Store::in_memory().unwrap())))
}

pub fn req(method: &str, path: &str, auth: Option<&str>, body: Vec<u8>) -> Request<Body> {
    let mut builder = Request::builder().method(method).uri(path);
    if let Some(auth) = auth {
        builder = builder.header(header::AUTHORIZATION, format!("Bearer {auth}"));
    }
    // The inbox POST names its link in a header; the body is the sealed item.
    if method == "POST" && path.starts_with("/api/inbox/") {
        builder = builder.header("X-PLSFIX-Link-Id", ID);
    }
    builder.body(Body::from(body)).unwrap()
}

pub async fn send(app: &Router, request: Request<Body>) -> (StatusCode, String) {
    let response = app.clone().oneshot(request).await.unwrap();
    let status = response.status();
    let body = response.into_body().collect().await.unwrap().to_bytes();
    (status, String::from_utf8_lossy(&body).to_string())
}

pub async fn call(
    app: &Router,
    method: &str,
    path: &str,
    auth: Option<&str>,
    body: &[u8],
) -> (StatusCode, String) {
    send(app, req(method, path, auth, body.to_vec())).await
}

pub async fn get(app: &Router, path: &str) -> (StatusCode, String) {
    send(app, Request::get(path).body(Body::empty()).unwrap()).await
}

pub async fn push(app: &Router, blob: &[u8]) {
    let link = format!("/api/links/{ID}");
    assert_eq!(call(app, "PUT", &link, Some(AUTH), blob).await.0, 200);
}

/// One batch item for the three bearer-less routes.
pub fn one(auth: &str) -> Vec<u8> {
    format!(r#"[{{"id":"{ID}","auth":"{auth}"}}]"#).into_bytes()
}

pub fn json(body: &str) -> serde_json::Value {
    serde_json::from_str(body).unwrap()
}

/// A static tree of its own per test, so one test's dotfile cannot be another's
/// missing file. The sibling outside it is what a traversal would reach.
pub fn static_dir(name: &str) -> PathBuf {
    let root = std::env::temp_dir().join(format!("modelis-e2-{}-{name}", std::process::id()));
    let dir = root.join("dist");
    std::fs::create_dir_all(dir.join("assets")).unwrap();
    std::fs::write(dir.join("taskpane.html"), "<title>pls,fix</title>").unwrap();
    std::fs::write(dir.join("pptpane.html"), "<title>pls,fix deck</title>").unwrap();
    std::fs::write(dir.join(".env"), "MODELIS_SECRET=1").unwrap();
    std::fs::write(root.join("outside.txt"), "not the pane").unwrap();
    dir
}
