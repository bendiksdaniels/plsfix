//! Hunt pass 1 (edges + repeats): static-host paths audit_e2.rs does not try
//! (%252e, a backslash, a doubled slash around a dotfile or `..`). The path is
//! Access-bypassed, so it never serves a dotfile or anything above `dist/`.
//! Each test has its own `static_dir`: tests on separate threads would race.

mod common;

use std::sync::Arc;

use axum::Router;
use common::*;
use plsfix_server::{app, relay::AppState, store::Store};

const SECRET: &str = "MODELIS_SECRET";
const OUTSIDE: &str = "not the pane";

fn app_over(name: &str) -> Router {
    app(
        static_dir(name),
        Arc::new(AppState::new(Store::in_memory().unwrap())),
    )
}

/// Every one of these must answer without the two markers a real escape would
/// carry, whatever status code it happens to pick.
async fn refuses(app: &Router, path: &str) {
    let (status, body) = get(app, path).await;
    assert!(
        !body.contains(SECRET),
        "{path} served a dotfile ({status}): {body}"
    );
    assert!(
        !body.contains(OUTSIDE),
        "{path} escaped dist/ ({status}): {body}"
    );
}

#[tokio::test]
async fn double_percent_encoded_dot_segments_are_never_served() {
    let app = app_over("hunt-static-double-pct");
    for path in [
        "/%252e%252e/outside.txt",
        "/%252e%252e%2foutside.txt",
        "/assets/%252e%252e/.env",
        "/%252eenv",
    ] {
        refuses(&app, path).await;
    }
}

#[tokio::test]
async fn backslash_forms_are_never_served() {
    let app = app_over("hunt-static-backslash");
    for path in [
        "/..\\..\\outside.txt",
        "/assets\\..\\.env",
        "/%5c.env",
        "/assets/%5c..%5c.env",
    ] {
        refuses(&app, path).await;
    }
}

#[tokio::test]
async fn a_doubled_slash_does_not_hide_a_dotfile_or_reach_outside_dist() {
    let app = app_over("hunt-static-double-slash");
    for path in [
        "//.env",
        "/assets//.env",
        "//outside.txt",
        "/assets//../../outside.txt",
    ] {
        refuses(&app, path).await;
    }
}

/// The dotfile guard matches by segment (`starts_with('.')`), so a doubled
/// slash around one still leaves it caught - this pins the status, not just
/// the content, for the cases the guard is actually built to catch.
#[tokio::test]
async fn a_dotfile_behind_a_doubled_slash_is_still_a_404() {
    let app = app_over("hunt-static-double-slash-404");
    for path in ["//.env", "/assets//.env"] {
        let (status, _) = get(&app, path).await;
        assert_eq!(
            status,
            axum::http::StatusCode::NOT_FOUND,
            "{path} answered {status}"
        );
    }
    // The pane itself is unaffected by any of the above.
    assert_eq!(
        get(&app, "/taskpane.html").await.0,
        axum::http::StatusCode::OK
    );
}
