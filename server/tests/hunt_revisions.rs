//! Hunt pass 1 (edges + repeats): revision-number edges near the limits of
//! `i64` and the pane's local-rev space, concurrent writers to one link, the
//! link DELETE route (untested at the HTTP layer before this file), and the
//! exact second at which the TTL sweep is supposed to take a row.

mod common;

use std::sync::Arc;

use axum::http::StatusCode;
use common::*;
use plsfix_server::{
    relay::*,
    store::{auth_hash, Get, Store, INBOX_TTL, LINK_TTL},
};

/// `head.rev + 1` in `Store::put_link` is unchecked. A link cannot reach
/// `i64::MAX` through the API (that is 2^63 pushes), so the only way to prove
/// what happens at the edge is to seed a row directly, the way a restored
/// backup or a hand-edited row could.
fn seed_row_at_rev(path: &std::path::Path, id: &str, auth: &[u8; 32], rev: i64, now: i64) {
    let conn = rusqlite::Connection::open(path).unwrap();
    conn.execute(
        "INSERT INTO links (id, rev, auth_hash, pushed_at, expires_at, blob) VALUES (?1, ?2, ?3, ?4, ?5, ?6)",
        rusqlite::params![id, rev, auth.as_slice(), now, now + LINK_TTL, b"seeded".as_slice()],
    )
    .unwrap();
}

fn temp_db(name: &str) -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!(
        "modelis-hunt-revisions-{}-{name}.sqlite",
        std::process::id()
    ));
    let _ = std::fs::remove_file(&path);
    path
}

/// A link already sitting at the highest revision `i64` can hold must not
/// crash the request that tries to push past it: it has to come back as an
/// ordinary refusal, the same shape a full store already answers with.
#[tokio::test]
async fn a_push_past_i64_max_revisions_is_refused_not_a_panic() {
    let path = temp_db("i64-max");
    let auth = auth_hash(AUTH);
    {
        let store = Store::open(&path).unwrap();
        drop(store); // just to create the schema
    }
    seed_row_at_rev(&path, ID, &auth, i64::MAX, now() - 10);

    let store = Store::open(&path).unwrap();
    let state = Arc::new(AppState::new(store));
    let app = routes(state);

    let (status, body) = call(&app, "PUT", &format!("/api/links/{ID}"), Some(AUTH), b"new").await;
    // The same refusal shape a full store already answers with - never the
    // generic store error a caught panic would have left behind.
    assert_eq!(status, StatusCode::INSUFFICIENT_STORAGE, "{body}");
    assert_eq!(body, r#"{"error":"storage full"}"#);

    let _ = std::fs::remove_file(&path);
    let _ = std::fs::remove_file(format!("{}-wal", path.display()));
    let _ = std::fs::remove_file(format!("{}-shm", path.display()));
}

/// The GET revert path takes `?rev=` straight off the query string as an
/// `i64`: a value at the pane's local-rev base (2^40) or right below
/// `i64::MAX` must parse and simply find no such row, never error out.
#[tokio::test]
async fn a_revision_query_at_2_40_or_near_i64_max_is_missing_not_an_error() {
    let app = relay_app();
    push(&app, b"one").await;

    for rev in ["1099511627776", "9223372036854775807"] {
        let path = format!("/api/links/{ID}?rev={rev}");
        let (status, _) = call(&app, "GET", &path, Some(AUTH), b"").await;
        assert_eq!(status, StatusCode::NOT_FOUND, "rev={rev}");
    }
    // One past what an i64 can hold is a bad request, not a silent wrap.
    let (status, body) = call(
        &app,
        "GET",
        &format!("/api/links/{ID}?rev=9223372036854775808"),
        Some(AUTH),
        b"",
    )
    .await;
    assert_eq!(status, StatusCode::BAD_REQUEST);
    assert_eq!(body, r#"{"error":"bad rev"}"#);
}

/// Two pushes to the same link from the same key, released at once: the
/// store's one connection mutex spans the whole read-then-write, so they
/// cannot both read rev 1 as the head and both try to write rev 2. Both must
/// land, in some order, as sequential revisions - never a lost update, never
/// a duplicate revision number.
#[tokio::test(flavor = "multi_thread", worker_threads = 4)]
async fn two_concurrent_pushes_to_the_same_link_land_as_sequential_revisions() {
    let state = Arc::new(AppState::new(Store::in_memory().unwrap()));
    push(&routes(state.clone()), b"first").await;

    let app = routes(state.clone());
    let gate = Arc::new(tokio::sync::Barrier::new(2));
    let mut racing = Vec::new();
    for blob in [b"second".to_vec(), b"third".to_vec()] {
        let (app, gate) = (app.clone(), Arc::clone(&gate));
        racing.push(tokio::spawn(async move {
            gate.wait().await;
            call(&app, "PUT", &format!("/api/links/{ID}"), Some(AUTH), &blob).await
        }));
    }
    let mut revs = Vec::new();
    for task in racing {
        let (status, body) = task.await.unwrap();
        assert_eq!(status, StatusCode::OK, "{body}");
        revs.push(json(&body)["rev"].as_i64().unwrap());
    }
    revs.sort_unstable();
    assert_eq!(
        revs,
        vec![2, 3],
        "two racing writers must not both become rev 2"
    );
    // Retention still holds to exactly two revisions after the race.
    assert_eq!(state.store.rev_count(ID), 2);
}

/// The DELETE route for a link itself (not an inbox row) had no test at the
/// HTTP layer at all: only auth-code checks and the store-level Forbidden
/// case existed. This pins the happy path - every revision goes, and a
/// repeat delete is a 404 like any other missing link, never a 200 that
/// quietly does nothing and never a 500.
#[tokio::test]
async fn deleting_a_link_removes_every_revision_and_a_repeat_delete_is_missing() {
    let state = Arc::new(AppState::new(Store::in_memory().unwrap()));
    let app = routes(state.clone());
    push(&app, b"one").await;
    push(&app, b"two").await;
    assert_eq!(state.store.rev_count(ID), 2);

    let path = format!("/api/links/{ID}");
    let (status, body) = call(&app, "DELETE", &path, Some(AUTH), b"").await;
    assert_eq!(status, StatusCode::OK);
    assert_eq!(body, r#"{"ok":true}"#);
    assert_eq!(state.store.rev_count(ID), 0, "every revision must be gone");

    let (status, _) = call(&app, "GET", &path, Some(AUTH), b"").await;
    assert_eq!(status, StatusCode::NOT_FOUND);

    // Repeat: the same delete a second time is "not found", never "ok" again
    // and never the generic store error.
    let (status, body) = call(&app, "DELETE", &path, Some(AUTH), b"").await;
    assert_eq!(status, StatusCode::NOT_FOUND);
    assert_eq!(body, r#"{"error":"not found"}"#);
}

/// `expires_at > now` (read) and `expires_at <= now` (sweep) must agree at the
/// exact second a link's TTL runs out; the existing tests only prove one
/// second either side of it.
#[test]
fn a_link_expires_at_the_exact_ttl_boundary_second() {
    let store = Store::in_memory().unwrap();
    let auth = auth_hash(AUTH);
    store.put_link(ID, &auth, b"a", 0).unwrap();
    // Exactly LINK_TTL seconds later: the boundary itself, not TTL-1 or TTL+1.
    assert!(matches!(
        store.get_link(ID, &auth, LINK_TTL).unwrap(),
        Get::Missing
    ));
    assert_eq!(store.sweep(LINK_TTL).unwrap(), 1);
}

/// The same boundary, one second earlier, is still live - the pair the exact
/// second sits between.
#[test]
fn a_link_is_still_live_one_second_before_the_ttl_boundary() {
    let store = Store::in_memory().unwrap();
    let auth = auth_hash(AUTH);
    store.put_link(ID, &auth, b"a", 0).unwrap();
    assert!(matches!(
        store.get_link(ID, &auth, LINK_TTL - 1).unwrap(),
        Get::Found(_)
    ));
    assert_eq!(store.sweep(LINK_TTL - 1).unwrap(), 0);
}

/// The inbox's own TTL has the same `>` / `<=` pairing at INBOX_TTL exactly.
#[test]
fn an_inbox_item_expires_at_the_exact_ttl_boundary_second() {
    let store = Store::in_memory().unwrap();
    let auth = auth_hash(AUTH);
    store.post_inbox("WS", &auth, ID, b"x", 0).unwrap();
    assert_eq!(store.list_inbox("WS", &auth, INBOX_TTL).unwrap().len(), 0);
    assert_eq!(
        store.list_inbox("WS", &auth, INBOX_TTL - 1).unwrap().len(),
        1
    );
}
