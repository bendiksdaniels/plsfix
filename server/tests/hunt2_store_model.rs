//! Hunt pass 2 (model-based): a tiny reference model of the store, run
//! alongside the real `Store` over a random operation sequence and compared
//! after every step. The model lives in `common/store_model.rs`, the harness
//! in `common/store_check.rs`; this file holds the tests themselves.

mod common;

use common::store_check::{apply, expect_put, ops_strategy, Op};
use common::store_model::{auths, id_str, Model, ModelLink, ModelPut};
use plsfix_server::store::{auth_hash, Store, LINK_TTL};
use plsfix_server::store_inbox::Posted;
use plsfix_server::store_room::Caps;
use proptest::prelude::*;
use proptest::test_runner::{FileFailurePersistence, RngSeed};

/// Reproduces the mechanism first, in isolation: a workspace already at its
/// row cap must still accept a re-post of a link it already holds, because an
/// upsert of an existing (ws, id, auth) row does not grow the row count. The
/// doc comment on `post_inbox` promises exactly this ("a re-export from the
/// same key replaces its own row"); the cap check did not special-case it.
#[test]
fn a_repost_of_an_existing_inbox_item_is_not_blocked_by_a_full_workspace() {
    let mut store = Store::in_memory().unwrap();
    store.set_caps(Caps {
        max_bytes: i64::MAX,
        inbox_rows: 1,
    });
    let auth = auth_hash("key");
    assert!(matches!(
        store.post_inbox("WS", &auth, "id-a", b"first", 0).unwrap(),
        Posted::Stored
    ));
    // The workspace is now at its one-row cap with that single link. Posting
    // the SAME link again is a replace, not a new row, and must succeed.
    let second = store.post_inbox("WS", &auth, "id-a", b"second", 1).unwrap();
    assert!(
        matches!(second, Posted::Stored),
        "a re-export of the only row a full workspace holds must replace it, got {second:?}"
    );
    let rows = store.list_inbox("WS", &auth, 1).unwrap();
    assert_eq!(rows.len(), 1, "still exactly one row, not two");
    assert_eq!(rows[0].blob, b"second", "the row was replaced, not skipped");
}

fn config() -> ProptestConfig {
    ProptestConfig {
        cases: 200,
        rng_seed: RngSeed::Fixed(0xC0FFEE),
        failure_persistence: Some(Box::new(FileFailurePersistence::Off)),
        ..ProptestConfig::default()
    }
}

proptest! {
    #![proptest_config(config())]

    /// The property: for any sequence of the store's own operations, against
    /// a store far smaller than its ceiling, the real `Store` never answers
    /// anything the reference model would not - same outcome variant, same
    /// revision, same bytes back, same counts - at every single step, not
    /// only at the end.
    #[test]
    fn the_real_store_never_disagrees_with_the_reference_model(ops in ops_strategy()) {
        let caps = Caps { max_bytes: 300, inbox_rows: 3 };
        let mut model = Model::new(caps);
        let mut store = Store::in_memory().unwrap();
        store.set_caps(caps);
        for op in &ops {
            apply(&mut model, &store, op);
        }
    }
}

/// Pins the exact sequence the property above's shrinker found at seed 1 with
/// 4000 cases before the fix in `store_model.rs`'s `Model::sweep`: two live
/// revisions of one link aged past `LINK_TTL` in a single step, swept as "1"
/// dead link by the model but "2" dead rows by `sweep_locked` (`store.rs`
/// deletes from `links`, keyed `(id, rev)`) - "sweep count 1 vs 2", a false
/// alarm on the model, never on the product it was checking.
#[test]
fn sweep_counts_both_revisions_of_one_expired_link_not_the_one_link() {
    let caps = Caps {
        max_bytes: 300,
        inbox_rows: 3,
    };
    let mut model = Model::new(caps);
    let mut store = Store::in_memory().unwrap();
    store.set_caps(caps);

    apply(&mut model, &store, &Op::PutLink(2, 1, Vec::new()));
    apply(&mut model, &store, &Op::PutLink(2, 1, Vec::new()));
    apply(&mut model, &store, &Op::Advance(LINK_TTL));
    apply(&mut model, &store, &Op::Sweep);
}

/// `i64::MAX` pushes are not reachable through 2^63 real calls, so this seeds
/// one row directly (the technique `hunt_revisions.rs` uses for the HTTP
/// layer) and asks the model to agree at the STORE level: `Put::Exhausted`
/// itself, one layer under the 507 the other file already pins.
#[test]
fn the_model_agrees_the_store_is_exhausted_at_i64_max() {
    let path = std::env::temp_dir().join(format!(
        "modelis-hunt2-store-model-exhausted-{}.sqlite",
        std::process::id()
    ));
    let _ = std::fs::remove_file(&path);
    let auth = auths()[0];
    {
        let store = Store::open(&path).unwrap(); // creates the schema
        drop(store);
    }
    {
        let conn = rusqlite::Connection::open(&path).unwrap();
        conn.execute(
            "INSERT INTO links (id, rev, auth_hash, pushed_at, expires_at, blob) VALUES (?1, ?2, ?3, ?4, ?5, ?6)",
            rusqlite::params![id_str(0), i64::MAX, auth.as_slice(), 0i64, LINK_TTL, b"seeded".as_slice()],
        )
        .unwrap();
    }
    let store = Store::open(&path).unwrap();

    let mut model = Model::new(Caps::OPEN);
    model.links.insert(
        0,
        ModelLink {
            auth: 0,
            revs: vec![(i64::MAX, 0, b"seeded".to_vec())],
            expires_at: LINK_TTL,
        },
    );

    let expect = model.put_link(0, 0, b"one more");
    let real = store.put_link(&id_str(0), &auth, b"one more", 0).unwrap();
    expect_put(&expect, &real, &Op::PutLink(0, 0, b"one more".to_vec()));
    assert_eq!(
        expect,
        ModelPut::Exhausted,
        "the model itself must reach Exhausted here"
    );

    let _ = std::fs::remove_file(&path);
    let _ = std::fs::remove_file(format!("{}-wal", path.display()));
    let _ = std::fs::remove_file(format!("{}-shm", path.display()));
}
