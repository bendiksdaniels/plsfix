//! Hunt pass 2 shared harness: random operation sequences over the model in
//! `store_model`, and applying one step to both the model and the real
//! `Store`, asserting they agree - the return value AND the store's own
//! aggregate counters (byte total, per-id revision count, per-workspace live
//! inbox count) - after every single step.
//! Invariant: a mismatch here names the exact op it happened after.

use plsfix_server::store::{Delete, Get, Put, Store};
use plsfix_server::store_inbox::{InboxRow, Posted};
use proptest::prelude::*;

use super::store_model::{
    auths, id_str, ws_str, Model, ModelDelete, ModelGet, ModelPosted, ModelPut, N_AUTHS, N_IDS,
    N_WS,
};

#[derive(Debug, Clone)]
pub enum Op {
    PutLink(usize, usize, Vec<u8>),
    GetLink(usize, usize),
    GetLinkRev(usize, usize, i64),
    DeleteLink(usize, usize),
    Status(Vec<(usize, usize)>),
    TouchLinks(Vec<(usize, usize)>),
    PostInbox(usize, usize, usize, Vec<u8>),
    ListInbox(usize, usize),
    DeleteInbox(usize, usize, usize),
    Sweep,
    Advance(i64),
}

fn blob_strategy() -> impl Strategy<Value = Vec<u8>> {
    prop::collection::vec(any::<u8>(), 0..40)
}

fn pairs_strategy() -> impl Strategy<Value = Vec<(usize, usize)>> {
    prop::collection::vec((0..N_IDS, 0..N_AUTHS), 0..4)
}

/// Mostly small forward steps, with the exact TTL boundaries of both tables
/// folded in so expiry (and the byte cap it frees up) is exercised, not just
/// theorized about.
fn advance_strategy() -> impl Strategy<Value = i64> {
    use plsfix_server::store::{INBOX_TTL, LINK_TTL};
    prop_oneof![
        4 => 0i64..6,
        1 => Just(LINK_TTL - 1),
        1 => Just(LINK_TTL),
        1 => Just(LINK_TTL + 1),
        1 => Just(INBOX_TTL - 1),
        1 => Just(INBOX_TTL),
        1 => Just(INBOX_TTL + 1),
    ]
}

pub fn op_strategy() -> impl Strategy<Value = Op> {
    prop_oneof![
        3 => (0..N_IDS, 0..N_AUTHS, blob_strategy())
            .prop_map(|(id, auth, blob)| Op::PutLink(id, auth, blob)),
        2 => (0..N_IDS, 0..N_AUTHS).prop_map(|(id, auth)| Op::GetLink(id, auth)),
        2 => (0..N_IDS, 0..N_AUTHS, -5i64..30)
            .prop_map(|(id, auth, rev)| Op::GetLinkRev(id, auth, rev)),
        1 => (0..N_IDS, 0..N_AUTHS).prop_map(|(id, auth)| Op::DeleteLink(id, auth)),
        1 => pairs_strategy().prop_map(Op::Status),
        1 => pairs_strategy().prop_map(Op::TouchLinks),
        8 => (0..N_WS, 0..N_AUTHS, 0..N_IDS, blob_strategy())
            .prop_map(|(ws, auth, id, blob)| Op::PostInbox(ws, auth, id, blob)),
        2 => (0..N_WS, 0..N_AUTHS).prop_map(|(ws, auth)| Op::ListInbox(ws, auth)),
        1 => (0..N_WS, 0..N_AUTHS, 0..N_IDS).prop_map(|(ws, auth, id)| Op::DeleteInbox(ws, auth, id)),
        1 => Just(Op::Sweep),
        2 => advance_strategy().prop_map(Op::Advance),
    ]
}

pub fn ops_strategy() -> impl Strategy<Value = Vec<Op>> {
    prop::collection::vec(op_strategy(), 1..30)
}

pub fn expect_put(expect: &ModelPut, real: &Put, op: &Op) {
    let matches = match (expect, real) {
        (ModelPut::Created(e), Put::Created(r)) => e == r,
        (ModelPut::Updated(e), Put::Updated(r)) => e == r,
        (ModelPut::Forbidden, Put::Forbidden) => true,
        (ModelPut::Full, Put::Full) => true,
        (ModelPut::Exhausted, Put::Exhausted) => true,
        _ => false,
    };
    assert!(
        matches,
        "put mismatch after {op:?}: model {expect:?}, store {real:?}"
    );
}

fn expect_get(expect: &ModelGet, real: &Get, op: &Op) {
    let matches = match (expect, real) {
        (ModelGet::Found(rev, pushed_at, blob), Get::Found(found)) => {
            *rev == found.rev && *pushed_at == found.pushed_at && blob == &found.blob
        }
        (ModelGet::Forbidden, Get::Forbidden) => true,
        (ModelGet::Missing, Get::Missing) => true,
        _ => false,
    };
    assert!(
        matches,
        "get mismatch after {op:?}: model {expect:?}, store {real:?}"
    );
}

fn expect_delete(expect: &ModelDelete, real: &Delete, op: &Op) {
    let matches = matches!(
        (expect, real),
        (ModelDelete::Deleted, Delete::Deleted)
            | (ModelDelete::Forbidden, Delete::Forbidden)
            | (ModelDelete::Missing, Delete::Missing)
    );
    assert!(
        matches,
        "delete mismatch after {op:?}: model {expect:?}, store {real:?}"
    );
}

fn expect_posted(expect: &ModelPosted, real: &Posted, op: &Op) {
    let matches = matches!(
        (expect, real),
        (ModelPosted::Stored, Posted::Stored) | (ModelPosted::Full, Posted::Full)
    );
    assert!(
        matches,
        "post_inbox mismatch after {op:?}: model {expect:?}, store {real:?}"
    );
}

fn expect_list(expect: &[(usize, i64, Vec<u8>)], real: &[InboxRow], op: &Op) {
    let real_ids: Vec<String> = real.iter().map(|row| row.id.clone()).collect();
    let expect_ids: Vec<String> = expect.iter().map(|(id, _, _)| id_str(*id)).collect();
    assert_eq!(expect_ids, real_ids, "list_inbox order/ids after {op:?}");
    for ((_, created_at, blob), row) in expect.iter().zip(real.iter()) {
        assert_eq!(*created_at, row.created_at, "created_at after {op:?}");
        assert_eq!(blob, &row.blob, "blob after {op:?}");
    }
}

/// Applies one op to both sides, asserts they agree, then re-checks the
/// aggregate counters `/version` and the store's own test helpers expose -
/// every id's revision count and every workspace's live inbox count - so a
/// divergence that a single call's return value would not show (a phantom
/// row, a byte miscount) still surfaces on the very next step.
pub fn apply(model: &mut Model, store: &Store, op: &Op) {
    match op {
        Op::PutLink(id, auth, blob) => {
            let expect = model.put_link(*id, *auth, blob);
            let real = store
                .put_link(&id_str(*id), &auths()[*auth], blob, model.now)
                .unwrap();
            expect_put(&expect, &real, op);
        }
        Op::GetLink(id, auth) => {
            let expect = model.get_link(*id, *auth);
            let real = store
                .get_link(&id_str(*id), &auths()[*auth], model.now)
                .unwrap();
            expect_get(&expect, &real, op);
        }
        Op::GetLinkRev(id, auth, rev) => {
            let expect = model.get_link_rev(*id, *auth, *rev);
            let real = store
                .get_link_rev(&id_str(*id), &auths()[*auth], *rev, model.now)
                .unwrap();
            expect_get(&expect, &real, op);
        }
        Op::DeleteLink(id, auth) => {
            let expect = model.delete_link(*id, *auth);
            let real = store
                .delete_link(&id_str(*id), &auths()[*auth], model.now)
                .unwrap();
            expect_delete(&expect, &real, op);
        }
        Op::Status(items) => {
            let expect = model.status(items);
            let queries: Vec<(String, [u8; 32])> = items
                .iter()
                .map(|&(id, auth)| (id_str(id), auths()[auth]))
                .collect();
            let real = store.status(&queries, model.now).unwrap();
            assert_eq!(expect.len(), real.len());
            for ((rev, pushed_at, auth_error), row) in expect.iter().zip(real.iter()) {
                assert_eq!(*rev, row.rev, "status rev after {op:?}");
                assert_eq!(*pushed_at, row.pushed_at, "status pushed_at after {op:?}");
                assert_eq!(
                    *auth_error, row.auth_error,
                    "status auth_error after {op:?}"
                );
            }
        }
        Op::TouchLinks(items) => {
            let expect = model.touch_links(items);
            let queries: Vec<(String, [u8; 32])> = items
                .iter()
                .map(|&(id, auth)| (id_str(id), auths()[auth]))
                .collect();
            let real = store.touch_links(&queries, model.now).unwrap();
            assert_eq!(expect, real, "touch_links count after {op:?}");
        }
        Op::PostInbox(ws, auth, id, blob) => {
            let expect = model.post_inbox(*ws, *auth, *id, blob);
            let real = store
                .post_inbox(&ws_str(*ws), &auths()[*auth], &id_str(*id), blob, model.now)
                .unwrap();
            expect_posted(&expect, &real, op);
        }
        Op::ListInbox(ws, auth) => {
            let expect = model.list_inbox(*ws, *auth);
            let real = store
                .list_inbox(&ws_str(*ws), &auths()[*auth], model.now)
                .unwrap();
            expect_list(&expect, &real, op);
        }
        Op::DeleteInbox(ws, auth, id) => {
            let expect = model.delete_inbox(*ws, *auth, *id);
            let real = store
                .delete_inbox(&ws_str(*ws), &auths()[*auth], &id_str(*id), model.now)
                .unwrap();
            assert_eq!(expect, real, "delete_inbox after {op:?}");
        }
        Op::Sweep => {
            let expect = model.sweep();
            let real = store.sweep(model.now).unwrap();
            assert_eq!(expect, real, "sweep count after {op:?}");
        }
        Op::Advance(delta) => model.now += delta,
    }

    assert_eq!(
        store.total_bytes().unwrap(),
        model.total_bytes(),
        "byte total diverged after {op:?}"
    );
    for id in 0..N_IDS {
        let expect = model.links.get(&id).map_or(0, |link| link.revs.len());
        assert_eq!(
            store.rev_count(&id_str(id)),
            expect,
            "rev_count[{id}] diverged after {op:?}"
        );
    }
    for ws in 0..N_WS {
        assert_eq!(
            store.inbox_count(&ws_str(ws), model.now).unwrap(),
            model.live_inbox_count(ws),
            "inbox_count[{ws}] diverged after {op:?}"
        );
    }
}
