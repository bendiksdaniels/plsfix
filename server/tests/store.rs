//! Store tests: rev retention, auth-hash ownership, TTL expiry and the inbox.
//! Everything runs on an in-memory sqlite so the suite touches no data dir.

use plsfix_server::store::*;

mod tests {
    use super::*;

    const ID: &str = "0123456789abcdef0123456789abcdef";
    fn hash(n: u8) -> [u8; 32] {
        [n; 32]
    }

    #[test]
    fn put_creates_then_updates_and_keeps_two_revs() {
        let store = Store::in_memory().unwrap();
        assert!(matches!(
            store.put_link(ID, &hash(1), b"a", 100).unwrap(),
            Put::Created(1)
        ));
        assert!(matches!(
            store.put_link(ID, &hash(1), b"b", 101).unwrap(),
            Put::Updated(2)
        ));
        assert!(matches!(
            store.put_link(ID, &hash(1), b"c", 102).unwrap(),
            Put::Updated(3)
        ));
        let found = match store.get_link(ID, &hash(1), 103).unwrap() {
            Get::Found(f) => f,
            _ => panic!(),
        };
        assert_eq!((found.rev, found.blob.as_slice()), (3, b"c".as_slice()));
        assert_eq!(store.rev_count(ID), 2);
    }

    #[test]
    fn wrong_auth_is_forbidden_everywhere() {
        let store = Store::in_memory().unwrap();
        store.put_link(ID, &hash(1), b"a", 100).unwrap();
        assert!(matches!(
            store.put_link(ID, &hash(2), b"b", 101).unwrap(),
            Put::Forbidden
        ));
        assert!(matches!(
            store.get_link(ID, &hash(2), 101).unwrap(),
            Get::Forbidden
        ));
        assert!(matches!(
            store.delete_link(ID, &hash(2), 101).unwrap(),
            Delete::Forbidden
        ));
        let rows = store.status(&[(ID.to_string(), hash(2))], 101).unwrap();
        assert!(rows[0].auth_error && rows[0].rev.is_none());
    }

    #[test]
    fn links_expire_seven_days_after_the_last_put() {
        let store = Store::in_memory().unwrap();
        store.put_link(ID, &hash(1), b"a", 0).unwrap();
        assert!(matches!(
            store.get_link(ID, &hash(1), LINK_TTL - 1).unwrap(),
            Get::Found(_)
        ));
        assert!(matches!(
            store.get_link(ID, &hash(1), LINK_TTL + 1).unwrap(),
            Get::Missing
        ));
        assert_eq!(store.sweep(LINK_TTL + 1).unwrap(), 1);
    }

    #[test]
    fn a_later_put_refreshes_the_ttl_of_every_kept_rev() {
        // "7 d from the *last* put" is one UPDATE in put_link; without it rev 1
        // still dies on the first push's clock and the sweep below finds it.
        let store = Store::in_memory().unwrap();
        store.put_link(ID, &hash(1), b"a", 0).unwrap();
        store.put_link(ID, &hash(1), b"b", LINK_TTL - 10).unwrap();
        assert_eq!(store.sweep(LINK_TTL + 10).unwrap(), 0);
        assert_eq!(store.rev_count(ID), 2);
        assert!(matches!(
            store.get_link(ID, &hash(1), LINK_TTL + 10).unwrap(),
            Get::Found(found) if found.rev == 2
        ));
    }

    #[test]
    fn a_put_after_expiry_starts_again_at_rev_one() {
        // The other half of the pane's "any rev change is an update" rule: a
        // swept link has no head, so its next push is Created(1), below the rev
        // the deck's tag still holds.
        let store = Store::in_memory().unwrap();
        store.put_link(ID, &hash(1), b"a", 0).unwrap();
        store.put_link(ID, &hash(1), b"b", 10).unwrap();
        store.sweep(LINK_TTL + 20).unwrap();
        assert!(matches!(
            store.put_link(ID, &hash(1), b"c", LINK_TTL + 20).unwrap(),
            Put::Created(1)
        ));
    }

    #[test]
    fn inbox_lists_only_matching_auth_and_expires_after_a_day() {
        let store = Store::in_memory().unwrap();
        store.post_inbox("WS", &hash(1), ID, b"x", 0).unwrap();
        assert_eq!(store.list_inbox("WS", &hash(1), 10).unwrap().len(), 1);
        assert_eq!(store.list_inbox("WS", &hash(2), 10).unwrap().len(), 0);
        assert_eq!(
            store
                .list_inbox("WS", &hash(1), INBOX_TTL + 1)
                .unwrap()
                .len(),
            0
        );
        assert!(store.delete_inbox("WS", &hash(1), ID, 10).unwrap());
    }
}

#[test]
fn a_foreign_auth_gets_its_own_inbox_row_and_touches_nobody_elses() {
    // ws and the link id both travel in URL paths, so a log reader can POST to
    // a workspace with a bearer of its own. Keying the row by the writer's hash
    // makes that write harmless: it neither replaces nor blocks the pane's row,
    // and it cannot be listed or deleted by anyone else.
    let store = Store::in_memory().unwrap();
    let id = "0123456789abcdef0123456789abcdef";
    store.post_inbox("WS", &[1u8; 32], id, b"x", 0).unwrap();
    store.post_inbox("WS", &[2u8; 32], id, b"y", 1).unwrap();

    let mine = store.list_inbox("WS", &[1u8; 32], 2).unwrap();
    assert_eq!(mine.len(), 1);
    assert_eq!(mine[0].blob, b"x");
    let theirs = store.list_inbox("WS", &[2u8; 32], 2).unwrap();
    assert_eq!(theirs.len(), 1);
    assert_eq!(theirs[0].blob, b"y");

    // A key that wrote nothing here deletes nothing.
    assert!(!store.delete_inbox("WS", &[3u8; 32], id, 2).unwrap());
    // The squatter's own delete takes only the squatter's row.
    assert!(store.delete_inbox("WS", &[2u8; 32], id, 2).unwrap());
    assert_eq!(store.list_inbox("WS", &[1u8; 32], 2).unwrap()[0].blob, b"x");
    // The owner re-exports and replaces only its own row.
    store.post_inbox("WS", &[1u8; 32], id, b"z", 3).unwrap();
    assert_eq!(store.list_inbox("WS", &[1u8; 32], 4).unwrap()[0].blob, b"z");
}
