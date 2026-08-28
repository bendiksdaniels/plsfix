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
        assert!(matches!(
            store.delete_inbox("WS", &hash(1), ID, 10).unwrap(),
            Delete::Deleted
        ));
    }
}

#[test]
fn inbox_post_with_a_foreign_auth_is_refused() {
    let store = Store::in_memory().unwrap();
    let id = "0123456789abcdef0123456789abcdef";
    assert!(store.post_inbox("WS", &[1u8; 32], id, b"x", 0).unwrap());
    assert!(!store.post_inbox("WS", &[2u8; 32], id, b"y", 1).unwrap());
    let rows = store.list_inbox("WS", &[1u8; 32], 2).unwrap();
    assert_eq!(rows[0].blob, b"x");
    assert!(store.list_inbox("WS", &[2u8; 32], 2).unwrap().is_empty());
}
