//! Bucket maths, the client key and the read/write split, driven with an
//! injected clock so nothing here sleeps: the allowance, the refill, the whole
//! seconds a refusal asks for, and which header names the client.

use std::net::SocketAddr;

use axum::{
    extract::ConnectInfo,
    http::{Extensions, HeaderMap, Method},
};
use plsfix_server::limits::{client_key, is_write, Allowed, RateLimiter};

fn headers(pairs: &[(&'static str, &str)]) -> HeaderMap {
    let mut headers = HeaderMap::new();
    for (name, value) in pairs {
        headers.insert(
            axum::http::HeaderName::from_static(name),
            value.parse().unwrap(),
        );
    }
    headers
}

#[test]
fn a_full_bucket_spends_a_minutes_allowance_then_refuses() {
    let limiter = RateLimiter::new(3);
    for _ in 0..3 {
        assert_eq!(limiter.take("client", 100), Allowed::Yes);
    }
    // 3 a minute is one every 20 seconds, and a refusal always waits whole
    // seconds - never zero, which would spin against the same empty bucket.
    assert_eq!(limiter.take("client", 100), Allowed::No { retry_after: 20 });
    // Another client has its own bucket and is untouched by the first.
    assert_eq!(limiter.take("other", 100), Allowed::Yes);
}

#[test]
fn tokens_refill_with_the_clock_and_never_past_capacity() {
    let limiter = RateLimiter::new(60);
    for _ in 0..60 {
        assert_eq!(limiter.take("client", 0), Allowed::Yes);
    }
    assert_eq!(limiter.take("client", 0), Allowed::No { retry_after: 1 });
    // 60 a minute is one a second.
    assert_eq!(limiter.take("client", 1), Allowed::Yes);
    assert_eq!(limiter.take("client", 1), Allowed::No { retry_after: 1 });
    // An hour later the bucket is full, not an hour's worth over.
    for _ in 0..60 {
        assert_eq!(limiter.take("client", 3601), Allowed::Yes);
    }
    assert_eq!(limiter.take("client", 3601), Allowed::No { retry_after: 1 });
}

#[test]
fn a_clock_that_steps_back_refills_nothing() {
    let limiter = RateLimiter::new(1);
    assert_eq!(limiter.take("client", 500), Allowed::Yes);
    assert_eq!(limiter.take("client", 400), Allowed::No { retry_after: 60 });
}

#[test]
fn idle_buckets_are_pruned_and_come_back_full() {
    let limiter = RateLimiter::new(2);
    limiter.take("stale", 0);
    limiter.take("busy", 3500);
    assert_eq!(limiter.len(), 2);
    assert_eq!(limiter.prune(3600, 3600), 1);
    assert_eq!(limiter.len(), 1);
    // The dropped bucket is simply a new, full one.
    assert_eq!(limiter.take("stale", 3600), Allowed::Yes);
    assert!(!limiter.is_empty());
}

#[test]
fn the_client_key_prefers_cloudflare_then_the_first_forwarded_hop() {
    let peer = {
        let mut extensions = Extensions::new();
        extensions.insert(ConnectInfo::<SocketAddr>("10.0.0.9:443".parse().unwrap()));
        extensions
    };
    assert_eq!(
        client_key(
            &headers(&[
                ("cf-connecting-ip", "1.2.3.4"),
                ("x-forwarded-for", "9.9.9.9, 8.8.8.8"),
            ]),
            &peer,
        ),
        "1.2.3.4"
    );
    assert_eq!(
        client_key(
            &headers(&[("x-forwarded-for", " 9.9.9.9 , 8.8.8.8")]),
            &peer
        ),
        "9.9.9.9"
    );
    // An empty header names nobody, so the next source answers.
    assert_eq!(
        client_key(&headers(&[("cf-connecting-ip", "")]), &peer),
        "10.0.0.9"
    );
    // No header and no ConnectInfo (the route tests): one shared bucket.
    assert_eq!(client_key(&HeaderMap::new(), &Extensions::new()), "unknown");
}

#[test]
fn only_the_three_batch_routes_are_reads_among_the_posts() {
    assert!(is_write(&Method::PUT, "/api/links/abc"));
    assert!(is_write(&Method::DELETE, "/api/inbox/ws/abc"));
    assert!(is_write(&Method::POST, "/api/inbox/ws"));
    assert!(!is_write(&Method::POST, "/api/links/status"));
    assert!(!is_write(&Method::POST, "/api/links/fetch"));
    assert!(!is_write(&Method::POST, "/api/links/touch"));
    assert!(!is_write(&Method::GET, "/api/links/abc"));
}
