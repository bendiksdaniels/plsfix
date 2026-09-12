//! Bucket maths, the client key and the read/write split, driven with an
//! injected clock so nothing here sleeps: the allowance, the refill, the whole
//! seconds a refusal asks for, what a body costs in tokens, which header the
//! deployment lets name the client, and the cap on the bucket map itself.

use std::net::SocketAddr;

use axum::{
    extract::ConnectInfo,
    http::{Extensions, HeaderMap, Method},
};
use plsfix_server::limits::{
    client_key, is_write, token_cost, Allowed, RateLimiter, TrustedProxy, RATE_TOKEN_BYTES,
};

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

fn peer() -> Extensions {
    let mut extensions = Extensions::new();
    extensions.insert(ConnectInfo::<SocketAddr>("10.0.0.9:443".parse().unwrap()));
    extensions
}

#[test]
fn nothing_is_trusted_by_default_so_the_peer_names_the_client() {
    let forged = headers(&[
        ("cf-connecting-ip", "1.2.3.4"),
        ("x-forwarded-for", "9.9.9.9, 8.8.8.8"),
    ]);
    // A header any client may write must never pick the bucket it is counted
    // in: with no proxy declared, the socket is the only thing not forgeable.
    assert_eq!(client_key(TrustedProxy::None, &forged, &peer()), "10.0.0.9");
    // No header and no ConnectInfo (the route tests): one shared bucket.
    assert_eq!(
        client_key(TrustedProxy::None, &HeaderMap::new(), &Extensions::new()),
        "unknown"
    );
}

#[test]
fn cloudflare_mode_reads_cf_connecting_ip_and_falls_back_to_the_peer() {
    assert_eq!(
        client_key(
            TrustedProxy::Cloudflare,
            &headers(&[
                ("cf-connecting-ip", "1.2.3.4"),
                ("x-forwarded-for", "9.9.9.9, 8.8.8.8"),
            ]),
            &peer(),
        ),
        "1.2.3.4"
    );
    // An empty header names nobody, and X-Forwarded-For is not this mode's
    // header, so the peer answers.
    assert_eq!(
        client_key(
            TrustedProxy::Cloudflare,
            &headers(&[("cf-connecting-ip", ""), ("x-forwarded-for", "9.9.9.9")]),
            &peer(),
        ),
        "10.0.0.9"
    );
}

#[test]
fn xff_mode_reads_the_last_hop_never_the_first() {
    // The last hop is the one the adjacent proxy wrote; every hop before it
    // came from the client, so the first is exactly the forgeable one.
    assert_eq!(
        client_key(
            TrustedProxy::ForwardedFor,
            &headers(&[("x-forwarded-for", " 9.9.9.9 , 8.8.8.8 ")]),
            &peer(),
        ),
        "8.8.8.8"
    );
    // Cloudflare's header is not trusted in this mode.
    assert_eq!(
        client_key(
            TrustedProxy::ForwardedFor,
            &headers(&[("cf-connecting-ip", "1.2.3.4")]),
            &peer(),
        ),
        "10.0.0.9"
    );
}

#[test]
fn a_trust_mode_is_read_from_its_name_and_anything_else_trusts_nothing() {
    assert_eq!("cloudflare".parse(), Ok(TrustedProxy::Cloudflare));
    assert_eq!("xff".parse(), Ok(TrustedProxy::ForwardedFor));
    assert_eq!("none".parse(), Ok(TrustedProxy::None));
    assert_eq!(TrustedProxy::default(), TrustedProxy::None);
    assert!("xff-last".parse::<TrustedProxy>().is_err());
}

#[test]
fn a_forged_header_cannot_grow_one_map_row_without_bound() {
    let long = "9".repeat(500);
    let key = client_key(
        TrustedProxy::Cloudflare,
        &headers(&[("cf-connecting-ip", &long)]),
        &peer(),
    );
    assert_eq!(key.len(), 64);
}

#[test]
fn the_bucket_map_is_capped_and_drops_the_oldest_idle_client() {
    let limiter = RateLimiter::new(10).max_clients(3);
    for (index, key) in ["a", "b", "c"].iter().enumerate() {
        limiter.take(key, index as i64);
    }
    assert_eq!(limiter.len(), 3);
    // The fourth client evicts "a", the one nobody has used for longest.
    limiter.take("d", 3);
    assert_eq!(limiter.len(), 3);
    // An evicted bucket is simply a full one again, so nothing is punished by
    // the cap - only the map stays bounded.
    assert_eq!(limiter.take("a", 3), Allowed::Yes);
    assert_eq!(limiter.len(), 3);
}

#[test]
fn a_body_costs_one_token_per_kib_and_never_less_than_one() {
    assert_eq!(RATE_TOKEN_BYTES, 1024);
    assert_eq!(token_cost(0), 1);
    assert_eq!(token_cost(1), 1);
    assert_eq!(token_cost(1024), 1);
    assert_eq!(token_cost(1025), 2);
    assert_eq!(token_cost(4 * 1024 * 1024), 4096);
}

#[test]
fn tokens_are_spent_by_the_body_and_a_refusal_waits_for_the_whole_cost() {
    let limiter = RateLimiter::new(60);
    assert_eq!(limiter.take_tokens("client", 0, 50), Allowed::Yes);
    // 10 tokens left, 20 wanted: 60 a minute is one a second, so ten seconds.
    assert_eq!(
        limiter.take_tokens("client", 0, 20),
        Allowed::No { retry_after: 10 }
    );
    assert_eq!(limiter.take_tokens("client", 10, 20), Allowed::Yes);
    // A cost past a whole minute's allowance is clamped to it: a body bigger
    // than the budget waits a minute, it is not refused for ever.
    assert_eq!(limiter.take_tokens("client", 70, 1_000), Allowed::Yes);
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
