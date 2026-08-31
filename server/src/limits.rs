//! Per-client rate limiting: one token bucket per client key, refilled from an
//! injected clock so the maths is testable without sleeping. Owns the buckets,
//! the key a request is counted under and whether a route counts as a write.
//! Invariant: pure arithmetic - no HTTP, no store, no wall clock of its own.

use std::{
    collections::HashMap,
    net::SocketAddr,
    sync::{Mutex, MutexGuard},
};

use axum::{
    extract::ConnectInfo,
    http::{Extensions, HeaderMap, Method},
};

/// Cloudflare's own header wins: behind the tunnel every peer address is the
/// tunnel's, so without it one bucket would hold the whole internet.
const CF_IP: &str = "cf-connecting-ip";
const FORWARDED_FOR: &str = "x-forwarded-for";
/// A request whose client cannot be named at all shares this bucket. Better
/// one crowded bucket than none: it still caps what an unnamed caller costs.
const UNKNOWN: &str = "unknown";

/// A bucket idle this long is dropped by the hourly sweeper; a client that
/// comes back simply gets a full one.
pub const BUCKET_IDLE_SECS: i64 = 3600;

/// One client's bucket: whole and partial tokens left, and the second they
/// were last refilled at.
#[derive(Clone, Copy)]
struct Bucket {
    tokens: f64,
    seen: i64,
}

/// Whether a request may proceed, and how long a refused one should wait.
#[derive(Debug, PartialEq, Eq)]
pub enum Allowed {
    Yes,
    No { retry_after: u64 },
}

/// A token bucket per client key. Capacity is one minute's allowance, so a
/// pane that opens a deck and polls every link at once is not punished for
/// bursting, while a caller that keeps going is held to the rate.
pub struct RateLimiter {
    per_minute: u32,
    buckets: Mutex<HashMap<String, Bucket>>,
}

impl RateLimiter {
    /// `per_minute` is both the capacity and the refill rate; below one the
    /// limiter would refuse for ever, so it is clamped.
    pub fn new(per_minute: u32) -> RateLimiter {
        RateLimiter {
            per_minute: per_minute.max(1),
            buckets: Mutex::new(HashMap::new()),
        }
    }

    /// The allowance this limiter was built with: one minute's worth of
    /// tokens, which is also the burst a full bucket holds.
    pub fn per_minute(&self) -> u32 {
        self.per_minute
    }

    fn buckets(&self) -> MutexGuard<'_, HashMap<String, Bucket>> {
        self.buckets.lock().expect("rate limiter mutex poisoned")
    }

    /// Spends one token for `key` at second `now`, or says how long the client
    /// has to wait for the next one. Time only ever moves forward here: a
    /// clock that steps back refills nothing rather than draining a bucket.
    pub fn take(&self, key: &str, now: i64) -> Allowed {
        let capacity = f64::from(self.per_minute);
        let mut buckets = self.buckets();
        let bucket = buckets.entry(key.to_string()).or_insert(Bucket {
            tokens: capacity,
            seen: now,
        });
        let elapsed = f64::from(i32::try_from(now - bucket.seen).unwrap_or(i32::MAX)).max(0.0);
        bucket.tokens = (bucket.tokens + elapsed * capacity / 60.0).min(capacity);
        bucket.seen = now;
        if bucket.tokens >= 1.0 {
            bucket.tokens -= 1.0;
            return Allowed::Yes;
        }
        // Whole seconds, never zero: a client told to retry immediately would
        // spin against the same empty bucket.
        let wait = (1.0 - bucket.tokens) * 60.0 / capacity;
        Allowed::No {
            retry_after: (wait.ceil() as u64).max(1),
        }
    }

    /// Drops buckets nobody has used for `idle` seconds, so a long-running
    /// server does not grow one entry per client it ever saw. A dropped bucket
    /// is a full one, and only an idle client can lose it.
    pub fn prune(&self, now: i64, idle: i64) -> usize {
        let mut buckets = self.buckets();
        let before = buckets.len();
        buckets.retain(|_, bucket| now - bucket.seen < idle);
        before - buckets.len()
    }

    /// Buckets held right now - test support, and the sweeper's log line.
    pub fn len(&self) -> usize {
        self.buckets().len()
    }

    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }
}

/// Which client a request is counted under: Cloudflare's header, then the
/// first hop of `X-Forwarded-For`, then the peer address if the server was
/// started with `ConnectInfo`, else one shared bucket. The peer is optional so
/// the route tests, which drive the router directly, still run.
pub fn client_key(headers: &HeaderMap, extensions: &Extensions) -> String {
    if let Some(ip) = header_ip(headers, CF_IP) {
        return ip;
    }
    if let Some(ip) = header_ip(headers, FORWARDED_FOR) {
        return ip;
    }
    extensions
        .get::<ConnectInfo<SocketAddr>>()
        .map_or_else(|| UNKNOWN.to_string(), |peer| peer.0.ip().to_string())
}

/// The first entry of a comma-separated forwarding header, trimmed. An empty
/// or unreadable header names nobody, so the next source is tried.
fn header_ip(headers: &HeaderMap, name: &str) -> Option<String> {
    headers
        .get(name)
        .and_then(|value| value.to_str().ok())
        .and_then(|value| value.split(',').next())
        .map(str::trim)
        .filter(|ip| !ip.is_empty())
        .map(str::to_string)
}

/// Which allowance a route spends. Writes are the expensive half: a PUT or a
/// DELETE, and every POST except the three batch routes, which only read.
/// Matched on the route's own name so nesting cannot change the answer.
pub fn is_write(method: &Method, path: &str) -> bool {
    match *method {
        Method::PUT | Method::DELETE => true,
        Method::POST => !["/status", "/fetch", "/touch"]
            .iter()
            .any(|batch| path.ends_with(batch)),
        _ => false,
    }
}
