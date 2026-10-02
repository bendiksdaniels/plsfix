//! Per-client rate limiting: one token bucket per client key, refilled from an
//! injected clock so the maths is testable without sleeping. Owns the buckets
//! and their cap, whom a request is counted under (`TrustedProxy`), what a
//! body costs in tokens and whether a route counts as a write.
//! Invariant: pure arithmetic - no HTTP, no store, no wall clock of its own.

use std::{
    collections::HashMap,
    net::SocketAddr,
    str::FromStr,
    sync::{Mutex, MutexGuard, PoisonError},
};

use axum::{
    extract::ConnectInfo,
    http::{Extensions, HeaderMap, Method},
};

const CF_IP: &str = "cf-connecting-ip";
const FORWARDED_FOR: &str = "x-forwarded-for";
/// A request whose client cannot be named at all shares this bucket. Better
/// one crowded bucket than none: it still caps what an unnamed caller costs.
pub(crate) const UNKNOWN: &str = "unknown";

/// A bucket idle this long is dropped by the hourly sweeper, and by the cap
/// below before it evicts anyone; a client that comes back gets a full one.
pub const BUCKET_IDLE_SECS: i64 = 3600;

/// A write spends one token per KiB of body on top of its one request token,
/// so a minute of pushes is bounded by bytes and not only by requests.
pub const RATE_TOKEN_BYTES: usize = 1024;

/// Buckets held before the oldest idle client is evicted
/// (`MODELIS_RATE_MAX_CLIENTS`). Ten thousand keys is well past any real
/// deployment and still a bounded map.
pub const DEFAULT_MAX_CLIENTS: usize = 10_000;

/// A client key is an IP address, so nothing legitimate is near this: a longer
/// one is a header chosen to make the map row expensive, and is truncated.
const CLIENT_KEY_MAX: usize = 64;

/// Which forwarding header, if any, this deployment lets a client be named by.
/// Nothing is trusted by default: only the operator knows whether the proxy in
/// front rewrites the header, and a header the client may write is a bucket
/// the client may choose.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum TrustedProxy {
    #[default]
    None,
    Cloudflare,
    ForwardedFor,
}

impl FromStr for TrustedProxy {
    type Err = ();

    fn from_str(text: &str) -> Result<TrustedProxy, ()> {
        match text.trim().to_ascii_lowercase().as_str() {
            "" | "none" => Ok(TrustedProxy::None),
            "cloudflare" => Ok(TrustedProxy::Cloudflare),
            "xff" => Ok(TrustedProxy::ForwardedFor),
            _ => Err(()),
        }
    }
}

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
    max_clients: usize,
    buckets: Mutex<HashMap<String, Bucket>>,
}

impl RateLimiter {
    /// `per_minute` is both the capacity and the refill rate; below one the
    /// limiter would refuse for ever, so it is clamped.
    pub fn new(per_minute: u32) -> RateLimiter {
        RateLimiter {
            per_minute: per_minute.max(1),
            max_clients: DEFAULT_MAX_CLIENTS,
            buckets: Mutex::new(HashMap::new()),
        }
    }

    /// How many clients the map may hold at once; zero would evict every
    /// bucket it just made, so it is clamped.
    pub fn max_clients(mut self, clients: usize) -> RateLimiter {
        self.max_clients = clients.max(1);
        self
    }

    /// The allowance this limiter was built with: one minute's worth of
    /// tokens, which is also the burst a full bucket holds.
    pub fn per_minute(&self) -> u32 {
        self.per_minute
    }

    /// A poisoned map is a bucket that lost a race, never corrupt state: one
    /// panic elsewhere must not turn every later request into a panic too.
    fn buckets(&self) -> MutexGuard<'_, HashMap<String, Bucket>> {
        self.buckets.lock().unwrap_or_else(PoisonError::into_inner)
    }

    /// Spends one token for `key` at second `now`, or says how long the client
    /// has to wait for the next one.
    pub fn take(&self, key: &str, now: i64) -> Allowed {
        self.take_tokens(key, now, 1)
    }

    /// Spends `tokens` at once - a write charged by the size of its body. Time
    /// only ever moves forward here: a clock that steps back refills nothing
    /// rather than draining a bucket. A cost past a whole minute's allowance
    /// is clamped to it, so an oversized body waits a minute instead of being
    /// refused for ever.
    pub fn take_tokens(&self, key: &str, now: i64, tokens: u32) -> Allowed {
        let capacity = f64::from(self.per_minute);
        let cost = f64::from(tokens).min(capacity);
        let max_clients = self.max_clients;
        let mut buckets = self.buckets();
        if !buckets.contains_key(key) {
            make_room(&mut buckets, now, max_clients);
        }
        let bucket = buckets.entry(key.to_string()).or_insert(Bucket {
            tokens: capacity,
            seen: now,
        });
        let elapsed = f64::from(i32::try_from(now - bucket.seen).unwrap_or(i32::MAX)).max(0.0);
        bucket.tokens = (bucket.tokens + elapsed * capacity / 60.0).min(capacity);
        bucket.seen = now;
        if bucket.tokens >= cost {
            bucket.tokens -= cost;
            return Allowed::Yes;
        }
        // Whole seconds, never zero: a client told to retry immediately would
        // spin against the same empty bucket.
        let wait = (cost - bucket.tokens) * 60.0 / capacity;
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

/// Room for one more client: the idle buckets go first, and only if the map is
/// still full does the oldest-idle live one go. A flood of one-request clients
/// therefore costs a bounded map rather than an unbounded one, and the price
/// of an eviction is a full bucket, never a refusal.
fn make_room(buckets: &mut HashMap<String, Bucket>, now: i64, max_clients: usize) {
    if buckets.len() < max_clients {
        return;
    }
    buckets.retain(|_, bucket| now - bucket.seen < BUCKET_IDLE_SECS);
    while buckets.len() >= max_clients {
        let oldest = buckets
            .iter()
            .min_by_key(|(_, bucket)| bucket.seen)
            .map(|(key, _)| key.clone());
        match oldest {
            Some(key) => drop(buckets.remove(&key)),
            None => break,
        }
    }
}

/// What one body costs the write budget: one token per started KiB, and never
/// zero, so an empty write still counts as a write.
pub fn token_cost(bytes: usize) -> u32 {
    let tokens = bytes.div_ceil(RATE_TOKEN_BYTES).max(1);
    u32::try_from(tokens).unwrap_or(u32::MAX)
}

/// Which client a request is counted under: the header this deployment
/// declared a proxy for, else the peer address if the server was started with
/// `ConnectInfo`, else one shared bucket. The peer is optional so the route
/// tests, which drive the router directly, still run.
pub fn client_key(trust: TrustedProxy, headers: &HeaderMap, extensions: &Extensions) -> String {
    let trusted = match trust {
        TrustedProxy::None => None,
        TrustedProxy::Cloudflare => header_ip(headers, CF_IP, Hop::First),
        TrustedProxy::ForwardedFor => header_ip(headers, FORWARDED_FOR, Hop::Last),
    };
    let key = trusted.unwrap_or_else(|| {
        extensions
            .get::<ConnectInfo<SocketAddr>>()
            .map_or_else(|| UNKNOWN.to_string(), |peer| peer.0.ip().to_string())
    });
    key.chars().take(CLIENT_KEY_MAX).collect()
}

/// Which hop of a comma-separated forwarding header names the client. The LAST
/// one is the address the adjacent proxy wrote; every hop before it was handed
/// over by the client and is forgeable. `CF-Connecting-IP` carries one address,
/// so first and last are the same there.
enum Hop {
    First,
    Last,
}

/// One hop of a forwarding header, trimmed. An empty or unreadable header
/// names nobody, so the peer answers instead.
fn header_ip(headers: &HeaderMap, name: &str, hop: Hop) -> Option<String> {
    headers
        .get(name)
        .and_then(|value| value.to_str().ok())
        .and_then(|value| match hop {
            Hop::First => value.split(',').next(),
            Hop::Last => value.rsplit(',').next(),
        })
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

#[cfg(test)]
mod tests {
    use super::*;

    /// M3 of the N1 security review: one panic while the map was locked used
    /// to turn every later request into a panic of its own.
    #[test]
    fn a_panic_under_the_lock_does_not_wedge_the_limiter() {
        let limiter = std::sync::Arc::new(RateLimiter::new(10));
        let poisoner = std::sync::Arc::clone(&limiter);
        let panicked = std::thread::spawn(move || {
            let _held = poisoner.buckets();
            panic!("a handler died holding the bucket map");
        })
        .join();
        assert!(panicked.is_err());
        assert_eq!(limiter.take("client", 0), Allowed::Yes);
    }
}
