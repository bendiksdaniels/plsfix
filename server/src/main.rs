//! Binary for the pls,fix host: environment, bind and the hourly
//! sweeper. The router lives in `lib.rs`, the relay in `relay.rs`; this file
//! only reads `MODELIS_PORT`, `MODELIS_BIND`, `MODELIS_STATIC`, `MODELIS_DATA`,
//! `MODELIS_MANIFEST`, `MODELIS_PUBLIC_URL`, `MODELIS_TRUSTED_PROXY` and the
//! limits (`MODELIS_MAX_BYTES`, `MODELIS_RATE_WRITE_PER_MIN`,
//! `MODELIS_RATE_WRITE_KIB_PER_MIN`, `MODELIS_RATE_READ_PER_MIN`,
//! `MODELIS_RATE_MAX_CLIENTS`, `MODELIS_MAX_INFLIGHT_WRITES`), and starts the
//! server, so the interesting parts stay testable without a socket.

use std::{
    env,
    net::{IpAddr, SocketAddr},
    path::{Path, PathBuf},
    str::FromStr,
    sync::Arc,
    time::Duration,
};

use plsfix_server::{
    app,
    limits::{RateLimiter, TrustedProxy, BUCKET_IDLE_SECS, DEFAULT_MAX_CLIENTS},
    manifest::ManifestSource,
    relay::{
        now, AppState, DEFAULT_MAX_BYTES, DEFAULT_MAX_INFLIGHT_WRITES, DEFAULT_READ_PER_MIN,
        DEFAULT_WRITE_KIB_PER_MIN, DEFAULT_WRITE_PER_MIN, INBOX_MAX_PER_WS,
    },
    store::Store,
    store_room::Caps,
};

/// Expired rows and idle rate-limit buckets die on every write too; this only
/// catches an idle server. The sweep is a blocking sqlite call like every
/// other, so it goes to the blocking pool rather than a runtime worker.
fn spawn_sweeper(state: Arc<AppState>) {
    tokio::spawn(async move {
        let mut tick = tokio::time::interval(Duration::from_secs(3600));
        loop {
            tick.tick().await;
            let sweeping = Arc::clone(&state);
            let swept = tokio::task::spawn_blocking(move || sweeping.store.sweep(now())).await;
            match swept {
                Ok(Err(error)) => eprintln!("store sweep hourly: {error}"),
                Err(error) => eprintln!("store sweep hourly: blocking task {error}"),
                Ok(Ok(_)) => {}
            }
            // The sweep just changed what the store holds, so the counters
            // `/version` reuses are no longer the truth.
            state.counts.clear().await;
            state.writes.prune(now(), BUCKET_IDLE_SECS);
            state.write_bytes.prune(now(), BUCKET_IDLE_SECS);
            state.reads.prune(now(), BUCKET_IDLE_SECS);
        }
    });
}

/// One env var, parsed or ignored: a typo in the unit file must not silently
/// halve a limit or move the bind, so an unparsable value keeps the default
/// and says so.
fn env_parsed<T: FromStr + Copy>(name: &str, default: T) -> T {
    match env::var(name).ok() {
        None => default,
        Some(text) => text.parse().unwrap_or_else(|_| {
            eprintln!("warning: {name}={text} cannot be parsed, using the default");
            default
        }),
    }
}

/// `MODELIS_TRUSTED_PROXY`: which forwarding header, if any, may name a client
/// for the rate limiter. Unset trusts nothing, so a deployment behind a proxy
/// must say so - a header a client can write is a bucket a client can choose.
fn trusted_proxy_from_env() -> TrustedProxy {
    match env::var("MODELIS_TRUSTED_PROXY").ok() {
        None => TrustedProxy::None,
        Some(text) => text.parse().unwrap_or_else(|()| {
            eprintln!("warning: MODELIS_TRUSTED_PROXY={text} is not none, cloudflare or xff - trusting nothing");
            TrustedProxy::None
        }),
    }
}

/// One limiter: its per-minute allowance from `name`, and the same bucket-map
/// cap as every other limiter, so no one map can grow while another is bounded.
fn limiter_from_env(name: &str, default: u32, clients: usize) -> RateLimiter {
    RateLimiter::new(env_parsed(name, default)).max_clients(clients)
}

/// The store, the limits the routes gate on and the manifest source, straight
/// from the unit file (or the container's environment). The storage ceiling is
/// set on the store itself, which is where a write is held to it.
fn state_from_env(store: Store) -> AppState {
    let clients = env_parsed("MODELIS_RATE_MAX_CLIENTS", DEFAULT_MAX_CLIENTS);
    let mut state = AppState {
        writes: limiter_from_env("MODELIS_RATE_WRITE_PER_MIN", DEFAULT_WRITE_PER_MIN, clients),
        write_bytes: limiter_from_env(
            "MODELIS_RATE_WRITE_KIB_PER_MIN",
            DEFAULT_WRITE_KIB_PER_MIN,
            clients,
        ),
        reads: limiter_from_env("MODELIS_RATE_READ_PER_MIN", DEFAULT_READ_PER_MIN, clients),
        trusted_proxy: trusted_proxy_from_env(),
        max_inflight_writes: env_parsed("MODELIS_MAX_INFLIGHT_WRITES", DEFAULT_MAX_INFLIGHT_WRITES),
        manifest: ManifestSource::from_env(),
        ..AppState::new(store)
    };
    state.store.set_caps(Caps {
        max_bytes: env_parsed("MODELIS_MAX_BYTES", DEFAULT_MAX_BYTES),
        inbox_rows: INBOX_MAX_PER_WS,
    });
    state
}

/// A deployment that trusts no forwarding header counts every request under
/// the peer address, which behind a proxy is the proxy: one bucket for everyone.
/// Bound to loopback it IS behind something, so that is a warning; bound to the
/// world (the container image) a proxy may or may not sit in front, so that is
/// a note. Never a refusal: an operator may mean it (I1 of the N1 security review).
fn one_bucket_line(trust: TrustedProxy, bind: IpAddr) -> Option<&'static str> {
    match (trust, bind.is_loopback()) {
        (TrustedProxy::None, true) => Some(
            "warning: behind a proxy and trusting nothing: every client shares one bucket; set MODELIS_TRUSTED_PROXY",
        ),
        (TrustedProxy::None, false) => Some(
            "note: trusting no forwarding header: behind a reverse proxy every client would share one bucket; set MODELIS_TRUSTED_PROXY if one sits in front",
        ),
        _ => None,
    }
}

/// What this process is, on three lines, before the first request.
fn announce(state: &AppState, static_dir: &Path, data_dir: &Path, addr: SocketAddr, bind: IpAddr) {
    println!(
        "plsfix-server serving {} on {addr}, data in {}",
        static_dir.display(),
        data_dir.display()
    );
    println!(
        "limits: {} bytes max, {} writes/min, {} KiB/min, {} reads/min per client, {} writes at once, client key from {:?}",
        state.max_bytes(),
        state.writes.per_minute(),
        state.write_bytes.per_minute(),
        state.reads.per_minute(),
        state.inflight_writes(),
        state.trusted_proxy,
    );
    println!(
        "manifest: {}{}",
        state.manifest.file.display(),
        match &state.manifest.public_url {
            Some(url) => format!(", re-pointed at {url}"),
            None => String::from(" verbatim"),
        }
    );
    if let Some(line) = one_bucket_line(state.trusted_proxy, bind) {
        eprintln!("{line}");
    }
}

#[tokio::main]
async fn main() {
    let port: u16 = env_parsed("MODELIS_PORT", 8804);
    // Loopback behind the gateway's nginx; a container sets MODELIS_BIND=0.0.0.0.
    let bind: IpAddr = env_parsed("MODELIS_BIND", IpAddr::from([127, 0, 0, 1]));
    let static_dir =
        PathBuf::from(env::var("MODELIS_STATIC").unwrap_or_else(|_| "dist".to_string()));
    if !static_dir.join("taskpane.html").is_file() {
        eprintln!(
            "warning: {} has no taskpane.html - build the pane first (npm run build)",
            static_dir.display()
        );
    }

    let data_dir = PathBuf::from(env::var("MODELIS_DATA").unwrap_or_else(|_| "data".to_string()));
    let database = data_dir.join("relay.sqlite");
    // Store::open creates data_dir (create_dir_all) before opening the file.
    let store = Store::open(&database)
        .unwrap_or_else(|error| panic!("cannot open {}: {error}", database.display()));
    let state = Arc::new(state_from_env(store));
    spawn_sweeper(state.clone());

    let addr = SocketAddr::new(bind, port);
    let listener = tokio::net::TcpListener::bind(addr)
        .await
        .unwrap_or_else(|error| panic!("cannot bind {addr}: {error}"));
    announce(&state, &static_dir, &data_dir, addr, bind);
    // ConnectInfo so the rate limiter can fall back to the peer address when a
    // request carries no forwarding header (a direct call, not through the
    // gateway); the extractor is optional, so the route tests still run.
    let service = app(static_dir, state).into_make_service_with_connect_info::<SocketAddr>();
    axum::serve(listener, service).await.expect("server error");
}

#[cfg(test)]
mod tests {
    use super::*;

    const LOOPBACK: IpAddr = IpAddr::V4(std::net::Ipv4Addr::LOCALHOST);
    const OPEN: IpAddr = IpAddr::V4(std::net::Ipv4Addr::UNSPECIFIED);

    /// The hosted unit's shape - loopback behind the gateway, no trust mode -
    /// is the one that quietly puts every client in one bucket. The container
    /// image binds the world behind Caddy or nginx and gets the conditional line.
    #[test]
    fn a_bind_that_trusts_nothing_gets_its_line() {
        assert!(one_bucket_line(TrustedProxy::None, LOOPBACK)
            .unwrap()
            .starts_with("warning:"));
        assert!(one_bucket_line(TrustedProxy::None, OPEN)
            .unwrap()
            .starts_with("note:"));
        assert!(one_bucket_line(TrustedProxy::Cloudflare, LOOPBACK).is_none());
        assert!(one_bucket_line(TrustedProxy::ForwardedFor, LOOPBACK).is_none());
        assert!(one_bucket_line(TrustedProxy::Cloudflare, OPEN).is_none());
    }

    /// A trust mode nobody can read is not a reason to start trusting one.
    #[test]
    fn an_unreadable_trust_mode_falls_back_to_trusting_nothing() {
        assert_eq!("cloudflare".parse(), Ok(TrustedProxy::Cloudflare));
        assert!("cf".parse::<TrustedProxy>().is_err());
    }
}
