//! Binary for the pls,fix host: environment, bind and the hourly
//! sweeper. The router lives in `lib.rs`, the relay in `relay.rs`; this file
//! only reads `MODELIS_PORT`, `MODELIS_STATIC`, `MODELIS_DATA` and the three
//! limits, and starts the server, so the interesting parts stay testable
//! without a socket.

use std::{env, net::SocketAddr, path::PathBuf, str::FromStr, sync::Arc, time::Duration};

use plsfix_server::{
    app,
    limits::{RateLimiter, BUCKET_IDLE_SECS},
    relay::{now, AppState, DEFAULT_MAX_BYTES, DEFAULT_READ_PER_MIN, DEFAULT_WRITE_PER_MIN},
    store::Store,
};

/// Expired rows and idle rate-limit buckets die on every write too; this only
/// catches an idle server.
fn spawn_sweeper(state: Arc<AppState>) {
    tokio::spawn(async move {
        let mut tick = tokio::time::interval(Duration::from_secs(3600));
        loop {
            tick.tick().await;
            if let Err(error) = state.store.sweep(now()) {
                eprintln!("store sweep hourly: {error}");
            }
            state.writes.prune(now(), BUCKET_IDLE_SECS);
            state.reads.prune(now(), BUCKET_IDLE_SECS);
        }
    });
}

/// One env var, parsed or ignored: a typo in the unit file must not silently
/// halve a limit, so an unparsable value keeps the default and says so.
fn env_number<T: FromStr + Copy>(name: &str, default: T) -> T {
    match env::var(name).ok() {
        None => default,
        Some(text) => text.parse().unwrap_or_else(|_| {
            eprintln!("warning: {name}={text} is not a number, using the default");
            default
        }),
    }
}

/// The store plus the limits the routes gate on, straight from the unit file.
fn state_from_env(store: Store) -> AppState {
    AppState {
        max_bytes: env_number("MODELIS_MAX_BYTES", DEFAULT_MAX_BYTES),
        writes: RateLimiter::new(env_number(
            "MODELIS_RATE_WRITE_PER_MIN",
            DEFAULT_WRITE_PER_MIN,
        )),
        reads: RateLimiter::new(env_number(
            "MODELIS_RATE_READ_PER_MIN",
            DEFAULT_READ_PER_MIN,
        )),
        ..AppState::new(store)
    }
}

#[tokio::main]
async fn main() {
    let port: u16 = env_number("MODELIS_PORT", 8804);
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

    let addr = SocketAddr::from(([127, 0, 0, 1], port));
    let listener = tokio::net::TcpListener::bind(addr)
        .await
        .unwrap_or_else(|error| panic!("cannot bind {addr}: {error}"));
    println!(
        "plsfix-server serving {} on {addr}, data in {}",
        static_dir.display(),
        data_dir.display()
    );
    println!(
        "limits: {} bytes max, {} writes/min, {} reads/min per client",
        state.max_bytes,
        state.writes.per_minute(),
        state.reads.per_minute()
    );
    // ConnectInfo so the rate limiter can fall back to the peer address when a
    // request carries no forwarding header (a direct call, not through the
    // gateway); the extractor is optional, so the route tests still run.
    let service = app(static_dir, state).into_make_service_with_connect_info::<SocketAddr>();
    axum::serve(listener, service).await.expect("server error");
}
