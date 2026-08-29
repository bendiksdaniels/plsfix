//! Binary for the pls,fix host: environment, bind and the hourly
//! sweeper. The router lives in `lib.rs`, the relay in `relay.rs`; this file
//! only reads `MODELIS_PORT`, `MODELIS_STATIC` and `MODELIS_DATA` and starts
//! the server, so the interesting parts stay testable without a socket.

use std::{env, net::SocketAddr, path::PathBuf, sync::Arc, time::Duration};

use plsfix_server::{
    app,
    relay::{now, AppState},
    store::Store,
};

/// Expired rows die on every write too; this only catches an idle server.
fn spawn_sweeper(state: Arc<AppState>) {
    tokio::spawn(async move {
        let mut tick = tokio::time::interval(Duration::from_secs(3600));
        loop {
            tick.tick().await;
            if let Err(error) = state.store.sweep(now()) {
                eprintln!("store sweep hourly: {error}");
            }
        }
    });
}

#[tokio::main]
async fn main() {
    let port: u16 = env::var("MODELIS_PORT")
        .ok()
        .and_then(|value| value.parse().ok())
        .unwrap_or(8804);
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
    let state = Arc::new(AppState { store });
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
    axum::serve(listener, app(static_dir, state))
        .await
        .expect("server error");
}
