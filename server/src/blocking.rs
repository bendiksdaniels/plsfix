//! The one way a route reaches the store. sqlite blocks the thread it runs on
//! and the request threads are the same ones answering `/healthz`, so a store
//! call is handed to `spawn_blocking` instead of run inline; the 30 s request
//! deadline cannot preempt a blocked worker, only an awaiting one.
//! Invariant: no handler in this crate touches `state.store` directly, and a
//! blocking task that dies takes its own request down, never the runtime.

use std::sync::Arc;

use crate::relay::{failed, short, AppState, Refused};
use crate::store::Store;

/// Runs one store call off the runtime's worker threads. `stage` and `id` are
/// the log line a failure gets, the same shape every other store failure uses.
pub(crate) async fn store_call<T, F>(
    state: &Arc<AppState>,
    stage: &'static str,
    id: &str,
    work: F,
) -> Result<T, Refused>
where
    F: FnOnce(&Store) -> rusqlite::Result<T> + Send + 'static,
    T: Send + 'static,
{
    let state = Arc::clone(state);
    let id = id.to_string();
    match tokio::task::spawn_blocking(move || work(&state.store)).await {
        Ok(Ok(value)) => Ok(value),
        Ok(Err(error)) => Err(failed(stage, &id, &error)),
        // A panic inside the call, or a runtime shutting down: one 500, and
        // the next request finds the store where it left it (the lock is
        // released poisoned, and both lock sites read through a poisoning).
        Err(join) => {
            eprintln!("store {stage} {}: blocking task {join}", short(&id));
            Err(Refused::Store)
        }
    }
}

#[cfg(test)]
mod tests {
    use std::{
        path::PathBuf,
        sync::{mpsc, Arc},
        time::{Duration, Instant},
    };

    use axum::{
        body::Body,
        http::{header, Request, StatusCode},
    };
    use tower::ServiceExt;

    use crate::{app_with_timeout, relay::AppState, store::Store};

    const ID: &str = "0123456789abcdef0123456789abcdef";
    const AUTH: &str = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";

    /// I2/I3 of the N1 security review: the store is one `Mutex<Connection>`
    /// called from async handlers, so a slow query used to hold a worker
    /// thread and every other route with it. Single-threaded runtime on
    /// purpose - if the store call ran inline there would be no thread left
    /// for `/healthz` at all.
    #[tokio::test]
    async fn a_blocked_store_call_does_not_stall_healthz() {
        let state = Arc::new(AppState::new(Store::in_memory().unwrap()));
        let app = app_with_timeout(
            PathBuf::from("dist"),
            Arc::clone(&state),
            Duration::from_secs(30),
        );

        // Hold the store's one connection on an OS thread of its own, and only
        // start timing once it really is held.
        let holder = Arc::clone(&state);
        let (locked, held) = mpsc::channel();
        let keeper = std::thread::spawn(move || {
            let _conn = holder.store.conn();
            locked.send(()).unwrap();
            std::thread::sleep(Duration::from_millis(300));
        });
        held.recv().unwrap();

        let started = Instant::now();
        let waiting = tokio::spawn(
            app.clone().oneshot(
                Request::get(format!("/api/links/{ID}"))
                    .header(header::AUTHORIZATION, format!("Bearer {AUTH}"))
                    .body(Body::empty())
                    .unwrap(),
            ),
        );
        // Let the store call reach the point where it waits. Inline, this is
        // where the single thread disappears for 300 ms.
        for _ in 0..8 {
            tokio::task::yield_now().await;
        }
        let health = app
            .oneshot(Request::get("/healthz").body(Body::empty()).unwrap())
            .await
            .unwrap();

        assert_eq!(health.status(), StatusCode::OK);
        assert!(
            started.elapsed() < Duration::from_millis(200),
            "healthz waited {:?} for a busy store",
            started.elapsed()
        );
        assert_eq!(waiting.await.unwrap().unwrap().status(), StatusCode::NOT_FOUND);
        keeper.join().unwrap();
    }
}
