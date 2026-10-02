//! Hunt pass 2 (failures): a real sqlite lock held from outside the process,
//! a body that trickles in slower than the deadline overall but never stalls
//! one frame, a burst across three client keys per trust mode, and the
//! hourly sweeper racing a batch of writers.

mod common;

use std::{
    future::Future,
    pin::Pin,
    sync::Arc,
    task::{Context, Poll},
    time::Duration,
};

use axum::{
    body::{Body, Bytes},
    http::{header, Request, StatusCode},
};
use common::*;
use http_body_util::BodyExt;
use plsfix_server::{
    app_with_timeout,
    limits::{RateLimiter, TrustedProxy},
    relay::*,
    store::{auth_hash, Get, Store, LINK_TTL},
};
use tower::ServiceExt;

fn temp_db(name: &str) -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!(
        "modelis-hunt2-failures-{}-{name}.sqlite",
        std::process::id()
    ));
    let _ = std::fs::remove_file(&path);
    path
}

fn remove_db(path: &std::path::Path) {
    let _ = std::fs::remove_file(path);
    let _ = std::fs::remove_file(format!("{}-wal", path.display()));
    let _ = std::fs::remove_file(format!("{}-shm", path.display()));
}

/// A real sqlite lock, not an invented failure: a second connection holds the
/// database's one write lock (`BEGIN IMMEDIATE`; WAL allows many readers but
/// one writer) for about as long as a brief external touch would - a backup,
/// `sqlite3` opened by hand - then commits. Nothing in the hosted deployment
/// holds this lock from outside the process anywhere near that long
/// (`deploy.sh` restarts the unit, never two overlapping), so the push must
/// simply wait it out: `rusqlite::Connection::open` sets
/// `sqlite3_busy_timeout(db, 5000)` on every connection unconditionally, and
/// that five-second allowance is what lets the write retry past a hold this
/// short instead of answering `SQLITE_BUSY`. A shorter override here once
/// turned that wait into a guaranteed failure - the reason this test exists.
#[tokio::test]
async fn a_write_waits_out_a_lock_held_briefly_from_outside_the_process() {
    let path = temp_db("busy");
    {
        let store = Store::open(&path).unwrap(); // creates the schema
        drop(store);
    }
    let store = Store::open(&path).unwrap();
    let state = Arc::new(AppState::new(store));
    let app = routes(state.clone());

    const HOLD: Duration = Duration::from_millis(300);
    let locker = rusqlite::Connection::open(&path).unwrap();
    locker.execute_batch("BEGIN IMMEDIATE;").unwrap();
    let releaser = std::thread::spawn(move || {
        std::thread::sleep(HOLD);
        locker.execute_batch("COMMIT;").unwrap();
    });

    let started = std::time::Instant::now();
    let (status, body) = call(&app, "PUT", &format!("/api/links/{ID}"), Some(AUTH), b"x").await;
    let waited = started.elapsed();
    releaser.join().unwrap();

    assert_eq!(status, StatusCode::OK, "{body}");
    assert!(
        waited >= HOLD,
        "the push answered in {waited:?}, before the {HOLD:?} hold was even released"
    );

    remove_db(&path);
}

/// A body that yields one small frame every `PERIOD`, well under
/// `body_deadline` each time, for long enough that the SUM easily exceeds it.
struct SlowSteady {
    frames_left: u32,
    period: Duration,
    sleep: Pin<Box<tokio::time::Sleep>>,
}

impl SlowSteady {
    fn new(frames: u32, period: Duration) -> SlowSteady {
        SlowSteady {
            frames_left: frames,
            period,
            sleep: Box::pin(tokio::time::sleep(period)),
        }
    }
}

impl http_body::Body for SlowSteady {
    type Data = Bytes;
    type Error = std::io::Error;

    fn poll_frame(
        mut self: Pin<&mut Self>,
        context: &mut Context<'_>,
    ) -> Poll<Option<Result<http_body::Frame<Bytes>, Self::Error>>> {
        if self.frames_left == 0 {
            return Poll::Ready(None);
        }
        match self.sleep.as_mut().poll(context) {
            Poll::Pending => Poll::Pending,
            Poll::Ready(()) => {
                self.frames_left -= 1;
                self.sleep = Box::pin(tokio::time::sleep(self.period));
                Poll::Ready(Some(Ok(http_body::Frame::data(Bytes::from_static(
                    b"chunk",
                )))))
            }
        }
    }
}

/// The comment on `DEFAULT_BODY_DEADLINE_SECS` promises a per-frame deadline,
/// not a per-body one ("a slow but steady upload never meets it"); nothing
/// pinned that promise until now. Six frames forty ms apart is 240 ms total,
/// twice the 120 ms deadline, yet every single gap is well inside it. Both
/// the frame sleeps and tower-http's `TimeoutBody` run on `tokio::time`, so a
/// paused clock keeps the six 40 ms gaps exact while costing no real wall time
/// and no jitter-flake on a loaded runner.
#[tokio::test(start_paused = true)]
async fn a_slow_but_steady_body_succeeds_because_the_deadline_is_per_frame() {
    let mut state = AppState::new(Store::in_memory().unwrap());
    state.body_deadline = Duration::from_millis(120);
    let app = app_with_timeout(
        static_dir("hunt2-slow-steady"),
        Arc::new(state),
        Duration::from_secs(5),
    );

    let body = Body::new(SlowSteady::new(6, Duration::from_millis(40)));
    let request = Request::builder()
        .method("PUT")
        .uri(format!("/api/links/{ID}"))
        .header(header::AUTHORIZATION, format!("Bearer {AUTH}"))
        .body(body)
        .unwrap();
    let response = app.clone().oneshot(request).await.unwrap();
    assert_eq!(response.status(), StatusCode::OK);

    // Every frame arrived, concatenated in order - nothing was truncated by
    // whichever frame happened to straddle the deadline's reset.
    let get = app
        .oneshot(
            Request::get(format!("/api/links/{ID}"))
                .header(header::AUTHORIZATION, format!("Bearer {AUTH}"))
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    let blob = get.into_body().collect().await.unwrap().to_bytes();
    assert_eq!(blob, Bytes::from_static(b"chunkchunkchunkchunkchunkchunk"));
}

/// Three real client keys, nine requests released at once, a two-a-minute
/// write allowance: under a trust mode that can tell the three apart, each
/// gets its own budget (two through, one refused); trusting nothing, the
/// route tests carry no peer address either, so all nine share one bucket
/// and only its own two tokens get through, in whichever order they land.
async fn burst_counts(trust: TrustedProxy) -> (usize, usize) {
    let state = Arc::new(AppState {
        writes: RateLimiter::new(2),
        trusted_proxy: trust,
        ..AppState::new(Store::in_memory().unwrap())
    });
    let app = routes(state);
    let gate = Arc::new(tokio::sync::Barrier::new(9));

    let mut tasks = Vec::new();
    let mut next_id = 0u32;
    for client in ["9.9.9.1", "9.9.9.2", "9.9.9.3"] {
        for _ in 0..3 {
            let id = format!("{next_id:032x}");
            next_id += 1;
            let (app, gate, client) = (app.clone(), Arc::clone(&gate), client.to_string());
            tasks.push(tokio::spawn(async move {
                let request = Request::builder()
                    .method("PUT")
                    .uri(format!("/api/links/{id}"))
                    .header(header::AUTHORIZATION, format!("Bearer {AUTH}"))
                    .header("cf-connecting-ip", &client)
                    .header("x-forwarded-for", &client)
                    .body(Body::from(b"x".to_vec()))
                    .unwrap();
                gate.wait().await;
                app.oneshot(request).await.unwrap().status()
            }));
        }
    }

    let (mut ok, mut too_many) = (0, 0);
    for task in tasks {
        match task.await.unwrap() {
            StatusCode::OK => ok += 1,
            StatusCode::TOO_MANY_REQUESTS => too_many += 1,
            other => panic!("unexpected status {other}"),
        }
    }
    (ok, too_many)
}

#[tokio::test(flavor = "multi_thread", worker_threads = 4)]
async fn a_burst_from_three_client_keys_partitions_by_trust_mode() {
    assert_eq!(
        burst_counts(TrustedProxy::Cloudflare).await,
        (6, 3),
        "cloudflare: three independent budgets, two through each"
    );
    assert_eq!(
        burst_counts(TrustedProxy::ForwardedFor).await,
        (6, 3),
        "xff: three independent budgets, two through each"
    );
    assert_eq!(
        burst_counts(TrustedProxy::None).await,
        (2, 7),
        "untrusted: one shared budget of two, over all nine requests"
    );
}

/// The hourly sweeper is a plain `Store::sweep` call from another task, racing
/// live pushes on the same connection mutex - production shape, since the
/// unit does exactly this while the relay keeps serving. Sixteen writers each
/// push their own id three times while the sweeper churns 500 times over a
/// batch of already-dead rows; every writer must keep exactly its last two
/// revisions with the final push's content, and the dead rows must be gone.
#[tokio::test(flavor = "multi_thread", worker_threads = 4)]
async fn the_hourly_sweeper_racing_a_batch_of_writers_loses_nothing() {
    let store = Arc::new(Store::in_memory().unwrap());
    let auth = auth_hash(AUTH);

    const DEAD_ROWS: usize = 20;
    for dead in 0..DEAD_ROWS {
        store
            .put_link(
                &format!("dead-{dead:032x}"),
                &auth,
                b"old",
                -(LINK_TTL) - 10,
            )
            .unwrap();
    }

    const WRITERS: usize = 16;
    const PUSHES: usize = 3;
    let gate = Arc::new(tokio::sync::Barrier::new(WRITERS + 1));

    let sweeper = {
        let store = Arc::clone(&store);
        let gate = Arc::clone(&gate);
        tokio::spawn(async move {
            gate.wait().await;
            for _ in 0..500 {
                store.sweep(now()).unwrap();
            }
        })
    };

    let mut writers = Vec::new();
    for writer in 0..WRITERS {
        let store = Arc::clone(&store);
        let gate = Arc::clone(&gate);
        writers.push(tokio::spawn(async move {
            let id = format!("writer-{writer:032x}");
            gate.wait().await;
            for round in 0..PUSHES {
                store
                    .put_link(&id, &auth, format!("round {round}").as_bytes(), now())
                    .unwrap();
            }
            id
        }));
    }

    sweeper.await.unwrap();
    let mut ids = Vec::new();
    for writer in writers {
        ids.push(writer.await.unwrap());
    }

    for id in &ids {
        assert_eq!(
            store.rev_count(id),
            2,
            "writer {id} lost a revision to the concurrent sweeper"
        );
        match store.get_link(id, &auth, now()).unwrap() {
            Get::Found(found) => assert_eq!(
                found.blob,
                format!("round {}", PUSHES - 1).into_bytes(),
                "writer {id} did not read back its own last push"
            ),
            other => panic!("writer {id} became {other:?} under a concurrent sweep"),
        }
    }
    for dead in 0..DEAD_ROWS {
        assert_eq!(
            store.rev_count(&format!("dead-{dead:032x}")),
            0,
            "dead row {dead} survived 500 sweeps"
        );
    }
}
