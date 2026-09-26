//! frames.rs tests: the timeline's frame count and times, and worker chunks that cover every
//! frame exactly once.
use plsfix_video::core::frames::{chunks, Timeline};

#[test]
fn timeline_counts_whole_frames() {
    let t = Timeline::new(57.0, 60);
    assert_eq!(t.frames, 3420);
    assert_eq!(t.time_of(60), 1.0);
    assert_eq!(t.seconds(), 57.0);
}

#[test]
fn chunks_cover_every_frame_once_in_order() {
    let c = chunks(3420, 7);
    assert_eq!(c.len(), 7);
    assert_eq!(c.first().unwrap().start, 0);
    assert_eq!(c.last().unwrap().end, 3420);
    for pair in c.windows(2) {
        assert_eq!(pair[0].end, pair[1].start);
    }
    let sizes: Vec<u32> = c.iter().map(|r| r.end - r.start).collect();
    assert!(sizes.iter().max().unwrap() - sizes.iter().min().unwrap() <= 1);
}

#[test]
fn chunks_never_make_empty_workers() {
    assert_eq!(chunks(3, 8).len(), 3);
    assert_eq!(chunks(10, 0), vec![0..10]);
}
