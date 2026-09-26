//! frames.rs: the render timeline as whole frames, and how frames split across workers.
//! Invariant: `chunks` covers 0..frames exactly once, in order, sizes differing by at most one.

use std::ops::Range;

/// A timeline sampled at `fps`: frame `f` shows time `f / fps` seconds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Timeline {
    pub fps: u32,
    pub frames: u32,
}

impl Timeline {
    /// The frames needed to show `duration_s` seconds at `fps` (rounded to the nearest frame).
    pub fn new(duration_s: f64, fps: u32) -> Timeline {
        let frames = (duration_s * fps as f64).round().max(0.0) as u32;
        Timeline { fps, frames }
    }

    /// The time in seconds that frame `frame` shows.
    pub fn time_of(&self, frame: u32) -> f64 {
        frame as f64 / self.fps as f64
    }

    /// The length in seconds of the whole timeline.
    pub fn seconds(&self) -> f64 {
        self.frames as f64 / self.fps as f64
    }
}

/// Splits `frames` into `workers` contiguous ranges (fewer when there are fewer frames).
pub fn chunks(frames: u32, workers: u32) -> Vec<Range<u32>> {
    let workers = workers.max(1).min(frames.max(1));
    let base = frames / workers;
    let extra = frames % workers;
    let mut out = Vec::with_capacity(workers as usize);
    let mut start = 0;
    for w in 0..workers {
        let len = base + u32::from(w < extra);
        out.push(start..start + len);
        start += len;
    }
    out.retain(|r| !r.is_empty());
    out
}
