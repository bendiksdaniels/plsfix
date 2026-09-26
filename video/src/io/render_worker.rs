//! render_worker.rs: one render worker: its own Chrome on the composition and its own ffmpeg,
//! stepping through a contiguous range of frames. A page error fails the worker with the time.

use crate::core::viewport::Viewport;
use crate::io::comp_page::CompPage;
use crate::io::ffmpeg::Encoder;
use anyhow::{Context, Result};
use std::ops::Range;
use std::path::PathBuf;
use std::time::Instant;

pub struct Job {
    pub index: usize,
    pub frames: Range<u32>,
    pub fps: u32,
    pub vp: Viewport,
    pub port: u16,
    pub out: PathBuf,
}

/// Renders the job's frames into its part file; returns the frames written.
pub fn run(job: &Job) -> Result<u32> {
    let started = Instant::now();
    let mut page = CompPage::open(job.port, job.vp).with_context(|| format!("render worker {}", job.index))?;
    let mut enc = Encoder::start(job.fps, &job.out)?;
    for f in job.frames.clone() {
        let png = page.frame_png(f as f64 / job.fps as f64).with_context(|| format!("render worker {}: frame {f}", job.index))?;
        enc.write(&png)?;
    }
    enc.finish()?;
    let n = job.frames.end - job.frames.start;
    println!("render: worker {} frames {}..{} ({n}) in {:.1}s", job.index, job.frames.start, job.frames.end, started.elapsed().as_secs_f64());
    Ok(n)
}
