//! still.rs: `plsfix-video still <t>...`: renders the frames at the given times to
//! video/build/stills/still-<t>.png (the English cut still-en-<t>.png), one Chrome for all of
//! them. For QA of a scene in progress.

use crate::cli::{build_dir, video_dir};
use crate::core::lang::Lang;
use crate::core::viewport::Viewport;
use crate::io::comp_page::{write_file, CompPage};
use crate::io::static_server;
use anyhow::{bail, Result};
use std::time::Instant;

pub fn run(times: &[f64], lang: Lang) -> Result<()> {
    if times.is_empty() {
        bail!("still: give at least one time in seconds");
    }
    let port = static_server::serve(video_dir())?;
    let mut page = CompPage::open(port, Viewport::STAGE, lang)?;
    let started = Instant::now();
    for &t in times {
        let png = page.frame_png(t)?;
        let path = build_dir().join("stills").join(lang.still(t));
        write_file(&path, &png)?;
        println!("{}", path.display());
    }
    let per = started.elapsed().as_secs_f64() * 1000.0 / times.len() as f64;
    println!("still: {} frame(s), {per:.0} ms each, timeline {:.2}s at {} fps", times.len(), page.meta.duration, page.meta.fps);
    Ok(())
}
