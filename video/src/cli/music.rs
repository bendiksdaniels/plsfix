//! music.rs: `plsfix-video music [--track <file>]`: the soundtrack. By default the score and
//! the sound design, written from the picture's cues; with --track, that audio file instead (cut
//! to length, faded). Either way the level is set to -16 LUFS with peaks under -2 dBFS and the
//! result is video/build/music.wav. `render` calls it first, so sound follows picture.

use crate::core::lang::Lang;
use crate::cli::{build_dir, video_dir};
use crate::core::dsp::dynamics::{db_to_gain, gain, limit};
use crate::core::dsp::SR;
use crate::core::sound::bus::Bus;
use crate::core::sound::{cues, mixdown, score, track};
use crate::core::viewport::Viewport;
use crate::io::comp_page::CompPage;
use crate::io::loudness::{self, Loudness};
use crate::io::{ffmpeg, static_server, wav};
use anyhow::{bail, Result};
use std::path::{Path, PathBuf};

/// Loudness for a video played on a laptop or in a meeting room (online-video level).
pub const TARGET_LUFS: f64 = -16.0;
pub const CEILING_DBFS: f64 = -2.0;

pub fn path() -> PathBuf {
    build_dir().join("music.wav")
}

pub fn run(own: Option<&Path>) -> Result<()> {
    let port = static_server::serve(video_dir())?;
    // the cues come from the timing tables, the same in every cut
    let mut page = CompPage::open(port, Viewport::STAGE, Lang::Lv)?;
    let seconds = page.meta.duration;
    let master = match own {
        Some(file) => picked(file, seconds)?,
        None => {
            let cues = cues::parse(&page.cdp.eval("window.__cues()")?, seconds).map_err(anyhow::Error::msg)?;
            let sections = score::sections(&cues).map_err(anyhow::Error::msg)?;
            let notes = score::notes(&sections, seconds);
            let sfx = cues.iter().filter(|c| c.kind != cues::Kind::Section).count();
            println!("music: {} notes, {sfx} sound cues, sections at bars {:?}", notes.len(), sections.iter().map(|(b, _)| b).collect::<Vec<_>>());
            mixdown::render(&notes, &cues, seconds)
        }
    };
    let l = level(master)?;
    println!("music: {:.1} LUFS, true peak {:.1} dBTP -> {}", l.integrated, l.true_peak, path().display());
    Ok(())
}

/// A picked audio file as the master, cut to the video's length.
fn picked(file: &Path, seconds: f64) -> Result<Bus> {
    if !file.is_file() {
        bail!("music: the track {} is not a file", file.display());
    }
    let (l, r) = ffmpeg::decode(file, seconds)?;
    let len = l.len() as f64 / SR;
    if len < 1.0 {
        bail!("music: the track {} has no audio (decoded {len:.2}s)", file.display());
    }
    println!("music: track {} ({len:.1}s of the video's {seconds:.1}s)", file.display());
    if len < seconds - 0.1 {
        println!("music: note: the track ends {:.1}s before the video does; the rest is silent", seconds - len);
    }
    Ok(track::fit(&l, &r, seconds))
}

/// Writes the master to music.wav at the target loudness: measure, correct, limit, again.
fn level(mut master: Bus) -> Result<Loudness> {
    let out = path();
    for round in 1..=3 {
        wav::write(&out, &master, SR as u32)?;
        let l = loudness::measure(&out)?;
        if (l.integrated - TARGET_LUFS).abs() <= 0.3 && l.true_peak <= CEILING_DBFS + 0.6 {
            return Ok(l);
        }
        let over = l.true_peak + TARGET_LUFS - l.integrated - CEILING_DBFS;
        println!("music: round {round}: {:.1} LUFS, true peak {:.1} dBTP; the limiter takes up to {:.1} dB", l.integrated, l.true_peak, over.max(0.0));
        if round == 3 {
            bail!("music: {:.1} LUFS, true peak {:.1} dBTP after 3 rounds (target {TARGET_LUFS} LUFS, peak under {CEILING_DBFS} dBFS)", l.integrated, l.true_peak);
        }
        gain(&mut master.l, &mut master.r, db_to_gain(TARGET_LUFS - l.integrated));
        limit(&mut master.l, &mut master.r, db_to_gain(CEILING_DBFS) as f32, SR, 0.002, 0.08);
    }
    unreachable!("the third round returns or bails")
}
