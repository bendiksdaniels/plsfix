//! cues.rs: the sounds the picture asks for, read from the composition's `__cues()` (every scene
//! registers its own from its timing table, so sound and picture share one clock).
//! Invariant: an unknown kind or a time outside the video is an error naming the cue.

use serde::Deserialize;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    Section,
    Whoosh,
    Click,
    Tap,
    Pop,
    Impact,
    Riser,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Cue {
    pub t: f64,
    pub kind: Kind,
    /// seconds (whoosh, riser)
    pub dur: f64,
    /// -1 left .. 1 right
    pub pan: f64,
    /// MIDI note (pop)
    pub pitch: f64,
    /// dB relative to the kind's own level
    pub gain: f64,
    /// section name (section)
    pub name: String,
}

#[derive(Deserialize)]
struct Raw {
    t: f64,
    kind: String,
    dur: Option<f64>,
    pan: Option<f64>,
    pitch: Option<f64>,
    gain: Option<f64>,
    name: Option<String>,
}

fn kind(s: &str) -> Option<Kind> {
    Some(match s {
        "section" => Kind::Section,
        "whoosh" => Kind::Whoosh,
        "click" => Kind::Click,
        "tap" => Kind::Tap,
        "pop" => Kind::Pop,
        "impact" => Kind::Impact,
        "riser" => Kind::Riser,
        _ => return None,
    })
}

/// Parses and checks the page's cue list for a video of `seconds`.
pub fn parse(json: &serde_json::Value, seconds: f64) -> Result<Vec<Cue>, String> {
    let raw: Vec<Raw> = serde_json::from_value(json.clone()).map_err(|e| format!("cues: not a cue list ({e})"))?;
    let mut out = Vec::with_capacity(raw.len());
    for (i, r) in raw.into_iter().enumerate() {
        let k = kind(&r.kind).ok_or_else(|| format!("cue {i} at {}s: unknown kind {}", r.t, r.kind))?;
        if !(0.0..seconds).contains(&r.t) {
            return Err(format!("cue {i} ({}) at {}s: outside the video (0-{seconds}s)", r.kind, r.t));
        }
        if k == Kind::Section && r.name.is_none() {
            return Err(format!("cue {i} at {}s: a section needs a name", r.t));
        }
        out.push(Cue {
            t: r.t, kind: k, dur: r.dur.unwrap_or(0.5).max(0.05), pan: r.pan.unwrap_or(0.0).clamp(-1.0, 1.0),
            pitch: r.pitch.unwrap_or(76.0), gain: r.gain.unwrap_or(0.0), name: r.name.unwrap_or_default(),
        });
    }
    out.sort_by(|a, b| a.t.total_cmp(&b.t));
    Ok(out)
}
