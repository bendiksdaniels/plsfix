//! track.rs: a track someone picked instead of the score (`--track <file>`): cut to the video's
//! length, faded in with the picture's opening and out over its last seconds, so a song cut
//! mid-phrase still ends gently. Invariant: the result is exactly the video's length.

use crate::core::dsp::env::ramp;
use crate::core::dsp::SR;
use crate::core::sound::bus::{at, Bus};

const FADE_IN: f64 = 0.35;
/// A song cut mid-phrase needs a longer goodbye than the score, which ends on its own chord.
const FADE_OUT: f64 = 2.5;

/// The decoded track (any length, 48 kHz) as a master of exactly `seconds`; a shorter track
/// leaves silence after its own end.
pub fn fit(l: &[f32], r: &[f32], seconds: f64) -> Bus {
    let mut out = Bus::new(at(seconds, SR));
    let n = out.len().min(l.len()).min(r.len());
    for i in 0..n {
        let t = i as f64 / SR;
        let g = (ramp(t, 0.0, FADE_IN) * (1.0 - ramp(t, seconds - FADE_OUT, seconds - 0.05))) as f32;
        out.l[i] = l[i] * g;
        out.r[i] = r[i] * g;
    }
    out
}
