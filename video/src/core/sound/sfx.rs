//! sfx.rs: the sound design the picture cues: a whoosh for a moving window, a click for the
//! pointer, a soft tap for a finger, a pop for a card landing (pitched to the music), an impact
//! under a reveal and a riser into one. Each renders one cue into the buses at its own level.

use crate::core::dsp::dynamics::db_to_gain;
use crate::core::dsp::env::decay;
use crate::core::dsp::filter::{Biquad, Svf};
use crate::core::dsp::osc::{midi_hz, Noise};
use crate::core::dsp::SR;
use crate::core::sound::bus::{at, Bus};
use crate::core::sound::cues::{Cue, Kind};
use std::f64::consts::{PI, TAU};

/// Relative levels, one place: whoosh, click, tap, pop, impact, riser.
pub const LEVELS: [f64; 6] = [0.35, 0.35, 0.45, 0.22, 0.8, 0.3];

pub fn render(sfx: &mut Bus, send: &mut Bus, c: &Cue, noise: &mut Noise) {
    match c.kind {
        Kind::Whoosh => whoosh(sfx, send, c, noise),
        Kind::Click => click(sfx, c, noise),
        Kind::Tap => tap(sfx, c, noise),
        Kind::Pop => pop(sfx, send, c),
        Kind::Impact => impact(sfx, send, c, noise),
        Kind::Riser => riser(sfx, send, c, noise),
        Kind::Section => {}
    }
}

/// band-passed noise whose centre rises and falls with the move, panned across it
fn whoosh(sfx: &mut Bus, send: &mut Bus, c: &Cue, noise: &mut Noise) {
    let (mut fl, mut fr) = (Svf::default(), Svf::default());
    // no hiss above 5 kHz: air, not static
    let (mut ll, mut lr) = (Biquad::lowpass(SR, 5000.0, 0.7), Biquad::lowpass(SR, 5000.0, 0.7));
    let g = LEVELS[0] * db_to_gain(c.gain);
    let i0 = at(c.t, SR);
    for k in 0..at(c.dur, SR) {
        let p = k as f64 / (c.dur * SR);
        let shape = (PI * p).sin();
        let centre = 300.0 + 2600.0 * shape.powf(1.5);
        let amp = shape * shape * g;
        let pan = (c.pan - 0.4 + 0.8 * p).clamp(-1.0, 1.0);
        let l = ll.tick(fl.tick(noise.next() as f64, centre, 0.9, SR).bp as f32) as f64 * amp * (1.0 - pan).min(1.0);
        let r = lr.tick(fr.tick(noise.next() as f64, centre, 0.9, SR).bp as f32) as f64 * amp * (1.0 + pan).min(1.0);
        sfx.add(i0 + k, l as f32, r as f32);
        send.add(i0 + k, (l * 0.15) as f32, (r * 0.15) as f32);
    }
}

fn click(sfx: &mut Bus, c: &Cue, noise: &mut Noise) {
    let mut bp = Biquad::bandpass(SR, 3200.0, 2.5);
    let g = LEVELS[1] * db_to_gain(c.gain);
    let i0 = at(c.t, SR);
    for k in 0..at(0.04, SR) {
        let t = k as f64 / SR;
        let y = (bp.tick(noise.next()) as f64 * decay(t, 0.0012) + 0.5 * (TAU * 1900.0 * t).sin() * decay(t, 0.005)) * g;
        sfx.add_pan(i0 + k, y as f32, 0.0);
    }
}

fn tap(sfx: &mut Bus, c: &Cue, noise: &mut Noise) {
    let mut bp = Biquad::bandpass(SR, 2400.0, 1.5);
    let g = LEVELS[2] * db_to_gain(c.gain);
    let i0 = at(c.t, SR);
    let mut phase = 0.0;
    for k in 0..at(0.08, SR) {
        let t = k as f64 / SR;
        phase += (380.0 + 220.0 * decay(t, 0.01)) / SR;
        let y = ((TAU * phase).sin() * decay(t, 0.02) + 0.4 * bp.tick(noise.next()) as f64 * decay(t, 0.0015)) * g;
        sfx.add_pan(i0 + k, y as f32, 0.0);
    }
}

/// a soft mallet note at the cue's pitch, so the cards and chips play the chord
fn pop(sfx: &mut Bus, send: &mut Bus, c: &Cue) {
    let f = midi_hz(c.pitch);
    let g = LEVELS[3] * db_to_gain(c.gain);
    let i0 = at(c.t, SR);
    let mut phase = 0.0;
    for k in 0..at(0.5, SR) {
        let t = k as f64 / SR;
        phase += f * (1.0 + 0.04 * decay(t, 0.01)) / SR;
        let x = (TAU * phase).sin() + 0.25 * (TAU * 2.0 * phase).sin() + 0.08 * (TAU * 3.0 * phase).sin();
        let y = x * decay(t, 0.12) * (1.0 - decay(t, 0.002)) * g;
        sfx.add_pan(i0 + k, y as f32, c.pan);
        send.add(i0 + k, (y * 0.3) as f32, (y * 0.3) as f32);
    }
}

/// a low boom with a dark burst of noise
fn impact(sfx: &mut Bus, send: &mut Bus, c: &Cue, noise: &mut Noise) {
    let mut lp = Biquad::lowpass(SR, 500.0, 0.7);
    let g = LEVELS[4] * db_to_gain(c.gain);
    let i0 = at(c.t, SR);
    let mut phase = 0.0;
    for k in 0..at(1.8, SR) {
        let t = k as f64 / SR;
        phase += (40.0 + 55.0 * decay(t, 0.07)) / SR;
        let y = ((TAU * phase).sin() * decay(t, 0.55) + 0.5 * lp.tick(noise.next()) as f64 * decay(t, 0.1)) * (1.0 - decay(t, 0.002)) * g;
        sfx.add_pan(i0 + k, y as f32, 0.0);
        send.add(i0 + k, (y * 0.35) as f32, (y * 0.35) as f32);
    }
}

/// noise swept up from 180 Hz to 5,4 kHz, growing into the moment it leads to
fn riser(sfx: &mut Bus, send: &mut Bus, c: &Cue, noise: &mut Noise) {
    let (mut fl, mut fr) = (Svf::default(), Svf::default());
    let g = LEVELS[5] * db_to_gain(c.gain);
    let i0 = at(c.t, SR);
    let n = at(c.dur, SR);
    let mut phase = 0.0;
    for k in 0..n {
        let p = k as f64 / n as f64;
        let centre = 180.0 * 30f64.powf(p);
        let tail = ((n - k) as f64 / (0.025 * SR)).min(1.0);
        let amp = p.powf(2.2) * tail * g;
        phase += 220.0 * 4f64.powf(p) / SR;
        let tone = 0.15 * (TAU * phase).sin();
        let l = (fl.tick(noise.next() as f64, centre, 1.2, SR).bp + tone) * amp;
        let r = (fr.tick(noise.next() as f64, centre, 1.2, SR).bp + tone) * amp;
        sfx.add(i0 + k, l as f32, r as f32);
        send.add(i0 + k, (l * 0.2) as f32, (r * 0.2) as f32);
    }
}
