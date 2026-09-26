//! instruments.rs: the band. A warm pad (three detuned saws through a low-pass that opens with
//! the note), a round sub bass, a soft kick, hats, a clap and a plucked melody, each rendering
//! one note into the buses. Levels live in LEVELS; the master is normalised afterwards.

use crate::core::dsp::env::{adsr, decay};
use crate::core::dsp::filter::{Biquad, Svf};
use crate::core::dsp::osc::{midi_hz, Noise, Saw};
use crate::core::dsp::SR;
use crate::core::sound::bus::{at, pan_gains, Bus};
use crate::core::sound::score::Note;
use std::f64::consts::TAU;

/// Relative levels, one place: pad (per voice), bass, kick, hat, clap, pluck.
pub const LEVELS: [f32; 6] = [0.05, 0.12, 0.3, 0.19, 0.2, 0.17];
const PAD_RELEASE: f64 = 1.1;

pub fn pad(music: &mut Bus, send: &mut Bus, n: &Note) {
    let f = midi_hz(n.midi);
    let mut saws = [Saw::new(0.13 * n.midi), Saw::new(0.57), Saw::new(0.91 - 0.01 * n.midi)];
    let inc = [-0.07, 0.0, 0.07].map(|d: f64| f * 2f64.powf(d / 12.0) / SR);
    let (mut fl, mut fr) = (Svf::default(), Svf::default());
    // below 180 Hz belongs to the bass and the kick; the pad only muddies it
    let (mut hl, mut hr) = (Biquad::highpass(SR, 180.0, 0.7), Biquad::highpass(SR, 180.0, 0.7));
    let (gl, gr) = pan_gains(n.pan);
    let i0 = at(n.t, SR);
    for k in 0..at(n.len + PAD_RELEASE, SR) {
        let t = k as f64 / SR;
        let env = adsr(t, n.len, 0.45, 0.4, 0.8, PAD_RELEASE);
        let (s0, s1, s2) = (saws[0].next(inc[0]) as f64, saws[1].next(inc[1]) as f64, saws[2].next(inc[2]) as f64);
        let cutoff = n.bright * (0.65 + 0.35 * env);
        let l = hl.tick(fl.tick(s0 * 0.6 + s1 * 0.5 + s2 * 0.2, cutoff, 0.7, SR).lp as f32) as f64;
        let r = hr.tick(fr.tick(s2 * 0.6 + s1 * 0.5 + s0 * 0.2, cutoff, 0.7, SR).lp as f32) as f64;
        let g = (env * n.vel) as f32 * LEVELS[0];
        music.add(i0 + k, l as f32 * g * gl * 1.4, r as f32 * g * gr * 1.4);
        send.add(i0 + k, l as f32 * g * 0.35, r as f32 * g * 0.35);
    }
}

pub fn bass(music: &mut Bus, n: &Note) {
    let f = midi_hz(n.midi);
    let i0 = at(n.t, SR);
    for k in 0..at(n.len + 0.08, SR) {
        let t = k as f64 / SR;
        let env = adsr(t, n.len, 0.004, 0.08, 0.7, 0.06);
        let x = (TAU * f * t).sin() + 0.3 * (TAU * 2.0 * f * t).sin() + 0.12 * (TAU * 3.0 * f * t).sin();
        let y = (1.4 * x).tanh() * env * n.vel;
        music.add_pan(i0 + k, y as f32 * LEVELS[1], 0.0);
    }
}

pub fn kick(drums: &mut Bus, n: &Note, noise: &mut Noise) {
    let i0 = at(n.t, SR);
    let mut phase = 0.0;
    for k in 0..at(0.5, SR) {
        let t = k as f64 / SR;
        phase += (50.0 + 110.0 * decay(t, 0.03)) / SR;
        let body = (TAU * phase).sin() * decay(t, 0.17) * (1.0 - decay(t, 0.001));
        let click = noise.next() as f64 * decay(t, 0.0015) * 0.4;
        let y = (1.3 * (body + click)).tanh() * n.vel;
        drums.add_pan(i0 + k, y as f32 * LEVELS[2], 0.0);
    }
}

pub fn hat(drums: &mut Bus, n: &Note, noise: &mut Noise) {
    let mut hp = Biquad::highpass(SR, 7500.0, 0.7);
    let mut lp = Biquad::lowpass(SR, 11_000.0, 0.7);   // shimmer, not fizz
    let i0 = at(n.t, SR);
    for k in 0..at(0.15, SR) {
        let t = k as f64 / SR;
        let y = lp.tick(hp.tick(noise.next())) as f64 * decay(t, 0.035) * n.vel;
        drums.add_pan(i0 + k, y as f32 * LEVELS[3], n.pan);
    }
}

pub fn clap(drums: &mut Bus, send: &mut Bus, n: &Note, noise: &mut Noise) {
    let (mut bl, mut br) = (Biquad::bandpass(SR, 1300.0, 1.0), Biquad::bandpass(SR, 1350.0, 1.0));
    let i0 = at(n.t, SR);
    for k in 0..at(0.4, SR) {
        let t = k as f64 / SR;
        let env = [0.0, 0.011, 0.022].iter().map(|d| decay(t - d, 0.006)).sum::<f64>() + 0.6 * decay(t - 0.022, 0.13);
        let g = (env * n.vel) as f32 * LEVELS[4];
        let (l, r) = (bl.tick(noise.next()) * g, br.tick(noise.next()) * g);
        drums.add(i0 + k, l, r);
        send.add(i0 + k, l * 0.25, r * 0.25);
    }
}

pub fn pluck(music: &mut Bus, send: &mut Bus, echo: &mut [f32], n: &Note) {
    let f = midi_hz(n.midi);
    let mut saws = [Saw::new(0.3), Saw::new(0.71)];
    let mut lp = Svf::default();
    let (gl, gr) = pan_gains(n.pan);
    let i0 = at(n.t, SR);
    for k in 0..at(0.9, SR) {
        let t = k as f64 / SR;
        let x = saws[0].next(f / SR) as f64 + 0.5 * saws[1].next(f * 1.007 / SR) as f64;
        let y = lp.tick(x, 900.0 + 5200.0 * decay(t, 0.09), 0.8, SR).lp * decay(t, 0.28) * (1.0 - decay(t, 0.002)) * n.vel;
        let y = y as f32 * LEVELS[5];
        music.add(i0 + k, y * gl, y * gr);
        send.add(i0 + k, y * 0.25, y * 0.25);
        if let Some(e) = echo.get_mut(i0 + k) {
            *e += y * 0.35;
        }
    }
}
