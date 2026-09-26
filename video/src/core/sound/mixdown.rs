//! mixdown.rs: plays every note and cue into its bus, ducks the music under the kick (the
//! gentle "pump" of modern tracks), adds the room (reverb) and the echo, and fades in and out
//! with the picture. Output is the raw stereo master; loudness is set afterwards.

use crate::core::dsp::delay::PingPong;
use crate::core::dsp::env::{decay, ramp};
use crate::core::dsp::filter::Biquad;
use crate::core::dsp::osc::Noise;
use crate::core::dsp::reverb::Reverb;
use crate::core::dsp::SR;
use crate::core::sound::bus::{at, Bus};
use crate::core::sound::cues::Cue;
use crate::core::sound::instruments as band;
use crate::core::sound::score::{Inst, Note, BEAT};
use crate::core::sound::sfx;

/// How deep the music ducks under each kick (0..1).
const DUCK: f32 = 0.32;
const REVERB_WET: f32 = 0.9;
const ECHO_WET: f32 = 0.5;
/// The master's high shelf: this much of everything above ~2,5 kHz is added back (+4 dB of air).
const AIR: f32 = 0.6;

/// The master for a video of `seconds`: notes (the music) plus cues (the sound design).
pub fn render(notes: &[Note], cues: &[Cue], seconds: f64) -> Bus {
    let n = at(seconds, SR);
    let (mut music, mut drums, mut fx, mut send) = (Bus::new(n), Bus::new(n), Bus::new(n), Bus::new(n));
    let mut echo = vec![0.0f32; n];
    let mut noise = Noise::new(20260925);
    for note in notes {
        match note.inst {
            Inst::Pad => band::pad(&mut music, &mut send, note),
            Inst::Bass => band::bass(&mut music, note),
            Inst::Kick => band::kick(&mut drums, note, &mut noise),
            Inst::Hat => band::hat(&mut drums, note, &mut noise),
            Inst::Clap => band::clap(&mut drums, &mut send, note, &mut noise),
            Inst::Pluck => band::pluck(&mut music, &mut send, &mut echo, note),
        }
    }
    for c in cues {
        sfx::render(&mut fx, &mut send, c, &mut noise);
    }
    duck(&mut music, notes);
    let mut out = Bus::new(n);
    let mut room = Reverb::new(SR, 0.82, 0.45);
    let (mut hl, mut hr) = (Biquad::highpass(SR, 250.0, 0.7), Biquad::highpass(SR, 250.0, 0.7));
    let mut delay = PingPong::new(SR, BEAT * 0.75, 0.38, 0.35);
    for i in 0..n {
        let (el, er) = delay.tick(echo[i]);
        let (wl, wr) = room.tick(hl.tick(send.l[i] + el * 0.3), hr.tick(send.r[i] + er * 0.3));
        out.l[i] = music.l[i] + drums.l[i] + fx.l[i] + el * ECHO_WET + wl * REVERB_WET;
        out.r[i] = music.r[i] + drums.r[i] + fx.r[i] + er * ECHO_WET + wr * REVERB_WET;
    }
    air(&mut out);
    fade(&mut out, seconds);
    out
}

/// A gentle high shelf: the top end a laptop speaker needs to sound open, not muffled.
fn air(out: &mut Bus) {
    let (mut hl, mut hr) = (Biquad::highpass(SR, 2500.0, 0.7), Biquad::highpass(SR, 2500.0, 0.7));
    for i in 0..out.len() {
        out.l[i] += AIR * hl.tick(out.l[i]);
        out.r[i] += AIR * hr.tick(out.r[i]);
    }
}

/// The music bus dips under every kick and comes back over ~0.15 s.
fn duck(music: &mut Bus, notes: &[Note]) {
    let mut env = vec![0.0f32; music.len()];
    for k in notes.iter().filter(|x| x.inst == Inst::Kick) {
        let i0 = at(k.t, SR);
        for j in 0..at(0.4, SR) {
            if let Some(e) = env.get_mut(i0 + j) {
                *e = e.max((decay(j as f64 / SR, 0.1) * k.vel) as f32);
            }
        }
    }
    for (i, e) in env.iter().enumerate() {
        let g = 1.0 - DUCK * e.min(1.0);
        music.l[i] *= g;
        music.r[i] *= g;
    }
}

/// In over the first 0,35 s with the picture's opening fade, out over its closing one.
fn fade(out: &mut Bus, seconds: f64) {
    for i in 0..out.len() {
        let t = i as f64 / SR;
        let g = (ramp(t, 0.0, 0.35) * (1.0 - ramp(t, seconds - 0.65, seconds - 0.05))) as f32;
        out.l[i] *= g;
        out.r[i] *= g;
    }
}
