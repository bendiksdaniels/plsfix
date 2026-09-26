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
use crate::core::sound::score::{Inst, Note, Style, BEAT};
use crate::core::sound::sfx;

/// How each music is mixed: `duck` = how deep the music dips under each kick (0..1), the reverb
/// and echo returns, `air` = how much of everything above ~2,5 kHz is added back (0,6 = +4 dB),
/// `fx` = the sound design's gain. The bed sits under a voice: no pump, little air, UI sounds
/// 6 dB down.
struct Mix {
    duck: f32,
    reverb_wet: f32,
    echo_wet: f32,
    air: f32,
    fx: f32,
}

fn mix_of(style: Style) -> Mix {
    match style {
        Style::Groove => Mix { duck: 0.32, reverb_wet: 0.9, echo_wet: 0.5, air: 0.6, fx: 1.0 },
        Style::Bed => Mix { duck: 0.0, reverb_wet: 0.8, echo_wet: 0.3, air: 0.15, fx: 0.5 },
    }
}

/// The master for a video of `seconds`: notes (the music) plus cues (the sound design).
pub fn render(notes: &[Note], cues: &[Cue], seconds: f64, style: Style) -> Bus {
    let mix = mix_of(style);
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
            Inst::Keys => band::keys(&mut music, &mut send, &mut echo, note),
        }
    }
    for c in cues {
        sfx::render(&mut fx, &mut send, c, &mut noise);
    }
    duck(&mut music, notes, mix.duck);
    let mut out = Bus::new(n);
    let mut room = Reverb::new(SR, 0.82, 0.45);
    let (mut hl, mut hr) = (Biquad::highpass(SR, 250.0, 0.7), Biquad::highpass(SR, 250.0, 0.7));
    let mut delay = PingPong::new(SR, BEAT * 0.75, 0.38, 0.35);
    for i in 0..n {
        let (el, er) = delay.tick(echo[i]);
        let (wl, wr) = room.tick(hl.tick(send.l[i] + el * 0.3), hr.tick(send.r[i] + er * 0.3));
        out.l[i] = music.l[i] + drums.l[i] + fx.l[i] * mix.fx + el * mix.echo_wet + wl * mix.reverb_wet;
        out.r[i] = music.r[i] + drums.r[i] + fx.r[i] * mix.fx + er * mix.echo_wet + wr * mix.reverb_wet;
    }
    air(&mut out, mix.air);
    fade(&mut out, seconds);
    out
}

/// A gentle high shelf: the top end a laptop speaker needs to sound open, not muffled.
fn air(out: &mut Bus, amount: f32) {
    let (mut hl, mut hr) = (Biquad::highpass(SR, 2500.0, 0.7), Biquad::highpass(SR, 2500.0, 0.7));
    for i in 0..out.len() {
        out.l[i] += amount * hl.tick(out.l[i]);
        out.r[i] += amount * hr.tick(out.r[i]);
    }
}

/// The music bus dips under every kick and comes back over ~0.15 s.
fn duck(music: &mut Bus, notes: &[Note], depth: f32) {
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
        let g = 1.0 - depth * e.min(1.0);
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
