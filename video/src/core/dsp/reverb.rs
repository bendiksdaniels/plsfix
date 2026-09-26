//! reverb.rs: a stereo Freeverb (eight damped combs into four allpasses per side, the right side
//! detuned by 23 samples), the room the pad, plucks and hits sit in.
//! Invariant: the comb feedback stays below 1, so the tail always dies away.

struct Comb {
    buf: Vec<f32>,
    i: usize,
    store: f32,
}

impl Comb {
    fn tick(&mut self, x: f32, feedback: f32, damp: f32) -> f32 {
        let out = self.buf[self.i];
        self.store = out * (1.0 - damp) + self.store * damp;
        self.buf[self.i] = x + self.store * feedback;
        self.i = (self.i + 1) % self.buf.len();
        out
    }
}

struct Allpass {
    buf: Vec<f32>,
    i: usize,
}

impl Allpass {
    fn tick(&mut self, x: f32) -> f32 {
        let b = self.buf[self.i];
        self.buf[self.i] = x + b * 0.5;
        self.i = (self.i + 1) % self.buf.len();
        b - x
    }
}

const COMBS: [usize; 8] = [1116, 1188, 1277, 1356, 1422, 1491, 1557, 1617];
const ALLPASSES: [usize; 4] = [556, 441, 341, 225];
const SPREAD: usize = 23;

pub struct Reverb {
    combs: [Vec<Comb>; 2],
    aps: [Vec<Allpass>; 2],
    feedback: f32,
    damp: f32,
}

impl Reverb {
    /// `room` 0..1 (size), `damp` 0..1 (how fast the highs die).
    pub fn new(sr: f64, room: f32, damp: f32) -> Reverb {
        let scale = |n: usize, side: usize| ((n + side * SPREAD) as f64 * sr / 44_100.0).round() as usize;
        let combs = [0, 1].map(|side| COMBS.iter().map(|&n| Comb { buf: vec![0.0; scale(n, side)], i: 0, store: 0.0 }).collect());
        let aps = [0, 1].map(|side| ALLPASSES.iter().map(|&n| Allpass { buf: vec![0.0; scale(n, side)], i: 0 }).collect());
        Reverb { combs, aps, feedback: room * 0.28 + 0.7, damp: damp * 0.4 }
    }

    /// One stereo sample in, the wet stereo sample out.
    pub fn tick(&mut self, l: f32, r: f32) -> (f32, f32) {
        let input = (l + r) * 0.015;
        let mut out = [0.0f32; 2];
        for side in 0..2 {
            let mut acc = 0.0;
            for c in self.combs[side].iter_mut() {
                acc += c.tick(input, self.feedback, self.damp);
            }
            for a in self.aps[side].iter_mut() {
                acc = a.tick(acc);
            }
            out[side] = acc * 3.0;
        }
        (out[0], out[1])
    }
}
