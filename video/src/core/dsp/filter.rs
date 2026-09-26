//! filter.rs: a biquad (RBJ cookbook low-, high- and band-pass) for fixed filters, and a
//! topology-preserving state-variable filter for swept ones (cutoff may change every sample).
//! Invariant: stable for cutoff within (10 Hz, 0.45 x sample rate) and q above 0.3.

use std::f64::consts::PI;

pub struct Biquad {
    b0: f64,
    b1: f64,
    b2: f64,
    a1: f64,
    a2: f64,
    z1: f64,
    z2: f64,
}

impl Biquad {
    fn from(b0: f64, b1: f64, b2: f64, a0: f64, a1: f64, a2: f64) -> Biquad {
        Biquad { b0: b0 / a0, b1: b1 / a0, b2: b2 / a0, a1: a1 / a0, a2: a2 / a0, z1: 0.0, z2: 0.0 }
    }

    fn parts(sr: f64, f: f64, q: f64) -> (f64, f64) {
        let w = 2.0 * PI * f.clamp(10.0, 0.45 * sr) / sr;
        (w.cos(), w.sin() / (2.0 * q.max(0.3)))
    }

    pub fn lowpass(sr: f64, f: f64, q: f64) -> Biquad {
        let (c, alpha) = Self::parts(sr, f, q);
        Self::from((1.0 - c) / 2.0, 1.0 - c, (1.0 - c) / 2.0, 1.0 + alpha, -2.0 * c, 1.0 - alpha)
    }

    pub fn highpass(sr: f64, f: f64, q: f64) -> Biquad {
        let (c, alpha) = Self::parts(sr, f, q);
        Self::from((1.0 + c) / 2.0, -(1.0 + c), (1.0 + c) / 2.0, 1.0 + alpha, -2.0 * c, 1.0 - alpha)
    }

    /// Band-pass with 0 dB at the centre.
    pub fn bandpass(sr: f64, f: f64, q: f64) -> Biquad {
        let (c, alpha) = Self::parts(sr, f, q);
        Self::from(alpha, 0.0, -alpha, 1.0 + alpha, -2.0 * c, 1.0 - alpha)
    }

    pub fn tick(&mut self, x: f32) -> f32 {
        let x = x as f64;
        let y = self.b0 * x + self.z1;
        self.z1 = self.b1 * x - self.a1 * y + self.z2;
        self.z2 = self.b2 * x - self.a2 * y;
        y as f32
    }
}

/// Outputs of one state-variable filter step.
pub struct Svf3 {
    pub lp: f64,
    pub bp: f64,
    pub hp: f64,
}

/// Andrew Simper's trapezoidal SVF.
#[derive(Default)]
pub struct Svf {
    ic1: f64,
    ic2: f64,
}

impl Svf {
    pub fn tick(&mut self, x: f64, cutoff: f64, q: f64, sr: f64) -> Svf3 {
        let g = (PI * cutoff.clamp(10.0, 0.45 * sr) / sr).tan();
        let k = 1.0 / q.max(0.3);
        let a1 = 1.0 / (1.0 + g * (g + k));
        let a2 = g * a1;
        let a3 = g * a2;
        let v3 = x - self.ic2;
        let v1 = a1 * self.ic1 + a2 * v3;
        let v2 = self.ic2 + a2 * self.ic1 + a3 * v3;
        self.ic1 = 2.0 * v1 - self.ic1;
        self.ic2 = 2.0 * v2 - self.ic2;
        Svf3 { lp: v2, bp: v1, hp: x - k * v1 - v2 }
    }
}
