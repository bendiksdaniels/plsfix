//! osc.rs: oscillators: a band-limited sawtooth (polyBLEP, no aliasing whine), the MIDI-to-Hz
//! table and a seeded white noise, so every render of the soundtrack is identical.
//! Invariant: phases run in [0, 1); outputs stay within [-1, 1].

/// A sawtooth whose corners are smoothed by polyBLEP.
pub struct Saw {
    phase: f64,
}

impl Saw {
    pub fn new(phase: f64) -> Saw {
        Saw { phase: phase.rem_euclid(1.0) }
    }

    /// One sample; `inc` = frequency / sample rate.
    pub fn next(&mut self, inc: f64) -> f32 {
        let t = self.phase;
        let y = 2.0 * t - 1.0 - poly_blep(t, inc);
        self.phase += inc;
        if self.phase >= 1.0 {
            self.phase -= 1.0;
        }
        y as f32
    }
}

fn poly_blep(t: f64, dt: f64) -> f64 {
    if t < dt {
        let x = t / dt;
        2.0 * x - x * x - 1.0
    } else if t > 1.0 - dt {
        let x = (t - 1.0) / dt;
        x * x + 2.0 * x + 1.0
    } else {
        0.0
    }
}

/// The frequency of a MIDI note (69 = A4 = 440 Hz; fractional notes allowed).
pub fn midi_hz(m: f64) -> f64 {
    440.0 * 2f64.powf((m - 69.0) / 12.0)
}

/// Seeded white noise (xorshift64*), uniform in [-1, 1).
pub struct Noise(u64);

impl Noise {
    pub fn new(seed: u64) -> Noise {
        Noise(seed.max(1).wrapping_mul(0x9E37_79B9_7F4A_7C15) | 1)
    }

    pub fn next(&mut self) -> f32 {
        let mut x = self.0;
        x ^= x >> 12;
        x ^= x << 25;
        x ^= x >> 27;
        self.0 = x;
        let v = x.wrapping_mul(0x2545_F491_4F6C_DD1D) >> 40;
        (v as f64 / (1u64 << 23) as f64 - 1.0) as f32
    }
}
