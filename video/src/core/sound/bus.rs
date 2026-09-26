//! bus.rs: a stereo buffer the voices add into, and the equal-power pan law they share.
//! Invariant: writes past the end are dropped (a tail that runs past the video is cut, never grown).

pub struct Bus {
    pub l: Vec<f32>,
    pub r: Vec<f32>,
}

impl Bus {
    pub fn new(n: usize) -> Bus {
        Bus { l: vec![0.0; n], r: vec![0.0; n] }
    }

    pub fn len(&self) -> usize {
        self.l.len()
    }

    pub fn is_empty(&self) -> bool {
        self.l.is_empty()
    }

    pub fn add(&mut self, i: usize, l: f32, r: f32) {
        if i < self.l.len() {
            self.l[i] += l;
            self.r[i] += r;
        }
    }

    /// Adds a mono sample at `pan` (-1 left .. 1 right).
    pub fn add_pan(&mut self, i: usize, x: f32, pan: f64) {
        let (gl, gr) = pan_gains(pan);
        self.add(i, x * gl, x * gr);
    }
}

/// Equal-power pan: both sides at -3 dB in the middle.
pub fn pan_gains(pan: f64) -> (f32, f32) {
    let a = (pan.clamp(-1.0, 1.0) + 1.0) * std::f64::consts::FRAC_PI_4;
    (a.cos() as f32, a.sin() as f32)
}

/// The sample index of a time in seconds.
pub fn at(t: f64, sr: f64) -> usize {
    (t.max(0.0) * sr).round() as usize
}
