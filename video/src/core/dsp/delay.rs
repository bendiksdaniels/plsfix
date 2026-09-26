//! delay.rs: a ping-pong echo (left then right), each repeat darker than the last, for the
//! plucked melody. Invariant: feedback below 1, so the echoes fade out.

pub struct PingPong {
    l: Vec<f32>,
    r: Vec<f32>,
    i: usize,
    feedback: f32,
    tone: f32,
    lp: [f32; 2],
}

impl PingPong {
    /// `seconds` between echoes, `feedback` 0..0.9, `tone` 0..1 (1 = bright).
    pub fn new(sr: f64, seconds: f64, feedback: f32, tone: f32) -> PingPong {
        let n = ((seconds * sr).round() as usize).max(1);
        PingPong { l: vec![0.0; n], r: vec![0.0; n], i: 0, feedback: feedback.min(0.9), tone, lp: [0.0; 2] }
    }

    /// A mono sample in, the echoes out (dry not included).
    pub fn tick(&mut self, x: f32) -> (f32, f32) {
        let (dl, dr) = (self.l[self.i], self.r[self.i]);
        self.lp[0] += (dl - self.lp[0]) * self.tone;
        self.lp[1] += (dr - self.lp[1]) * self.tone;
        self.l[self.i] = x + self.lp[1] * self.feedback;
        self.r[self.i] = self.lp[0] * self.feedback;
        self.i = (self.i + 1) % self.l.len();
        (dl, dr)
    }
}
