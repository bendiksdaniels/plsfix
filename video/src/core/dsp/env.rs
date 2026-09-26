//! env.rs: amplitude envelopes as pure functions of time since a note began.
//! Invariant: every envelope is 0 before its start and never exceeds 1.

/// Attack, decay to sustain while the gate is open (`gate` s), then release.
pub fn adsr(t: f64, gate: f64, a: f64, d: f64, s: f64, r: f64) -> f64 {
    if t < 0.0 {
        return 0.0;
    }
    let held = |t: f64| {
        if t < a {
            t / a
        } else if t < a + d {
            1.0 - (1.0 - s) * (t - a) / d
        } else {
            s
        }
    };
    if t < gate {
        held(t)
    } else {
        let at = held(gate);
        (at * (1.0 - (t - gate) / r)).max(0.0)
    }
}

/// An exponential decay with time constant `tau` from a note's start.
pub fn decay(t: f64, tau: f64) -> f64 {
    if t < 0.0 {
        0.0
    } else {
        (-t / tau).exp()
    }
}

/// 0 before `a`, 1 after `b`, a smooth (cosine) ramp between.
pub fn ramp(t: f64, a: f64, b: f64) -> f64 {
    if t <= a {
        0.0
    } else if t >= b {
        1.0
    } else {
        0.5 - 0.5 * (std::f64::consts::PI * (t - a) / (b - a)).cos()
    }
}
