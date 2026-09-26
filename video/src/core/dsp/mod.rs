//! dsp: the soundtrack's building blocks, pure and sample-accurate: oscillators, envelopes,
//! filters, a reverb, a delay and a limiter. Nothing here reads a clock or touches a file.
pub mod delay;
pub mod dynamics;
pub mod env;
pub mod filter;
pub mod osc;
pub mod reverb;

/// The soundtrack's sample rate (the AAC track is 48 kHz too).
pub const SR: f64 = 48_000.0;
