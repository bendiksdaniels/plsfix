//! core: pure rules for the video builder; no I/O, no clock.
//! frames = the timeline and worker chunks; viewport = a window size; shots = the capture
//! record's rules; probe_check, colour, copy_lint = the audit's rules;
//! dsp + sound = the soundtrack, synthesised from the picture's cues.
pub mod colour;
pub mod copy_lint;
pub mod dsp;
pub mod frames;
pub mod probe_check;
pub mod shots;
pub mod sound;
pub mod sync;
pub mod viewport;
