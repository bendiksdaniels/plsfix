//! wav.rs: writes the stereo master as a 32-bit float WAV (IEEE float, with the fact chunk the
//! format calls for), the file ffmpeg measures and muxes. Invariant: left and right same length.

use crate::core::sound::bus::Bus;
use anyhow::{bail, Context, Result};
use std::io::Write;
use std::path::Path;

pub fn write(path: &Path, bus: &Bus, sr: u32) -> Result<()> {
    if bus.l.len() != bus.r.len() {
        bail!("wav {}: channels differ in length", path.display());
    }
    let frames = bus.l.len() as u32;
    let data = frames * 8;
    let mut b: Vec<u8> = Vec::with_capacity(data as usize + 58);
    b.extend_from_slice(b"RIFF");
    b.extend_from_slice(&(50 + data).to_le_bytes());
    b.extend_from_slice(b"WAVEfmt ");
    b.extend_from_slice(&18u32.to_le_bytes());
    b.extend_from_slice(&3u16.to_le_bytes()); // IEEE float
    b.extend_from_slice(&2u16.to_le_bytes());
    b.extend_from_slice(&sr.to_le_bytes());
    b.extend_from_slice(&(sr * 8).to_le_bytes());
    b.extend_from_slice(&8u16.to_le_bytes());
    b.extend_from_slice(&32u16.to_le_bytes());
    b.extend_from_slice(&0u16.to_le_bytes());
    b.extend_from_slice(b"fact");
    b.extend_from_slice(&4u32.to_le_bytes());
    b.extend_from_slice(&frames.to_le_bytes());
    b.extend_from_slice(b"data");
    b.extend_from_slice(&data.to_le_bytes());
    for (l, r) in bus.l.iter().zip(bus.r.iter()) {
        b.extend_from_slice(&l.to_le_bytes());
        b.extend_from_slice(&r.to_le_bytes());
    }
    if let Some(dir) = path.parent() {
        std::fs::create_dir_all(dir)?;
    }
    std::fs::File::create(path).and_then(|mut f| f.write_all(&b)).with_context(|| format!("wav: write {}", path.display()))
}
