// Screenshots: reads a PNG from manual/shots and works out how big it may be
// on an A4 page. Owns the picture size rule (at most 15 cm wide and 11 cm
// tall, aspect kept); the invariant is that a missing file fails loudly.

use crate::style::{EMU_PER_CM, MAX_IMAGE_H_CM, MAX_IMAGE_W_CM};

pub struct Shot {
    pub bytes: Vec<u8>,
    /// Width and height in EMU, already scaled to fit the page.
    pub size: (u32, u32),
}

/// The folder the screenshots live in, beside this crate's Cargo.toml.
pub fn folder() -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("shots")
}

pub fn load(file: &str) -> Shot {
    let path = folder().join(file);
    let bytes =
        std::fs::read(&path).unwrap_or_else(|e| panic!("read screenshot {}: {e}", path.display()));
    let (w, h) = png_size(&bytes)
        .unwrap_or_else(|| panic!("not a PNG with a readable header: {}", path.display()));
    Shot {
        size: scaled(w, h),
        bytes,
    }
}

/// Width and height in pixels from a PNG's IHDR chunk.
fn png_size(bytes: &[u8]) -> Option<(u32, u32)> {
    const SIGNATURE: [u8; 8] = [137, 80, 78, 71, 13, 10, 26, 10];
    if bytes.len() < 24 || bytes[..8] != SIGNATURE {
        return None;
    }
    let width = u32::from_be_bytes(bytes[16..20].try_into().ok()?);
    let height = u32::from_be_bytes(bytes[20..24].try_into().ok()?);
    if width == 0 || height == 0 {
        return None;
    }
    Some((width, height))
}

/// The picture's size in EMU: as large as both caps allow, aspect kept.
fn scaled(width_px: u32, height_px: u32) -> (u32, u32) {
    let aspect = f64::from(height_px) / f64::from(width_px);
    let mut w_cm = MAX_IMAGE_W_CM;
    if w_cm * aspect > MAX_IMAGE_H_CM {
        w_cm = MAX_IMAGE_H_CM / aspect;
    }
    let h_cm = w_cm * aspect;
    (
        (w_cm * EMU_PER_CM).round() as u32,
        (h_cm * EMU_PER_CM).round() as u32,
    )
}

/// A smaller cap for the cover logo, which is a square icon.
pub fn cover_size(shot: &Shot) -> (u32, u32) {
    let (w, h) = shot.size;
    let target = (2.6 * EMU_PER_CM).round() as u32;
    if w == 0 {
        return (target, target);
    }
    let height = (u64::from(h) * u64::from(target) / u64::from(w)) as u32;
    (target, height)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn wide_pictures_stop_at_the_width_cap() {
        let (w, h) = scaled(2880, 1434);
        assert_eq!(w, (MAX_IMAGE_W_CM * EMU_PER_CM) as u32);
        assert!(h < (MAX_IMAGE_H_CM * EMU_PER_CM) as u32);
    }

    #[test]
    fn tall_pictures_stop_at_the_height_cap() {
        let (w, h) = scaled(698, 1076);
        assert_eq!(h, (MAX_IMAGE_H_CM * EMU_PER_CM).round() as u32);
        assert!(w < (MAX_IMAGE_W_CM * EMU_PER_CM) as u32);
    }
}
