//! viewport.rs: the size of a browser window used for a capture or a render.
//! CSS pixels times `dpr` gives the screenshot's pixel size; `mobile` turns on touch layout.

/// A browser window: `w` x `h` CSS pixels at `dpr` device pixels per CSS pixel.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Viewport {
    pub w: u32,
    pub h: u32,
    pub dpr: f64,
    pub mobile: bool,
}

impl Viewport {
    /// The stage the composition is authored at: 1920 x 1080 CSS pixels.
    pub const STAGE: Viewport = Viewport { w: 1920, h: 1080, dpr: 1.0, mobile: false };

    /// The screenshot's size in pixels.
    pub fn pixels(&self) -> (u32, u32) {
        ((self.w as f64 * self.dpr).round() as u32, (self.h as f64 * self.dpr).round() as u32)
    }
}
