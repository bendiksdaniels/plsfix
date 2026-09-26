//! comp_page.rs: the composition page (video/comp/index.html) open in one headless Chrome.
//! Owns the page contract: `__ready` resolves to {duration, fps} once every image and font is in;
//! `__seek(t)` sets the frame for time t. Any page error seen here fails with the time it showed.

use crate::core::viewport::Viewport;
use crate::io::cdp::Cdp;
use crate::io::chrome::Chrome;
use anyhow::{bail, Context, Result};
use serde::Deserialize;
use std::path::Path;

/// What the page reports once ready.
#[derive(Debug, Clone, Copy, Deserialize)]
pub struct Meta {
    pub duration: f64,
    pub fps: u32,
}

pub struct CompPage {
    _chrome: Chrome,
    pub cdp: Cdp,
    pub meta: Meta,
}

impl CompPage {
    /// Opens the composition served on `port` in a fresh Chrome of size `vp` and waits for it.
    pub fn open(port: u16, vp: Viewport) -> Result<CompPage> {
        let (chrome, mut cdp) = Chrome::launch(vp)?;
        cdp.navigate(&format!("http://127.0.0.1:{port}/comp/index.html"))?;
        let meta = cdp.eval_async("window.__ready").context("comp page: __ready")?;
        let errors = cdp.take_errors();
        if !errors.is_empty() {
            bail!("comp page: errors while loading: {}", errors.join(" | "));
        }
        let meta: Meta = serde_json::from_value(meta).context("comp page: __ready did not return {duration, fps}")?;
        Ok(CompPage { _chrome: chrome, cdp, meta })
    }

    /// The PNG of the frame at time `t` seconds.
    pub fn frame_png(&mut self, t: f64) -> Result<Vec<u8>> {
        self.cdp.eval(&format!("window.__seek({t})")).with_context(|| format!("comp page: seek {t:.3}"))?;
        let png = self.cdp.screenshot_png().with_context(|| format!("comp page: frame at {t:.3}s"))?;
        let errors = self.cdp.take_errors();
        if !errors.is_empty() {
            bail!("comp page: errors at {t:.3}s: {}", errors.join(" | "));
        }
        Ok(png)
    }
}

/// Writes `bytes` to `path`, creating its folder.
pub fn write_file(path: &Path, bytes: &[u8]) -> Result<()> {
    if let Some(dir) = path.parent() {
        std::fs::create_dir_all(dir).with_context(|| format!("mkdir {}", dir.display()))?;
    }
    std::fs::write(path, bytes).with_context(|| format!("write {}", path.display()))
}
