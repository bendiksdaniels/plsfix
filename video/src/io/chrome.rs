//! chrome.rs: one headless Chrome process with its own profile and debugging port.
//! Owns launch, the page target lookup and the window size; Drop kills the process and deletes
//! the profile, so no Chrome outlives the command that started it.

use crate::core::viewport::Viewport;
use crate::io::cdp::Cdp;
use crate::io::http;
use anyhow::{bail, Context, Result};
use serde_json::{json, Value};
use std::net::TcpListener;
use std::path::PathBuf;
use std::process::{Child, Command, Stdio};
use std::thread::sleep;
use std::time::{Duration, Instant};

pub const CHROME: &str = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";

pub struct Chrome {
    child: Child,
    profile: PathBuf,
}

impl Chrome {
    /// Starts headless Chrome, connects to its blank page and sizes the window to `vp`.
    pub fn launch(vp: Viewport) -> Result<(Chrome, Cdp)> {
        let port = free_port()?;
        let profile = std::env::temp_dir().join(format!("plsfix-video-{}-{port}", std::process::id()));
        let child = Command::new(CHROME)
            .args(["--headless=new", &format!("--remote-debugging-port={port}"), &format!("--user-data-dir={}", profile.display())])
            .args(["--no-first-run", "--no-default-browser-check", "--disable-extensions", "--disable-background-networking"])
            .args(["--disable-sync", "--hide-scrollbars", "--mute-audio", "--force-color-profile=srgb"])
            .args(["--disable-renderer-backgrounding", "--disable-background-timer-throttling"])
            .arg("about:blank")
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .spawn()
            .context("chrome launch: spawn")?;
        let chrome = Chrome { child, profile };
        let mut cdp = Cdp::connect(&page_ws_url(port)?)?;
        for m in ["Page.enable", "Runtime.enable", "Log.enable"] {
            cdp.call(m, json!({}))?;
        }
        set_viewport(&mut cdp, vp)?;
        Ok((chrome, cdp))
    }
}

impl Drop for Chrome {
    fn drop(&mut self) {
        let _ = self.child.kill();
        let _ = self.child.wait();
        let _ = std::fs::remove_dir_all(&self.profile);
    }
}

/// Sizes the page's window: CSS pixels, device pixel ratio and touch layout.
pub fn set_viewport(cdp: &mut Cdp, vp: Viewport) -> Result<()> {
    cdp.call("Emulation.setDeviceMetricsOverride", json!({
        "width": vp.w, "height": vp.h, "deviceScaleFactor": vp.dpr, "mobile": vp.mobile,
    }))?;
    cdp.call("Emulation.setTouchEmulationEnabled", json!({ "enabled": vp.mobile }))?;
    Ok(())
}

/// A port nothing listens on right now (bound, read, released).
pub fn free_port() -> Result<u16> {
    let l = TcpListener::bind("127.0.0.1:0").context("free port")?;
    Ok(l.local_addr()?.port())
}

/// Waits for Chrome's DevTools endpoint and returns the first page's WebSocket URL.
fn page_ws_url(port: u16) -> Result<String> {
    let start = Instant::now();
    loop {
        if let Ok((200, body)) = http::get(port, "/json/list") {
            let targets: Value = serde_json::from_str(&body).unwrap_or(Value::Null);
            // new headless also lists browser UI targets; ours is the page on about:blank
            let page = targets.as_array().into_iter().flatten().find(|t| t["type"] == "page" && t["url"] == "about:blank");
            if let Some(url) = page.and_then(|t| t["webSocketDebuggerUrl"].as_str()) {
                return Ok(url.to_string());
            }
        }
        if start.elapsed() > Duration::from_secs(20) {
            bail!("chrome launch: no page target on port {port} after 20s");
        }
        sleep(Duration::from_millis(150));
    }
}
