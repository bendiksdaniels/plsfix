//! http.rs: a tiny HTTP/1.1 GET against 127.0.0.1, for Chrome's /json/list and /json/version (the
//! rig's check). Reads the headers, then exactly Content-Length bytes (Chrome keeps the connection
//! open and ignores `Connection: close`), or to the end when no length is given.

use anyhow::{Context, Result};
use std::io::{BufRead, BufReader, Read, Write};
use std::net::TcpStream;
use std::time::Duration;

/// GETs `path` from 127.0.0.1:`port`; returns the status code and the body as text.
pub fn get(port: u16, path: &str) -> Result<(u16, String)> {
    let what = || format!("http get :{port}{path}");
    let mut s = TcpStream::connect(("127.0.0.1", port)).with_context(what)?;
    s.set_read_timeout(Some(Duration::from_secs(10)))?;
    write!(s, "GET {path} HTTP/1.1\r\nHost: 127.0.0.1:{port}\r\nAccept: */*\r\n\r\n").with_context(what)?;
    let mut reader = BufReader::new(s);
    let mut status_line = String::new();
    reader.read_line(&mut status_line).with_context(what)?;
    let status = status_line.split_whitespace().nth(1).and_then(|c| c.parse().ok()).unwrap_or(0);
    let mut length: Option<usize> = None;
    loop {
        let mut line = String::new();
        if reader.read_line(&mut line).with_context(what)? == 0 || line == "\r\n" {
            break;
        }
        if let Some((k, v)) = line.split_once(':') {
            if k.trim().eq_ignore_ascii_case("content-length") {
                length = v.trim().parse().ok();
            }
        }
    }
    let mut body = Vec::new();
    match length {
        Some(n) => {
            body.resize(n, 0);
            reader.read_exact(&mut body).with_context(what)?;
        }
        None => {
            reader.read_to_end(&mut body).with_context(what)?;
        }
    }
    Ok((status, String::from_utf8_lossy(&body).into_owned()))
}
