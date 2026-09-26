//! static_server.rs: serves the video/ folder over HTTP on 127.0.0.1 for the composition page
//! (module-free scripts, images, fonts, JSON). GET only; a path that climbs out of the root
//! (`..`, raw or percent-encoded) is a 404. Every answer is no-store and closes the connection.

use anyhow::{Context, Result};
use std::io::{BufRead, BufReader, Write};
use std::net::{TcpListener, TcpStream};
use std::path::{Path, PathBuf};
use std::thread;

/// Starts serving `root` on a free port in background threads; returns the port.
pub fn serve(root: PathBuf) -> Result<u16> {
    let listener = TcpListener::bind("127.0.0.1:0").context("static server: bind")?;
    let port = listener.local_addr()?.port();
    thread::spawn(move || {
        for conn in listener.incoming().flatten() {
            let root = root.clone();
            thread::spawn(move || {
                let _ = answer(conn, &root);
            });
        }
    });
    Ok(port)
}

fn answer(stream: TcpStream, root: &Path) -> std::io::Result<()> {
    let mut reader = BufReader::new(stream.try_clone()?);
    let mut line = String::new();
    reader.read_line(&mut line)?;
    let mut header = String::new();
    while reader.read_line(&mut header)? > 2 {
        header.clear();
    }
    let mut parts = line.split_whitespace();
    let (method, target) = (parts.next().unwrap_or(""), parts.next().unwrap_or(""));
    let file = if method == "GET" { resolve(root, target) } else { None };
    let mut out = stream;
    match file.and_then(|f| std::fs::read(&f).ok().map(|b| (f, b))) {
        Some((f, body)) => {
            let head = format!(
                "HTTP/1.1 200 OK\r\nContent-Type: {}\r\nContent-Length: {}\r\nCache-Control: no-store\r\nConnection: close\r\n\r\n",
                content_type(&f), body.len()
            );
            out.write_all(head.as_bytes())?;
            out.write_all(&body)
        }
        None => out.write_all(b"HTTP/1.1 404 Not Found\r\nContent-Length: 0\r\nConnection: close\r\n\r\n"),
    }
}

/// The file under `root` that the request target names, if it exists and stays inside `root`.
pub fn resolve(root: &Path, target: &str) -> Option<PathBuf> {
    let path = target.split(['?', '#']).next()?;
    let decoded = percent_decode(path)?;
    if decoded.split(['/', '\\']).any(|seg| seg == "..") {
        return None;
    }
    let full = root.join(decoded.trim_start_matches('/'));
    full.is_file().then_some(full)
}

fn percent_decode(s: &str) -> Option<String> {
    let bytes = s.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' {
            let hex = std::str::from_utf8(bytes.get(i + 1..i + 3)?).ok()?;
            out.push(u8::from_str_radix(hex, 16).ok()?);
            i += 3;
        } else {
            out.push(bytes[i]);
            i += 1;
        }
    }
    String::from_utf8(out).ok()
}

/// The Content-Type for a file, by extension.
pub fn content_type(path: &Path) -> &'static str {
    match path.extension().and_then(|e| e.to_str()).unwrap_or("") {
        "html" => "text/html; charset=utf-8",
        "js" => "text/javascript; charset=utf-8",
        "css" => "text/css; charset=utf-8",
        "json" => "application/json",
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "svg" => "image/svg+xml",
        "ttf" => "font/ttf",
        _ => "application/octet-stream",
    }
}
