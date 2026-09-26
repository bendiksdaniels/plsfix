//! static_server.rs tests: files inside the root are served with their type; anything that
//! climbs out of the root, raw or percent-encoded, is a 404.
use plsfix_video::io::{http, static_server};
use std::fs;

fn root_with_page() -> std::path::PathBuf {
    let root = std::env::temp_dir().join(format!("pv-static-{}", std::process::id()));
    fs::create_dir_all(root.join("comp")).unwrap();
    fs::write(root.join("comp/index.html"), "<p>ok</p>").unwrap();
    fs::write(root.join("secret.txt"), "no").unwrap();
    root
}

#[test]
fn serves_files_and_refuses_escapes() {
    let root = root_with_page();
    let port = static_server::serve(root.join("comp")).unwrap();
    assert_eq!(http::get(port, "/index.html").unwrap(), (200, "<p>ok</p>".to_string()));
    assert_eq!(http::get(port, "/index.html?v=1").unwrap().0, 200);
    assert_eq!(http::get(port, "/../secret.txt").unwrap().0, 404);
    assert_eq!(http::get(port, "/%2e%2e/secret.txt").unwrap().0, 404);
    assert_eq!(http::get(port, "/missing.png").unwrap().0, 404);
}

#[test]
fn content_types_by_extension() {
    use std::path::Path;
    assert_eq!(static_server::content_type(Path::new("a.js")), "text/javascript; charset=utf-8");
    assert_eq!(static_server::content_type(Path::new("a.png")), "image/png");
    assert_eq!(static_server::content_type(Path::new("a.ttf")), "font/ttf");
}
