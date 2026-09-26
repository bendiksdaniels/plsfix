//! io: the only code that touches processes, sockets and files.
//! cdp = one page's DevTools socket; chrome = headless Chrome processes; comp_page = the
//! composition page contract; rig = the signed-in scratch Chrome and the manifest server the
//! capture drives; ffmpeg + render_worker = frames into parts; video_scan = reading a render
//! back; http = a tiny GET; static_server = video/; wav + loudness = the soundtrack; tts = the
//! narrator's lines through edge-tts; sha = a file's SHA-256.
pub mod cdp;
pub mod chrome;
pub mod comp_page;
pub mod ffmpeg;
pub mod http;
pub mod loudness;
pub mod render_worker;
pub mod rig;
pub mod sha;
pub mod static_server;
pub mod tts;
pub mod video_scan;
pub mod wav;
