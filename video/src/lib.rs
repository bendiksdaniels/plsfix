//! plsfix-video: builds the pls,fix motion-graphic video.
//! cli (commands) -> core (pure: frames, shots, copy and probe rules) -> io (Chrome over CDP,
//! the web rig, the static server, ffmpeg). Screens come only from the demo workbook and deck
//! running in Excel and PowerPoint for the web.
pub mod cli;
pub mod core;
pub mod io;
