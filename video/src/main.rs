//! main.rs: the `plsfix-video` entry point; all behaviour lives in `cli`.
//! Exit code 1 on any error, with the stage and identifier in the message.
fn main() -> std::process::ExitCode {
    plsfix_video::cli::run()
}
