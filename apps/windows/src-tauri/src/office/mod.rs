#[path = "platform/non_macos_background.rs"]
pub mod background;
pub mod certificate;
pub mod formula_cache;
#[path = "platform/non_macos_installer.rs"]
pub mod installer;
pub mod lifecycle;
#[path = "platform/non_macos_manifest.rs"]
pub mod manifest;
pub mod platform;
#[path = "platform/non_macos_powerpoint_native.rs"]
pub mod powerpoint_native;
pub mod server;
pub mod sessions;
pub mod state;
#[path = "platform/non_macos_word_native.rs"]
pub mod word_native;
pub mod windows_backend;
pub mod windows_pipe;

pub use lifecycle::{initialize, start};
