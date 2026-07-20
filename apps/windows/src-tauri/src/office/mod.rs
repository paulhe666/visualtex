#[cfg(target_os = "macos")]
pub mod background;
#[cfg(not(target_os = "macos"))]
#[path = "platform/non_macos_background.rs"]
pub mod background;

pub mod certificate;
pub mod formula_cache;

#[cfg(target_os = "macos")]
pub mod installer;
#[cfg(not(target_os = "macos"))]
#[path = "platform/non_macos_installer.rs"]
pub mod installer;

pub mod lifecycle;

#[cfg(target_os = "macos")]
pub mod manifest;
#[cfg(not(target_os = "macos"))]
#[path = "platform/non_macos_manifest.rs"]
pub mod manifest;

pub mod platform;

#[cfg(target_os = "macos")]
pub mod powerpoint_native;
#[cfg(not(target_os = "macos"))]
#[path = "platform/non_macos_powerpoint_native.rs"]
pub mod powerpoint_native;

pub mod server;
pub mod sessions;
pub mod state;

#[cfg(target_os = "macos")]
pub mod word_native;
#[cfg(not(target_os = "macos"))]
#[path = "platform/non_macos_word_native.rs"]
pub mod word_native;

#[cfg(target_os = "windows")]
pub mod windows_backend;
#[cfg(target_os = "windows")]
pub mod windows_pipe;

pub use lifecycle::{initialize, start};
