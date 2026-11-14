//! WPS Office for Mac. JS add-ins in WPS Writer and WPS Presentation talk to
//! VisualTeX over a local HTTP bridge and reuse the Word-for-Mac and
//! PowerPoint-for-Mac formula pipelines: the same Sessions, editor, artifacts
//! and document state.

pub(crate) mod bridge;
pub(crate) mod numbering;
pub mod installer;
#[cfg(target_os = "macos")]
pub(crate) mod pasteboard;
pub(crate) mod payload;
pub(crate) mod presentation;
