//! WPS Office for Mac. A JS add-in in WPS Writer talks to VisualTeX over a
//! local HTTP bridge and reuses the Word-for-Mac formula pipeline: the same
//! Sessions, editor, artifacts and document state.

pub(crate) mod bridge;
pub(crate) mod numbering;
pub mod installer;
pub(crate) mod payload;
