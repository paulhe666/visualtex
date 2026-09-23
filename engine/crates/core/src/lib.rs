//! VisualTeX's independent math engine. This crate owns document state and
//! mathematical rules; hosts supply inputs and resources, never DOM geometry.
//! P02 intentionally supports empty sessions only. Capabilities are added with
//! explicit tests; successful lifecycle tests do not imply formula support.
#![forbid(unsafe_code)]

pub mod error;
pub mod protocol;
pub mod session;
pub mod source;
pub mod model;
pub mod parser;
pub mod symbols;
pub mod symbol_catalog;
pub mod rare_operators;
pub mod source_edit;
pub mod positions;
pub mod projection;
pub mod position_map;
pub mod transactions;
pub mod styles;
pub mod spacing;
pub mod paint;
pub mod history;
pub mod composition;
pub mod ocr;
pub mod editing;
pub mod fonts;
pub mod stretch;
pub mod delimiters;
pub mod accents;
pub mod operators;
pub mod fractions;
pub mod grids;
pub mod macros;
pub mod grid_selection;
pub mod scene;
pub mod scene_queries;
pub mod clipboard;
pub mod custom_symbols;
pub mod artwork_scene;
pub mod presentation_mathml;
mod scene_validation;
mod grid_scene_validation;
mod scene_navigation;
pub mod layout;

pub use error::{EngineError, EngineResult, ErrorCode};
pub use protocol::{Runtime, MAX_WIRE_BYTES, PROTOCOL_VERSION, error_json};
