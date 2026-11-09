//! VisualTeX's LaTeX documents, the TeX side: find TeX, keep a per-document build folder, run the engine (with a
//! warm engine holding the preamble), turn the XDV/DVI pages into SVG with dvisvgm after rewriting the
//! (x)dvipdfmx specials into SVG (so the pages show exactly what the PDF of the same run shows), make that PDF.

pub mod cjk_table;
pub mod commands;
pub mod dvi;
pub mod engine;
pub mod emu;
pub mod fontfix;
pub mod fonts;
pub mod images;
pub mod live;
pub mod locate;
pub mod outlines;
pub mod pdfobj;
pub mod process;
pub mod project;
pub mod screen;
pub mod service;
pub mod svg;
pub mod tfm;
pub mod type1;
pub mod util;
