//! The screen copy of a TeX run's DVI/XDV (see emu.rs): written next to it, with each page's size, hash and
//! what could not be drawn.

use std::collections::HashMap;
use std::path::{Path, PathBuf};

use serde::Serialize;

use crate::dvi::Dvi;
use crate::emu::Emu;
use crate::fonts;
use crate::images::ImageHost;
use crate::tfm::Tfm;

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PageMeta {
    /// page size in bp
    pub width: f64,
    pub height: f64,
    /// changes when the page's content changes
    pub hash: String,
    #[serde(skip)]
    pub hash_value: u64,
    /// what the SVG cannot show exactly (the page is shown from the PDF then)
    pub unsupported: Vec<String>,
}

/// Rewrite `dvi_bytes` into `screen_path`.
#[allow(clippy::too_many_arguments)]
pub fn make_screen(
    dvi_bytes: &[u8],
    screen_path: &Path,
    dir: &Path,
    inputs: &[PathBuf],
    kpsewhich: Option<&Path>,
    tfm_cache: &mut HashMap<String, Option<Tfm>>,
    images: &mut ImageHost,
    paper: (f64, f64),
) -> Result<Vec<PageMeta>, String> {
    let dvi = Dvi::parse(dvi_bytes).map_err(|e| e.to_string())?;
    let metrics = fonts::load_metrics(&dvi, inputs, kpsewhich, dir, tfm_cache);
    images.new_run();
    let pages = {
        let mut emu = Emu::new(&dvi, &metrics, images, paper);
        emu.rewrite_all()
    };
    let bodies: Vec<Vec<u8>> = pages.iter().map(|p| p.body.clone()).collect();
    let bytes = dvi.write_with(&bodies);
    std::fs::write(screen_path, bytes).map_err(|e| format!("{}: {e}", screen_path.display()))?;
    Ok(pages
        .into_iter()
        .map(|p| PageMeta { width: p.width, height: p.height, hash: crate::util::hex(p.hash), hash_value: p.hash, unsupported: p.unsupported })
        .collect())
}
