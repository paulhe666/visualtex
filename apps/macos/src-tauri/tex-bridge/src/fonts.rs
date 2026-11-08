//! The TFM files of a DVI's fonts: found from the engine's recorder file (`-recorder`, the files TeX read),
//! else with kpsewhich.

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::process::Command;

use crate::dvi::{Dvi, FontKind};
use crate::tfm::{Metrics, Tfm};

/// The paths TeX read (INPUT lines of a .fls file).
pub fn recorded_inputs(fls: &str, dir: &Path) -> Vec<PathBuf> {
    let mut pwd = dir.to_path_buf();
    let mut out = vec![];
    for line in fls.lines() {
        if let Some(p) = line.strip_prefix("PWD ") {
            pwd = PathBuf::from(p.trim());
        } else if let Some(p) = line.strip_prefix("INPUT ") {
            let p = Path::new(p.trim());
            out.push(if p.is_absolute() { p.to_path_buf() } else { pwd.join(p) });
        }
    }
    out
}

/// TFM widths for every TFM font of the DVI; `known` caches parsed files between runs.
pub fn load_metrics(dvi: &Dvi, inputs: &[PathBuf], kpsewhich: Option<&Path>, dir: &Path, known: &mut HashMap<String, Option<Tfm>>) -> Metrics {
    let mut by_name: HashMap<String, PathBuf> = HashMap::new();
    for p in inputs {
        if let Some(name) = p.file_name().and_then(|n| n.to_str()) {
            if let Some(stem) = name.strip_suffix(".tfm") {
                by_name.entry(stem.to_string()).or_insert_with(|| p.clone());
            }
        }
    }
    let mut missing: Vec<String> = vec![];
    for def in dvi.fonts.values() {
        if let FontKind::Tfm { name, .. } = &def.kind {
            if !known.contains_key(name) && !by_name.contains_key(name) {
                missing.push(name.clone());
            }
        }
    }
    missing.sort();
    missing.dedup();
    if !missing.is_empty() {
        if let Some(kpse) = kpsewhich {
            let mut cmd = Command::new(kpse);
            cmd.current_dir(dir);
            for m in &missing {
                cmd.arg(format!("{m}.tfm"));
            }
            crate::process::hide_window(&mut cmd);
            if let Ok(out) = cmd.output() {
                for line in String::from_utf8_lossy(&out.stdout).lines() {
                    let p = PathBuf::from(line.trim());
                    if let Some(stem) = p.file_stem().and_then(|s| s.to_str()) {
                        by_name.entry(stem.to_string()).or_insert(p);
                    }
                }
            }
        }
    }
    let mut metrics = Metrics::default();
    for (k, def) in &dvi.fonts {
        if let FontKind::Tfm { name, scale, .. } = &def.kind {
            if !known.contains_key(name) {
                let tfm = by_name.get(name).and_then(|p| std::fs::read(p).ok()).and_then(|b| Tfm::parse(&b));
                known.insert(name.clone(), tfm);
            }
            if let Some(Some(tfm)) = known.get(name) {
                metrics.fonts.insert(*k, (tfm.clone(), *scale));
            }
        }
    }
    metrics
}
