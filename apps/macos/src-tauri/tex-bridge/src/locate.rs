//! Finding the TeX distribution: PATH first (a terminal start), then the places TeX Live, MacTeX and MiKTeX
//! install to (an app started from the Dock or Start menu does not have the TeX folder on its PATH).

use std::path::{Path, PathBuf};
use std::process::Command;

use serde::Serialize;

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TexDist {
    /// the folder of xelatex, pdflatex, dvisvgm, …
    pub bin: PathBuf,
    /// "TeX Live 2025", "MiKTeX …"
    pub version: String,
    pub engines: Vec<String>,
    pub dvisvgm: bool,
    pub biber: bool,
    /// draws PDF/EPS images and pages the SVG cannot show
    pub mutool: Option<PathBuf>,
    pub gs: Option<PathBuf>,
}

fn exe(name: &str) -> String {
    if cfg!(windows) {
        format!("{name}.exe")
    } else {
        name.to_string()
    }
}

impl TexDist {
    pub fn tool(&self, name: &str) -> PathBuf {
        self.bin.join(exe(name))
    }
    pub fn has(&self, name: &str) -> bool {
        self.tool(name).is_file()
    }
    /// PATH for the TeX programs: the TeX folder first (they call each other: extractbb, kpsewhich, rungs).
    pub fn path_env(&self) -> std::ffi::OsString {
        let mut paths = vec![self.bin.clone()];
        if let Some(old) = std::env::var_os("PATH") {
            paths.extend(std::env::split_paths(&old));
        }
        std::env::join_paths(paths).unwrap_or_default()
    }
}

fn years_desc(dir: &Path) -> Vec<PathBuf> {
    let mut years: Vec<PathBuf> = std::fs::read_dir(dir)
        .map(|r| r.flatten().map(|e| e.path()).filter(|p| p.file_name().and_then(|n| n.to_str()).map(|n| n.len() == 4 && n.chars().all(|c| c.is_ascii_digit())).unwrap_or(false)).collect())
        .unwrap_or_default();
    years.sort();
    years.reverse();
    years
}

/// Folders that may hold xelatex, best first.
pub fn candidate_bins() -> Vec<PathBuf> {
    let mut out = vec![];
    if let Some(v) = std::env::var_os("VISUALTEX_TEX_BIN") {
        out.push(PathBuf::from(v));
    }
    if let Some(path) = std::env::var_os("PATH") {
        out.extend(std::env::split_paths(&path));
    }
    if cfg!(windows) {
        for root in ["C:\\texlive", "D:\\texlive"] {
            for y in years_desc(Path::new(root)) {
                for arch in ["windows", "win64", "win32"] {
                    out.push(y.join("bin").join(arch));
                }
            }
        }
        if let Some(local) = std::env::var_os("LOCALAPPDATA") {
            out.push(PathBuf::from(local).join("Programs\\MiKTeX\\miktex\\bin\\x64"));
        }
        for pf in ["C:\\Program Files\\MiKTeX\\miktex\\bin\\x64", "C:\\Program Files (x86)\\MiKTeX\\miktex\\bin"] {
            out.push(PathBuf::from(pf));
        }
    } else {
        out.push(PathBuf::from("/Library/TeX/texbin"));
        for y in years_desc(Path::new("/usr/local/texlive")) {
            for arch in ["universal-darwin", "x86_64-darwinlegacy", "aarch64-linux", "x86_64-linux"] {
                out.push(y.join("bin").join(arch));
            }
        }
        for p in ["/opt/homebrew/bin", "/usr/local/bin", "/usr/bin"] {
            out.push(PathBuf::from(p));
        }
    }
    out
}

fn find_in(dirs: &[PathBuf], name: &str) -> Option<PathBuf> {
    dirs.iter().map(|d| d.join(exe(name))).find(|p| p.is_file())
}

pub fn locate() -> Option<TexDist> {
    let dirs = candidate_bins();
    let bin = dirs.iter().find(|d| d.join(exe("xelatex")).is_file() || d.join(exe("pdflatex")).is_file())?.clone();
    let mut version = String::new();
    let mut cmd = Command::new(bin.join(exe(if bin.join(exe("xelatex")).is_file() { "xelatex" } else { "pdflatex" })));
    cmd.arg("--version");
    crate::process::hide_window(&mut cmd);
    if let Ok(out) = cmd.output() {
        let text = String::from_utf8_lossy(&out.stdout);
        let first = text.lines().next().unwrap_or("");
        // "XeTeX 3.141592653-2.6-0.999997 (TeX Live 2025)"
        version = match (first.find('('), first.rfind(')')) {
            (Some(a), Some(b)) if b > a => first[a + 1..b].to_string(),
            _ => first.to_string(),
        };
    }
    let engines = ["xelatex", "pdflatex", "lualatex"].iter().filter(|e| bin.join(exe(e)).is_file()).map(|s| s.to_string()).collect();
    let mut gs = None;
    if cfg!(windows) {
        // TeX Live's own Ghostscript, through its wrapper (it sets the library path)
        if bin.join("rungs.exe").is_file() {
            gs = Some(bin.join("rungs.exe"));
        }
        if gs.is_none() {
            gs = find_in(&dirs, "gswin64c");
        }
    } else {
        gs = find_in(&dirs, "gs");
    }
    let mutool = find_in(&dirs, "mutool");
    Some(TexDist {
        dvisvgm: bin.join(exe("dvisvgm")).is_file(),
        biber: bin.join(exe("biber")).is_file(),
        bin,
        version,
        engines,
        mutool,
        gs,
    })
}

#[cfg(test)]
mod tests {
    #[test]
    fn finds_tex_here() {
        // this machine has TeX Live; a machine without it simply gets None
        if let Some(d) = super::locate() {
            assert!(d.engines.iter().any(|e| e == "xelatex" || e == "pdflatex"));
            assert!(d.tool("kpsewhich").exists());
        }
    }
}
