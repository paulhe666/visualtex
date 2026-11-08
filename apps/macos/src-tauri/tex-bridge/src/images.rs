//! The images `pdf:image` specials name, for the screen copy: PNG/JPEG as they are (data: URLs), PDF and EPS
//! pages drawn by MuPDF (SVG) when it is there, else by Ghostscript (PNG at 3× resolution).

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::Arc;

use base64::engine::general_purpose::STANDARD as BASE64;
use base64::Engine;

use crate::emu::{Host, ImageData, ImageSource};
use crate::util::{hash_bytes, hex};

/// Pixel size and density (bp per pixel) of a PNG.
pub fn png_info(b: &[u8]) -> Option<(f64, f64, f64, f64)> {
    if b.len() < 24 || &b[..8] != b"\x89PNG\r\n\x1a\n" {
        return None;
    }
    let w = u32::from_be_bytes([b[16], b[17], b[18], b[19]]) as f64;
    let h = u32::from_be_bytes([b[20], b[21], b[22], b[23]]) as f64;
    let (mut xd, mut yd) = (1.0, 1.0);
    let mut i = 8;
    while i + 8 <= b.len() {
        let len = u32::from_be_bytes([b[i], b[i + 1], b[i + 2], b[i + 3]]) as usize;
        let kind = &b[i + 4..i + 8];
        if kind == b"pHYs" && i + 17 <= b.len() {
            let px = u32::from_be_bytes([b[i + 8], b[i + 9], b[i + 10], b[i + 11]]) as f64;
            let py = u32::from_be_bytes([b[i + 12], b[i + 13], b[i + 14], b[i + 15]]) as f64;
            if b[i + 16] == 1 && px > 0.0 && py > 0.0 {
                // pixels per metre → bp per pixel
                xd = 72.0 / (px * 0.0254);
                yd = 72.0 / (py * 0.0254);
            }
        }
        if kind == b"IDAT" {
            break;
        }
        i += 12 + len;
    }
    Some((w, h, xd, yd))
}

/// Pixel size and density of a JPEG (JFIF density when given).
pub fn jpeg_info(b: &[u8]) -> Option<(f64, f64, f64, f64)> {
    if b.len() < 4 || b[0] != 0xFF || b[1] != 0xD8 {
        return None;
    }
    let (mut xd, mut yd) = (1.0, 1.0);
    let mut i = 2;
    while i + 4 <= b.len() {
        if b[i] != 0xFF {
            i += 1;
            continue;
        }
        let marker = b[i + 1];
        if marker == 0xD8 || marker == 0x01 || (0xD0..=0xD7).contains(&marker) {
            i += 2;
            continue;
        }
        let len = ((b[i + 2] as usize) << 8) | b[i + 3] as usize;
        let seg = &b[i + 4..(i + 2 + len).min(b.len())];
        if marker == 0xE0 && seg.starts_with(b"JFIF\0") && seg.len() >= 12 {
            let unit = seg[7];
            let dx = ((seg[8] as f64) * 256.0) + seg[9] as f64;
            let dy = ((seg[10] as f64) * 256.0) + seg[11] as f64;
            if dx > 0.0 && dy > 0.0 && unit != 0 {
                let f = if unit == 1 { 1.0 } else { 2.54 };
                xd = 72.0 / (dx * f);
                yd = 72.0 / (dy * f);
            }
        }
        if matches!(marker, 0xC0..=0xC3 | 0xC5..=0xC7 | 0xC9..=0xCB | 0xCD..=0xCF) && seg.len() >= 5 {
            let h = ((seg[1] as f64) * 256.0) + seg[2] as f64;
            let w = ((seg[3] as f64) * 256.0) + seg[4] as f64;
            return Some((w, h, xd, yd));
        }
        i += 2 + len;
    }
    None
}

/// The first /CropBox (else /MediaBox) of a PDF, when it can be read without decompressing.
pub fn pdf_box(b: &[u8]) -> Option<[f64; 4]> {
    let text = String::from_utf8_lossy(b);
    for key in ["/CropBox", "/MediaBox"] {
        if let Some(at) = text.find(key) {
            let rest = &text[at + key.len()..];
            let open = rest.find('[')?;
            let close = rest.find(']')?;
            let v: Vec<f64> = rest[open + 1..close].split_whitespace().filter_map(|s| s.parse().ok()).collect();
            if v.len() == 4 {
                return Some([v[0], v[1], v[2], v[3]]);
            }
        }
    }
    None
}

/// `%%BoundingBox:` of an EPS file.
pub fn eps_box(b: &[u8]) -> Option<[f64; 4]> {
    let text = String::from_utf8_lossy(&b[..b.len().min(64 * 1024)]);
    for key in ["%%HiResBoundingBox:", "%%BoundingBox:"] {
        if let Some(at) = text.find(key) {
            let line = text[at + key.len()..].lines().next()?;
            let v: Vec<f64> = line.split_whitespace().filter_map(|s| s.parse().ok()).collect();
            if v.len() == 4 {
                return Some([v[0], v[1], v[2], v[3]]);
            }
        }
    }
    None
}

/// External programs that draw PDF/EPS pages.
#[derive(Clone, Debug, Default)]
pub struct Drawers {
    pub mutool: Option<PathBuf>,
    pub gs: Option<PathBuf>,
    /// TeX's extractbb: a PDF page's crop box (what dvipdfmx places)
    pub extractbb: Option<PathBuf>,
}

/// The images of one build folder, cached by file content (kept between compiles).
pub struct ImageHost {
    pub dir: PathBuf,
    pub cache_dir: PathBuf,
    pub drawers: Drawers,
    cache: HashMap<String, Option<Arc<ImageData>>>,
    by_content: HashMap<u64, Arc<ImageData>>,
}

impl ImageHost {
    pub fn new(dir: &Path, cache_dir: &Path, drawers: Drawers) -> ImageHost {
        ImageHost { dir: dir.to_path_buf(), cache_dir: cache_dir.to_path_buf(), drawers, cache: HashMap::new(), by_content: HashMap::new() }
    }

    /// Forget names (a new compile: files may have changed); content stays cached.
    pub fn new_run(&mut self) {
        self.cache.clear();
    }

    fn load(&mut self, file: &str, page: u32) -> Option<Arc<ImageData>> {
        let path = {
            let p = Path::new(file);
            if p.is_absolute() {
                p.to_path_buf()
            } else {
                self.dir.join(p)
            }
        };
        let bytes = std::fs::read(&path).ok()?;
        let mut key = hash_bytes(&bytes);
        key ^= page as u64;
        if let Some(found) = self.by_content.get(&key) {
            return Some(found.clone());
        }
        let lower = file.to_ascii_lowercase();
        let data = if let Some((w, h, xd, yd)) = png_info(&bytes) {
            ImageData { source: ImageSource::Raster { width: w, height: h, xdensity: xd, ydensity: yd }, href: data_url("image/png", &bytes) }
        } else if let Some((w, h, xd, yd)) = jpeg_info(&bytes) {
            ImageData { source: ImageSource::Raster { width: w, height: h, xdensity: xd, ydensity: yd }, href: data_url("image/jpeg", &bytes) }
        } else if bytes.starts_with(b"%PDF") || lower.ends_with(".pdf") {
            let b = self.extract_box(&path, page).or_else(|| pdf_box(&bytes)).unwrap_or([0.0, 0.0, 612.0, 792.0]);
            let href = self.draw(&path, page, false, &hex(key))?;
            ImageData { source: ImageSource::Box { llx: b[0], lly: b[1], urx: b[2], ury: b[3] }, href }
        } else if bytes.starts_with(b"%!") || lower.ends_with(".eps") || lower.ends_with(".ps") {
            // as dvipdfmx does it: Ghostscript makes a PDF of the bounding box (-dEPSCrop), which is placed
            match self.eps_to_pdf(&path, &hex(key)) {
                Some(pdf) => {
                    let b = self.extract_box(&pdf, 1).or_else(|| eps_box(&bytes).map(|b| [0.0, 0.0, b[2] - b[0], b[3] - b[1]]))?;
                    let href = self.draw(&pdf, 1, false, &hex(key))?;
                    ImageData { source: ImageSource::Box { llx: b[0], lly: b[1], urx: b[2], ury: b[3] }, href }
                }
                None => {
                    let b = eps_box(&bytes)?;
                    let href = self.draw(&path, 1, true, &hex(key))?;
                    ImageData { source: ImageSource::Box { llx: b[0], lly: b[1], urx: b[2], ury: b[3] }, href }
                }
            }
        } else if lower.ends_with(".svg") {
            ImageData { source: ImageSource::Raster { width: 100.0, height: 100.0, xdensity: 1.0, ydensity: 1.0 }, href: data_url("image/svg+xml", &bytes) }
        } else {
            return None;
        };
        let data = Arc::new(data);
        self.by_content.insert(key, data.clone());
        Some(data)
    }

    /// An EPS file as PDF (dvipdfmx's conversion: Ghostscript, -dEPSCrop).
    fn eps_to_pdf(&self, path: &Path, key: &str) -> Option<PathBuf> {
        let gs = self.drawers.gs.as_ref()?;
        let _ = std::fs::create_dir_all(&self.cache_dir);
        let out = self.cache_dir.join(format!("eps-{key}.pdf"));
        if !out.exists() {
            let mut cmd = Command::new(gs);
            cmd.args(["-q", "-dNOPAUSE", "-dBATCH", "-dSAFER", "-dEPSCrop", "-sPAPERSIZE=a0", "-sDEVICE=pdfwrite", "-dAutoRotatePages=/None"])
                .arg(format!("-sOutputFile={}", out.display()))
                .arg(path);
            crate::process::hide_window(&mut cmd);
            if !matches!(cmd.status(), Ok(s) if s.success()) {
                let _ = std::fs::remove_file(&out);
                return None;
            }
        }
        Some(out)
    }

    /// The crop box of a PDF page, from extractbb.
    fn extract_box(&self, path: &Path, page: u32) -> Option<[f64; 4]> {
        let tool = self.drawers.extractbb.as_ref()?;
        let mut cmd = Command::new(tool);
        cmd.args(["-O", "-B", "cropbox", "-p", &page.to_string()]).arg(path);
        crate::process::hide_window(&mut cmd);
        let out = cmd.output().ok()?;
        let text = String::from_utf8_lossy(&out.stdout).to_string();
        eps_box(text.as_bytes())
    }

    /// A PDF page or EPS file drawn: SVG from MuPDF, else PNG from Ghostscript.
    fn draw(&self, path: &Path, page: u32, eps: bool, key: &str) -> Option<String> {
        let _ = std::fs::create_dir_all(&self.cache_dir);
        if let (Some(mutool), false) = (&self.drawers.mutool, eps) {
            // mutool names the file by the page drawn: img-KEY-<page>.svg
            let out = self.cache_dir.join(format!("img-{key}-{page}.svg"));
            if !out.exists() {
                let mut cmd = Command::new(mutool);
                cmd.args(["draw", "-q", "-F", "svg", "-o"]).arg(self.cache_dir.join(format!("img-{key}-%d.svg"))).arg(path).arg(page.to_string());
                crate::process::hide_window(&mut cmd);
                let status = cmd.status();
                if !matches!(status, Ok(s) if s.success()) {
                    let _ = std::fs::remove_file(&out);
                }
            }
            if let Ok(svg) = std::fs::read(&out) {
                return Some(data_url("image/svg+xml", &svg));
            }
        }
        if let Some(gs) = &self.drawers.gs {
            let out = self.cache_dir.join(format!("img-{key}.png"));
            if !out.exists() {
                let mut cmd = Command::new(gs);
                cmd.args(["-q", "-dSAFER", "-dBATCH", "-dNOPAUSE", "-sDEVICE=pngalpha", "-r216", "-dTextAlphaBits=4", "-dGraphicsAlphaBits=4"]);
                if eps {
                    cmd.arg("-dEPSCrop");
                } else {
                    cmd.arg("-dUseCropBox").arg(format!("-dFirstPage={page}")).arg(format!("-dLastPage={page}"));
                }
                cmd.arg(format!("-sOutputFile={}", out.display())).arg(path);
                crate::process::hide_window(&mut cmd);
                let status = cmd.status();
                if !matches!(status, Ok(s) if s.success()) {
                    let _ = std::fs::remove_file(&out);
                }
            }
            if let Ok(png) = std::fs::read(&out) {
                return Some(data_url("image/png", &png));
            }
        }
        None
    }
}

impl Host for ImageHost {
    fn image(&mut self, file: &str, page: u32) -> Option<Arc<ImageData>> {
        let key = format!("{file}\u{0}{page}");
        if let Some(found) = self.cache.get(&key) {
            return found.clone();
        }
        let found = self.load(file, page);
        self.cache.insert(key, found.clone());
        found
    }
}

pub fn data_url(mime: &str, bytes: &[u8]) -> String {
    format!("data:{mime};base64,{}", BASE64.encode(bytes))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn png_header() {
        let mut b = b"\x89PNG\r\n\x1a\n".to_vec();
        b.extend_from_slice(&13u32.to_be_bytes());
        b.extend_from_slice(b"IHDR");
        b.extend_from_slice(&640u32.to_be_bytes());
        b.extend_from_slice(&480u32.to_be_bytes());
        b.extend_from_slice(&[8, 6, 0, 0, 0, 0, 0, 0, 0]);
        // pHYs: 300 dpi
        b.extend_from_slice(&9u32.to_be_bytes());
        b.extend_from_slice(b"pHYs");
        b.extend_from_slice(&11811u32.to_be_bytes());
        b.extend_from_slice(&11811u32.to_be_bytes());
        b.push(1);
        b.extend_from_slice(&[0; 4]);
        let (w, h, xd, _) = png_info(&b).unwrap();
        assert_eq!((w, h), (640.0, 480.0));
        assert!((xd - 72.0 / 300.0).abs() < 1e-3);
    }

    #[test]
    fn boxes() {
        assert_eq!(pdf_box(b"%PDF-1.5 << /Type /Page /MediaBox [0 0 200 100] >>"), Some([0.0, 0.0, 200.0, 100.0]));
        assert_eq!(eps_box(b"%!PS-Adobe-3.0 EPSF-3.0\n%%BoundingBox: 1 2 30 40\n"), Some([1.0, 2.0, 30.0, 40.0]));
    }
}
