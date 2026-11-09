//! A live paragraph's page of a DVI/XDV, read into glyphs (font file, glyph, position) and rules, so that it is
//! drawn from the fonts' outlines (outlines.rs) at once rather than through dvisvgm: XDV glyph runs carry their
//! glyph ids and positions; a TFM font's characters advance by its TFM widths and are drawn from its Type 1 file
//! through the map file (pdftex.map), a virtual font's from its packets. What cannot be drawn this way (a font
//! without outlines here, specials that draw) is counted: the page then goes through dvisvgm instead.

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::process::Command;

use crate::dvi::{Dvi, FontKind, Reader};
use crate::outlines::FontKey;
use crate::tfm::Tfm;

/// A glyph: font (index into `fonts`), x, y (baseline) in bp from the page's origin, code, colour (index into
/// `colors`, 0 black).
pub type Glyph = (usize, f64, f64, u32, usize);

/// What a glyph stands for (the caret's places in a paragraph come from these): its advance (bp) and where it
/// comes from — `TEXT` a word of a native font (its run's characters are in `runs`), `CHAR` a TFM font's
/// character, `GLYPH` a native font's glyph outside a word (a formula's).
pub const TEXT: u8 = 1;
pub const CHAR: u8 = 2;
pub const GLYPH: u8 = 3;

#[derive(Default, Debug)]
pub struct Drawn {
    pub glyphs: Vec<Glyph>,
    /// for each glyph: advance (bp), kind (TEXT, CHAR, GLYPH)
    pub advances: Vec<(f64, u8)>,
    /// the words of native fonts (XDV text and glyphs, \XeTeXgenerateactualtext): first glyph, glyph count, text
    pub runs: Vec<(usize, usize, String)>,
    /// specials "vtx:<name>" (other than the paragraph's own marker and padding) and where they are (bp)
    pub marks: Vec<(String, f64, f64)>,
    /// x, y (top), width, height (bp), colour
    pub rules: Vec<([f64; 4], usize)>,
    /// (outline set, size in bp, slant, extend) of each font used
    pub fonts: Vec<(FontKey, f64, f64, f64)>,
    /// CSS colours ("#rrggbb"); 0 is black
    pub colors: Vec<String>,
    /// what could not be drawn here
    pub undrawable: Vec<String>,
}

/// Where fonts' files are (kpsewhich, the map file), kept for the server's life.
#[derive(Default)]
pub struct FontFiles {
    kpse: Option<PathBuf>,
    found: HashMap<(String, String), Option<PathBuf>>,
    map: Option<HashMap<String, MapEntry>>,
    tfm: HashMap<String, Option<Tfm>>,
    vf: HashMap<String, Option<Vf>>,
}

#[derive(Clone, Debug, Default)]
struct MapEntry {
    font: String,
    enc: String,
    slant: f64,
    extend: f64,
}

#[derive(Clone, Debug)]
struct VfFont {
    checksum: u32,
    scale: i32,
    design: i32,
    name: String,
}

#[derive(Clone, Debug, Default)]
struct Vf {
    design: i32,
    fonts: Vec<(u32, VfFont)>,
    chars: HashMap<u32, Vec<u8>>,
}

impl FontFiles {
    pub fn new(kpse: Option<PathBuf>) -> FontFiles {
        FontFiles { kpse, ..Default::default() }
    }

    fn find(&mut self, name: &str, format: &str) -> Option<PathBuf> {
        let key = (name.to_string(), format.to_string());
        if let Some(f) = self.found.get(&key) {
            return f.clone();
        }
        let p = Path::new(name);
        let found = if p.is_absolute() && p.is_file() {
            Some(p.to_path_buf())
        } else {
            self.kpse.as_ref().and_then(|k| {
                let mut cmd = Command::new(k);
                if !format.is_empty() {
                    cmd.arg(format!("-format={format}"));
                }
                cmd.arg(name);
                crate::process::hide_window(&mut cmd);
                let out = cmd.output().ok()?;
                let line = String::from_utf8_lossy(&out.stdout).lines().next()?.trim().to_string();
                (!line.is_empty()).then(|| PathBuf::from(line)).filter(|p| p.is_file())
            })
        };
        self.found.insert(key, found.clone());
        found
    }

    fn map_entry(&mut self, tfm: &str) -> Option<MapEntry> {
        if self.map.is_none() {
            let mut map = HashMap::new();
            if let Some(file) = self.find("pdftex.map", "map") {
                for line in std::fs::read_to_string(file).unwrap_or_default().lines() {
                    let first = line.chars().next().unwrap_or('%');
                    if first == '%' || first == '#' || line.trim().is_empty() {
                        continue;
                    }
                    let name = line.split_whitespace().next().unwrap_or("").to_string();
                    if name.is_empty() || map.contains_key(&name) {
                        continue;
                    }
                    let mut entry = MapEntry { extend: 1.0, ..Default::default() };
                    if let Some(a) = line.find('"') {
                        if let Some(b) = line[a + 1..].find('"') {
                            let special: Vec<&str> = line[a + 1..a + 1 + b].split_whitespace().collect();
                            for w in special.windows(2) {
                                if w[1] == "SlantFont" {
                                    entry.slant = w[0].parse().unwrap_or(0.0);
                                }
                                if w[1] == "ExtendFont" {
                                    entry.extend = w[0].parse().unwrap_or(1.0);
                                }
                            }
                        }
                    }
                    for word in line.split_whitespace() {
                        if let Some(file) = word.strip_prefix('<') {
                            let file = file.trim_start_matches(['<', '[']);
                            let lower = file.to_ascii_lowercase();
                            if lower.ends_with(".enc") {
                                entry.enc = file.to_string();
                            } else if lower.ends_with(".pfb") || lower.ends_with(".pfa") || lower.ends_with(".otf") || lower.ends_with(".ttf") {
                                entry.font = file.to_string();
                            }
                        }
                    }
                    map.insert(name, entry);
                }
            }
            self.map = Some(map);
        }
        let mut entry = self.map.as_ref()?.get(tfm)?.clone();
        if entry.font.is_empty() {
            return None;
        }
        let lower = entry.font.to_ascii_lowercase();
        let format = if lower.ends_with(".otf") { "opentype fonts" } else if lower.ends_with(".ttf") { "truetype fonts" } else { "type1 fonts" };
        entry.font = self.find(&entry.font.clone(), format)?.to_string_lossy().to_string();
        if !entry.enc.is_empty() {
            entry.enc = self.find(&entry.enc.clone(), "enc files")?.to_string_lossy().to_string();
        }
        Some(entry)
    }

    fn tfm(&mut self, name: &str) -> Option<Tfm> {
        if !self.tfm.contains_key(name) {
            let t = self.find(&format!("{name}.tfm"), "tfm").and_then(|p| std::fs::read(p).ok()).and_then(|b| Tfm::parse(&b));
            self.tfm.insert(name.to_string(), t);
        }
        self.tfm.get(name).cloned().flatten()
    }

    fn vf(&mut self, name: &str) -> Option<Vf> {
        if !self.vf.contains_key(name) {
            let v = self.find(&format!("{name}.vf"), "vf").and_then(|p| std::fs::read(p).ok()).and_then(|b| parse_vf(&b));
            self.vf.insert(name.to_string(), v);
        }
        self.vf.get(name).cloned().flatten()
    }
}

fn parse_vf(data: &[u8]) -> Option<Vf> {
    let mut r = Reader::new(data);
    if r.u(1).ok()? != 247 || r.u(1).ok()? != 202 {
        return None;
    }
    let k = r.u(1).ok()? as usize;
    r.bytes(k).ok()?;
    r.u(4).ok()?; // checksum
    let design = r.s(4).ok()?;
    let mut vf = Vf { design, ..Default::default() };
    loop {
        let o = r.u(1).ok()?;
        match o {
            0..=241 => {
                let pl = o as usize;
                let cc = r.u(1).ok()?;
                r.u(3).ok()?;
                vf.chars.insert(cc, r.bytes(pl).ok()?.to_vec());
            }
            242 => {
                let pl = r.u(4).ok()? as usize;
                let cc = r.u(4).ok()?;
                r.u(4).ok()?;
                vf.chars.insert(cc, r.bytes(pl).ok()?.to_vec());
            }
            243..=246 => {
                let k = r.u((o - 242) as usize).ok()?;
                let checksum = r.u(4).ok()?;
                let scale = r.s(4).ok()?;
                let design = r.s(4).ok()?;
                let a = r.u(1).ok()? as usize;
                let l = r.u(1).ok()? as usize;
                let name = String::from_utf8_lossy(&r.bytes(a + l).ok()?[a..]).to_string();
                vf.fonts.push((k, VfFont { checksum, scale, design, name }));
            }
            248 => break,
            _ => return None,
        }
    }
    Some(vf)
}

/// One font as the interpreter draws it.
#[derive(Clone)]
enum Font {
    /// XDV native: its outline set (glyph ids), size in bp, slant, extend
    Native { key: FontKey, size: f64, slant: f64, extend: f64 },
    /// TFM: its widths, its scale (DVI units), its outline set when it has one, or its virtual font
    Tfm { name: String, tfm: Tfm, scale: i32, key: Option<(FontKey, f64, f64)>, vf: Option<(Vf, Vec<(u32, Font)>)> },
}

struct Machine {
    unit: f64,
    out: Drawn,
    font_index: HashMap<(FontKey, i64, i64, i64), usize>,
    color: Vec<usize>,
}

impl Machine {
    fn font_slot(&mut self, key: &FontKey, size: f64, slant: f64, extend: f64) -> usize {
        let k = (key.clone(), (size * 1000.0).round() as i64, (slant * 1000.0).round() as i64, (extend * 1000.0).round() as i64);
        if let Some(&i) = self.font_index.get(&k) {
            return i;
        }
        let i = self.out.fonts.len();
        self.out.fonts.push((key.clone(), size, slant, extend));
        self.font_index.insert(k, i);
        i
    }

    fn color_now(&self) -> usize {
        *self.color.last().unwrap_or(&0)
    }

    /// A colour special ("color push rgb 0 0 1", "color pop", "color set ..."), a mark of ours ("vtx:<name>") at
    /// (h, v) in bp
    fn special(&mut self, text: &[u8], h: f64, v: f64) {
        let s = String::from_utf8_lossy(text);
        let s = s.trim();
        if let Some(name) = s.strip_prefix("vtx:") {
            // (the paragraph's own marker "vtx:q<n>" and the padding page's are not marks)
            if !name.starts_with("pad") && !(name.starts_with('q') && name[1..].chars().all(|c| c.is_ascii_digit())) {
                self.out.marks.push((name.to_string(), h, v));
            }
            return;
        }
        if let Some(rest) = s.strip_prefix("color push").or_else(|| s.strip_prefix("color set")) {
            let css = css_color(rest.trim()).unwrap_or_else(|| "#000000".into());
            let index = match self.out.colors.iter().position(|c| *c == css) {
                Some(i) => i,
                None => {
                    self.out.colors.push(css);
                    self.out.colors.len() - 1
                }
            };
            if s.starts_with("color set") {
                if let Some(top) = self.color.last_mut() {
                    *top = index;
                } else {
                    self.color.push(index);
                }
            } else {
                self.color.push(index);
            }
        } else if s.starts_with("color pop") {
            self.color.pop();
        } else if let Some(rest) = s.strip_prefix("pdf:uxobj") {
            // a form drawn by reference (a list's bullet drawn once in the preamble): not drawn here; where it is,
            // for the page to show it there
            let name = rest.split_whitespace().next().unwrap_or("").to_string();
            self.out.marks.push((format!("xobj {name}"), h, v));
        } else if s.starts_with("vtx:") || s.starts_with("pdf:bann") || s.starts_with("pdf:eann") || s.starts_with("pdf:dest") || s.starts_with("pdf:outline") || s.starts_with("pdf:docinfo") || s.starts_with("pdf:pagesize") || s.starts_with("papersize") || s.starts_with("header=") || s.starts_with("src:") || s.is_empty() {
            // links, destinations, our marker: nothing drawn
        } else {
            self.out.undrawable.push(format!("special {}", s.chars().take(40).collect::<String>()));
        }
    }

    /// Runs DVI operations `ops` with `fonts` (font number -> font), from (h, v) in DVI units.
    fn run(&mut self, ops: &[u8], fonts: &HashMap<u32, Font>, first_font: Option<u32>, h0: f64, v0: f64, scale: f64, depth: usize) {
        if depth > 6 {
            self.out.undrawable.push("virtual fonts nested too deep".into());
            return;
        }
        let mut r = Reader::new(ops);
        let (mut h, mut v, mut w, mut x, mut y, mut z) = (h0, v0, 0.0, 0.0, 0.0, 0.0);
        let mut stack: Vec<(f64, f64, f64, f64, f64, f64)> = vec![];
        let mut font = first_font;
        while r.pos < ops.len() {
            let o = ops[r.pos];
            // XDV glyph runs and native fonts: read here (the shared reader skips their glyphs)
            if o == 253 || o == 254 {
                r.pos += 1;
                let Some(()) = (|| -> Option<()> {
                    let mut text = None;
                    if o == 254 {
                        let l = r.u(2).ok()? as usize;
                        let units: Vec<u16> = r.bytes(l * 2).ok()?.chunks(2).map(|c| u16::from_be_bytes([c[0], c[1]])).collect();
                        text = Some(String::from_utf16_lossy(&units));
                    }
                    let wd = r.s(4).ok()? as f64;
                    let n = r.u(2).ok()? as usize;
                    let mut xy = Vec::with_capacity(n);
                    for _ in 0..n {
                        xy.push((r.s(4).ok()? as f64, r.s(4).ok()? as f64));
                    }
                    let current = font.and_then(|f| fonts.get(&f)).cloned();
                    let first = self.out.glyphs.len();
                    for (i, &(gx, gy)) in xy.iter().enumerate() {
                        let g = r.u(2).ok()?;
                        match &current {
                            Some(Font::Native { key, size, slant, extend }) => {
                                let fi = self.font_slot(key, *size, *slant, *extend);
                                let c = self.color_now();
                                self.out.glyphs.push((fi, (h + gx * scale) * self.unit, (v + gy * scale) * self.unit, g, c));
                                // a glyph's advance: to the next one's place, the last to the run's end
                                let next = xy.get(i + 1).map(|p| p.0).unwrap_or(wd);
                                self.out.advances.push(((next - gx) * scale * self.unit, if text.is_some() { TEXT } else { GLYPH }));
                            }
                            _ => self.out.undrawable.push("glyphs without a native font".into()),
                        }
                    }
                    if let Some(text) = text {
                        self.out.runs.push((first, self.out.glyphs.len() - first, text));
                    }
                    h += wd * scale;
                    Some(())
                })() else {
                    self.out.undrawable.push("bad glyph run".into());
                    return;
                };
                continue;
            }
            let Ok(op) = r.op() else {
                self.out.undrawable.push(format!("opcode {o}"));
                return;
            };
            use crate::dvi::Op;
            match op {
                Op::SetChar(c) | Op::PutChar(c) => {
                    let advance = self.char(fonts, font, c, h, v, depth);
                    if matches!(op, Op::SetChar(_)) {
                        h += advance;
                    }
                }
                Op::SetRule(a, b) | Op::PutRule(a, b) => {
                    let (a, b) = (a as f64 * scale, b as f64 * scale);
                    if a > 0.0 && b > 0.0 {
                        let c = self.color_now();
                        self.out.rules.push(([h * self.unit, (v - a) * self.unit, b * self.unit, a * self.unit], c));
                    }
                    if matches!(op, Op::SetRule(..)) {
                        h += b;
                    }
                }
                Op::Nop => {}
                Op::Push => stack.push((h, v, w, x, y, z)),
                Op::Pop => {
                    if let Some(s) = stack.pop() {
                        (h, v, w, x, y, z) = s;
                    }
                }
                Op::Right(d) => h += d as f64 * scale,
                Op::W0 => h += w,
                Op::W(d) => {
                    w = d as f64 * scale;
                    h += w;
                }
                Op::X0 => h += x,
                Op::X(d) => {
                    x = d as f64 * scale;
                    h += x;
                }
                Op::Down(d) => v += d as f64 * scale,
                Op::Y0 => v += y,
                Op::Y(d) => {
                    y = d as f64 * scale;
                    v += y;
                }
                Op::Z0 => v += z,
                Op::Z(d) => {
                    z = d as f64 * scale;
                    v += z;
                }
                Op::Fnt(k) => font = Some(k),
                Op::Xxx(text) => self.special(text, h * self.unit, v * self.unit),
                Op::FntDef(..) => {}
                Op::Glyphs(_) => {}
                Op::Pic => self.out.undrawable.push("picture".into()),
                Op::Dir(_) => self.out.undrawable.push("direction".into()),
            }
        }
    }

    /// A TFM character at (h, v): drawn from its font's outlines or its virtual font; its advance.
    fn char(&mut self, fonts: &HashMap<u32, Font>, font: Option<u32>, c: u32, h: f64, v: f64, depth: usize) -> f64 {
        let Some(f) = font.and_then(|f| fonts.get(&f)) else {
            self.out.undrawable.push("character without a font".into());
            return 0.0;
        };
        match f {
            Font::Native { .. } => {
                self.out.undrawable.push("character in a native font".into());
                0.0
            }
            Font::Tfm { name, tfm, scale, key, vf } => {
                let advance = tfm.width(c, *scale);
                if let Some((vf, local)) = vf {
                    match vf.chars.get(&c) {
                        Some(packet) => {
                            let local: HashMap<u32, Font> = local.iter().cloned().collect();
                            let first = vf.fonts.first().map(|(k, _)| *k);
                            // packet dimensions are fix_words of the virtual font's size
                            let s = *scale as f64 / (1u64 << 20) as f64;
                            let packet = packet.clone();
                            self.run(&packet, &local, first, h, v, s, depth + 1);
                        }
                        None => self.out.undrawable.push(format!("{name}: no character {c}")),
                    }
                } else if let Some((k, slant, extend)) = key {
                    let size = *scale as f64 * self.unit;
                    let fi = self.font_slot(k, size, *slant, *extend);
                    let col = self.color_now();
                    self.out.glyphs.push((fi, h * self.unit, v * self.unit, c, col));
                    self.out.advances.push((advance * self.unit, CHAR));
                } else {
                    self.out.undrawable.push(format!("{name}: no outlines"));
                }
                advance
            }
        }
    }
}

fn css_color(spec: &str) -> Option<String> {
    let w: Vec<&str> = spec.split_whitespace().collect();
    let n = |s: &str| s.parse::<f64>().ok().map(|v| v.clamp(0.0, 1.0));
    let (r, g, b) = match w.first().copied()? {
        "rgb" => (n(w.get(1)?)?, n(w.get(2)?)?, n(w.get(3)?)?),
        "gray" => {
            let g = n(w.get(1)?)?;
            (g, g, g)
        }
        "cmyk" => {
            let (c, m, y, k) = (n(w.get(1)?)?, n(w.get(2)?)?, n(w.get(3)?)?, n(w.get(4)?)?);
            ((1.0 - c) * (1.0 - k), (1.0 - m) * (1.0 - k), (1.0 - y) * (1.0 - k))
        }
        _ => return None,
    };
    Some(format!("#{:02x}{:02x}{:02x}", (r * 255.0).round() as u8, (g * 255.0).round() as u8, (b * 255.0).round() as u8))
}

/// The native font definition's own fields: file, face index, slant and extend (XDV 252).
fn native_def(raw: &[u8]) -> Option<(String, u32, f64, f64, i32)> {
    let mut r = Reader::new(raw);
    if r.u(1).ok()? != 252 {
        return None;
    }
    r.u(4).ok()?;
    let size = r.s(4).ok()?;
    let flags = r.u(2).ok()?;
    let len = r.u(1).ok()? as usize;
    let name = String::from_utf8_lossy(r.bytes(len).ok()?).to_string();
    let index = r.u(4).ok()?;
    let (mut extend, mut slant) = (1.0, 0.0);
    if flags & 0x0200 != 0 {
        r.u(4).ok()?;
    }
    if flags & 0x1000 != 0 {
        extend = r.s(4).ok()? as f64 / 65536.0;
    }
    if flags & 0x2000 != 0 {
        slant = r.s(4).ok()? as f64 / 65536.0;
    }
    if flags & 0x4000 != 0 {
        // emboldened: not drawn the same here
        return None;
    }
    if flags & 0x0100 != 0 {
        return None; // vertical
    }
    Some((name, index, slant, extend, size))
}

fn tfm_font(files: &mut FontFiles, name: &str, scale: i32, depth: usize) -> Option<Font> {
    let tfm = files.tfm(name)?;
    let key = files.map_entry(name).map(|e| {
        let kind = if e.font.to_ascii_lowercase().ends_with(".pfb") || e.font.to_ascii_lowercase().ends_with(".pfa") { "type1" } else { "otf" };
        (FontKey { kind: kind.into(), file: e.font, enc: e.enc, sub: 0 }, e.slant, e.extend)
    });
    let vf = if key.is_none() && depth < 4 {
        files.vf(name).map(|vf| {
            let local: Vec<(u32, Font)> = vf
                .fonts
                .iter()
                .filter_map(|(k, f)| {
                    // the local font's size: its scale (a fix_word of the virtual font's size)
                    let s = (f.scale as f64 * scale as f64 / (1u64 << 20) as f64).round() as i32;
                    let _ = (f.checksum, f.design, vf.design);
                    tfm_font(files, &f.name, s, depth + 1).map(|font| (*k, font))
                })
                .collect();
            (vf, local)
        })
    } else {
        None
    };
    Some(Font::Tfm { name: name.to_string(), tfm, scale, key, vf })
}

/// The glyphs and rules of a page body of `dvi` (its fonts as defined in the file so far).
pub fn draw_page(dvi: &Dvi, body: &[u8], files: &mut FontFiles) -> Drawn {
    let unit = dvi.unit_bp();
    let mut fonts: HashMap<u32, Font> = HashMap::new();
    let mut missing = vec![];
    for (k, def) in &dvi.fonts {
        match &def.kind {
            FontKind::Native { .. } => match native_def(&def.raw) {
                Some((name, index, slant, extend, size)) => {
                    let lower = name.to_ascii_lowercase();
                    let format = if lower.ends_with(".ttf") || lower.ends_with(".ttc") { "truetype fonts" } else { "opentype fonts" };
                    match files.find(&name, format) {
                        Some(path) => {
                            let key = FontKey { kind: "otf".into(), file: path.to_string_lossy().to_string(), enc: String::new(), sub: index };
                            fonts.insert(*k, Font::Native { key, size: size as f64 * unit, slant, extend });
                        }
                        None => missing.push(name),
                    }
                }
                None => missing.push(format!("font {k}")),
            },
            FontKind::Tfm { scale, name, .. } => match tfm_font(files, name, *scale, 0) {
                Some(f) => {
                    fonts.insert(*k, f);
                }
                None => missing.push(name.clone()),
            },
        }
    }
    let mut m = Machine { unit, out: Drawn { colors: vec!["#000000".into()], ..Default::default() }, font_index: HashMap::new(), color: vec![] };
    m.run(body, &fonts, None, 0.0, 0.0, 1.0, 0);
    // a font of the page that could not be read: only when the page used it (it shows as undrawable glyphs)
    let _ = missing;
    m.out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn colours_from_specials() {
        assert_eq!(css_color("rgb 0 0 1").as_deref(), Some("#0000ff"));
        assert_eq!(css_color("gray 0").as_deref(), Some("#000000"));
        assert_eq!(css_color("cmyk 0 0 0 1").as_deref(), Some("#000000"));
        assert_eq!(css_color("named Red"), None);
    }
}
