//! Glyph outlines for drawing live paragraphs (live.rs) in the page's own fonts: OpenType and TrueType files
//! (glyph indices, through ttf-parser) and Type 1 files (TFM slots through the map file's encoding, type1.rs).
//! Paths are SVG path data in font units, y up; `scale` turns font units into bp per bp of font size.

use std::collections::HashMap;
use std::path::Path;

use crate::type1::{self, Type1};

enum Face {
    Otf { data: &'static [u8], index: u32, upem: f64 },
    Type1 { font: Type1, enc: Option<Vec<Option<String>>> },
}

/// A font file the paths come from: kind ("otf" | "type1"), file, encoding file, face index.
#[derive(Clone, Debug, Hash, PartialEq, Eq)]
pub struct FontKey {
    pub kind: String,
    pub file: String,
    pub enc: String,
    pub sub: u32,
}

impl FontKey {
    /// A short stable name for the key (the front end keeps paths by it).
    pub fn id(&self) -> String {
        let mut h = crate::util::Fnv::default();
        h.write(self.kind.as_bytes());
        h.write(self.file.as_bytes());
        h.write(self.enc.as_bytes());
        h.write_u64(self.sub as u64);
        format!("{:012x}", h.finish() & 0xffff_ffff_ffff)
    }
}

#[derive(Default)]
pub struct Outlines {
    faces: HashMap<FontKey, Option<Face>>,
    paths: HashMap<(FontKey, u32), Option<String>>,
    /// an OpenType font's characters by glyph (its cmap read backwards)
    unicode: HashMap<FontKey, HashMap<u16, char>>,
}

/// Of two characters one glyph stands for, the one text has: not a compatibility form (Kangxi radicals, CJK
/// compatibility ideographs, presentation forms), else the lower.
fn better(a: char, b: char) -> bool {
    let compat = |c: char| matches!(c as u32, 0x2E80..=0x2FDF | 0xF900..=0xFAFF | 0xFB00..=0xFB4F | 0xFE30..=0xFE4F | 0xFF00..=0xFFEF | 0x2F800..=0x2FA1F | 0x1D400..=0x1D7FF);
    match (compat(a), compat(b)) {
        (false, true) => true,
        (true, false) => false,
        _ => a < b,
    }
}

struct PathOut(String);

fn num(v: f32) -> String {
    let r = (v * 10.0).round() / 10.0;
    if r == r.trunc() {
        format!("{}", r as i64)
    } else {
        format!("{r:.1}")
    }
}

impl ttf_parser::OutlineBuilder for PathOut {
    fn move_to(&mut self, x: f32, y: f32) {
        self.0.push_str(&format!("M{} {}", num(x), num(y)));
    }
    fn line_to(&mut self, x: f32, y: f32) {
        self.0.push_str(&format!("L{} {}", num(x), num(y)));
    }
    fn quad_to(&mut self, x1: f32, y1: f32, x: f32, y: f32) {
        self.0.push_str(&format!("Q{} {} {} {}", num(x1), num(y1), num(x), num(y)));
    }
    fn curve_to(&mut self, x1: f32, y1: f32, x2: f32, y2: f32, x: f32, y: f32) {
        self.0.push_str(&format!("C{} {} {} {} {} {}", num(x1), num(y1), num(x2), num(y2), num(x), num(y)));
    }
    fn close(&mut self) {
        self.0.push('Z');
    }
}

impl Outlines {
    fn face(&mut self, key: &FontKey) -> Option<&Face> {
        if !self.faces.contains_key(key) {
            let face = load(key);
            self.faces.insert(key.clone(), face);
        }
        self.faces.get(key).and_then(|f| f.as_ref())
    }

    /// Font units per bp of font size: (x, y) scale of the path data.
    pub fn scale(&mut self, key: &FontKey) -> Option<(f64, f64)> {
        match self.face(key)? {
            Face::Otf { upem, .. } => Some((1.0 / upem, 1.0 / upem)),
            Face::Type1 { font, .. } => Some((font.matrix[0], font.matrix[3])),
        }
    }

    /// The characters glyph `code` (a glyph index, or a slot of the TFM font) stands for ("" when unknown): an
    /// OpenType glyph's from the font's cmap, a Type 1 character's from its glyph name.
    pub fn chars(&mut self, key: &FontKey, code: u32) -> String {
        if key.kind == "otf" {
            if !self.unicode.contains_key(key) {
                let mut map: HashMap<u16, char> = HashMap::new();
                if let Some(Face::Otf { data, index, .. }) = self.face(key) {
                    if let Ok(face) = ttf_parser::Face::parse(data, *index) {
                        if let Some(cmap) = face.tables().cmap {
                            for sub in cmap.subtables {
                                if !sub.is_unicode() {
                                    continue;
                                }
                                sub.codepoints(|cp| {
                                    let (Some(c), Some(g)) = (char::from_u32(cp), sub.glyph_index(cp)) else { return };
                                    match map.get(&g.0) {
                                        Some(&old) if !better(c, old) => {}
                                        _ => {
                                            map.insert(g.0, c);
                                        }
                                    }
                                });
                            }
                        }
                    }
                }
                self.unicode.insert(key.clone(), map);
            }
            return self.unicode.get(key).and_then(|m| m.get(&(code as u16))).map(|c| c.to_string()).unwrap_or_default();
        }
        match self.face(key) {
            Some(Face::Type1 { font, enc }) => {
                let names = enc.as_deref().unwrap_or(&font.encoding);
                names.get(code as usize).and_then(|n| n.as_deref()).map(crate::agl::chars).unwrap_or_default()
            }
            _ => String::new(),
        }
    }

    /// The outline of glyph `code` (a glyph index, or a slot of the TFM font): SVG path data, None when the
    /// font or glyph cannot be read ("" for a glyph with no ink, a space).
    pub fn path(&mut self, key: &FontKey, code: u32) -> Option<String> {
        let k = (key.clone(), code);
        if let Some(p) = self.paths.get(&k) {
            return p.clone();
        }
        let path = match self.face(key) {
            Some(Face::Otf { data, index, .. }) => ttf_parser::Face::parse(data, *index).ok().map(|face| {
                let mut out = PathOut(String::new());
                let _ = face.outline_glyph(ttf_parser::GlyphId(code as u16), &mut out);
                out.0
            }),
            Some(Face::Type1 { font, enc }) => font.path_by_code(code, enc.as_deref()),
            None => None,
        };
        self.paths.insert(k, path.clone());
        path
    }
}

fn load(key: &FontKey) -> Option<Face> {
    let bytes = std::fs::read(Path::new(&key.file)).ok()?;
    match key.kind.as_str() {
        "otf" => {
            // fonts stay for the app's life (one copy each, read once)
            let data: &'static [u8] = Box::leak(bytes.into_boxed_slice());
            let face = ttf_parser::Face::parse(data, key.sub).ok()?;
            let upem = face.units_per_em() as f64;
            Some(Face::Otf { data, index: key.sub, upem: if upem > 0.0 { upem } else { 1000.0 } })
        }
        "type1" => {
            let font = Type1::parse(&bytes)?;
            let enc = if key.enc.is_empty() {
                None
            } else {
                Some(type1::parse_encoding(&String::from_utf8_lossy(&std::fs::read(&key.enc).ok()?)))
            };
            Some(Face::Type1 { font, enc })
        }
        _ => None,
    }
}
