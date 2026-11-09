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
