//! Font resources, not an imported typesetter. TeX-family faces encode their
//! alphabets at ordinary character positions and their size variants in separate
//! fonts. All geometry is still computed by VisualTeX and all painted glyphs
//! retain the SHA/face identity of the actual font that supplied the outline.
use std::collections::BTreeMap;
use crate::fonts::FontAsset;
use crate::model::AtomClass;
use crate::error::EngineResult;
use crate::styles::MathVariant;
use super::boxes::Builder;

pub(super) const ROLES: &[&str] = &[
    "main", "mainItalic", "mainBold", "mainBoldItalic", "mathItalic", "mathBoldItalic",
    "ams", "calligraphic", "calligraphicBold", "script", "fraktur", "frakturBold",
    "sans", "sansBold", "sansItalic", "monospace", "size1", "size2", "size3", "size4",
];
pub(super) type ResolvedTexFonts<'a> = BTreeMap<String, &'a FontAsset>;

impl<'a> Builder<'a> {
    pub(super) fn tex_glyph(&self, role: &str, ch: char) -> EngineResult<Option<(&'a FontAsset, u16)>> {
        let Some(font) = self.tex_fonts.get(role) else { return Ok(None) };
        Ok(font.glyph_index(ch)?.map(|id| (*font, id)))
    }

    pub(super) fn tex_symbol_glyph(&self, ch: char, class: AtomClass) -> EngineResult<Option<(&'a FontAsset, u16)>> {
        if self.tex_fonts.is_empty() { return Ok(None) }
        use MathVariant::*;
        let letter = ch.is_ascii_alphabetic() || matches!(u32::from(ch), 0x0370..=0x03ff);
        let math_letter = class == AtomClass::Ordinary && letter;
        let lower_greek = matches!(u32::from(ch), 0x03b1..=0x03c9 | 0x03d1 | 0x03d5 | 0x03d6 | 0x03f0 | 0x03f1 | 0x03f5);
        let role = match self.variant {
            // Match MathLive 0.109.2's KaTeX font-class policy without importing
            // MathLive itself. The VisualTeX core still owns syntax, layout,
            // cursor movement and editing; this table only chooses glyph faces.
            Default if class == AtomClass::Ordinary && (ch.is_ascii_alphabetic() || lower_greek) => "mathItalic",
            Default | Upright => "main",
            // MathLive's .ML__mathit uses KaTeX_Math, not KaTeX_Main-Italic.
            Italic if math_letter => "mathItalic",
            // TeX's idiosyncratic lowercase Greek \mathbf stays in the math
            // family and selects its bold face; Latin/uppercase Greek use Main.
            Bold if lower_greek => "mathBoldItalic",
            Bold if letter || ch.is_ascii_digit() => "mainBold",
            BoldItalic if math_letter => "mathBoldItalic",
            BoldItalic if ch.is_ascii_digit() => "mainBoldItalic",
            Sans if ch.is_ascii_alphanumeric() => "sans",
            SansBold if ch.is_ascii_alphanumeric() => "sansBold",
            SansItalic if ch.is_ascii_alphanumeric() => "sansItalic",
            Monospace if ch.is_ascii_alphanumeric() => "monospace",
            DoubleStruck if ch.is_ascii_uppercase() => "ams",
            Script if ch.is_ascii_uppercase() => if self.script_alphabet { "script" } else { "calligraphic" },
            BoldScript if ch.is_ascii_uppercase() => "calligraphicBold",
            Fraktur if ch.is_ascii_alphanumeric() => "fraktur",
            BoldFraktur if ch.is_ascii_alphanumeric() => "frakturBold",
            // Do not synthesize MathLive editing or font-style behavior here.
            // Unsupported semantic alphabets fall through to the core's real
            // mathematical Unicode/font fallback path.
            _ if ch.is_ascii_alphanumeric() || letter => return Ok(None),
            _ => "main",
        };
        if let Some(value) = self.tex_glyph(role, ch)? { return Ok(Some(value)) }
        // Math-Italic includes the lowercase Greek repertoire absent from Main.
        if lower_greek && matches!(self.variant, Default | Upright | Bold | Italic) {
            if let Some(value) = self.tex_glyph(if self.variant == Bold {"mathBoldItalic"} else {"mathItalic"}, ch)? {
                return Ok(Some(value))
            }
        }
        if !ch.is_ascii_alphanumeric() && !letter {
            // Extended symbols in AMS are at their Unicode values. AMS ASCII
            // capitals are blackboard symbols, never a generic text fallback.
            if let Some(value) = self.tex_glyph("ams", ch)? { return Ok(Some(value)) }
        }
        Ok(None)
    }

    pub(super) fn tex_vertical_variant(&self, ch: char, target: f64) -> EngineResult<Option<(&'a FontAsset, crate::stretch::StretchedGlyph)>> {
        let mut best: Option<(&FontAsset, crate::stretch::StretchedGlyph)> = None;
        for role in ["main", "size1", "size2", "size3", "size4"] {
            let Some((font, gid)) = self.tex_glyph(role, ch)? else { continue };
            let shape = font.glyph(gid)?;
            let Some(ink) = shape.ink else { continue };
            let height = ink[3] - ink[1];
            if height + 1e-9 < target || best.as_ref().is_some_and(|(_, b)| b.height <= height) { continue }
            best = Some((font, crate::stretch::StretchedGlyph {
                width: shape.advance, height, italic_correction: shape.italic_correction,
                pieces: vec![crate::stretch::StretchPiece { y: -ink[1], shape }],
                assembled: false, overlaps: Vec::new(), minimum_overlap: 0.0,
            }));
        }
        Ok(best)
    }
}
