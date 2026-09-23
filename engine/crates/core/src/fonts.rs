//! Font bytes are injected by the host. No filesystem, network, DOM or Office
//! access is permitted here. MATH values and outlines use em units, y down.
use std::cell::RefCell;
use std::collections::BTreeMap;
use std::sync::Arc;
use serde::Serialize;
use sha2::{Digest, Sha256};
use ttf_parser::{Face, GlyphId, OutlineBuilder};
use crate::error::{EngineError, EngineResult, ErrorCode};

pub const MAX_FONT_BYTES: usize = 32 * 1024 * 1024;
// Complete app-private Latin/math/CJK catalog plus bounded optional system faces.
// Resources remain lazy, deduplicated by bytes+face, and explicitly budgeted.
const MAX_FONT_TOTAL: usize = 128 * 1024 * 1024;
// One product-facing family may contribute regular/italic/bold/bold-italic
// faces in addition to the structural MATH face and text/custom-symbol fonts.
// Keep the byte budget authoritative while allowing several user font choices
// to coexist in one long-lived desktop runtime without a false resource fault.
const MAX_FONTS: usize = 64;
const MAX_OUTLINE_OPS: usize = 16_384;
const MAX_CACHED_GLYPHS: usize = 512;

fn bad(message: impl Into<String>) -> EngineError {
    EngineError::new(ErrorCode::InvalidInput, message)
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FontIdentity {
    pub id: String,
    pub sha256: String,
    pub face_index: u32,
    pub units_per_em: u16,
    /// K0 uses the font's default axis coordinates and records them explicitly.
    /// Custom variation settings require a separate, versioned registration API.
    pub variation_axes: Vec<(String, f32)>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MathMetrics {
    pub script_scale: f64, pub scriptscript_scale: f64, pub axis: f64,
    pub display_operator_min_height: f64,
    pub fraction_rule: f64,
    pub numerator_shift: f64, pub numerator_display_shift: f64,
    pub denominator_shift: f64, pub denominator_display_shift: f64,
    pub numerator_gap: f64, pub numerator_display_gap: f64,
    pub denominator_gap: f64, pub denominator_display_gap: f64,
    pub stack_top_shift: f64, pub stack_top_display_shift: f64,
    pub stack_bottom_shift: f64, pub stack_bottom_display_shift: f64,
    pub stack_gap: f64, pub stack_display_gap: f64,
    pub skewed_horizontal_gap: f64, pub skewed_vertical_gap: f64,
    pub superscript_shift: f64, pub superscript_cramped_shift: f64,
    pub superscript_bottom_min: f64, pub superscript_drop_max: f64,
    pub subscript_shift: f64, pub subscript_top_max: f64, pub subscript_drop_min: f64,
    pub sub_sup_gap: f64, pub superscript_bottom_max_with_sub: f64,
    pub space_after_script: f64,
    pub radical_gap: f64, pub radical_display_gap: f64, pub radical_rule: f64,
    pub radical_extra_ascender: f64, pub radical_kern_before_degree: f64,
    pub radical_kern_after_degree: f64, pub radical_degree_raise: f64,
    pub accent_base_height: f64,
    pub overbar_gap: f64, pub overbar_rule: f64, pub overbar_extra: f64,
    pub underbar_gap: f64, pub underbar_rule: f64, pub underbar_extra: f64,
    pub upper_limit_gap: f64, pub upper_limit_rise: f64,
    pub lower_limit_gap: f64, pub lower_limit_drop: f64,
}

impl MathMetrics {
    fn read(face: &Face<'_>) -> Option<Self> {
        let c = face.tables().math?.constants?;
        let em = f64::from(face.units_per_em());
        Some(Self {
            script_scale: f64::from(c.script_percent_scale_down()) / 100.0,
            scriptscript_scale: f64::from(c.script_script_percent_scale_down()) / 100.0,
            axis: f64::from(c.axis_height().value) / em,
            display_operator_min_height: f64::from(c.display_operator_min_height()) / em,
            fraction_rule: f64::from(c.fraction_rule_thickness().value) / em,
            numerator_shift: f64::from(c.fraction_numerator_shift_up().value) / em,
            numerator_display_shift: f64::from(c.fraction_numerator_display_style_shift_up().value) / em,
            denominator_shift: f64::from(c.fraction_denominator_shift_down().value) / em,
            denominator_display_shift: f64::from(c.fraction_denominator_display_style_shift_down().value) / em,
            numerator_gap: f64::from(c.fraction_numerator_gap_min().value) / em,
            numerator_display_gap: f64::from(c.fraction_num_display_style_gap_min().value) / em,
            denominator_gap: f64::from(c.fraction_denominator_gap_min().value) / em,
            denominator_display_gap: f64::from(c.fraction_denom_display_style_gap_min().value) / em,
            stack_top_shift: f64::from(c.stack_top_shift_up().value) / em,
            stack_top_display_shift: f64::from(c.stack_top_display_style_shift_up().value) / em,
            stack_bottom_shift: f64::from(c.stack_bottom_shift_down().value) / em,
            stack_bottom_display_shift: f64::from(c.stack_bottom_display_style_shift_down().value) / em,
            stack_gap: f64::from(c.stack_gap_min().value) / em,
            stack_display_gap: f64::from(c.stack_display_style_gap_min().value) / em,
            skewed_horizontal_gap: f64::from(c.skewed_fraction_horizontal_gap().value) / em,
            skewed_vertical_gap: f64::from(c.skewed_fraction_vertical_gap().value) / em,
            superscript_shift: f64::from(c.superscript_shift_up().value) / em,
            superscript_cramped_shift: f64::from(c.superscript_shift_up_cramped().value) / em,
            superscript_bottom_min: f64::from(c.superscript_bottom_min().value) / em,
            superscript_drop_max: f64::from(c.superscript_baseline_drop_max().value) / em,
            subscript_shift: f64::from(c.subscript_shift_down().value) / em,
            subscript_top_max: f64::from(c.subscript_top_max().value) / em,
            subscript_drop_min: f64::from(c.subscript_baseline_drop_min().value) / em,
            sub_sup_gap: f64::from(c.sub_superscript_gap_min().value) / em,
            superscript_bottom_max_with_sub: f64::from(c.superscript_bottom_max_with_subscript().value) / em,
            space_after_script: f64::from(c.space_after_script().value) / em,
            radical_gap: f64::from(c.radical_vertical_gap().value) / em,
            radical_display_gap: f64::from(c.radical_display_style_vertical_gap().value) / em,
            radical_rule: f64::from(c.radical_rule_thickness().value) / em,
            radical_extra_ascender: f64::from(c.radical_extra_ascender().value) / em,
            radical_kern_before_degree: f64::from(c.radical_kern_before_degree().value) / em,
            radical_kern_after_degree: f64::from(c.radical_kern_after_degree().value) / em,
            radical_degree_raise: f64::from(c.radical_degree_bottom_raise_percent()) / 100.0,
            accent_base_height: f64::from(c.accent_base_height().value) / em,
            overbar_gap: f64::from(c.overbar_vertical_gap().value) / em,
            overbar_rule: f64::from(c.overbar_rule_thickness().value) / em,
            overbar_extra: f64::from(c.overbar_extra_ascender().value) / em,
            underbar_gap: f64::from(c.underbar_vertical_gap().value) / em,
            underbar_rule: f64::from(c.underbar_rule_thickness().value) / em,
            underbar_extra: f64::from(c.underbar_extra_descender().value) / em,
            upper_limit_gap: f64::from(c.upper_limit_gap_min().value) / em,
            upper_limit_rise: f64::from(c.upper_limit_baseline_rise_min().value) / em,
            lower_limit_gap: f64::from(c.lower_limit_gap_min().value) / em,
            lower_limit_drop: f64::from(c.lower_limit_baseline_drop_min().value) / em,
        })
    }
    fn validate(&self) -> EngineResult<()> {
        if !(0.0 < self.scriptscript_scale && self.scriptscript_scale <= self.script_scale
            && self.script_scale <= 1.0 && self.fraction_rule > 0.0
            && self.axis >= 0.0 && self.sub_sup_gap >= 0.0
            && self.radical_rule > 0.0 && self.radical_gap >= 0.0
            && self.radical_display_gap >= 0.0 && self.radical_extra_ascender >= 0.0
            && (0.0..=1.0).contains(&self.radical_degree_raise))
        { return Err(bad("Font contains invalid required MATH constants")); }
        Ok(())
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(tag = "op", content = "points", rename_all = "camelCase")]
pub enum PathOp {
    Move([f64; 2]), Line([f64; 2]), Quad([f64; 4]), Cubic([f64; 6]), Close,
}
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct GlyphShape {
    pub glyph_id: u16,
    pub advance: f64,
    /// xMin, yMin, xMax, yMax; y increases downward. None for whitespace.
    pub ink: Option<[f64; 4]>,
    pub italic_correction: f64,
    pub path: Vec<PathOp>,
}
struct Outline { ops: Vec<PathOp>, scale: f64, overflow: bool }
impl Outline {
    fn push(&mut self, op: PathOp) {
        if self.ops.len() >= MAX_OUTLINE_OPS { self.overflow = true; }
        else { self.ops.push(op); }
    }
}
impl OutlineBuilder for Outline {
    fn move_to(&mut self, x: f32, y: f32) { self.push(PathOp::Move([f64::from(x)*self.scale, -f64::from(y)*self.scale])); }
    fn line_to(&mut self, x: f32, y: f32) { self.push(PathOp::Line([f64::from(x)*self.scale, -f64::from(y)*self.scale])); }
    fn quad_to(&mut self, x1: f32, y1: f32, x: f32, y: f32) { self.push(PathOp::Quad([f64::from(x1)*self.scale, -f64::from(y1)*self.scale, f64::from(x)*self.scale, -f64::from(y)*self.scale])); }
    fn curve_to(&mut self, x1: f32, y1: f32, x2: f32, y2: f32, x: f32, y: f32) { self.push(PathOp::Cubic([f64::from(x1)*self.scale, -f64::from(y1)*self.scale, f64::from(x2)*self.scale, -f64::from(y2)*self.scale, f64::from(x)*self.scale, -f64::from(y)*self.scale])); }
    fn close(&mut self) { self.push(PathOp::Close); }
}

pub struct FontAsset {
    bytes: Arc<[u8]>,
    pub identity: FontIdentity,
    pub math: Option<MathMetrics>,
    pub ascent: f64,
    pub descent: f64,
    glyph_cache: RefCell<BTreeMap<u16, Arc<GlyphShape>>>,
}
impl FontAsset {
    pub fn from_bytes(bytes: &[u8], face_index: u32) -> EngineResult<Self> {
        if bytes.is_empty() || bytes.len() > MAX_FONT_BYTES {
            return Err(EngineError::new(ErrorCode::ResourceLimit, "Font byte budget exceeded or empty font"));
        }
        let face = Face::parse(bytes, face_index).map_err(|_| bad("Invalid OpenType font or collection face index"))?;
        let variation_axes = face.variation_axes().into_iter()
            .map(|axis| (axis.tag.to_string(), axis.def_value)).collect();
        let upem = face.units_per_em();
        let math = MathMetrics::read(&face);
        if let Some(math) = &math { math.validate()?; }
        let sha256 = format!("{:x}", Sha256::digest(bytes));
        let identity = FontIdentity { id: format!("{sha256}:{face_index}"), sha256, face_index,
            units_per_em: upem, variation_axes };
        Ok(Self { bytes: Arc::from(bytes), identity, math,
            ascent: f64::from(face.ascender()).max(0.0)/f64::from(upem),
            descent: (-f64::from(face.descender())).max(0.0)/f64::from(upem),
            glyph_cache: RefCell::new(BTreeMap::new()) })
    }
    pub(crate) fn face(&self) -> EngineResult<Face<'_>> {
        Face::parse(&self.bytes, self.identity.face_index).map_err(|_| bad("Previously validated font no longer parses"))
    }
    pub fn glyph_index(&self, ch: char) -> EngineResult<Option<u16>> {
        Ok(self.face()?.glyph_index(ch).filter(|id| id.0 != 0).map(|id| id.0))
    }
    pub fn glyph_id(&self, ch: char) -> EngineResult<u16> {
        self.glyph_index(ch)?.ok_or_else(|| bad(format!("Font {} is missing glyph U+{:04X}", self.identity.id, u32::from(ch))))
    }
    /// A display operator uses a designed vertical variant. It is not stretched
    /// to the operand's height and never borrows a delimiter assembly.
    pub fn operator_glyph(&self,id:u16,display:bool)->EngineResult<Arc<GlyphShape>> {
        let original=self.glyph(id)?;
        if !display{return Ok(original);}
        let target=self.math.as_ref().ok_or_else(||bad("Operator font lacks MATH constants"))?.display_operator_min_height;
        let face=self.face()?;
        let construction=face.tables().math.and_then(|m|m.variants).and_then(|v|v.vertical_constructions.get(GlyphId(id)));
        let mut candidates=vec![original];
        if let Some(construction)=construction {
            if construction.variants.len()>256{return Err(EngineError::new(ErrorCode::ResourceLimit,"Operator glyph variant budget exceeded"));}
            for variant in construction.variants{candidates.push(self.glyph(variant.variant_glyph.0)?);}
        }
        let height=|g:&GlyphShape|g.ink.map_or(0.0,|b|b[3]-b[1]);
        // When the font has no sufficiently tall variant, retain the largest
        // designed glyph. No anisotropic or arbitrary enlarged substitute.
        Ok(candidates.iter().filter(|g|height(g)+1e-9>=target)
            .min_by(|a,b|height(a).total_cmp(&height(b)).then(a.glyph_id.cmp(&b.glyph_id)))
            .or_else(||candidates.iter().max_by(|a,b|height(a).total_cmp(&height(b))))
            .cloned().ok_or_else(||bad("Operator has no glyph"))?)
    }
    pub fn accent_attachment(&self, id: u16) -> EngineResult<f64> {
        let shape = self.glyph(id)?;
        Ok(self.face()?.tables().math.and_then(|m| m.glyph_info)
            .and_then(|g| g.top_accent_attachments).and_then(|values| values.get(GlyphId(id)))
            .map_or(shape.advance / 2.0, |value| f64::from(value.value) / f64::from(self.identity.units_per_em)))
    }
    pub fn glyph(&self, id: u16) -> EngineResult<Arc<GlyphShape>> {
        if let Some(glyph) = self.glyph_cache.borrow().get(&id) { return Ok(Arc::clone(glyph)); }
        let face = self.face()?;
        if id == 0 || id >= face.number_of_glyphs() { return Err(bad("Invalid or missing glyph id")); }
        let gid = GlyphId(id);
        let em = f64::from(self.identity.units_per_em);
        let advance = f64::from(face.glyph_hor_advance(gid).ok_or_else(|| bad("Glyph lacks horizontal metrics"))?)/em;
        let mut outline = Outline { ops: Vec::new(), scale: 1.0/em, overflow: false };
        let rect = face.outline_glyph(gid, &mut outline);
        if outline.overflow { return Err(EngineError::new(ErrorCode::ResourceLimit, "Glyph outline budget exceeded")); }
        let ink = rect.map(|r| [f64::from(r.x_min)/em, -f64::from(r.y_max)/em, f64::from(r.x_max)/em, -f64::from(r.y_min)/em]);
        if rect.is_none() && !outline.ops.is_empty() { return Err(bad("Font outline is incomplete")); }
        let italic_correction = face.tables().math.and_then(|m| m.glyph_info)
            .and_then(|g| g.italic_corrections).and_then(|v| v.get(gid))
            .map(|v| f64::from(v.value)/em)
            .unwrap_or_else(|| ink.map_or(0.0, |bounds| (bounds[2]-advance).max(0.0)));
        let glyph = Arc::new(GlyphShape { glyph_id: id, advance, ink, italic_correction, path: outline.ops });
        let mut cache = self.glyph_cache.borrow_mut();
        if cache.len() >= MAX_CACHED_GLYPHS { cache.clear(); }
        cache.insert(id, Arc::clone(&glyph));
        Ok(glyph)
    }
    pub fn shape_text(&self, text: &str) -> EngineResult<ShapedRun> {
        if text.len() > 65_536 { return Err(EngineError::new(ErrorCode::ResourceLimit, "Text shaping budget exceeded")); }
        // Bidi paragraph resolution is a separate host/core capability. Never
        // silently lay out an RTL or mixed-direction text run as left-to-right.
        if text.chars().any(|c| matches!(u32::from(c), 0x0590..=0x08FF | 0x202A..=0x202E | 0x2066..=0x2069 | 0xFB1D..=0xFDFF | 0xFE70..=0xFEFF)) {
            return Err(EngineError::new(ErrorCode::UnsupportedSyntax, "K0 text shaping does not yet support bidirectional text"));
        }
        let face = rustybuzz::Face::from_slice(&self.bytes, self.identity.face_index).ok_or_else(|| bad("Text shaper rejected font"))?;
        let mut buffer = rustybuzz::UnicodeBuffer::new();
        buffer.push_str(text); buffer.guess_segment_properties();
        let features = ["liga=0".parse().unwrap(), "clig=0".parse().unwrap()];
        let shaped = rustybuzz::shape(&face, &features, buffer);
        if shaped.len() > 16_384 { return Err(EngineError::new(ErrorCode::ResourceLimit, "Shaped glyph budget exceeded")); }
        let em = f64::from(self.identity.units_per_em);
        let mut x = 0.0;
        let mut glyphs = Vec::with_capacity(shaped.len());
        for (info, pos) in shaped.glyph_infos().iter().zip(shaped.glyph_positions()) {
            if info.glyph_id == 0 { return Err(bad(format!("Text font {} has missing glyphs; no implicit font fallback", self.identity.id))); }
            if pos.y_advance != 0 || pos.x_advance < 0 { return Err(bad("K0 expects horizontal forward text shaping")); }
            let shape = self.glyph(u16::try_from(info.glyph_id).map_err(|_| bad("Glyph id overflow"))?)?;
            glyphs.push(ShapedGlyph { shape, cluster: info.cluster,
                x: x + f64::from(pos.x_offset)/em, y: -f64::from(pos.y_offset)/em,
                pen_x: x, advance: f64::from(pos.x_advance)/em });
            x += f64::from(pos.x_advance)/em;
        }
        Ok(ShapedRun { glyphs, advance: x })
    }
}

pub struct ShapedGlyph { pub shape: Arc<GlyphShape>, pub cluster: u32, pub x: f64, pub y: f64, pub pen_x: f64, pub advance: f64 }
pub struct ShapedRun { pub glyphs: Vec<ShapedGlyph>, pub advance: f64 }

#[derive(Default)]
pub struct FontRegistry { assets: BTreeMap<String, FontAsset>, total_bytes: usize }
impl FontRegistry {
    pub fn register(&mut self, bytes: &[u8], face_index: u32) -> EngineResult<FontIdentity> {
        // Hash once per host injection, not per glyph or per keystroke.
        let asset = FontAsset::from_bytes(bytes, face_index)?;
        if let Some(existing) = self.assets.get(&asset.identity.id) { return Ok(existing.identity.clone()); }
        if self.assets.len() >= MAX_FONTS || self.total_bytes.saturating_add(bytes.len()) > MAX_FONT_TOTAL {
            return Err(EngineError::new(ErrorCode::ResourceLimit, "Font registry budget exceeded"));
        }
        let identity = asset.identity.clone();
        self.total_bytes += bytes.len(); self.assets.insert(identity.id.clone(), asset);
        Ok(identity)
    }
    pub fn get(&self, id: &str) -> EngineResult<&FontAsset> {
        self.assets.get(id).ok_or_else(|| bad("Font resource identity is not registered in this runtime"))
    }
    pub fn len(&self) -> usize { self.assets.len() }
    pub fn is_empty(&self) -> bool { self.assets.is_empty() }
}
