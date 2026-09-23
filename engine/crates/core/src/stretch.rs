//! Font-driven vertical glyph construction, shared by radicals and later fences.
//! OpenType MATH variants/assemblies provide data, not a rendering engine.
//! Pieces retain their real outlines and uniform em scale; no anisotropic stretch.
use std::sync::Arc;
use ttf_parser::{GlyphId, math::GlyphPart};
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::fonts::{FontAsset, GlyphShape};

const MAX_VARIANTS: usize = 256;
const MAX_PARTS: usize = 1024;
const EPS: f64 = 1e-9;

#[derive(Debug, Clone)]
pub struct StretchPiece {
    pub shape: Arc<GlyphShape>,
    /// Baseline relative to the assembly's ink top, in em, y down.
    pub y: f64,
}
#[derive(Debug, Clone)]
pub struct StretchedGlyph {
    pub pieces: Vec<StretchPiece>,
    pub width: f64,
    pub height: f64,
    pub italic_correction: f64,
    pub assembled: bool,
    pub overlaps: Vec<f64>,
    pub minimum_overlap: f64,
}
#[derive(Debug, Clone)]
pub struct HorizontalPiece { pub shape: Arc<GlyphShape>, pub x: f64 }
#[derive(Debug, Clone)]
pub struct HorizontalGlyph {
    pub pieces: Vec<HorizontalPiece>,
    pub width: f64,
    pub ink_top: f64,
    pub ink_bottom: f64,
    pub attachment: f64,
    /// False only under the explicit accent policy when the font exhausts its
    /// designed variants and has no assembly. No outline is rescaled.
    pub target_met: bool,
    pub assembled: bool,
    pub overlaps: Vec<f64>,
    pub minimum_overlap: f64,
}
fn invalid(message: &str) -> EngineError { EngineError::new(ErrorCode::InvalidInput, message) }
fn limit(message: &str) -> EngineError { EngineError::new(ErrorCode::ResourceLimit, message) }

/// Stretch on the horizontal axis using designed glyphs or real connector parts.
/// x is normalized to the first ink edge while baseline y stays font-defined.
pub fn horizontal(font: &FontAsset, glyph_id: u16, target: f64) -> EngineResult<HorizontalGlyph> {
    horizontal_impl(font, glyph_id, target, false)
}
/// MathML Core 5.3.2 final variant rule for accents without an assembly.
/// This is deliberately separate from strict connector/surd sizing.
pub fn horizontal_accent(font: &FontAsset, glyph_id: u16, target: f64) -> EngineResult<HorizontalGlyph> {
    horizontal_impl(font, glyph_id, target, true)
}
fn horizontal_impl(font: &FontAsset, glyph_id: u16, target: f64, allow_last_variant: bool) -> EngineResult<HorizontalGlyph> {
    if !target.is_finite() || target <= 0.0 { return Err(invalid("Invalid horizontal stretch target")); }
    if target > 2048.0 { return Err(limit("Horizontal stretch target exceeds its budget")); }
    let face = font.face()?;
    let em = f64::from(face.units_per_em());
    let variants = face.tables().math.and_then(|m| m.variants);
    let construction = variants.and_then(|v| v.horizontal_constructions.get(GlyphId(glyph_id)));
    let mut candidates = vec![font.glyph(glyph_id)?];
    let width = |shape: &GlyphShape| shape.ink.map_or(0.0, |b| b[2]-b[0]);
    if let Some(c) = &construction {
        if usize::from(c.variants.len()) > MAX_VARIANTS { return Err(limit("Too many horizontal variants")); }
        for v in c.variants { candidates.push(font.glyph(v.variant_glyph.0)?); }
    }
    let fitting = candidates.iter().filter(|s| width(s)+EPS >= target)
        .min_by(|a,b| width(a).total_cmp(&width(b)).then(a.glyph_id.cmp(&b.glyph_id))).cloned();
    let last_variant = if allow_last_variant && construction.as_ref().is_some_and(|c|c.assembly.is_none()) {
        candidates.iter().max_by(|a,b|width(a).total_cmp(&width(b))).cloned()
    } else { None };
    if let Some(shape) = fitting.or(last_variant) {
        let ink = shape.ink.ok_or_else(|| invalid("Horizontal variant has no outline"))?;
        return Ok(HorizontalGlyph { width: ink[2]-ink[0], ink_top: ink[1], ink_bottom: ink[3],
            attachment: font.accent_attachment(shape.glyph_id)?-ink[0],
            pieces: vec![HorizontalPiece { x: -ink[0], shape }], target_met: ink[2]-ink[0]+EPS >= target, assembled: false,
            overlaps: Vec::new(), minimum_overlap: 0.0 });
    }
    let assembly = construction.and_then(|c| c.assembly)
        .ok_or_else(|| invalid("Font has no sufficiently wide glyph or horizontal assembly"))?;
    let min_overlap = f64::from(variants.unwrap().min_connector_overlap)/em;
    let parts: Vec<GlyphPart> = assembly.parts.into_iter().collect();
    if parts.is_empty() || parts.len()>MAX_PARTS { return Err(limit("Horizontal assembly part budget exceeded")); }
    let mut fixed_count=0usize;let mut ext_count=0usize;let mut fixed_size=0.0;let mut ext_size=0.0;
    for part in &parts {
        if part.full_advance==0 || part.start_connector_length>part.full_advance || part.end_connector_length>part.full_advance {
            return Err(invalid("Invalid horizontal MATH connectors"));
        }
        if part.part_flags.extender() { ext_count+=1;ext_size+=f64::from(part.full_advance)/em; }
        else { fixed_count+=1;fixed_size+=f64::from(part.full_advance)/em; }
    }
    let growth=ext_size-min_overlap*ext_count as f64;
    if ext_count==0 || growth<=0.0 { return Err(invalid("Horizontal assembly cannot grow")); }
    let repeats=((target-fixed_size+min_overlap*(fixed_count as f64-1.0))/growth).ceil().max(0.0);
    if repeats>MAX_PARTS as f64 { return Err(limit("Horizontal assembly repetition budget exceeded")); }
    let count=fixed_count+ext_count*repeats as usize;
    if count==0 || count>MAX_PARTS { return Err(limit("Horizontal assembly glyph budget exceeded")); }
    let expanded:Vec<&GlyphPart>=parts.iter().flat_map(|p|std::iter::repeat_n(p,if p.part_flags.extender(){repeats as usize}else{1})).collect();
    let max_overlap=expanded.windows(2).map(|p|f64::from(p[0].end_connector_length.min(p[1].start_connector_length))/em).fold(f64::INFINITY,f64::min);
    if count>1 && max_overlap+EPS<min_overlap { return Err(invalid("Horizontal connectors cannot overlap safely")); }
    let full: f64=expanded.iter().map(|p|f64::from(p.full_advance)/em).sum();
    let overlap=if count>1 {((full-target)/(count-1)as f64).clamp(min_overlap,max_overlap)}else{0.0};
    let mut x=0.0;let mut left=f64::INFINITY;let mut right=f64::NEG_INFINITY;
    let mut top=f64::INFINITY;let mut bottom=f64::NEG_INFINITY;let mut pieces=Vec::with_capacity(count);
    for (i,part) in expanded.iter().enumerate() {
        let shape=font.glyph(part.glyph_id.0)?;let ink=shape.ink.ok_or_else(||invalid("Horizontal part has no outline"))?;
        left=left.min(x+ink[0]);right=right.max(x+ink[2]);top=top.min(ink[1]);bottom=bottom.max(ink[3]);
        pieces.push(HorizontalPiece {shape,x});
        if i+1<count {x+=f64::from(part.full_advance)/em-overlap;}
    }
    let width=right-left;
    if !width.is_finite() || width+EPS<target { return Err(invalid("Horizontal assembly ink misses the requested extent")); }
    for piece in &mut pieces {piece.x-=left;}
    Ok(HorizontalGlyph {pieces,width,ink_top:top,ink_bottom:bottom,attachment:width/2.0,
        target_met:true,assembled:true,overlaps:vec![overlap;count-1],minimum_overlap:min_overlap})
}

/// Select the smallest usable designed variant, then assemble real glyph parts.
/// Missing data fails explicitly; it never substitutes a renderer or scales y.
pub fn vertical(font: &FontAsset, glyph_id: u16, target: f64) -> EngineResult<StretchedGlyph> {
    if !target.is_finite() || target <= 0.0 { return Err(invalid("Invalid vertical stretch target")); }
    if target > 2048.0 { return Err(limit("Vertical stretch target exceeds its budget")); }
    let face = font.face()?;
    let em = f64::from(face.units_per_em());
    let variants = face.tables().math.and_then(|table| table.variants);
    let construction = variants.and_then(|v| v.vertical_constructions.get(GlyphId(glyph_id)));
    let original = font.glyph(glyph_id)?;
    let height = |shape: &GlyphShape| shape.ink.map_or(0.0, |b| b[3] - b[1]);
    let mut candidates = vec![original];
    if let Some(construction) = &construction {
        if usize::from(construction.variants.len()) > MAX_VARIANTS { return Err(limit("Too many MATH glyph variants")); }
        for variant in construction.variants {
            // Validate the declared size as well as the actual glyph geometry.
            if variant.advance_measurement == 0 { return Err(invalid("MATH variant has zero extent")); }
            candidates.push(font.glyph(variant.variant_glyph.0)?);
        }
    }
    if let Some(shape) = candidates.into_iter().filter(|shape| height(shape) + EPS >= target)
        .min_by(|a, b| height(a).total_cmp(&height(b)).then(a.glyph_id.cmp(&b.glyph_id)))
    {
        let ink = shape.ink.ok_or_else(|| invalid("Stretch variant has no ink"))?;
        return Ok(StretchedGlyph { width: shape.advance, height: ink[3] - ink[1],
            italic_correction: shape.italic_correction, pieces: vec![StretchPiece { y: -ink[1], shape }],
            assembled: false, overlaps: Vec::new(), minimum_overlap: 0.0 });
    }
    let assembly = construction.and_then(|c| c.assembly)
        .ok_or_else(|| invalid("Font has no sufficiently large glyph or vertical assembly"))?;
    let min_overlap = f64::from(variants.unwrap().min_connector_overlap) / em;
    let parts: Vec<GlyphPart> = assembly.parts.into_iter().collect();
    if parts.is_empty() || parts.len() > MAX_PARTS { return Err(limit("MATH assembly part budget exceeded")); }
    let mut fixed_count = 0usize;
    let mut ext_count = 0usize;
    let mut fixed_size = 0.0;
    let mut ext_size = 0.0;
    for part in &parts {
        if part.full_advance == 0 || part.start_connector_length > part.full_advance
            || part.end_connector_length > part.full_advance {
            return Err(invalid("Invalid MATH assembly connectors or part advance"));
        }
        if part.part_flags.extender() { ext_count += 1; ext_size += f64::from(part.full_advance) / em; }
        else { fixed_count += 1; fixed_size += f64::from(part.full_advance) / em; }
    }
    let growth = ext_size - min_overlap * ext_count as f64;
    if ext_count == 0 || growth <= 0.0 { return Err(invalid("MATH assembly cannot grow")); }
    let repetitions = ((target - fixed_size + min_overlap * (fixed_count as f64 - 1.0)) / growth).ceil().max(0.0);
    if repetitions > MAX_PARTS as f64 { return Err(limit("MATH assembly repetition budget exceeded")); }
    let repetitions = repetitions as usize;
    let count = fixed_count.checked_add(ext_count.checked_mul(repetitions).ok_or_else(|| limit("Assembly size overflow"))?)
        .ok_or_else(|| limit("Assembly size overflow"))?;
    if count == 0 || count > MAX_PARTS { return Err(limit("MATH assembly glyph budget exceeded")); }
    let expanded: Vec<&GlyphPart> = parts.iter().flat_map(|part| {
        std::iter::repeat_n(part, if part.part_flags.extender() { repetitions } else { 1 })
    }).collect();
    let max_overlap = expanded.windows(2).map(|pair| {
        f64::from(pair[0].end_connector_length.min(pair[1].start_connector_length)) / em
    }).fold(f64::INFINITY, f64::min);
    if count > 1 && max_overlap + EPS < min_overlap { return Err(invalid("MATH assembly cannot satisfy minimum connector overlap")); }
    let full_size: f64 = expanded.iter().map(|part| f64::from(part.full_advance) / em).sum();
    let overlap = if count > 1 { ((full_size - target) / (count - 1) as f64).clamp(min_overlap, max_overlap) } else { 0.0 };
    let mut pieces = Vec::with_capacity(count);
    let mut baseline = 0.0;
    let mut top = f64::INFINITY;
    let mut bottom = f64::NEG_INFINITY;
    let mut width = 0.0f64;
    // MATH lists vertical parts from bottom to top. Their origins, rather than
    // individually normalized bounding boxes, define connector alignment.
    for (i, part) in expanded.iter().enumerate() {
        let shape = font.glyph(part.glyph_id.0)?;
        let ink = shape.ink.ok_or_else(|| invalid("MATH assembly part has no outline"))?;
        top = top.min(baseline + ink[1]); bottom = bottom.max(baseline + ink[3]);
        width = width.max(shape.advance);
        pieces.push(StretchPiece { shape, y: baseline });
        if i + 1 < count { baseline -= f64::from(part.full_advance) / em - overlap; }
    }
    let height = bottom - top;
    if !height.is_finite() || height + EPS < target || width <= 0.0 {
        return Err(invalid("MATH assembly ink does not reach its requested extent"));
    }
    for piece in &mut pieces { piece.y -= top; }
    Ok(StretchedGlyph { pieces, width, height,
        italic_correction: f64::from(assembly.italics_correction.value) / em,
        assembled: true, overlaps: vec![overlap; count - 1], minimum_overlap: min_overlap })
}
