//! Validation of the self-contained scene consumed by vector export and hosts.
//! It checks values and provenance, not current editor state or Office protocols.
use std::collections::BTreeSet;
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::fonts::PathOp;
use crate::scene::{LayoutSnapshot, Rect};

const EPSILON: f64 = 1e-9;
const MAX_SCENE_EXTENT_EM: f64 = 1_000_000.0;
const MAX_PATH_COORDINATE_EM: f64 = 1_000_000.0;
const MAX_PATH_OPS: usize = 16_384;
const MAX_GLYPHS: usize = 16_384;
const MAX_CARETS: usize = 32_768;

fn invalid(message: &str) -> EngineError { EngineError::new(ErrorCode::InvalidInput, message) }
fn close(a: f64, b: f64) -> bool { a.is_finite() && b.is_finite() && (a-b).abs() <= EPSILON*a.abs().max(b.abs()).max(1.0) }
fn inside(rect: Rect, s: &LayoutSnapshot) -> bool {
    let e = EPSILON * s.width_em.max(s.height_em).max(1.0);
    rect.valid() && rect.x_min >= -e && rect.y_min >= -e
        && rect.x_max <= s.width_em+e && rect.y_max <= s.height_em+e
}
fn points(op: &PathOp) -> &[f64] {
    match op { PathOp::Move(p)|PathOp::Line(p)=>p, PathOp::Quad(p)=>p,
        PathOp::Cubic(p)=>p, PathOp::Close=>&[] }
}

pub(crate) fn validate_snapshot(s: &LayoutSnapshot) -> EngineResult<()> {
    crate::grid_scene_validation::validate(s)?;
    if ![s.font_size_px,s.width,s.height,s.baseline,s.width_em,s.height_em,s.baseline_em,s.advance_em]
        .into_iter().all(f64::is_finite)
        || s.font_size_px <= 0.0 || s.width_em <= 0.0 || s.height_em <= 0.0 || s.advance_em.abs() > MAX_SCENE_EXTENT_EM
        || s.baseline_em < 0.0 || s.baseline_em > s.height_em
        || !close(s.width,s.width_em*s.font_size_px)
        || !close(s.height,s.height_em*s.font_size_px)
        || !close(s.baseline,s.baseline_em*s.font_size_px)
    { return Err(invalid("Natural CSS-pixel dimensions and em geometry disagree")); }
    // Signed advances are legal for authored negative glue. They are not an
    // independent caller-supplied dimension: validate against the actual root
    // row's logical extent and directed first/last caret displacement.
    let root=s.regions.first().ok_or_else(||invalid("Scene lacks its root mathematical row"))?;
    if root.role!="row" || !close(root.logical.width(),s.advance_em.abs()) {
        return Err(invalid("Root logical width and signed advance disagree"));
    }
    let mut root_carets=s.carets.iter().filter_map(|c|match c.position {
        crate::positions::Position::Row{node_id,offset,..} if node_id==root.node_id=>Some((offset,c)),_=>None,
    }).collect::<Vec<_>>();
    root_carets.sort_by_key(|(offset,_)|*offset);
    let (Some((0,start)),Some((end_offset,end)))=(root_carets.first().copied(),root_carets.last().copied()) else {
        return Err(invalid("Scene lacks root row caret boundaries"));
    };
    if end_offset>0 && !close(end.x-start.x,s.advance_em) {
        return Err(invalid("Root caret displacement and signed advance disagree"));
    }
    if s.width_em > MAX_SCENE_EXTENT_EM || s.height_em > MAX_SCENE_EXTENT_EM
        || s.glyphs.len() > MAX_GLYPHS || s.resources.len() > MAX_GLYPHS || s.carets.len() > MAX_CARETS
    { return Err(EngineError::new(ErrorCode::ResourceLimit,"Scene geometry or item budget exceeded")); }
    if s.ink.is_some_and(|r| !inside(r,s)) || s.rules.iter().any(|r| !inside(r.rect,s)) {
        return Err(invalid("Scene ink or a rule falls outside its natural export frame"));
    }
    let mut columns=BTreeSet::new();let mut grid_indices=std::collections::BTreeMap::new();
    for column in &s.grid_columns {
        let owner=s.regions.iter().find(|r|r.node_id==column.grid_id&&r.role=="grid")
            .ok_or_else(||invalid("Grid column owner is absent"))?;
        let expected=grid_indices.entry(column.grid_id).or_insert(0usize);
        let epsilon=EPSILON*owner.logical.width().abs().max(1.0);
        if column.column_id.0==0 || !columns.insert(column.column_id)
            || column.index!=*expected || column.index>=crate::grids::MAX_GRID_COLUMNS
            || !column.x.is_finite() || !column.width.is_finite() || column.width<0.0
            || column.x.abs()>MAX_SCENE_EXTENT_EM || column.width>MAX_SCENE_EXTENT_EM
            || column.x<owner.logical.x_min-epsilon || column.x+column.width>owner.logical.x_max+epsilon
            || s.regions.iter().any(|r|r.node_id.0==column.column_id.0) {
            return Err(invalid("Grid column geometry or identity is invalid"));
        }
        *expected+=1;
    }
    let mut identities = BTreeSet::new();
    for font in &s.fonts {
        if font.sha256.len()!=64 || !font.sha256.bytes().all(|c|c.is_ascii_hexdigit())
            || font.id!=format!("{}:{}",font.sha256,font.face_index)
            || font.units_per_em==0 || !identities.insert(font.id.as_str())
            || font.variation_axes.iter().any(|(_,v)| !v.is_finite())
        { return Err(invalid("Scene contains an invalid or duplicate font identity")); }
    }
    for resource in &s.resources {
        if !identities.contains(resource.font_id.as_str()) {
            return Err(invalid("Glyph resource is not bound to a font in this scene"));
        }
        let shape = &resource.shape;
        if shape.glyph_id==0 || !shape.advance.is_finite() || shape.advance<0.0
            || !shape.italic_correction.is_finite()
            || shape.ink.is_some_and(|[a,b,c,d]| !Rect::new(a,b,c,d).valid())
        { return Err(invalid("Scene glyph resource contains invalid metrics")); }
        if shape.path.len()>MAX_PATH_OPS { return Err(EngineError::new(ErrorCode::ResourceLimit,"Scene glyph path budget exceeded")); }
        for coordinate in shape.path.iter().flat_map(points) {
            if !coordinate.is_finite() { return Err(invalid("Scene glyph path contains non-finite coordinates")); }
            if coordinate.abs()>MAX_PATH_COORDINATE_EM { return Err(EngineError::new(ErrorCode::ResourceLimit,"Scene glyph coordinate budget exceeded")); }
        }
    }
    for glyph in &s.glyphs {
        let resource = s.resources.get(glyph.resource).ok_or_else(|| invalid("Unresolved glyph resource index"))?;
        if ![glyph.x,glyph.y,glyph.scale].into_iter().all(f64::is_finite) || glyph.scale<=0.0 {
            return Err(invalid("Invalid glyph placement or scale"));
        }
        if let Some([a,b,c,d])=resource.shape.ink {
            if !inside(Rect::new(a*glyph.scale+glyph.x,b*glyph.scale+glyph.y,
                c*glyph.scale+glyph.x,d*glyph.scale+glyph.y),s) {
                return Err(invalid("Glyph ink falls outside its natural export frame"));
            }
        }
        // Even a corrupt path with an innocent bounding box cannot overflow
        // the transformed SVG coordinates. Test arithmetic, never clamp it.
        for op in &resource.shape.path {
            for (i,v) in points(op).iter().enumerate() {
                if !(v*glyph.scale+if i%2==0 {glyph.x}else{glyph.y}).is_finite() {
                    return Err(invalid("Transformed glyph path has non-finite coordinates"));
                }
            }
        }
    }
    Ok(())
}
