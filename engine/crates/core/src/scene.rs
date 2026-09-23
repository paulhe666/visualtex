//! Immutable layout output: drawing, hit testing and export share one revision.
//! Geometry is in em, y down; exported natural sizes are 96-DPI CSS pixels.
use std::collections::BTreeMap;
use std::sync::Arc;
use serde::Serialize;
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::fonts::{FontIdentity, GlyphShape, PathOp};
use crate::positions::Position;
use crate::session::{NodeId, Revision};

#[derive(Debug, Clone, Copy, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Rect { pub x_min: f64, pub y_min: f64, pub x_max: f64, pub y_max: f64 }
impl Rect {
    pub fn new(x_min: f64, y_min: f64, x_max: f64, y_max: f64) -> Self { Self { x_min, y_min, x_max, y_max } }
    pub fn union(self, other: Self) -> Self { Self::new(self.x_min.min(other.x_min), self.y_min.min(other.y_min), self.x_max.max(other.x_max), self.y_max.max(other.y_max)) }
    pub fn translated(self, x: f64, y: f64) -> Self { Self::new(self.x_min+x, self.y_min+y, self.x_max+x, self.y_max+y) }
    pub fn width(self) -> f64 { self.x_max-self.x_min }
    pub fn height(self) -> f64 { self.y_max-self.y_min }
    pub fn contains(self, x: f64, y: f64) -> bool { x >= self.x_min && x <= self.x_max && y >= self.y_min && y <= self.y_max }
    pub fn valid(self) -> bool { [self.x_min,self.y_min,self.x_max,self.y_max].into_iter().all(f64::is_finite) && self.width() >= 0.0 && self.height() >= 0.0 }
}
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct GlyphResource { pub font_id: String, pub shape: Arc<GlyphShape> }
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct GlyphInstance { pub node_id: NodeId, pub resource: usize, pub x: f64, pub y: f64, pub scale: f64,
    #[serde(skip_serializing_if="Option::is_none")] pub color:Option<crate::paint::RgbColor> }
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Rule { pub node_id: NodeId, pub rect: Rect,
    #[serde(skip_serializing_if="Option::is_none")] pub color:Option<crate::paint::RgbColor> }
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CaretStop { pub position: Position, pub x: f64, pub baseline: f64, pub top: f64, pub bottom: f64, pub depth: usize }
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct NodeRegion { pub node_id: NodeId, pub role: String, pub logical: Rect, pub ink: Option<Rect>, pub depth: usize }
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct GridColumnAnchor {
    pub grid_id: NodeId, pub column_id: crate::grids::GridColumnId,
    pub index: usize, pub x: f64, pub width: f64,
}
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct LayoutSnapshot {
    pub document_id: String, pub revision: Revision, pub source_id: u32,
    pub font_size_px: f64,
    pub width: f64, pub height: f64, pub baseline: f64,
    pub width_em: f64, pub height_em: f64, pub baseline_em: f64,
    pub advance_em: f64, pub ink: Option<Rect>,
    pub fonts: Vec<FontIdentity>, pub resources: Vec<GlyphResource>,
    pub glyphs: Vec<GlyphInstance>, pub rules: Vec<Rule>,
    #[serde(skip_serializing_if="Vec::is_empty")]pub artworks:Vec<crate::artwork_scene::ArtworkInstance>,
    #[serde(skip_serializing_if="Vec::is_empty")]pub paint_order:Vec<crate::artwork_scene::PaintItem>,
    pub carets: Vec<CaretStop>, pub regions: Vec<NodeRegion>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub grid_columns: Vec<GridColumnAnchor>,
    pub export_ready: bool,
    pub rebuilt_fragments: usize, pub reused_fragments: usize,
}
impl LayoutSnapshot {
    pub fn validate(&self) -> EngineResult<()> {
        crate::scene_validation::validate_snapshot(self)?;
        crate::artwork_scene::validate(self)?;
        if ![self.font_size_px,self.width,self.height,self.baseline,self.width_em,self.height_em,self.baseline_em,self.advance_em].into_iter().all(f64::is_finite)
            || self.width <= 0.0 || self.height <= 0.0 || self.baseline < 0.0 || self.baseline > self.height
            || self.ink.is_some_and(|r| !r.valid())
            || self.rules.iter().any(|r| !r.rect.valid())
            || self.glyphs.iter().any(|g| g.resource >= self.resources.len() || ![g.x,g.y,g.scale].into_iter().all(f64::is_finite) || g.scale <= 0.0)
            || self.carets.iter().any(|c| ![c.x,c.baseline,c.top,c.bottom].into_iter().all(f64::is_finite) || c.top>c.bottom)
            || self.regions.iter().any(|r| !r.logical.valid() || r.ink.is_some_and(|i| !i.valid()))
        { return Err(EngineError::new(ErrorCode::InvalidInput, "Layout contains invalid geometry")); }
        Ok(())
    }
    pub fn hit_test(&self, x_em: f64, y_em: f64) -> EngineResult<Position> {
        if !x_em.is_finite() || !y_em.is_finite() { return Err(EngineError::new(ErrorCode::InvalidInput, "Hit-test coordinates must be finite")); }
        self.carets.iter().min_by(|a,b| {
            let score = |c: &CaretStop| {
                let dy = if y_em < c.top { c.top-y_em } else if y_em>c.bottom { y_em-c.bottom } else {0.0};
                (c.x-x_em).powi(2)+4.0*dy.powi(2)
            };
            score(a).total_cmp(&score(b)).then(b.depth.cmp(&a.depth))
        }).map(|c| c.position).ok_or_else(|| EngineError::new(ErrorCode::InvalidPosition, "Layout has no editable caret"))
    }
    pub fn caret(&self, position: Position) -> Option<&CaretStop> {
        // Affinity changes insertion ownership, not the geometric caret edge.
        self.carets.iter().find(|c| match (c.position,position) {
            (Position::Row {node_id:a,offset:x,..},Position::Row {node_id:b,offset:y,..}) => a==b && x==y,
            (Position::Text {node_id:a,grapheme:x,..},Position::Text {node_id:b,grapheme:y,..}) => a==b && x==y,
            _=>false,
        })
    }
    pub fn vertical(&self, from: Position, down: bool, desired_x: Option<f64>) -> EngineResult<Position> {
        let start = self.caret(from).ok_or_else(|| EngineError::new(ErrorCode::InvalidPosition, "Caret is absent from this layout"))?;
        let x = desired_x.unwrap_or(start.x);
        if !x.is_finite() { return Err(EngineError::new(ErrorCode::InvalidInput, "Desired caret coordinate must be finite")); }
        if let Some(target) = crate::scene_navigation::fraction_target(self, from, down, x) { return Ok(target); }
        let candidates: Vec<_> = self.carets.iter().filter(|c| if down {c.baseline>start.baseline+1e-8} else {c.baseline<start.baseline-1e-8}).collect();
        Ok(candidates.into_iter().min_by(|a,b| {
            let score=|c: &CaretStop| (c.x-x).abs() + (c.baseline-start.baseline).abs()*0.5;
            score(a).total_cmp(&score(b)).then(b.depth.cmp(&a.depth))
        }).map_or(from, |c| c.position))
    }
    /// Paths and rects only. No source text, scripts, URLs, fonts, foreignObject,
    /// raster images or editor overlays are emitted into this vector export.
    pub fn svg(&self) -> EngineResult<String> {
        self.validate()?;
        if !self.export_ready { return Err(EngineError::new(ErrorCode::IncompleteSource, "Fill all formula slots before export")); }
        if !self.artworks.is_empty(){return crate::artwork_scene::svg(self);}
        let mut out = format!("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{}\" height=\"{}\" viewBox=\"0 0 {} {}\">", number(self.width),number(self.height),number(self.width_em),number(self.height_em));
        for g in &self.glyphs {
            let resource = &self.resources[g.resource];
            if resource.shape.path.is_empty() { continue; }
            out.push_str("<path d=\"");
            for op in &resource.shape.path {
                let mut point = |prefix: &str, p: &[f64]| {
                    out.push_str(prefix);
                    for (i,v) in p.iter().enumerate() {
                        if i>0 {out.push(' ');}
                        out.push_str(&number(v*g.scale + if i%2==0 {g.x} else {g.y}));
                    }
                };
                match op {PathOp::Move(p)=>point("M",p),PathOp::Line(p)=>point("L",p),PathOp::Quad(p)=>point("Q",p),PathOp::Cubic(p)=>point("C",p),PathOp::Close=>out.push('Z')}
            }
            out.push('"');
            if let Some(color)=g.color {out.push_str(&format!(" fill=\"{}\"",color.hex()));}
            out.push_str("/>");
            if out.len()>16*1024*1024 { return Err(EngineError::new(ErrorCode::ResourceLimit, "SVG output budget exceeded")); }
        }
        for r in &self.rules {
            out.push_str(&format!("<rect x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\"",number(r.rect.x_min),number(r.rect.y_min),number(r.rect.width()),number(r.rect.height())));
            if let Some(color)=r.color {out.push_str(&format!(" fill=\"{}\"",color.hex()));}
            out.push_str("/>");
        }
        out.push_str("</svg>"); Ok(out)
    }
}
fn number(value: f64) -> String { let mut v=format!("{value:.6}"); while v.contains('.') && v.ends_with('0') {v.pop();} if v.ends_with('.') {v.pop();} if v=="-0" {v="0".into();} v }

pub(crate) struct SceneBuilder { pub snapshot: LayoutSnapshot, pub resource_index: BTreeMap<(String,u16),usize>,pub artwork_ops:usize }
impl SceneBuilder {
    pub fn glyph(&mut self, node_id: NodeId, font_id: &str, shape: Arc<GlyphShape>, x: f64, y:f64, scale:f64, color:Option<crate::paint::RgbColor>) {
        let key=(font_id.to_owned(),shape.glyph_id);
        let resource=if let Some(index)=self.resource_index.get(&key) {*index} else {
            let index=self.snapshot.resources.len(); self.snapshot.resources.push(GlyphResource {font_id:font_id.to_owned(),shape});
            self.resource_index.insert(key,index); index
        };
        self.snapshot.paint_order.push(crate::artwork_scene::PaintItem::Glyph{index:self.snapshot.glyphs.len()});
        self.snapshot.glyphs.push(GlyphInstance {node_id,resource,x,y,scale,color});
    }
}
