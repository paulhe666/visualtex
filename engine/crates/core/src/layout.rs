//! K0 mathematical layout and cache ownership; geometry is always in em.
mod boxes;
mod structures;
// Additional two-argument mathematical layouts: binomial and bevelled fraction.
mod fraction_forms;
mod grids;
mod macros;
mod custom_symbols;
mod radicals;
mod fences;
mod accents;
mod operators;
mod tex_fonts;
use std::collections::BTreeMap;
use std::sync::Arc;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::custom_symbols::ArtworkFontStyle;
use crate::fonts::{FontAsset,FontRegistry, MathMetrics};
use crate::model::{Node, NodeKind, MAX_DEPTH};
use crate::positions::TreeIndex;
use crate::scene::{LayoutSnapshot, Rect, SceneBuilder};
use crate::session::{EngineSession, NodeId};
use boxes::{Builder, Fragment, flatten};

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(rename_all="camelCase")]
pub enum MathStyle { Display, Text, Script, ScriptScript }
impl MathStyle {
    fn fraction_child(self)->Self {match self {Self::Display=>Self::Text,Self::Text=>Self::Script,_=>Self::ScriptScript}}
    fn script_child(self)->Self {match self {Self::Display|Self::Text=>Self::Script,_=>Self::ScriptScript}}
    fn scale(self,m:&MathMetrics)->f64 {match self {Self::Display|Self::Text=>1.0,Self::Script=>m.script_scale,Self::ScriptScript=>m.scriptscript_scale}}
}
#[derive(Debug,Clone,PartialEq,Eq,Serialize,Deserialize)]
#[serde(rename_all="camelCase",deny_unknown_fields)]
pub struct CustomSymbolTextFontBinding {
    pub font_family:String,pub font_style:ArtworkFontStyle,pub font_weight:u16,pub font_id:String,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all="camelCase",deny_unknown_fields)]
pub struct LayoutOptions {
    pub math_font:String, pub text_font:Option<String>, pub style:MathStyle,
    pub font_size_px:f64, pub padding_em:f64,
    /// Product-facing alphabet faces are verified immutable resources but do
    /// not need a MATH table. Structural geometry remains owned by math_font.
    #[serde(default)]
    pub alphabet_upright_font:Option<String>,
    #[serde(default)]
    pub alphabet_italic_font:Option<String>,
    #[serde(default)]
    pub alphabet_bold_font:Option<String>,
    #[serde(default)]
    pub alphabet_bold_italic_font:Option<String>,
    /// Optional verified faces for a multi-font TeX family (ordinary cmap positions).
    #[serde(default)]
    pub tex_font_set:BTreeMap<String,String>,
    /// Ordered explicit resource IDs; empty keeps the original strict policy.
    #[serde(default)]
    pub math_fallbacks:Vec<String>,
    /// Exact legacy text-style tuple -> already registered immutable font bytes.
    /// No CSS/system-family lookup is permitted in the core or browser view.
    #[serde(default)]
    pub custom_symbol_text_fonts:Vec<CustomSymbolTextFontBinding>,
}
impl LayoutOptions {
    pub fn new(math_font:String)->Self {Self {math_font,text_font:None,style:MathStyle::Display,font_size_px:16.0,padding_em:0.125,alphabet_upright_font:None,alphabet_italic_font:None,alphabet_bold_font:None,alphabet_bold_italic_font:None,tex_font_set:BTreeMap::new(),math_fallbacks:Vec::new(),custom_symbol_text_fonts:Vec::new()}}
    fn validate(&self)->EngineResult<()> {
        if !self.font_size_px.is_finite() || !(1.0..=2048.0).contains(&self.font_size_px)
            || !self.padding_em.is_finite() || !(0.0..=8.0).contains(&self.padding_em)
        {return Err(EngineError::new(ErrorCode::InvalidInput,"Invalid natural font size or layout padding"));}
        if self.math_fallbacks.len()>4 || self.math_fallbacks.iter().collect::<std::collections::BTreeSet<_>>().len()!=self.math_fallbacks.len()
            || self.math_fallbacks.iter().any(|id|id==&self.math_font) {
            return Err(EngineError::new(ErrorCode::InvalidInput,"Math fallback resource list must be bounded and unique"));
        }
        if self.tex_font_set.len()>20 || self.tex_font_set.iter().any(|(role,id)| !tex_fonts::ROLES.contains(&role.as_str()) || id.is_empty() || id.len()>256) {
            return Err(EngineError::new(ErrorCode::InvalidInput,"TeX font set must contain only bounded verified font roles"));
        }
        if self.custom_symbol_text_fonts.len()>32{return Err(EngineError::new(ErrorCode::ResourceLimit,"Custom-symbol text font binding budget exceeded"));}
        let mut keys=std::collections::BTreeSet::new();for binding in &self.custom_symbol_text_fonts{
            if binding.font_family.is_empty()||binding.font_family.len()>128||binding.font_family.chars().any(|c|c.is_control()||matches!(c,'\"'|'\''|'<'|'>'))
                ||!(100..=900).contains(&binding.font_weight)||binding.font_id.is_empty()||binding.font_id.len()>256
                ||!keys.insert((binding.font_family.as_str(),binding.font_style,binding.font_weight)){
                return Err(EngineError::new(ErrorCode::InvalidInput,"Custom-symbol text font bindings must be bounded, valid and unique"));
            }
        }Ok(())
    }
}
struct Cached {fingerprint:[u8;32],fragment:Arc<Fragment>}
type Key=(NodeId,MathStyle,bool,crate::styles::MathVariant,Option<u64>,bool);
pub(super) struct ResolvedCustomTextFont<'a>{pub family:&'a str,pub style:ArtworkFontStyle,pub weight:u16,pub font:&'a FontAsset}
#[derive(PartialEq,Eq)]
struct LayoutContext {
    document_id:String, math_font:String, text_font:String,
    alphabet_fonts:[Option<String>;4], tex_fonts:BTreeMap<String,String>,
    fallbacks:Vec<String>, custom_text_fonts:Vec<CustomSymbolTextFontBinding>,
}
#[derive(Default)]
pub struct LayoutEngine {cache:BTreeMap<Key,Cached>,context:Option<LayoutContext>}
impl LayoutEngine {
    pub fn layout(&mut self,session:&EngineSession,fonts:&FontRegistry,options:&LayoutOptions)->EngineResult<LayoutSnapshot> {
        options.validate()?;
        let document=&session.document;
        if !document.complete {return Err(EngineError::new(ErrorCode::IncompleteSource,"Layout cannot silently repair a source draft"));}
        if !document.supported {return Err(EngineError::new(ErrorCode::UnsupportedSyntax,"K0 layout does not support this source syntax"));}
        let index=TreeIndex::new(&document.tree)?;
        let math_font=fonts.get(&options.math_font)?;
        let text_font=fonts.get(options.text_font.as_deref().unwrap_or(&options.math_font))?;
        let alphabet_upright_font=options.alphabet_upright_font.as_deref().map(|id|fonts.get(id)).transpose()?;
        let alphabet_italic_font=options.alphabet_italic_font.as_deref().map(|id|fonts.get(id)).transpose()?;
        let alphabet_bold_font=options.alphabet_bold_font.as_deref().map(|id|fonts.get(id)).transpose()?;
        let alphabet_bold_italic_font=options.alphabet_bold_italic_font.as_deref().map(|id|fonts.get(id)).transpose()?;
        let math=math_font.math.as_ref().ok_or_else(|| EngineError::new(ErrorCode::InvalidInput,"Mathematical font must provide an OpenType MATH table"))?;
        let fallback_fonts=options.math_fallbacks.iter().map(|id|fonts.get(id)).collect::<EngineResult<Vec<_>>>()?;
        if fallback_fonts.iter().any(|font|font.math.is_none()) {
            return Err(EngineError::new(ErrorCode::InvalidInput,"Mathematical fallback fonts must provide a MATH table"));
        }
        let custom_text_fonts=options.custom_symbol_text_fonts.iter().map(|binding|Ok(ResolvedCustomTextFont{family:binding.font_family.as_str(),style:binding.font_style,weight:binding.font_weight,font:fonts.get(&binding.font_id)?})).collect::<EngineResult<Vec<_>>>()?;
        let tex_fonts=options.tex_font_set.iter().map(|(role,id)|Ok((role.clone(),fonts.get(id)?))).collect::<EngineResult<tex_fonts::ResolvedTexFonts<'_>>>()?;
        let context=LayoutContext {
            document_id:session.document_id().to_owned(),math_font:math_font.identity.id.clone(),text_font:text_font.identity.id.clone(),
            alphabet_fonts:[options.alphabet_upright_font.clone(),options.alphabet_italic_font.clone(),options.alphabet_bold_font.clone(),options.alphabet_bold_italic_font.clone()],
            tex_fonts:options.tex_font_set.clone(),fallbacks:options.math_fallbacks.clone(),custom_text_fonts:options.custom_symbol_text_fonts.clone(),
        };
        if self.context.as_ref()!=Some(&context) {self.cache.clear();self.context=Some(context);}
        self.cache.retain(|(id,_,_,_,_,_),_|index.nodes.contains_key(id));
        let mut hashes=BTreeMap::new();fingerprint(&document.tree,&mut hashes,0)?;
        let mut builder=Builder {tex_fonts:&tex_fonts,script_alphabet:false,math_font,text_font,alphabet_upright_font,alphabet_italic_font,alphabet_bold_font,alphabet_bold_italic_font,math,fallback_fonts:&fallback_fonts,custom_text_fonts:&custom_text_fonts,hashes:&hashes,cache:&mut self.cache,built:0,reused:0,
            variant:crate::styles::MathVariant::Default,middle_target:None};
        let fragment=builder.node(&document.tree,options.style,false,0)?;
        let rebuilt=builder.built;let reused=builder.reused;
        let logical=Rect::new(fragment.advance.min(0.0),-fragment.ascent,fragment.advance.max(0.0),fragment.descent);
        let extent=fragment.ink.map_or(logical,|i|logical.union(i));
        let width_em=(extent.width()+2.0*options.padding_em).max(1e-6);
        let height_em=(extent.height()+2.0*options.padding_em).max(1e-6);
        let x=options.padding_em-extent.x_min;let baseline_em=options.padding_em-extent.y_min;
        let font_size=options.font_size_px;
        let mut font_ids=vec![math_font.identity.clone()];
        if text_font.identity.id!=math_font.identity.id {font_ids.push(text_font.identity.clone());}
        for font in [alphabet_upright_font,alphabet_italic_font,alphabet_bold_font,alphabet_bold_italic_font].into_iter().flatten() {if !font_ids.iter().any(|id|id.id==font.identity.id) {font_ids.push(font.identity.clone());}}
        for font in tex_fonts.values() {if !font_ids.iter().any(|id|id.id==font.identity.id) {font_ids.push(font.identity.clone());}}
        for font in &fallback_fonts {if !font_ids.iter().any(|id|id.id==font.identity.id) {font_ids.push(font.identity.clone());}}
        for binding in &custom_text_fonts {if !font_ids.iter().any(|id|id.id==binding.font.identity.id){font_ids.push(binding.font.identity.clone());}}
        let snapshot=LayoutSnapshot {document_id:session.document_id().to_owned(),revision:session.revision,source_id:document.source.id(),
            font_size_px:font_size,width:width_em*font_size,height:height_em*font_size,baseline:baseline_em*font_size,
            width_em,height_em,baseline_em,advance_em:fragment.advance,ink:fragment.ink.map(|r|r.translated(x,baseline_em)),
            fonts:font_ids,resources:Vec::new(),glyphs:Vec::new(),rules:Vec::new(),artworks:Vec::new(),paint_order:Vec::new(),carets:Vec::new(),regions:Vec::new(),grid_columns:Vec::new(),
            export_ready:document.exportable,rebuilt_fragments:rebuilt,reused_fragments:reused};
        let mut scene=SceneBuilder {snapshot,resource_index:BTreeMap::new(),artwork_ops:0};
        flatten(&fragment,x,baseline_em,0,&mut scene)?;
        if scene.snapshot.artworks.is_empty(){scene.snapshot.paint_order.clear();}
        scene.snapshot.validate()?;
        for caret in &scene.snapshot.carets {index.validate_position(caret.position)?;}
        Ok(scene.snapshot)
    }
}
fn fingerprint(node:&Node,out:&mut BTreeMap<NodeId,[u8;32]>,depth:usize)->EngineResult<[u8;32]> {
    if depth>MAX_DEPTH {return Err(EngineError::new(ErrorCode::ResourceLimit,"Render fingerprint depth exceeded"));}
    let mut hash=Sha256::new();hash.update(node.id.0.to_le_bytes());
    // This hash identifies cached geometry, never an editing node. Source
    // trivia is excluded; IDs remain included to preserve hit-test targets.
    match &node.kind {
        NodeKind::Row {..}=>hash.update([0]),
        NodeKind::Grid {environment,columns,array_spec,..}=>{
            hash.update([19,*environment as u8]);for column in columns {hash.update(column.0.to_le_bytes());}
            if let Some(spec)=array_spec {hash.update(spec.source.as_bytes());hash.update([0]);for alignment in &spec.alignments{hash.update([*alignment as u8]);}for rule in &spec.vertical_rules{hash.update(rule.to_le_bytes());}}
        },
        NodeKind::GridRow {cells,..}=>{hash.update([20]);hash.update(cells.len().to_le_bytes());},
        NodeKind::Symbol {value,class,command}=>{hash.update([1,*class as u8]);hash.update(value.as_bytes());hash.update([0]);if let Some(command)=command {hash.update(command.as_bytes());}},
        NodeKind::StyleChange {command,..}=>{hash.update([26]);hash.update(command.as_bytes());},
        NodeKind::Text {command,text,..}=>{hash.update([2]);hash.update(command.as_bytes());hash.update([0]);hash.update(text.as_bytes());},
        NodeKind::Fraction {command,..}=>{hash.update([3]);hash.update(command.as_bytes());},
        NodeKind::Scripts {slots,..}=>{hash.update([4]);for s in slots {hash.update([s.script as u8]);}},
        NodeKind::Radical {degree,..}=>hash.update([7,u8::from(degree.is_some())]),
        NodeKind::Fence {left,right,..}=>{hash.update([8]);hash.update(left.as_bytes());hash.update([0]);hash.update(right.as_bytes());},
        NodeKind::MiddleDelimiter {delimiter,..}=>{hash.update([21]);hash.update(delimiter.as_bytes());},
        NodeKind::Accent {command,..}=>{hash.update([9]);hash.update(command.as_bytes());},
        NodeKind::LargeOperator {command,parts}=>{
            hash.update([11,crate::operators::effective_mode(parts) as u8]);hash.update(command.as_bytes());
            for slot in crate::operators::slots(parts){hash.update([slot.script as u8]);}
        },
        NodeKind::CustomSymbol {asset_sha256,parts,..}=>{
            hash.update([24,crate::operators::effective_mode(parts) as u8]);hash.update(asset_sha256.as_bytes());
            for slot in crate::operators::slots(parts){hash.update([slot.script as u8]);}
        },
        NodeKind::GeneralOperator {command,parts,..}=>{
            hash.update([18,crate::operators::effective_mode(parts) as u8]);hash.update(command.as_bytes());
            for slot in crate::operators::slots(parts){hash.update([slot.script as u8]);}
        },
        NodeKind::MacroCall {command,macro_kind,arguments,..}=>{
            hash.update([21,*macro_kind as u8]);hash.update(command.as_bytes());hash.update(arguments.len().to_le_bytes());
        },
        NodeKind::Style {command,variant,..}=>{hash.update([12,*variant as u8]);hash.update(command.as_bytes());},
        NodeKind::Space {command,width_mu}=>{hash.update([13]);hash.update(command.as_bytes());hash.update(width_mu.to_le_bytes());},
        NodeKind::Phantom {command,axis,..}=>{hash.update([14,*axis as u8]);hash.update(command.as_bytes());},
        NodeKind::Boxed {..}=>hash.update([15]),
        NodeKind::Color {spec,..}=>hash.update([16,spec.color.r,spec.color.g,spec.color.b]),
        NodeKind::ColorChange {spec}=>hash.update([17,spec.color.r,spec.color.g,spec.color.b]),
        NodeKind::Annotation {command,above,..}=>{hash.update([10,u8::from(*above)]);hash.update(command.as_bytes());},
        NodeKind::EmptySlot {..}=>hash.update([5]),
        NodeKind::UnparsedSource {raw}=>{hash.update([6]);hash.update(raw.as_bytes());},
    }
    for child in node.children() {hash.update(fingerprint(child,out,depth+1)?);}
    let result=hash.finalize().into();out.insert(node.id,result);Ok(result)
}
