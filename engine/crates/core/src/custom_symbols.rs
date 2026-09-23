//! Immutable custom-symbol resources. Legacy library v1 is input data, never
//! executable SVG/HTML or a live user-library lookup. Unsupported artwork stays
//! in the exact definition snapshot and receives an explicit capability issue.
mod geometry;
use std::{collections::BTreeMap,sync::Arc};
use serde::{Deserialize,Serialize};
use serde_json::Value;
use sha2::{Digest,Sha256};
use crate::{EngineError,EngineResult,ErrorCode,fonts::PathOp,model::AtomClass,scene::Rect};

pub const MAX_SYMBOL_BYTES:usize=1_048_576;
pub const MAX_SYMBOLS:usize=128;
pub const MAX_LIBRARY_BYTES:usize=4*1024*1024;
pub const MAX_ARTWORK_OPS:usize=65_536;

#[derive(Debug,Clone,Copy,PartialEq,Serialize,Deserialize)]
#[serde(rename_all="camelCase",deny_unknown_fields)]
pub struct SymbolMetrics {pub width_em:f64,pub ascent_em:f64,pub descent_em:f64}
#[derive(Debug,Clone,Copy,PartialEq,Eq,Serialize,Deserialize)]
#[serde(rename_all="lowercase")]
pub enum SymbolLimits {Auto,Limits,Nolimits}
#[derive(Debug,Clone,Serialize)]
#[serde(rename_all="camelCase")]
pub struct ArtworkIssue {pub shape_index:usize,pub code:String,pub message:String}
#[derive(Debug,Clone,Copy,PartialEq,Eq,Serialize)]
#[serde(rename_all="camelCase")]
pub enum ArtworkLineCap {Butt,Round,Square}
#[derive(Debug,Clone,Copy,PartialEq,Eq,Serialize)]
#[serde(rename_all="camelCase")]
pub enum ArtworkLineJoin {Miter,Round,Bevel}
#[derive(Debug,Clone,Copy,PartialEq,Eq,PartialOrd,Ord,Serialize,Deserialize)]
#[serde(rename_all="lowercase")]
pub enum ArtworkFontStyle {Normal,Italic}
#[derive(Debug,Clone,Serialize)]
#[serde(rename_all="camelCase")]
pub struct ArtworkTextSource {
    pub text:String,pub x:f64,pub y:f64,pub font_family:String,pub font_size:f64,
    pub font_style:ArtworkFontStyle,pub font_weight:u16,
}
#[derive(Debug,Clone,Copy,PartialEq,Eq,Serialize)]
#[serde(rename_all="lowercase")]
pub enum ArtworkOperation {Paint,Erase}
#[derive(Debug,Clone,Serialize)]
#[serde(rename_all="camelCase")]
pub struct ArtworkPath {
    /// Local path coordinates are baseline-relative em units. `transform` is
    /// applied before the enclosing formula style scale/translation, matching
    /// SVG transform semantics for fill and stroke without CSS participation.
    pub path:Vec<PathOp>,pub transform:[f64;6],pub inner_transform:[f64;6],pub ink:Rect,
    #[serde(skip_serializing_if="Option::is_none")]pub clip:Option<Rect>,
    #[serde(skip_serializing_if="Option::is_none")]pub text_source:Option<ArtworkTextSource>,
    #[serde(skip_serializing_if="Option::is_none")]pub outline_font_id:Option<String>,
    pub operation:ArtworkOperation,pub fill:bool,pub stroke_width:f64,pub line_cap:ArtworkLineCap,pub line_join:ArtworkLineJoin,
}
#[derive(Debug,Clone,Serialize)]
#[serde(rename_all="camelCase")]
pub struct SymbolAsset {
    sha256:String,command:String,role:AtomClass,limits_behavior:SymbolLimits,
    metrics:SymbolMetrics,definition:Value,geometry:Vec<Arc<ArtworkPath>>,diagnostics:Vec<ArtworkIssue>,
    #[serde(skip)]retained_weight:usize,
}
impl SymbolAsset {
    pub fn freeze(definition:Value)->EngineResult<Arc<Self>> {
        check_json_budget(&definition,MAX_SYMBOL_BYTES)?;
        let object=definition.as_object().ok_or_else(||bad("Custom symbol must be a definition object"))?;
        let command=string(&definition,"command")?.to_owned();
        if command.is_empty()||command.len()>48||!command.bytes().all(|b|b.is_ascii_alphabetic()) {
            return Err(bad("Custom symbol command must contain 1..48 ASCII letters"));
        }
        if reserved(&command){return Err(bad("Custom symbol cannot replace a reserved mathematical command"));}
        let id=string(&definition,"id")?;
        if id.is_empty()||id.len()>96||!id.as_bytes()[0].is_ascii_alphanumeric()
            ||!id.bytes().all(|b|b.is_ascii_alphanumeric()||b"._:-".contains(&b)){return Err(bad("Invalid custom symbol id"));}
        let name=string(&definition,"name")?;
        if name.is_empty()||name.chars().count()>64||name.chars().any(char::is_control){return Err(bad("Invalid custom symbol name"));}
        let role=match string(&definition,"role")? {
            "ordinary"=>AtomClass::Ordinary,"binary"=>AtomClass::Binary,"relation"=>AtomClass::Relation,
            "operator"=>AtomClass::Operator,"open"=>AtomClass::Open,"close"=>AtomClass::Close,
            "punctuation"=>AtomClass::Punctuation,_=>return Err(bad("Invalid custom symbol mathematical role")),
        };
        let limits_behavior:SymbolLimits=serde_json::from_value(definition["limitsBehavior"].clone()).map_err(|_|bad("Invalid custom symbol limits policy"))?;
        let metrics:SymbolMetrics=serde_json::from_value(definition["metrics"].clone()).map_err(|_|bad("Invalid custom symbol metrics"))?;
        if !(0.02..=64.0).contains(&metrics.width_em)||!(0.02..=64.0).contains(&metrics.ascent_em)
            ||!(0.0..=64.0).contains(&metrics.descent_em){return Err(bad("Custom symbol metrics exceed the legacy bounds"));}
        for key in ["createdAt","updatedAt"]{number(&definition,key)?;}
        if let Some(fallback)=object.get("ommlFallback").filter(|v|!v.is_null()) {
            if !fallback.is_string()||fallback.as_str().unwrap().len()>8000 {return Err(bad("Invalid custom symbol semantic fallback"));}
            // The string is retained only. It never participates in artwork
            // compilation and cannot trigger TeX expansion or a file read.
        }
        if let Some(designer)=object.get("designerSource").filter(|v|!v.is_null()) {
            if !designer.is_object()||designer["version"]!=1 {return Err(bad("Unknown custom symbol designer archive version"));}
            let assets=designer["assets"].as_array().ok_or_else(||bad("Missing designer assets"))?;
            let layers=designer["layers"].as_array().ok_or_else(||bad("Missing designer layers"))?;
            if assets.len()>96||layers.len()>192{return Err(limit("Designer archive exceeds the legacy item budget"));}
        }
        let shapes=definition["artwork"]["shapes"].as_array().ok_or_else(||bad("Missing custom symbol artwork"))?;
        if shapes.is_empty()||shapes.len()>4096{return Err(limit("Custom symbol artwork exceeds the shape budget"));}
        let mut geometry=Vec::new();let mut diagnostics=Vec::new();let mut ops=0usize;
        for (index,shape) in shapes.iter().enumerate(){
            match geometry::compile(shape,metrics.ascent_em) {
                Ok(paths)=>for path in paths {ops=ops.checked_add(path.path.len()).ok_or_else(||limit("Artwork budget overflow"))?;
                    if ops>MAX_ARTWORK_OPS{return Err(limit("Custom artwork exceeds the compiled path budget"));}geometry.push(Arc::new(path));},
                Err(error) if error.code==ErrorCode::UnsupportedSyntax=>diagnostics.push(ArtworkIssue{shape_index:index,code:"unsupportedArtwork".into(),message:error.message}),
                Err(error)=>return Err(error),
            }
        }
        // Never draw a partial symbol while dropping a layer. A retained,
        // unsupported snapshot cannot be promoted to an export-ready resource.
        if !diagnostics.is_empty(){geometry.clear();}
        let bytes=serde_json::to_vec(&definition).map_err(|_|bad("Custom symbol serialization failed"))?;
        let sha256=format!("{:x}",Sha256::digest(&bytes));
        let retained_weight=bytes.len()+geometry.iter().map(|p|p.path.len()*64).sum::<usize>();
        Ok(Arc::new(Self{sha256,command,role,limits_behavior,metrics,definition,geometry,diagnostics,retained_weight}))
    }
    pub fn sha256(&self)->&str{&self.sha256}
    pub fn command(&self)->&str{&self.command}
    pub fn role(&self)->AtomClass{self.role}
    pub fn limits(&self)->SymbolLimits{self.limits_behavior}
    pub fn metrics(&self)->SymbolMetrics{self.metrics}
    pub fn definition(&self)->&Value{&self.definition}
    pub fn geometry(&self)->&[Arc<ArtworkPath>]{&self.geometry}
    pub fn supported(&self)->bool{self.diagnostics.is_empty()}
    pub fn ensure_supported(&self)->EngineResult<()> {
        if self.supported(){Ok(())}else{Err(unsupported("Custom artwork is retained but needs an unsupported rendering capability"))}
    }
    pub fn retained_bytes(&self)->usize { self.retained_weight }
}
#[derive(Debug,Clone,Default,Serialize)]
#[serde(transparent)]
pub struct SymbolCatalog(BTreeMap<String,Arc<SymbolAsset>>);
impl SymbolCatalog {
    pub fn freeze(definitions:Vec<Value>)->EngineResult<Arc<Self>> {
        if definitions.len()>MAX_SYMBOLS{return Err(limit("Custom symbol count exceeds the catalog budget"));}
        let mut catalog=BTreeMap::new();let mut ids=std::collections::BTreeSet::new();let mut bytes=0;
        for definition in definitions {
            let asset=SymbolAsset::freeze(definition)?;bytes+=asset.retained_bytes();
            if bytes>MAX_LIBRARY_BYTES{return Err(limit("Custom symbol catalog exceeds its memory budget"));}
            if !ids.insert(asset.definition()["id"].as_str().unwrap().to_owned())||catalog.insert(asset.command().to_owned(),asset).is_some(){return Err(bad("Duplicate custom symbol id or command"));}
        }
        Ok(Arc::new(Self(catalog)))
    }
    pub fn from_library(library:&Value)->EngineResult<Arc<Self>> {
        check_json_budget(library,MAX_LIBRARY_BYTES)?;
        if library["version"]!=1{return Err(bad("Unknown custom symbol library version"));}
        Self::freeze(library["symbols"].as_array().ok_or_else(||bad("Missing custom symbol library records"))?.clone())
    }
    pub fn merge(&self,incoming:&Self)->EngineResult<Arc<Self>> {
        let mut all=self.0.clone();
        for (name,asset) in &incoming.0 {
            if let Some(old)=all.get(name){if old.sha256()!=asset.sha256(){return Err(EngineError::new(ErrorCode::Conflict,"Custom symbol name already belongs to a different frozen definition"));}continue;}
            if all.values().any(|old|old.definition()["id"]==asset.definition()["id"]){return Err(EngineError::new(ErrorCode::Conflict,"Custom symbol identifier already belongs to a different definition"));}
            all.insert(name.clone(),Arc::clone(asset));
        }
        let result=Self(all);if result.0.len()>MAX_SYMBOLS||result.retained_bytes()>MAX_LIBRARY_BYTES{return Err(limit("Merged symbol catalog exceeds its resource budget"));}Ok(Arc::new(result))
    }
    pub fn get(&self,command:&str)->Option<Arc<SymbolAsset>>{self.0.get(command.strip_prefix('\\').unwrap_or(command)).cloned()}
    pub fn is_empty(&self)->bool{self.0.is_empty()}
    pub fn definitions(&self)->Vec<Value>{self.0.values().map(|a|a.definition().clone()).collect()}
    pub fn retained_bytes(&self)->usize{self.0.values().map(|a|a.retained_bytes()).sum()}
}
pub fn empty_catalog(catalog:&Arc<SymbolCatalog>)->bool{catalog.is_empty()}
pub(crate) fn materialize_text(shape:&ArtworkPath,font:&crate::fonts::FontAsset,ascent:f64)->EngineResult<ArtworkPath>{geometry::materialize_text(shape,font,ascent)}
pub fn validate_bindings(tree:&crate::model::Node,catalog:&SymbolCatalog)->EngineResult<bool>{
    let mut error=None;let mut supported=true;
    tree.visit(&mut |node|if let crate::model::NodeKind::CustomSymbol{command,asset_sha256,asset,..}=&node.kind {
        if catalog.get(command).is_none_or(|bound|bound.sha256()!=asset_sha256||bound.sha256()!=asset.sha256()){
            error=Some(bad("Custom symbol node is not bound to this document's immutable resource snapshot"));
        }
        supported&=asset.supported();
    });
    match error{Some(error)=>Err(error),None=>Ok(supported)}
}
fn reserved(name:&str)->bool {
    let command=format!("\\{name}");
    matches!(name,"begin"|"end"|"left"|"right"|"middle"|"frac"|"dfrac"|"tfrac"|"sqrt"|"placeholder"|"text"|"operatorname"|"mathop"|"color"|"textcolor"|"boxed"|"def"|"gdef"|"let"|"newcommand"|"renewcommand"|"providecommand"|"input"|"include"|"write"|"openout"|"read"|"csname"|"endcsname"|"catcode"|"usepackage"|"class"|"href"|"htmlClass"|"cssId"|"htmlId"|"rule"|"raise"|"raisebox"|"mathord"|"mathbin"|"mathrel"|"mathopen"|"mathclose"|"mathpunct"|"kern"|"mkern"|"hskip"|"vskip"|"hspace"|"vspace"|"unicode"|"char"|"over"|"atop"|"above"|"choose")
        ||crate::symbols::named_symbol(&command).is_some()||crate::operators::spec(&command).is_some()
        ||crate::operators::LimitMode::from_command(&command).is_some()||crate::macros::spec(&command).is_some()
        ||crate::styles::command_variant(&command).is_some()||crate::accents::spec(&command).is_some()
        ||crate::fractions::spec(&command).is_some()||crate::spacing::space_mu(&command).is_some()
        ||crate::spacing::phantom_axis(&command).is_some()||crate::accents::annotation_above(&command).is_some()
        ||crate::delimiters::glyph(&command).is_some()
}
pub(super) fn bad(message:&str)->EngineError{EngineError::new(ErrorCode::InvalidInput,message)}
pub(super) fn limit(message:&str)->EngineError{EngineError::new(ErrorCode::ResourceLimit,message)}
pub(super) fn unsupported(message:&str)->EngineError{EngineError::new(ErrorCode::UnsupportedSyntax,message)}
pub(super) fn string<'a>(v:&'a Value,key:&str)->EngineResult<&'a str>{v[key].as_str().ok_or_else(||bad("Expected a custom symbol string field"))}
pub(super) fn number(v:&Value,key:&str)->EngineResult<f64>{v[key].as_f64().filter(|n|n.is_finite()&&n.abs()<=1e16).ok_or_else(||bad("Invalid custom symbol numeric field"))}
pub(super) fn optional_number(v:&Value,key:&str,default:f64)->EngineResult<f64>{if v.get(key).is_some(){number(v,key)}else{Ok(default)}}
pub(super) fn check_json_budget(v:&Value,max:usize)->EngineResult<()> {
    fn walk(v:&Value,depth:usize,count:&mut usize,bytes:&mut usize,max:usize)->EngineResult<()> {
        *count+=1;if depth>48||*count>131072{return Err(limit("Custom symbol data nesting/item budget exceeded"));}
        match v {Value::String(s)=>*bytes+=s.len(),Value::Array(a)=>for v in a{walk(v,depth+1,count,bytes,max)?;},
            Value::Object(m)=>for(k,v)in m{*bytes+=k.len();walk(v,depth+1,count,bytes,max)?;},Value::Number(n)=>{if n.as_f64().is_none_or(|n|!n.is_finite()){return Err(bad("Non-finite custom symbol number"));}*bytes+=32;},_=>*bytes+=8}
        if *bytes>max{return Err(limit("Custom symbol data byte budget exceeded"));}Ok(())
    }
    walk(v,0,&mut 0,&mut 0,max)
}
