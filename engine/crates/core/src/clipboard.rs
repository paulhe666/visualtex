//! Clipboard boundary for the self-owned engine.
//! Plain LaTeX remains the portable fallback. The optional structural payload is
//! versioned, contains no live NodeId/session identity and is never hydrated into
//! a tree directly: paste re-parses the bound LaTeX and requires the derived
//! identity-free shape to match byte-for-byte semantic metadata first.
use std::sync::Arc;
use serde::{Deserialize,Serialize};
use serde_json::{Map,Value};
use sha2::{Digest,Sha256};
use unicode_segmentation::UnicodeSegmentation;
use crate::error::{EngineError,EngineResult,ErrorCode};
use crate::model::{Grouping,IdAllocator,LatexMode,MAX_DEPTH,MAX_NODES,Node,NodeKind,latex};
use crate::parser;
use crate::positions::{NormalizedRange,Selection,TreeIndex};
use crate::session::{EngineSession,NodeId,Revision};
use crate::source::{MAX_SOURCE_BYTES,SourceText};

pub const STRUCTURAL_CLIPBOARD_SCHEMA:&str="visualtex-engine-fragment";
pub const STRUCTURAL_CLIPBOARD_VERSION:u32=1;
pub const MAX_STRUCTURAL_CLIPBOARD_BYTES:usize=512*1024;
const MAX_STRUCTURAL_STRING_BYTES:usize=64*1024;
const MAX_STRUCTURAL_OBJECTS:usize=MAX_NODES*4;
const MAX_STRUCTURAL_JSON_DEPTH:usize=MAX_DEPTH*4+32;

#[derive(Debug,Clone,PartialEq,Eq,Serialize,Deserialize)]
#[serde(rename_all="camelCase",deny_unknown_fields)]
pub struct ClipboardAssetRef {
    pub kind:String,
    pub asset_id:String,
    pub sha256:String,
    #[serde(default,skip_serializing_if="Option::is_none")]
    pub definition:Option<Value>,
}

#[derive(Debug,Clone,PartialEq,Serialize,Deserialize)]
#[serde(rename_all="camelCase",deny_unknown_fields)]
pub struct StructuredClipboardFragment {
    pub schema:String,
    pub version:u32,
    pub source_sha256:String,
    pub shape:Value,
    #[serde(default)]
    pub assets:Vec<ClipboardAssetRef>,
}

#[derive(Debug,Clone,Serialize)]
#[serde(rename_all="camelCase")]
pub struct ClipboardFragment {
    /// Protocol-level clipboard reply version. Kept at v1 for compatibility;
    /// the independent `structure.version` owns structural-format evolution.
    pub schema_version:u32,
    pub latex:Option<String>,
    #[serde(skip_serializing_if="Option::is_none")]
    pub structure:Option<StructuredClipboardFragment>,
    pub revision:Revision,
    pub selection_epoch:u32,
}

impl EngineSession {
    pub fn copy_selection(&self,expected_revision:Revision)->EngineResult<ClipboardFragment> {
        self.check_revision(expected_revision)?;
        self.ensure_no_composition()?;
        let index=TreeIndex::new(&self.document.tree)?;
        let selection=&self.editor.selection;
        index.validate_matrix_fragment(selection)?;
        let source=match selection {
            Selection::GridCells{grid_id,anchor_cell,focus_cell}=>Some(grid_fragment(&index,&index.grid_rectangle(*grid_id,*anchor_cell,*focus_cell)?)?),
            Selection::Node {node_id}=>Some(latex(index.node(*node_id)?,LatexMode::Lossless)),
            Selection::Range {..}=>match index.normalize_range(selection)? {
                NormalizedRange::GridCells{..}=>return Err(copy_error()),
                NormalizedRange::Row {node_id,start,end,..}=> {
                    if start==end {None} else {
                        let row=index.node(node_id)?;
                        let NodeKind::Row {entries,..}=&row.kind else {return Err(copy_error());};
                        let mut selected=entries.get(start as usize..end as usize).ok_or_else(copy_error)?.to_vec();
                        // Trivia preceding the first selected atom belongs to
                        // its old row, not to the selected mathematical object.
                        if let Some(first)=selected.first_mut() {first.leading.clear();}
                        let fragment=Node {id:row.id,origin:None,kind:NodeKind::Row {grouping:Grouping::Bare,entries:selected,trailing:String::new()}};
                        Some(latex(&fragment,LatexMode::Lossless))
                    }
                }
                NormalizedRange::Text {node_id,start,end}=> {
                    if start==end {None} else {
                        let text_node=index.node(node_id)?;
                        let NodeKind::Text {text,command,..}=&text_node.kind else {return Err(copy_error());};
                        let selected=text.graphemes(true).skip(start as usize).take((end-start) as usize).collect();
                        let fragment=Node {id:node_id,origin:None,kind:NodeKind::Text {command:command.clone(),before_body:String::new(),text:selected}};
                        Some(latex(&fragment,LatexMode::Lossless))
                    }
                }
            },
        };
        let structure=match source.as_deref(){Some(source)=>build_structural_fragment(source,Arc::clone(&self.document.symbol_assets))?,None=>None};
        Ok(ClipboardFragment {schema_version:1,latex:source,structure,revision:self.revision,selection_epoch:self.editor.selection_epoch})
    }
}

/// Validate the custom clipboard flavor before a normal paste transaction.
/// The payload is data only. No HTML, JavaScript, TeX definitions, live IDs or
/// caller-provided tree objects are executed/hydrated.
pub fn validate_structural_fragment(latex:&str,fragment:&StructuredClipboardFragment)->EngineResult<()> {
    validate_structural_fragment_with_symbols(latex,fragment,Arc::default()).map(|_|())
}
pub fn validate_structural_fragment_with_symbols(latex:&str,fragment:&StructuredClipboardFragment,target:Arc<crate::custom_symbols::SymbolCatalog>)->EngineResult<Arc<crate::custom_symbols::SymbolCatalog>> {
    if fragment.schema!=STRUCTURAL_CLIPBOARD_SCHEMA || fragment.version!=STRUCTURAL_CLIPBOARD_VERSION {
        return Err(EngineError::new(ErrorCode::InvalidInput,"Unsupported VisualTeX structural clipboard version"));
    }
    if latex.len()>MAX_SOURCE_BYTES {return Err(EngineError::new(ErrorCode::ResourceLimit,"Clipboard LaTeX exceeds the source budget"));}
    if fragment.source_sha256!=source_sha256(latex) {
        return Err(EngineError::new(ErrorCode::InvalidInput,"Structural clipboard source digest does not match its LaTeX fallback"));
    }
    let mut definitions=Vec::new();let mut resource_hashes=std::collections::BTreeSet::new();
    for reference in &fragment.assets {
        if reference.kind!="customSymbol"||reference.definition.is_none(){return Err(EngineError::new(ErrorCode::UnsupportedSyntax,"Clipboard resource requires a supported self-contained symbol snapshot"));}
        let asset=crate::custom_symbols::SymbolAsset::freeze(reference.definition.clone().unwrap())?;
        asset.ensure_supported()?;
        if asset.sha256()!=reference.sha256||asset.definition()["id"]!=reference.asset_id||!resource_hashes.insert(reference.sha256.clone()){
            return Err(EngineError::new(ErrorCode::InvalidInput,"Clipboard symbol resource digest/id is inconsistent or duplicated"));
        }
        definitions.push(asset.definition().clone());
    }
    let incoming=crate::custom_symbols::SymbolCatalog::freeze(definitions)?;
    let symbols=target.merge(&incoming)?;
    let encoded=serde_json::to_vec(fragment).map_err(|_|EngineError::new(ErrorCode::InvalidInput,"Structural clipboard cannot be encoded deterministically"))?;
    if encoded.len()>MAX_STRUCTURAL_CLIPBOARD_BYTES {
        return Err(EngineError::new(ErrorCode::ResourceLimit,"Structural clipboard exceeds its bounded payload size"));
    }
    validate_shape_budget(&fragment.shape)?;
    let mut ids=IdAllocator::new(2);
    let parsed=parser::parse_with_symbols(SourceText::new(1,latex)?,NodeId(1),&mut ids,Arc::clone(&symbols))?;
    let used=used_symbols(&parsed.tree);if used.keys().cloned().collect::<std::collections::BTreeSet<_>>()!=resource_hashes{return Err(EngineError::new(ErrorCode::InvalidInput,"Clipboard symbol snapshots must match exactly the referenced artwork"));}
    if !parsed.complete {return Err(EngineError::new(ErrorCode::IncompleteSource,"Structural clipboard LaTeX is incomplete"));}
    if !parsed.supported {return Err(EngineError::new(ErrorCode::UnsupportedSyntax,"Structural clipboard LaTeX is outside the current visual capability"));}
    let derived=portable_shape(&parsed.tree)?;
    if derived!=fragment.shape {
        return Err(EngineError::new(ErrorCode::InvalidInput,"Structural clipboard metadata does not match the bound LaTeX structure"));
    }
    Ok(symbols)
}

fn build_structural_fragment(latex:&str,symbols:Arc<crate::custom_symbols::SymbolCatalog>)->EngineResult<Option<StructuredClipboardFragment>> {
    if latex.len()>MAX_SOURCE_BYTES{return Ok(None);}
    let mut ids=IdAllocator::new(2);
    let parsed=match parser::parse_with_symbols(SourceText::new(1,latex)?,NodeId(1),&mut ids,symbols) {
        Ok(parsed)=>parsed,
        Err(error) if error.code==ErrorCode::ResourceLimit=>return Ok(None),
        Err(error)=>return Err(error),
    };
    // Unsupported/incomplete source still copies through text/plain/LaTeX, but
    // it must not receive a structural trust marker.
    if !parsed.complete||!parsed.supported{return Ok(None);}
    let shape=portable_shape(&parsed.tree)?;
    validate_shape_budget(&shape)?;
    let fragment=StructuredClipboardFragment{
        schema:STRUCTURAL_CLIPBOARD_SCHEMA.into(),version:STRUCTURAL_CLIPBOARD_VERSION,
        source_sha256:source_sha256(latex),shape,assets:used_symbols(&parsed.tree).into_values().map(|asset|ClipboardAssetRef{kind:"customSymbol".into(),asset_id:asset.definition()["id"].as_str().unwrap().into(),sha256:asset.sha256().into(),definition:Some(asset.definition().clone())}).collect(),
    };
    let encoded=serde_json::to_vec(&fragment).map_err(|_|EngineError::new(ErrorCode::InvalidInput,"Structural clipboard cannot be encoded deterministically"))?;
    if encoded.len()>MAX_STRUCTURAL_CLIPBOARD_BYTES{
        if !fragment.assets.is_empty(){return Err(EngineError::new(ErrorCode::ResourceLimit,"Custom symbol clipboard cannot discard its asset snapshot to fit the payload budget"));}
        return Ok(None);
    }
    Ok(Some(fragment))
}

fn used_symbols(tree:&Node)->std::collections::BTreeMap<String,Arc<crate::custom_symbols::SymbolAsset>> {
    let mut result=std::collections::BTreeMap::new();tree.visit(&mut|node|if let NodeKind::CustomSymbol{asset,..}=&node.kind{result.insert(asset.sha256().into(),Arc::clone(asset));});result
}
fn source_sha256(source:&str)->String {
    let digest=Sha256::digest(source.as_bytes());
    digest.iter().map(|byte|format!("{byte:02x}")).collect()
}

/// Convert the serializable core tree to a stable identity-free semantic shape.
/// Exact source spelling is already bound by sourceSha256; source trivia and
/// live identity are deliberately excluded from the structural trust record.
fn portable_shape(node:&Node)->EngineResult<Value> {
    let raw=serde_json::to_value(node).map_err(|_|EngineError::new(ErrorCode::InvalidInput,"Unable to project structural clipboard metadata"))?;
    Ok(strip_identity_and_trivia(raw))
}
fn strip_identity_and_trivia(value:Value)->Value {
    match value {
        Value::Array(values)=>Value::Array(values.into_iter().map(strip_identity_and_trivia).collect()),
        Value::Object(map)=>{
            let is_grid=map.get("kind").and_then(Value::as_str)==Some("grid");
            let mut out=Map::new();
            for (key,value) in map {
                if key=="id"||key=="origin"||key=="leading"||key=="trailing"||key.starts_with("before") {continue;}
                if is_grid&&key=="columns" {
                    out.insert(key,Value::from(value.as_array().map_or(0,Vec::len) as u64));
                }else{out.insert(key,strip_identity_and_trivia(value));}
            }
            Value::Object(out)
        }
        other=>other,
    }
}

fn validate_shape_budget(shape:&Value)->EngineResult<()> {
    fn allowed_kind(kind:&str)->bool {matches!(kind,
        "row"|"grid"|"gridRow"|"symbol"|"text"|"fraction"|"radical"|"fence"|"middleDelimiter"|
        "accent"|"annotation"|"largeOperator"|"generalOperator"|"customSymbol"|"macroCall"|"style"|"space"|"boxed"|
        "color"|"colorChange"|"phantom"|"scripts"|"emptySlot"|"script"|"limits")}
    fn walk(value:&Value,depth:usize,objects:&mut usize)->EngineResult<()> {
        if depth>MAX_STRUCTURAL_JSON_DEPTH{return Err(EngineError::new(ErrorCode::ResourceLimit,"Structural clipboard nesting budget exceeded"));}
        match value {
            Value::String(text)=>if text.len()>MAX_STRUCTURAL_STRING_BYTES{return Err(EngineError::new(ErrorCode::ResourceLimit,"Structural clipboard string budget exceeded"));},
            Value::Array(values)=>{
                if values.len()>MAX_STRUCTURAL_OBJECTS{return Err(EngineError::new(ErrorCode::ResourceLimit,"Structural clipboard array budget exceeded"));}
                for value in values{walk(value,depth+1,objects)?;}
            }
            Value::Object(map)=>{
                *objects+=1;if *objects>MAX_STRUCTURAL_OBJECTS{return Err(EngineError::new(ErrorCode::ResourceLimit,"Structural clipboard object budget exceeded"));}
                if map.len()>64{return Err(EngineError::new(ErrorCode::ResourceLimit,"Structural clipboard object field budget exceeded"));}
                if let Some(kind)=map.get("kind") {
                    let kind=kind.as_str().ok_or_else(||EngineError::new(ErrorCode::InvalidInput,"Structural clipboard node kind must be textual"))?;
                    if !allowed_kind(kind){return Err(EngineError::new(ErrorCode::InvalidInput,"Structural clipboard contains an unknown node kind"));}
                }
                for (key,value) in map {
                    if key.len()>128{return Err(EngineError::new(ErrorCode::ResourceLimit,"Structural clipboard key budget exceeded"));}
                    walk(value,depth+1,objects)?;
                }
            }
            Value::Number(number)=>if number.as_i64().is_none()&&number.as_u64().is_none(){return Err(EngineError::new(ErrorCode::InvalidInput,"Structural clipboard numeric metadata must be integral"));},
            Value::Bool(_) | Value::Null=>{},
        }Ok(())
    }
    let mut objects=0;walk(shape,0,&mut objects)
}

/// Portable rectangular payload. Only source is emitted, never live IDs.
fn grid_fragment(index:&TreeIndex<'_>,rect:&crate::grid_selection::GridRectangle)->EngineResult<String> {
    use std::sync::Arc;
    let original=index.node(rect.grid_id)?;
    let NodeKind::Grid{columns,rows,..}=&original.kind else{return Err(copy_error())};
    let mut selected=Vec::new();
    for (i,row) in rows[rect.row_start..=rect.row_end].iter().enumerate() {
        let mut row=row.as_ref().clone();row.origin=None;
        let NodeKind::GridRow{cells,separators,row_break}=&mut row.kind else{return Err(copy_error())};
        *cells=cells[rect.column_start..=rect.column_end].to_vec();
        *separators=vec!["&".into();rect.columns()-1];
        *row_break=if i+1<rect.rows(){"\\\\".into()}else{String::new()};
        selected.push(Arc::new(row));
    }
    let fragment=Node{id:original.id,origin:None,kind:NodeKind::Grid{
        environment:crate::grids::GridKind::Matrix,
        columns:columns[rect.column_start..=rect.column_end].to_vec(),array_spec:None,rows:selected,
        before_begin_name:String::new(),before_end_name:String::new()}};
    Ok(latex(&fragment,LatexMode::Lossless))
}
fn copy_error()->EngineError {EngineError::new(ErrorCode::InvalidPosition,"Clipboard selection is not a valid mathematical fragment")}
