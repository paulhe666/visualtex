//! Read-only queries against a revision- and resource-bound layout snapshot.
//! Hosts draw the returned geometry; they do not reinterpret math/source offsets.
use serde::{Deserialize, Serialize};
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::model::{Node, NodeKind, ScriptKind};
use crate::positions::{NormalizedRange, Position, Selection, TreeIndex};
use crate::scene::{LayoutSnapshot, Rect};
use crate::session::NodeId;

#[derive(Debug, Clone, Deserialize)]
#[serde(tag="kind",rename_all="camelCase",rename_all_fields="camelCase",deny_unknown_fields)]
pub enum SceneQuery {
    HitTest { x_em: f64, y_em: f64 },
    Vertical { from: Position, down: bool, desired_x: Option<f64> },
    Navigation { selection: Selection, down: bool, desired_x: Option<f64> },
    Boundary { selection: Selection, to_end: bool, document: bool },
    Selection { selection: Selection },
    ExportSvg,
}
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all="camelCase")]
pub struct PositionDescription {
    pub path: Vec<String>,
    pub offset: usize,
    pub length: usize,
    pub text: Option<String>,
}
#[derive(Debug, Clone, Serialize)]
#[serde(tag="kind",rename_all="camelCase",rename_all_fields="camelCase")]
pub enum SceneQueryResult {
    HitTest { position: Position },
    Vertical { position: Position, origin: Position, desired_x: f64 },
    Boundary { position: Position, origin: Position },
    Selection { caret: Option<Position>, rectangles: Vec<Rect>, description: PositionDescription },
    Svg { svg: String, width: f64, height: f64, baseline: f64 },
}
fn missing()->EngineError {EngineError::new(ErrorCode::InvalidPosition,"Selection is not covered by the current scene")}

pub(crate) fn query(scene:&LayoutSnapshot,tree:&Node,query:SceneQuery)->EngineResult<SceneQueryResult> {
    match query {
        SceneQuery::HitTest {x_em,y_em}=>Ok(SceneQueryResult::HitTest {position:scene.hit_test(x_em,y_em)?}),
        SceneQuery::Vertical {from,down,desired_x}=> {
            let x=desired_x.unwrap_or(scene.caret(from).ok_or_else(missing)?.x);
            Ok(SceneQueryResult::Vertical {position:scene.vertical(from,down,Some(x))?,origin:from,desired_x:x})
        }
        SceneQuery::Navigation {selection,down,desired_x}=> {
            let from=focus_position(tree,&selection,down)?;
            let x=desired_x.unwrap_or(scene.caret(from).ok_or_else(missing)?.x);
            Ok(SceneQueryResult::Vertical {position:scene.vertical(from,down,Some(x))?,origin:from,desired_x:x})
        }
        SceneQuery::Boundary {selection,to_end,document}=> {
            let from=focus_position(tree,&selection,to_end)?;
            let node=if document {tree} else {tree.find(from.node_id()).ok_or_else(missing)?};
            let position=match &node.kind {
                NodeKind::Row {entries,..}=>Position::row(node.id,if to_end {entries.len()} else {0}),
                NodeKind::Text {text,..}=> {
                    use unicode_segmentation::UnicodeSegmentation;
                    Position::text(node.id,if to_end {text.graphemes(true).count()} else {0})
                }
                _=>return Err(missing()),
            };
            scene.caret(position).ok_or_else(missing)?;
            Ok(SceneQueryResult::Boundary {position,origin:from})
        }
        SceneQuery::Selection {selection}=>selection_geometry(scene,tree,&selection),
        SceneQuery::ExportSvg=>Ok(SceneQueryResult::Svg {svg:scene.svg()?,width:scene.width,height:scene.height,baseline:scene.baseline}),
    }
}
fn focus_position(tree:&Node,selection:&Selection,after:bool)->EngineResult<Position> {
    let index=TreeIndex::new(tree)?;
    index.normalize_range(selection)?;
    match selection {
        Selection::GridCells {focus_cell,..}=>Ok(Position::row(*focus_cell,0)),
        Selection::Range {focus,..}=>Ok(*focus),
        Selection::Node {node_id}=> {
            let (row,offset)=index.row_owner(*node_id)?;
            Ok(Position::row(row,offset+usize::from(after)))
        }
    }
}
fn selection_geometry(scene:&LayoutSnapshot,tree:&Node,selection:&Selection)->EngineResult<SceneQueryResult> {
    let index=TreeIndex::new(tree)?;
    let normalized=index.normalize_range(selection)?;
    let mut rectangles=Vec::new();
    let mut caret=None;
    let focus=match selection {Selection::GridCells{focus_cell,..}=>Position::row(*focus_cell,0),Selection::Range {focus,..}=>*focus,Selection::Node {node_id}=>{
        let (row,offset)=index.row_owner(*node_id)?;Position::row(row,offset)
    }};
    let description=describe(tree,focus)?;
    if let Selection::Node {node_id}=selection {
        rectangles.push(node_rectangle(scene,*node_id)?);
    } else {
        match normalized {
            NormalizedRange::GridCells {rectangle}=> {
                for r in rectangle.row_start..=rectangle.row_end {
                    for c in rectangle.column_start..=rectangle.column_end {
                        let cell=index.cell_at(rectangle.grid_id,r,c)?;
                        let region=node_rectangle(scene,cell)?;
                        let column=scene.grid_columns.iter().find(|a|a.grid_id==rectangle.grid_id&&a.index==c).ok_or_else(missing)?;
                        // Cell highlight uses shared column geometry, including
                        // empty cells, and never an unrelated root rectangle.
                        rectangles.push(Rect::new(column.x,region.y_min,column.x+column.width,region.y_max));
                    }
                }
            }
            NormalizedRange::Row {node_id,start,end,..}=> {
                if start==end {caret=Some(focus);} else {
                    let row=index.node(node_id)?;
                    let NodeKind::Row {entries,..}=&row.kind else {return Err(missing());};
                    for entry in &entries[start as usize..end as usize] {rectangles.push(node_rectangle(scene,entry.node.id)?);}
                }
            }
            NormalizedRange::Text {node_id,start,end}=> {
                if start==end {caret=Some(focus);} else {
                    let a=scene.caret(Position::text(node_id,start as usize)).ok_or_else(missing)?;
                    let b=scene.caret(Position::text(node_id,end as usize)).ok_or_else(missing)?;
                    rectangles.push(Rect::new(a.x.min(b.x),a.top.min(b.top),a.x.max(b.x),a.bottom.max(b.bottom)));
                }
            }
        }
    }
    if let Some(position)=caret {scene.caret(position).ok_or_else(missing)?;}
    Ok(SceneQueryResult::Selection {caret,rectangles,description})
}
fn node_rectangle(scene:&LayoutSnapshot,id:NodeId)->EngineResult<Rect> {
    let region=scene.regions.iter().find(|region|region.node_id==id).ok_or_else(missing)?;
    Ok(region.ink.map_or(region.logical,|ink|region.logical.union(ink)))
}
fn describe(tree:&Node,position:Position)->EngineResult<PositionDescription> {
    let node=tree.find(position.node_id()).ok_or_else(missing)?;
    let mut path=vec!["root".to_owned()];
    if !describe_path(tree,node.id,&mut path) {return Err(missing());}
    let (offset,length,text)=match (position,&node.kind) {
        (Position::Row {offset,..},NodeKind::Row {entries,..})=>(offset as usize,entries.len(),None),
        (Position::Text {grapheme,..},NodeKind::Text {text,..})=> {
            use unicode_segmentation::UnicodeSegmentation;
            (grapheme as usize,text.graphemes(true).count(),Some(text.chars().take(128).collect()))
        }
        _=>return Err(missing()),
    };
    Ok(PositionDescription {path,offset,length,text})
}
fn describe_path(node:&Node,target:NodeId,path:&mut Vec<String>)->bool {
    if node.id==target {return true;}
    if let NodeKind::Grid{rows,..}=&node.kind {
        for (i,row) in rows.iter().enumerate(){path.push(format!("gridRow:{}",i+1));if describe_path(row,target,path){return true;}path.pop();}
        return false;
    }
    if let NodeKind::GridRow{cells,..}=&node.kind {
        for (i,cell) in cells.iter().enumerate(){path.push(format!("gridCell:{}",i+1));if describe_path(cell,target,path){return true;}path.pop();}
        return false;
    }
    let children:Vec<(&Node,&str)>=match &node.kind {
        NodeKind::Fraction {numerator,denominator,..}=>vec![(numerator,"numerator"),(denominator,"denominator")],
        NodeKind::Fence {body,..}=>vec![(body.as_ref(),"fenceBody")],
        NodeKind::Accent {body,..}=>vec![(body.as_ref(),"accentBody")],
        NodeKind::Style {body,..}=>vec![(body.as_ref(),"styleBody")],
        NodeKind::Phantom {body,..}=>vec![(body.as_ref(),"phantomBody")],
        NodeKind::Boxed {body,..}=>vec![(body.as_ref(),"boxedBody")],
        NodeKind::Color {body,..}=>vec![(body.as_ref(),"colorBody")],
        NodeKind::LargeOperator {parts,..}|NodeKind::CustomSymbol {parts,..}=>crate::operators::slots(parts).map(|slot|
            (slot.argument.as_ref(),if slot.script==ScriptKind::Subscript{"lowerLimit"}else{"upperLimit"})).collect(),
        NodeKind::GeneralOperator {body,parts,..}=>std::iter::once((body.as_ref(),"operatorBody"))
            .chain(crate::operators::slots(parts).map(|slot|(slot.argument.as_ref(),if slot.script==ScriptKind::Subscript{"lowerLimit"}else{"upperLimit"}))).collect(),
        NodeKind::MacroCall {arguments,..}=>arguments.iter().enumerate().map(|(index,argument)|
            (argument.as_ref(),match index{0=>"macroArgument:1",1=>"macroArgument:2",_=>"macroArgument:3"})).collect(),
        NodeKind::Annotation {annotation,base,..}=>vec![(annotation.as_ref(),"annotationLabel"),(base.as_ref(),"annotationBase")],
        NodeKind::Radical {degree,radicand,..}=>degree.iter().map(|d|(d.as_ref(),"degree"))
            .chain(std::iter::once((radicand.as_ref(),"radicand"))).collect(),
        NodeKind::Scripts {base,slots}=> {
            let mut children=vec![(base.as_ref(),"base")];
            children.extend(slots.iter().map(|slot|(slot.argument.as_ref(),if slot.script==ScriptKind::Superscript {"superscript"} else {"subscript"})));children
        }
        NodeKind::Row {entries,..}=>entries.iter().map(|entry| {
            let kind=match entry.node.kind {NodeKind::Row {..}=>"group",NodeKind::Grid{..}=>"grid",NodeKind::GridRow{..}=>"gridRow",NodeKind::Fraction {..}=>"fraction",NodeKind::Radical {..}=>"radical",NodeKind::Fence {..}=>"fence",NodeKind::Accent {..}=>"accent",NodeKind::Annotation {..}=>"annotation",NodeKind::LargeOperator {..}=>"largeOperator",NodeKind::GeneralOperator {..}=>"generalOperator",NodeKind::CustomSymbol {..}=>"customSymbol",NodeKind::MacroCall {..}=>"macroCall",NodeKind::MiddleDelimiter {..}=>"middleDelimiter",NodeKind::Style {..}=>"style",NodeKind::Phantom {..}=>"phantom",NodeKind::Boxed {..}=>"boxed",NodeKind::Color {..}=>"color",NodeKind::ColorChange {..}=>"colorChange",NodeKind::StyleChange {..}=>"styleChange",NodeKind::Space {..}=>"space",NodeKind::Scripts {..}=>"scripts",NodeKind::Text {..}=>"text",_=>"symbol"};
            (entry.node.as_ref(),kind)
        }).collect(),
        _=>Vec::new(),
    };
    for (child,role) in children {
        path.push(role.to_owned());if describe_path(child,target,path) {return true;}path.pop();
    }
    false
}
