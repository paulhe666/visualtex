//! P09 bounded array environment: authored l/c/r columns and vertical rules.
//! Arrays remain one Grid owner; fences around an augmented matrix remain normal
//! Fence structure. No TeX pre-expansion or DOM table geometry is involved.
use std::{collections::BTreeSet,path::PathBuf};
use visualtex_engine_core::{editing::{EditorAction,EditorActionResult},fonts::FontRegistry,
    grids::{GridColumnAlignment,GridKind},layout::{LayoutEngine,LayoutOptions},model::{self,LatexMode,Node,NodeKind},
    positions::{Position,Selection},session::EngineSession,transactions::{EditCommand,InputMode}};

fn find_grid_opt(node:&Node)->Option<&Node>{
    if matches!(node.kind,NodeKind::Grid{environment:GridKind::Array,..}){return Some(node);}
    for child in node.children(){if let Some(found)=find_grid_opt(child){return Some(found);}}None
}
fn find_grid(node:&Node)->&Node{find_grid_opt(node).expect("array grid")}
fn find_fence(node:&Node)->Option<&Node>{
    if matches!(node.kind,NodeKind::Fence{..}){return Some(node);}
    for child in node.children(){if let Some(found)=find_fence(child){return Some(found);}}None
}
fn cell(grid:&Node,row:usize,column:usize)->&Node{
    let NodeKind::Grid{rows,..}=&grid.kind else{panic!()};let NodeKind::GridRow{cells,..}=&rows[row].kind else{panic!()};&cells[column]
}
fn ids(node:&Node)->BTreeSet<u32>{let mut out=BTreeSet::new();node.visit(&mut|n|{assert!(out.insert(n.id.0));if let NodeKind::Grid{columns,..}=&n.kind{for c in columns{assert!(out.insert(c.0));}}});out}
fn font()->(FontRegistry,LayoutOptions){
    let path=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    let mut fonts=FontRegistry::default();let id=fonts.register(&std::fs::read(path).unwrap(),0).unwrap().id;(fonts,LayoutOptions::new(id))
}
struct Owned{session:EngineSession,sequence:u32}
impl Owned{
    fn new(source:&str)->Self{Self{session:EngineSession::from_source("array-review".into(),source.into()).unwrap(),sequence:0}}
    fn act(&mut self,action:EditorAction)->EditorActionResult{self.sequence+=1;self.session.apply_action(self.session.revision(),self.sequence,action).unwrap()}
    fn edit(&mut self,command:EditCommand)->EditorActionResult{self.act(EditorAction::Edit{command,group_id:None})}
    fn reject(&mut self,command:EditCommand){let before=serde_json::to_value(self.session.editor_snapshot().unwrap()).unwrap();self.sequence+=1;assert!(self.session.apply_action(self.session.revision(),self.sequence,EditorAction::Edit{command,group_id:None}).is_err());assert_eq!(serde_json::to_value(self.session.editor_snapshot().unwrap()).unwrap(),before);}
}

#[test]
fn arrays_preserve_authored_column_spec_and_source_byte_for_byte(){
    for source in [
        r"\begin{array}{cc|c}a&b&c\\d&e&f\end{array}",
        r"\begin{array}{l|cr}a&b&c\\d&e&f\end{array}",
        r"\begin {array} { c c || c }a&b&c\\d&e&f\end {array}",
    ]{
        let s=EngineSession::from_source("array-source".into(),source.into()).unwrap();let d=s.document_snapshot();
        assert!(d.analysis.complete&&d.analysis.supported&&d.analysis.exportable,"{:?}",d.analysis.diagnostics);
        assert_eq!(model::latex(&d.analysis.tree,LatexMode::Lossless),source);
        let g=find_grid(&d.analysis.tree);let NodeKind::Grid{environment,columns,array_spec,rows,..}=&g.kind else{panic!()};
        assert_eq!(*environment,GridKind::Array);let spec=array_spec.as_ref().expect("array spec");assert_eq!(spec.alignments.len(),columns.len());assert_eq!(rows.len(),2);
        assert_eq!(spec.canonical_source().chars().filter(|c|matches!(c,'l'|'c'|'r')).count(),columns.len());
    }
}

#[test]
fn augmented_matrix_is_real_outer_fence_around_one_array_grid(){
    let source=r"\left[\begin{array}{cc|c}a&b&c\\d&e&f\end{array}\right]";
    let s=EngineSession::from_source("augmented".into(),source.into()).unwrap();let d=s.document_snapshot();assert!(d.analysis.complete&&d.analysis.supported&&d.analysis.exportable,"{:?}",d.analysis.diagnostics);assert_eq!(model::latex(&d.analysis.tree,LatexMode::Lossless),source);assert_eq!(d.analysis.facts.grids,1);assert_eq!(d.analysis.facts.fences,1);
    let g=find_grid(&d.analysis.tree);let f=find_fence(&d.analysis.tree).unwrap();assert_ne!(g.id,f.id);
    let NodeKind::Grid{array_spec:Some(spec),columns,..}=&g.kind else{panic!()};assert_eq!(spec.source,"cc|c");assert_eq!(spec.vertical_rules,vec![2]);assert_eq!(columns.len(),3);
    let (fonts,options)=font();let mut engine=LayoutEngine::default();let scene=engine.layout(&s,&fonts,&options).unwrap();
    assert_eq!(scene.grid_columns.len(),3);assert_eq!(scene.rules.iter().filter(|r|r.node_id==g.id).count(),1,"array separator must be a real vector rule");assert!(scene.glyphs.iter().any(|glyph|glyph.node_id==f.id),"outer brackets must be real font glyphs");
    let rule=&scene.rules.iter().find(|r|r.node_id==g.id).unwrap().rect;let left=&scene.grid_columns[1];let right=&scene.grid_columns[2];assert!(rule.x_min>left.x&&rule.x_max<right.x+right.width);
    let svg=scene.svg().unwrap();for forbidden in ["<text","foreignObject","<image","NaN","Infinity"]{assert!(!svg.contains(forbidden));}
}

#[test]
fn array_alignment_and_rules_are_geometry_not_css_or_source_rewrites(){
    let source=r"\begin{array}{l|cr}a&b&c\\long&long&long\end{array}";let s=EngineSession::from_source("array-geometry".into(),source.into()).unwrap();let d=s.document_snapshot();let g=find_grid(&d.analysis.tree);
    let NodeKind::Grid{array_spec:Some(spec),..}=&g.kind else{panic!()};assert_eq!(spec.alignments,[GridColumnAlignment::Left,GridColumnAlignment::Center,GridColumnAlignment::Right]);assert_eq!(spec.vertical_rules,[1]);
    let (fonts,mut options)=font();let mut engine=LayoutEngine::default();let a=engine.layout(&s,&fonts,&options).unwrap();let rule=a.rules.iter().find(|r|r.node_id==g.id).unwrap();assert!(rule.rect.height()>1.0);assert_eq!(a.grid_columns.len(),3);assert!(rule.rect.x_min>a.grid_columns[0].x);assert!(rule.rect.x_max<a.grid_columns[1].x+a.grid_columns[1].width);
    options.font_size_px=32.0;let b=engine.layout(&s,&fonts,&options).unwrap();assert_eq!(b.width,a.width*2.0);assert_eq!(b.height,a.height*2.0);assert_eq!(serde_json::to_value(&b.grid_columns).unwrap(),serde_json::to_value(&a.grid_columns).unwrap());assert_eq!(serde_json::to_value(&b.rules).unwrap(),serde_json::to_value(&a.rules).unwrap());
}

#[test]
fn array_cells_are_editable_rows_and_row_insertion_keeps_column_spec(){
    let mut o=Owned::new(r"\begin{array}{cc|c}a&b&c\\d&e&f\end{array}");let before=o.session.editor_snapshot().unwrap();let g=find_grid(&before.document.analysis.tree);let grid_id=g.id;let original_ids=ids(g);let NodeKind::Grid{columns,array_spec:Some(spec),..}=&g.kind else{panic!()};let columns=columns.clone();let spec_source=spec.source.clone();let target=cell(g,0,1).id;
    o.act(EditorAction::SetSelection{selection:Selection::caret(Position::row(target,1))});let edited=o.edit(EditCommand::InsertText{text:"q".into(),mode:InputMode::Math}).editor;assert!(edited.document.source.contains("bq"));let g=find_grid(&edited.document.analysis.tree);assert_eq!(g.id,grid_id);let NodeKind::Grid{array_spec:Some(spec),columns:now,..}=&g.kind else{panic!()};assert_eq!(spec.source,spec_source);assert_eq!(*now,columns);
    let split=o.edit(EditCommand::SplitLine).editor;let g=find_grid(&split.document.analysis.tree);let NodeKind::Grid{array_spec:Some(spec),columns:now,rows,..}=&g.kind else{panic!()};assert_eq!(spec.source,"cc|c");assert_eq!(*now,columns);assert_eq!(rows.len(),3);assert_eq!(g.id,grid_id);assert!(original_ids.iter().all(|id|ids(g).contains(id)));
    o.reject(EditCommand::InsertGridColumn{after:true});o.reject(EditCommand::DeleteGridColumn);
    o.act(EditorAction::Undo);o.act(EditorAction::Undo);assert_eq!(o.session.source(),before.document.source);
}

#[test]
fn copied_array_has_fresh_internal_identities_and_same_column_spec(){
    let mut o=Owned::new(r"\begin{array}{cc|c}a&b&c\\d&e&f\end{array}");let d=o.session.document_snapshot();let g=find_grid(&d.analysis.tree);let old=ids(g);o.act(EditorAction::SetSelection{selection:Selection::Node{node_id:g.id}});let duplicated=o.edit(EditCommand::DuplicateSelection).editor;let NodeKind::Row{entries,..}=&duplicated.document.analysis.tree.kind else{panic!()};assert_eq!(entries.len(),2);let left=find_grid(&entries[0].node);let right=find_grid(&entries[1].node);assert!(ids(right).iter().all(|id|!old.contains(id)));let NodeKind::Grid{array_spec:Some(a),..}=&left.kind else{panic!()};let NodeKind::Grid{array_spec:Some(b),..}=&right.kind else{panic!()};assert_eq!(a.source,b.source);assert_eq!(a.alignments,b.alignments);assert_eq!(a.vertical_rules,b.vertical_rules);o.act(EditorAction::Undo);assert_eq!(o.session.source(),d.source);
}

#[test]
fn invalid_or_unsupported_array_specs_remain_lossless_nonexportable(){
    for source in [
        r"\begin{array}a&b\end{array}",
        r"\begin{array}{cc}a&b&c\end{array}",
        r"\begin{array}{cc@{=}c}a&b&c\end{array}",
        r"\begin{array}{cc|c}a&b&c\\d&e\end{array}",
        r"\begin{array}{cc|c}a&b&c\\[2pt]d&e&f\end{array}",
    ]{
        let s=EngineSession::from_source("bad-array".into(),source.into()).unwrap();let d=s.document_snapshot();assert_eq!(d.source,source);assert!(!d.analysis.exportable, "{source}");assert_eq!(model::latex(&d.analysis.tree,LatexMode::Lossless),source);
    }
}
