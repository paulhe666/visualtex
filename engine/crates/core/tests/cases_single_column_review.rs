use std::path::PathBuf;
use visualtex_engine_core::{editing::{EditorAction,EditorActionResult},fonts::FontRegistry,layout::{LayoutEngine,LayoutOptions},model::{self,LatexMode,Node,NodeKind},positions::{Position,Selection},session::EngineSession,transactions::{EditCommand,InputMode}};

fn grid(node:&Node)->&Node{if matches!(node.kind,NodeKind::Grid{..}){return node;}node.children().into_iter().find_map(|child|{if matches!(child.kind,NodeKind::Grid{..}){Some(child)}else{None}}).expect("grid")}
fn font()->(FontRegistry,LayoutOptions){let path=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));let mut fonts=FontRegistry::default();let id=fonts.register(&std::fs::read(path).unwrap(),0).unwrap().id;(fonts,LayoutOptions::new(id))}

#[test]
fn one_column_cases_roundtrip_without_inventing_condition_cells(){
 for source in [r"\begin{cases}a\\b\\c\end{cases}",r"\begin % open
 {cases}a% one
\\b\\c\end % end
 {cases}"]{
  let s=EngineSession::from_source("one-column-cases".into(),source.into()).unwrap();let d=s.document_snapshot();assert!(d.analysis.complete&&d.analysis.supported&&d.analysis.exportable,"{:?}",d.analysis.diagnostics);assert_eq!(model::latex(&d.analysis.tree,LatexMode::Lossless),source);let g=grid(&d.analysis.tree);let NodeKind::Grid{environment,columns,rows,..}=&g.kind else{panic!()};assert_eq!(environment.name(),"cases");assert_eq!(columns.len(),1);assert_eq!(rows.len(),3);for row in rows{let NodeKind::GridRow{cells,separators,..}=&row.kind else{panic!()};assert_eq!(cells.len(),1);assert!(separators.is_empty());}
 }
}

#[test]
fn one_column_cases_use_real_brace_and_one_shared_column_anchor(){
 let s=EngineSession::from_source("one-column-cases-layout".into(),r"\begin{cases}\frac{x}{y}\\\sqrt z\\q\end{cases}".into()).unwrap();let d=s.document_snapshot();let g=grid(&d.analysis.tree);let (fonts,options)=font();let mut engine=LayoutEngine::default();let scene=engine.layout(&s,&fonts,&options).unwrap();assert_eq!(scene.grid_columns.len(),1);assert_eq!(scene.grid_columns[0].grid_id,g.id);assert!(scene.grid_columns[0].x>scene.regions.iter().find(|r|r.node_id==g.id).unwrap().logical.x_min);assert!(scene.glyphs.iter().any(|glyph|glyph.node_id==g.id),"cases brace must be real font glyphs");
}

struct Owned{session:EngineSession,sequence:u32}
impl Owned{fn new(source:&str)->Self{Self{session:EngineSession::from_source("one-column-edit".into(),source.into()).unwrap(),sequence:0}}fn act(&mut self,action:EditorAction)->EditorActionResult{self.sequence+=1;self.session.apply_action(self.session.revision(),self.sequence,action).unwrap()}fn edit(&mut self,command:EditCommand)->EditorActionResult{self.act(EditorAction::Edit{command,group_id:None})}}

#[test]
fn shift_enter_and_row_crud_preserve_single_column_cases_schema(){
 let mut o=Owned::new(r"\begin{cases}a\\b\end{cases}");let before=o.session.editor_snapshot().unwrap();let g=grid(&before.document.analysis.tree);let NodeKind::Grid{columns,rows,..}=&g.kind else{panic!()};let grid_id=g.id;let column_id=columns[0];let NodeKind::GridRow{cells,..}=&rows[0].kind else{panic!()};let first_cell=cells[0].id;
 o.act(EditorAction::SetSelection{selection:Selection::caret(Position::row(first_cell,1))});let split=o.edit(EditCommand::SplitLine).editor;let g=grid(&split.document.analysis.tree);let NodeKind::Grid{columns,rows,..}=&g.kind else{panic!()};assert_eq!(g.id,grid_id);assert_eq!(columns, &[column_id]);assert_eq!(rows.len(),3);assert_eq!(split.selection.focus().unwrap().node_id(),match &rows[1].kind{NodeKind::GridRow{cells,..}=>cells[0].id,_=>panic!()});
 let changed=o.edit(EditCommand::InsertText{text:"q".into(),mode:InputMode::Math}).editor;assert!(changed.document.source.contains("q"));o.act(EditorAction::Undo);o.act(EditorAction::Undo);assert_eq!(o.session.source(),before.document.source);let restored_snapshot=o.session.document_snapshot();let restored=grid(&restored_snapshot.analysis.tree);let NodeKind::Grid{columns,..}=&restored.kind else{panic!()};assert_eq!(restored.id,grid_id);assert_eq!(columns,&[column_id]);
}

#[test]
fn one_and_two_column_cases_are_both_valid_but_ragged_mixing_is_not(){
 for source in [r"\begin{cases}a\\b\end{cases}",r"\begin{cases}a&x>0\\b&x<0\end{cases}"]{let s=EngineSession::from_source("valid-cases".into(),source.into()).unwrap();let d=s.document_snapshot();assert!(d.analysis.complete&&d.analysis.supported&&d.analysis.exportable);}
 let source=r"\begin{cases}a\\b&x<0\end{cases}";let s=EngineSession::from_source("ragged-cases".into(),source.into()).unwrap();let d=s.document_snapshot();assert_eq!(d.source,source);assert!(!d.analysis.supported&&!d.analysis.exportable);
}
