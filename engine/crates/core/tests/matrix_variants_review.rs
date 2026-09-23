//! P09 matrix environment aliases are one editable Grid with font-driven
//! presentation fences, not source preprocessing or nested editor owners.
use std::{collections::BTreeSet,path::PathBuf};
use visualtex_engine_core::{editing::{EditorAction,EditorActionResult},fonts::FontRegistry,
 grids::GridKind,layout::{LayoutEngine,LayoutOptions},model::{self,LatexMode,Node,NodeKind},
 positions::{Position,Selection},session::EngineSession,transactions::{EditCommand,InputMode}};

struct Owned{s:EngineSession,seq:u32}
impl Owned{
 fn new(source:&str)->Self{Self{s:EngineSession::from_source("matrix-variant".into(),source.into()).unwrap(),seq:0}}
 fn action(&mut self,action:EditorAction)->EditorActionResult{self.seq+=1;self.s.apply_action(self.s.revision(),self.seq,action).unwrap()}
 fn select(&mut self,p:Position){self.action(EditorAction::SetSelection{selection:Selection::caret(p)});}
 fn edit(&mut self,c:EditCommand)->EditorActionResult{self.action(EditorAction::Edit{command:c,group_id:None})}
}
fn grid(root:&Node)->&Node{if matches!(root.kind,NodeKind::Grid{..}){return root;}root.children().into_iter().find_map(|child|find_grid(child)).unwrap()}
fn find_grid(root:&Node)->Option<&Node>{if matches!(root.kind,NodeKind::Grid{..}){return Some(root);}root.children().into_iter().find_map(find_grid)}
fn cells(g:&Node)->Vec<Vec<u32>>{let NodeKind::Grid{rows,..}=&g.kind else{panic!()};rows.iter().map(|r|{let NodeKind::GridRow{cells,..}=&r.kind else{panic!()};cells.iter().map(|c|c.id.0).collect()}).collect()}
fn all_ids(n:&Node)->BTreeSet<u32>{let mut ids=BTreeSet::new();n.visit(&mut|n|{assert!(ids.insert(n.id.0));if let NodeKind::Grid{columns,..}=&n.kind{for c in columns{assert!(ids.insert(c.0));}}});ids}
fn font()->(FontRegistry,LayoutOptions){let path=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));let mut fonts=FontRegistry::default();let id=fonts.register(&std::fs::read(path).unwrap(),0).unwrap().id;(fonts,LayoutOptions::new(id))}
const VARIANTS:&[(&str,GridKind)]=&[("pmatrix",GridKind::PMatrix),("bmatrix",GridKind::BMatrix),("Bmatrix",GridKind::BraceMatrix),("vmatrix",GridKind::VMatrix),("Vmatrix",GridKind::DoubleVMatrix)];

#[test]
fn five_matrix_environment_spellings_roundtrip_with_one_grid_identity(){
 for (name,kind) in VARIANTS {let source=format!("\\begin{{{name}}}a&b\\\\c&d\\end{{{name}}}");let owned=Owned::new(&source);let d=owned.s.document_snapshot();assert!(d.analysis.complete&&d.analysis.supported&&d.analysis.exportable,"{name}: {:?}",d.analysis.diagnostics);assert_eq!(model::latex(&d.analysis.tree,LatexMode::Lossless),source);let g=grid(&d.analysis.tree);let NodeKind::Grid{environment,columns,rows,..}=&g.kind else{panic!()};assert_eq!(environment,kind);assert_eq!(environment.name(),*name);assert!(environment.is_matrix());assert_eq!(columns.len(),2);assert_eq!(rows.len(),2);assert_eq!(d.analysis.facts.grids,1);}
}

#[test]
fn wrappers_are_real_font_geometry_and_shift_shared_column_anchors(){
 let (fonts,options)=font();let bare=Owned::new(r"\begin{matrix}a&b\\c&d\end{matrix}");let mut layout=LayoutEngine::default();let bare_scene=layout.layout(&bare.s,&fonts,&options).unwrap();let bare_anchor=bare_scene.grid_columns[0].x;
 for (name,_) in VARIANTS {let source=format!("\\begin{{{name}}}a&b\\\\c&d\\end{{{name}}}");let owned=Owned::new(&source);let d=owned.s.document_snapshot();let g=grid(&d.analysis.tree);let scene=layout.layout(&owned.s,&fonts,&options).unwrap();assert!(scene.width>bare_scene.width,"{name}");assert!(scene.grid_columns[0].x>bare_anchor,"{name}: left presentation delimiter did not shift content");let owned_glyphs=scene.glyphs.iter().filter(|glyph|glyph.node_id==g.id).count();assert!(owned_glyphs>=2,"{name}: wrapper must use real grid-owned glyph geometry");let svg=scene.svg().unwrap();assert!(!svg.contains("<text")&&!svg.contains("foreignObject")&&!svg.contains("<image"));}
}

#[test]
fn row_column_edits_and_shift_enter_preserve_variant_and_survivor_ids(){
 for (name,kind) in VARIANTS {let source=format!("\\begin{{{name}}}a&b\\\\c&d\\end{{{name}}}");let mut owned=Owned::new(&source);let before=owned.s.document_snapshot();let g=grid(&before.analysis.tree);let gid=g.id;let before_cells=cells(g);let NodeKind::Grid{columns,..}=&g.kind else{panic!()};let old_columns=columns.clone();let target=before_cells[0][1];owned.select(Position::row(visualtex_engine_core::session::NodeId(target),1));let row_added=owned.edit(EditCommand::SplitLine).editor;let g=grid(&row_added.document.analysis.tree);let NodeKind::Grid{environment,columns,rows,..}=&g.kind else{panic!()};assert_eq!(*environment,*kind);assert_eq!(g.id,gid);assert_eq!(columns,&old_columns);assert_eq!(rows.len(),3);assert!(row_added.document.source.starts_with(&format!("\\begin{{{name}}}")));
  let new_row_cell=if let NodeKind::GridRow{cells,..}=&rows[1].kind{cells[1].id}else{panic!()};assert_eq!(row_added.selection.focus().unwrap().node_id(),new_row_cell);owned.edit(EditCommand::InsertText{text:"q".into(),mode:InputMode::Math});
  owned.edit(EditCommand::InsertGridColumn{after:true});let col=owned.s.document_snapshot();let g=grid(&col.analysis.tree);let NodeKind::Grid{environment,columns,..}=&g.kind else{panic!()};assert_eq!(*environment,*kind);assert_eq!(columns.len(),3);owned.edit(EditCommand::DeleteGridColumn);let back=owned.s.document_snapshot();let g=grid(&back.analysis.tree);let NodeKind::Grid{environment,columns,..}=&g.kind else{panic!()};assert_eq!(*environment,*kind);assert_eq!(columns.len(),2);assert_eq!(g.id,gid);}
}

#[test]
fn whole_variant_copy_is_independent_and_undo_restores_original(){
 let mut owned=Owned::new(r"\begin{bmatrix}a&b\\c&d\end{bmatrix}");let before=owned.s.document_snapshot();let root=&before.analysis.tree;let NodeKind::Row{entries,..}=&root.kind else{panic!()};let original=all_ids(&entries[0].node);owned.action(EditorAction::SetSelection{selection:Selection::Range{anchor:Position::row(root.id,0),focus:Position::row(root.id,1)}});let dup=owned.edit(EditCommand::DuplicateSelection).editor;let NodeKind::Row{entries,..}=&dup.document.analysis.tree.kind else{panic!()};assert_eq!(entries.len(),2);assert!(original.is_disjoint(&all_ids(&entries[1].node)));assert_eq!(dup.document.source,r"\begin{bmatrix}a&b\\c&d\end{bmatrix}\begin{bmatrix}a&b\\c&d\end{bmatrix}");let undo=owned.action(EditorAction::Undo).editor;assert_eq!(undo.document.source,before.source);let NodeKind::Row{entries,..}=&undo.document.analysis.tree.kind else{panic!()};assert_eq!(all_ids(&entries[0].node),original);
}

#[test]
fn ragged_and_row_gap_variant_source_remains_preserved_and_nonexportable(){
 for source in [r"\begin{bmatrix}a&b\\c\end{bmatrix}",r"\begin{pmatrix}a&b\\[2pt]c&d\end{pmatrix}",r"\begin{Vmatrix}a&b"] {let owned=Owned::new(source);let d=owned.s.document_snapshot();assert_eq!(model::latex(&d.analysis.tree,LatexMode::Lossless),source);assert!(!d.analysis.exportable,"{source}");}
}
