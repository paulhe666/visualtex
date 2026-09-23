//! P08 aligned/align: shared anchors, contained structural split/merge and history.
use std::{collections::BTreeSet,path::PathBuf};
use serde_json::{json,Value};
use visualtex_engine_core::{editing::{EditorAction,EditorActionResult},fonts::FontRegistry,
    layout::{LayoutEngine,LayoutOptions,MathStyle},model::{self,LatexMode,Node,NodeKind},
    positions::{Position,Selection,Horizontal,TreeIndex},session::EngineSession,
    transactions::{EditCommand,InputMode}};
struct Owned{s:EngineSession,seq:u32}
impl Owned{
 fn new(text:&str)->Self{Self{s:EngineSession::from_source("aligned-review".into(),text.into()).unwrap(),seq:0}}
 fn snapshot(&self)->Value{serde_json::to_value(self.s.editor_snapshot().unwrap()).unwrap()}
 fn action(&mut self,a:EditorAction)->EditorActionResult{
  self.seq+=1;let old=self.s.document_snapshot();let result=self.s.apply_action(self.s.revision(),self.seq,a).unwrap();
  let d=&result.editor.document;let idx=TreeIndex::new(&d.analysis.tree).unwrap();idx.validate_selection(&result.editor.selection).unwrap();
  if let Some(map)=&result.position_map{for p in TreeIndex::new(&old.analysis.tree).unwrap().carets{idx.validate_position(map.apply_position(p,&idx).unwrap()).unwrap();}}
  let fresh=Owned::new(&d.source).s.document_snapshot();assert_eq!(semantic(serde_json::to_value(&d.analysis.tree).unwrap()),semantic(serde_json::to_value(&fresh.analysis.tree).unwrap()),"{}",d.source);
  result
 }
 fn select(&mut self,p:Position){self.action(EditorAction::SetSelection{selection:Selection::caret(p)});}
 fn edit(&mut self,c:EditCommand)->EditorActionResult{self.action(EditorAction::Edit{command:c,group_id:None})}
 fn reject(&mut self,c:EditCommand){let before=self.snapshot();self.seq+=1;assert!(self.s.apply_action(self.s.revision(),self.seq,EditorAction::Edit{command:c,group_id:None}).is_err());assert_eq!(self.snapshot(),before);}
}
fn semantic(v:Value)->Value{match v{Value::Object(m)=>Value::Object(m.into_iter().filter(|(k,_)|!matches!(k.as_str(),"id"|"origin"|"leading"|"trailing"|"grouping")&&!k.starts_with("before"))
 .map(|(k,v)|{let v=if k=="columns"{json!(v.as_array().unwrap().len())}else{semantic(v)};(k,v)}).collect()),Value::Array(v)=>Value::Array(v.into_iter().map(semantic).collect()),v=>v}}
fn content(s:&Value)->Value{json!({"source":s["document"]["source"],"tree":s["document"]["analysis"]["tree"],"selection":s["selection"],"owner":s["document"]["documentId"]})}
fn grid(n:&Node)->&Node{if matches!(n.kind,NodeKind::Grid{..}){return n;}n.children().into_iter().find_map(|c|{let mut found=false;c.visit(&mut|n|found|=matches!(n.kind,NodeKind::Grid{..}));found.then(||grid(c))}).unwrap()}
fn cell(g:&Node,r:usize,c:usize)->&Node{let NodeKind::Grid{rows,..}=&g.kind else{panic!()};let NodeKind::GridRow{cells,..}=&rows[r].kind else{panic!()};&cells[c]}
fn entries(n:&Node)->&Vec<model::RowEntry>{let NodeKind::Row{entries,..}=&n.kind else{panic!()};entries}
fn ids(n:&Node)->BTreeSet<u32>{let mut all=BTreeSet::new();n.visit(&mut|n|{assert!(all.insert(n.id.0));if let NodeKind::Grid{columns,..}=&n.kind{for c in columns{assert!(all.insert(c.0));}}});all}
fn font()->(FontRegistry,LayoutOptions){let path=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));let mut f=FontRegistry::default();let id=f.register(&std::fs::read(path).unwrap(),0).unwrap().id;(f,LayoutOptions::new(id))}
const SOURCE:&str=r"\begin{aligned}ab&=cd\\ef&=gh\end{aligned}";

#[test]
fn aligned_spellings_multiple_pairs_and_empty_ampersand_cells_roundtrip(){
 for source in [SOURCE,r"\begin{align}a&=b\\c&=d\end{align}",r"\begin{align*}a&=b\\c&=d\end{align*}",
  r"\begin{aligned}a&=b&c&=d\\x&=y&z&=w\end{aligned}",r"\begin{aligned}a&=b&&c&=d\\x&=y&&z&=w\end{aligned}",
  r"\begin{aligned}\frac{x}{y}&=\sqrt{z}\\\sum_i^n x_i&=\int_0^1 x\,\mathrm{d}x\end{aligned}",
  "\\begin %header\n {aligned}a% & \\\\\n&=b\\\\c&=d\\end %end\n {aligned}"]{
  let t=Owned::new(source);let d=t.s.document_snapshot();assert!(d.analysis.complete&&d.analysis.supported,"{:?}",d.analysis.diagnostics);
  assert_eq!(model::latex(&d.analysis.tree,LatexMode::Lossless),source);
  let normalized=model::latex(&d.analysis.tree,LatexMode::Canonical);let reopened=Owned::new(&normalized).s.document_snapshot();
  assert_eq!(semantic(serde_json::to_value(&d.analysis.tree).unwrap()),semantic(serde_json::to_value(&reopened.analysis.tree).unwrap()));
  assert_eq!(ids(&d.analysis.tree).len(),d.analysis.facts.nodes+d.analysis.facts.grid_columns);
 }
}
#[test]
fn aligned_structure_insertion_preserves_nodes_and_one_undo(){
 let mut t=Owned::new("xy");let d=t.s.document_snapshot();t.action(EditorAction::SetSelection{selection:Selection::Range{anchor:Position::row(d.analysis.tree.id,0),focus:Position::row(d.analysis.tree.id,2)}});
 let before=t.snapshot();let original:Vec<_>=entries(&d.analysis.tree).iter().map(|e|e.node.id).collect();let added=t.edit(EditCommand::InsertAligned{rows:2,pairs:2}).editor;
 let g=grid(&added.document.analysis.tree);assert_eq!(entries(cell(g,0,0)).iter().map(|e|e.node.id).collect::<Vec<_>>(),original);
 let NodeKind::Grid{columns,..}=&g.kind else{panic!()};assert_eq!(columns.len(),4);let after=t.snapshot();
 t.action(EditorAction::Undo);assert_eq!(content(&t.snapshot()),content(&before));t.action(EditorAction::Redo);assert_eq!(content(&t.snapshot()),content(&after));
}
#[test]
fn every_cell_boundary_split_preserves_math_ids_and_merges_by_columns(){
 for c in 0..2{for at in 0..=if c==0{2}else{3}{for direction in [Horizontal::Left,Horizontal::Right]{
  let mut t=Owned::new(SOURCE);let old=t.s.document_snapshot();let g=grid(&old.analysis.tree);let old_ids=ids(g);
  t.select(Position::row(cell(g,0,c).id,at));let before=t.snapshot();let split=t.edit(EditCommand::SplitLine).editor;
  let next=grid(&split.document.analysis.tree);let NodeKind::Grid{rows,columns,..}=&next.kind else{panic!()};assert_eq!(rows.len(),3);assert_eq!(columns.len(),2);assert!(old_ids.is_subset(&ids(next)));
  let after=t.snapshot();t.action(EditorAction::Undo);assert_eq!(content(&t.snapshot()),content(&before));t.action(EditorAction::Redo);assert_eq!(content(&t.snapshot()),content(&after));
  let now=t.s.document_snapshot();let g=grid(&now.analysis.tree);
  t.select(if direction==Horizontal::Left{Position::row(cell(g,1,0).id,0)}else{Position::row(cell(g,0,1).id,entries(cell(g,0,1)).len())});
  t.edit(EditCommand::MergeLine{direction});assert_eq!(t.s.source(),SOURCE);assert_eq!(ids(grid(&t.s.document_snapshot().analysis.tree)),old_ids);
 }}}
}
#[test]
fn end_of_aligned_row_places_caret_in_first_cell_of_new_row(){
 let mut t=Owned::new(SOURCE);let d=t.s.document_snapshot();let g=grid(&d.analysis.tree);t.select(Position::row(cell(g,1,1).id,3));
 let added=t.edit(EditCommand::SplitLine).editor;let g=grid(&added.document.analysis.tree);assert_eq!(added.selection,Selection::caret(Position::row(cell(g,2,0).id,0)));
 t.edit(EditCommand::InsertText{text:"i".into(),mode:InputMode::Math});t.edit(EditCommand::MoveGridCell{direction:Horizontal::Right});t.edit(EditCommand::InsertText{text:"=j".into(),mode:InputMode::Math});
 assert_eq!(t.s.source(),r"\begin{aligned}ab&=cd\\ef&=gh\\i&=j\end{aligned}");
}
#[test]
fn aligned_interior_split_preserves_suffix_columns_and_untouched_rows(){
 let mut t=Owned::new(r"\begin{aligned}ab&=cd&ef&=gh\\i&=j&k&=l\end{aligned}");let old=t.s.document_snapshot();let g=grid(&old.analysis.tree);
 t.select(Position::row(cell(g,0,1).id,2));let d=t.edit(EditCommand::SplitLine).editor.document;let n=grid(&d.analysis.tree);
 assert_eq!(model::latex(cell(n,0,0),LatexMode::Lossless),"ab");assert_eq!(model::latex(cell(n,0,1),LatexMode::Lossless),"=c");
 assert_eq!(model::latex(cell(n,1,1),LatexMode::Lossless),"d");assert_eq!(model::latex(cell(n,1,2),LatexMode::Lossless),"ef");
 for c in 0..4{assert_eq!(cell(n,2,c).id,cell(g,1,c).id);assert_eq!(cell(n,0,c).id,cell(g,0,c).id);}
}
#[test]
fn nearest_aligned_with_fraction_interior_inserts_row_without_cutting_fraction(){
 let mut t=Owned::new(r"\begin{cases}\begin{aligned}\frac{xy}{z}&=a\\b&=c\end{aligned}&q\\d&e\end{cases}");
 let old=t.s.document_snapshot();let outer=grid(&old.analysis.tree);let inner=grid(cell(outer,0,0));let frac=&entries(cell(inner,0,0))[0].node;
 let NodeKind::Fraction{numerator,..}=&frac.kind else{panic!()};t.select(Position::row(numerator.id,1));
 let d=t.edit(EditCommand::SplitLine).editor.document;let next_outer=grid(&d.analysis.tree);let next_inner=grid(cell(next_outer,0,0));
 assert_eq!(next_outer.id,outer.id);assert_eq!(next_inner.id,inner.id);assert_eq!(d.analysis.facts.grid_rows,5);
 assert_eq!(model::latex(next_inner.find(frac.id).unwrap(),LatexMode::Lossless),r"\frac{xy}{z}");
}
#[test]
fn aligned_boundaries_never_merge_another_grid_or_formula(){
 let mut t=Owned::new(SOURCE);let d=t.s.document_snapshot();let g=grid(&d.analysis.tree);
 for (r,c,o,cmd) in [(0,0,0,EditCommand::DeleteBackward),(1,1,3,EditCommand::DeleteForward),(0,1,0,EditCommand::DeleteBackward)]{
  t.select(Position::row(cell(g,r,c).id,o));let before=t.snapshot();t.edit(cmd);assert_eq!(t.snapshot(),before);
 }
 t.action(EditorAction::SetSelection{selection:Selection::Range{anchor:Position::row(cell(g,0,0).id,0),focus:Position::row(cell(g,1,1).id,2)}});
 t.reject(EditCommand::DeleteBackward);assert!(t.s.copy_selection(t.s.revision()).is_err());
}
#[test]
fn shared_alignment_anchors_and_rhs_relation_glue_are_real_layout(){
 let t=Owned::new(r"\begin{aligned}long&=b&x&=y\\a&=cd&uvw&=z\end{aligned}");let d=t.s.document_snapshot();let g=grid(&d.analysis.tree);let(f,mut options)=font();let mut layout=LayoutEngine::default();
 for style in [MathStyle::Display,MathStyle::Text,MathStyle::Script,MathStyle::ScriptScript]{options.style=style;options.font_size_px=16.0;
  let s=layout.layout(&t.s,&f,&options).unwrap();s.validate().unwrap();assert_eq!(s.grid_columns.len(),4);
  for c in [1,3]{let a=cell(g,0,c);let b=cell(g,1,c);let ra=s.regions.iter().find(|r|r.node_id==entries(a)[0].node.id).unwrap();let rb=s.regions.iter().find(|r|r.node_id==entries(b)[0].node.id).unwrap();
   assert!((ra.logical.x_min-rb.logical.x_min).abs()<1e-9,"Equal signs do not share one anchor");
   let x=s.caret(Position::row(a.id,0)).unwrap().x;assert!((ra.logical.x_min-x-5.0/18.0).abs()<1e-9,"RHS relation lost its mathematical glue");
   assert!((s.grid_columns[c-1].x+s.grid_columns[c-1].width-s.grid_columns[c].x).abs()<1e-9);
  }
  assert!((s.grid_columns[2].x-(s.grid_columns[1].x+s.grid_columns[1].width)-2.0).abs()<1e-9);
  let cached=layout.layout(&t.s,&f,&options).unwrap();assert_eq!(cached.rebuilt_fragments,0);assert_eq!(cached.svg().unwrap(),s.svg().unwrap());
  options.font_size_px=32.0;let twice=layout.layout(&t.s,&f,&options).unwrap();assert_eq!(twice.width,s.width*2.0);assert_eq!(twice.height,s.height*2.0);assert_eq!(serde_json::to_value(twice.grid_columns).unwrap(),serde_json::to_value(s.grid_columns).unwrap());
 }
}
#[test]
fn rhs_binary_context_does_not_change_unrelated_matrix_or_inner_row(){
 let(f,options)=font();let mut layout=LayoutEngine::default();
 for (source,wanted) in [(r"\begin{aligned}a&+b\end{aligned}",4.0/18.0),(r"\begin{matrix}a&+b\end{matrix}",0.0)]{
  let t=Owned::new(source);let d=t.s.document_snapshot();let c=cell(grid(&d.analysis.tree),0,1);let s=layout.layout(&t.s,&f,&options).unwrap();
  let plus=s.regions.iter().find(|r|r.node_id==entries(c)[0].node.id).unwrap();assert!((plus.logical.x_min-s.caret(Position::row(c.id,0)).unwrap().x-wanted).abs()<1e-9);
 }
}
#[test]
fn color_and_comment_splits_preserve_reopenability_and_exact_history(){
 for source in [r"\begin{aligned}\color{red}ab&=c\\d&=e\end{aligned}","\\begin{aligned}a% keep\n b&=c\\\\d&=e\\end{aligned}"]{
  let mut t=Owned::new(source);let d=t.s.document_snapshot();let g=grid(&d.analysis.tree);t.select(Position::row(cell(g,0,0).id,1));let before=t.snapshot();
  let d=t.edit(EditCommand::SplitLine).editor.document;let(f,options)=font();LayoutEngine::default().layout(&t.s,&f,&options).unwrap().validate().unwrap();assert_eq!(d.analysis.facts.grids,1);
  t.action(EditorAction::Undo);assert_eq!(content(&t.snapshot()),content(&before));
 }
}
#[test]
fn aligned_rectangle_clipboard_and_invalid_options_are_atomic(){
 let mut t=Owned::new(SOURCE);let d=t.s.document_snapshot();let g=grid(&d.analysis.tree);t.action(EditorAction::SetSelection{selection:Selection::GridCells{grid_id:g.id,anchor_cell:cell(g,0,0).id,focus_cell:cell(g,1,1).id}});
 let copy=t.s.copy_selection(t.s.revision()).unwrap();t.edit(EditCommand::ClearGridCells);t.edit(EditCommand::PasteGridCells{latex:copy.latex.unwrap()});assert_eq!(t.s.source(),SOURCE);
 t.reject(EditCommand::InsertAligned{rows:0,pairs:1});t.reject(EditCommand::InsertAligned{rows:2,pairs:u32::MAX});
 for source in [r"\begin{aligned}[t]a&=b\end{aligned}",r"\begin{aligned}a&=b\\c\end{aligned}",r"\begin{aligned}a&=b\\[2pt]c&=d\end{aligned}"]{
  let t=Owned::new(source);assert_eq!(t.s.source(),source);assert!(!t.s.document_snapshot().analysis.exportable);
 }
}
