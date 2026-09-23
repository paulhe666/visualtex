//! Cell rectangle contract: one identified owner, exact history, source oracle,
//! disjoint copies and atomically refused shape/ownership errors.
use std::{collections::BTreeSet,path::PathBuf};
use serde_json::{json,Value};
use visualtex_engine_core::{editing::{EditorAction,EditorActionResult},
    grid_selection::GridDirection,model::{self,Node,NodeKind,IdAllocator,LatexMode},parser,
    positions::{Position,Selection,TreeIndex,Horizontal},session::{EngineSession,NodeId},
    source::SourceText,source_edit::SourceEdit,transactions::{EditCommand,InputMode},
    fonts::FontRegistry,layout::{LayoutEngine,LayoutOptions},Runtime};
struct Owned {s:EngineSession,n:u32}
impl Owned {
 fn new(source:&str)->Self{Self{s:EngineSession::from_source("rect-A".into(),source.into()).unwrap(),n:0}}
 fn action(&mut self,a:EditorAction)->EditorActionResult{
  let old=self.s.document_snapshot();self.n+=1;
  let r=self.s.apply_action(self.s.revision(),self.n,a).unwrap();
  let idx=TreeIndex::new(&r.editor.document.analysis.tree).unwrap();idx.validate_selection(&r.editor.selection).unwrap();
  if let Some(m)=&r.position_map{for p in TreeIndex::new(&old.analysis.tree).unwrap().carets{
   idx.validate_position(m.apply_position(p,&idx).unwrap()).unwrap();
  }}
  oracle(&r.editor.document.analysis.tree,&r.editor.document.source);r
 }
 fn edit(&mut self,c:EditCommand)->EditorActionResult{self.action(EditorAction::Edit{command:c,group_id:None})}
 fn snap(&self)->Value{serde_json::to_value(self.s.editor_snapshot().unwrap()).unwrap()}
 fn select(&mut self,a:(usize,usize),b:(usize,usize)){
  let t=self.s.document_snapshot().analysis.tree.clone();let g=grid(&t);
  self.action(EditorAction::SetSelection{selection:Selection::GridCells{grid_id:g.id,
   anchor_cell:cell(g,a.0,a.1).id,focus_cell:cell(g,b.0,b.1).id}});
 }
 fn fail(&mut self,a:EditorAction){let before=self.snap();self.n+=1;
  assert!(self.s.apply_action(self.s.revision(),self.n,a).is_err());assert_eq!(self.snap(),before);
 }
}
fn grid(n:&Node)->&Node{if matches!(n.kind,NodeKind::Grid{..}){return n}
 for c in n.children(){let mut has=false;c.visit(&mut |n|has|=matches!(n.kind,NodeKind::Grid{..}));if has{return grid(c)}}panic!("grid")}
fn cell(g:&Node,r:usize,c:usize)->&Node{let NodeKind::Grid{rows,..}=&g.kind else{panic!()};
 let NodeKind::GridRow{cells,..}=&rows[r].kind else{panic!()};&cells[c]}
fn shape(v:Value)->Value{match v{Value::Object(m)=>Value::Object(m.into_iter()
 .filter(|(k,_)|!matches!(k.as_str(),"id"|"origin"|"leading"|"trailing"|"grouping")&&!k.starts_with("before"))
 .map(|(k,v)|{let v=if k=="columns"{json!(v.as_array().unwrap().len())}else{shape(v)};(k,v)}).collect()),
 Value::Array(v)=>Value::Array(v.into_iter().map(shape).collect()),v=>v}}
fn oracle(t:&Node,src:&str){let p=parser::parse(SourceText::new(1,src).unwrap(),NodeId(1),&mut IdAllocator::new(2)).unwrap();
 assert!(p.complete&&p.supported,"{src}: {:?}",p.diagnostics);assert_eq!(model::latex(&p.tree,LatexMode::Lossless),src);
 assert_eq!(shape(serde_json::to_value(t).unwrap()),shape(serde_json::to_value(&p.tree).unwrap()),"{src}");}
fn without_origins(v:Value)->Value {match v {
 Value::Object(m)=>Value::Object(m.into_iter().filter(|(k,_)|k!="origin").map(|(k,v)|(k,without_origins(v))).collect()),
 Value::Array(a)=>Value::Array(a.into_iter().map(without_origins).collect()),v=>v}}
fn contents(v:&Value)->Value{json!({"source":v["document"]["source"],"tree":v["document"]["analysis"]["tree"],"selection":v["selection"]})}
const SRC:&str=r"\begin{matrix}a&b&c\\d&\frac{x}{y}&f\\g&h&i\end{matrix}";
fn ids(n:&Node)->BTreeSet<NodeId>{let mut ids=BTreeSet::new();n.visit(&mut|n|{assert!(ids.insert(n.id));
 if let NodeKind::Grid{columns,..}=&n.kind{for c in columns{assert!(ids.insert(NodeId(c.0)));}}});ids}

#[test]
fn all_reversed_rectangle_corners_clear_only_selected_cells_and_restore_exact_ids(){
 for ar in 0..3{for ac in 0..3{for br in 0..3{for bc in 0..3{
  let mut t=Owned::new(SRC);t.select((ar,ac),(br,bc));let before=t.snap();let old=t.s.document_snapshot();
  let r=t.edit(EditCommand::DeleteForward);let g=grid(&r.editor.document.analysis.tree);
  assert_eq!(g.id,grid(&old.analysis.tree).id);assert_eq!(r.editor.undo_depth,before["undoDepth"].as_u64().unwrap()as usize+1);
  for y in 0..3{for x in 0..3{let a=cell(grid(&old.analysis.tree),y,x);let b=cell(g,y,x);assert_eq!(a.id,b.id);
   if y>=ar.min(br)&&y<=ar.max(br)&&x>=ac.min(bc)&&x<=ac.max(bc){assert!(b.is_empty_row())}
   else{assert_eq!(without_origins(serde_json::to_value(a).unwrap()),without_origins(serde_json::to_value(b).unwrap()));
    assert_eq!(model::latex(a,LatexMode::Lossless),model::latex(b,LatexMode::Lossless));}
  }}
  let after=t.snap();t.action(EditorAction::Undo);assert_eq!(contents(&t.snap()),contents(&before));
  t.action(EditorAction::Redo);assert_eq!(contents(&t.snap()),contents(&after));
 }}}}
}
#[test]
fn rectangle_copy_is_read_only_reopenable_and_contains_no_live_identity(){
 let mut t=Owned::new(SRC);t.select((1,2),(0,1));let before=t.snap();
 let clip=t.s.copy_selection(t.s.revision()).unwrap();assert_eq!(t.snap(),before);
 let src=clip.latex.unwrap();let copied=Owned::new(&src);assert_eq!(grid(&copied.s.document_snapshot().analysis.tree).children().len(),2);
 assert_eq!(src,r"\begin{matrix}b&c\\\frac{x}{y}&f\end{matrix}");
 let v=serde_json::to_value(clip.schema_version).unwrap();assert_eq!(v,1);
}
#[test]
fn rectangular_paste_replaces_contents_preserves_cell_ids_and_clones_descendants(){
 let mut t=Owned::new(SRC);t.select((0,0),(1,1));let src=t.s.copy_selection(t.s.revision()).unwrap().latex.unwrap();
 let original=t.s.document_snapshot();t.select((1,1),(2,2));let before=t.snap();
 let r=t.edit(EditCommand::PasteLatex{latex:src});let g=grid(&r.editor.document.analysis.tree);
 for y in 0..3{for x in 0..3{assert_eq!(cell(g,y,x).id,cell(grid(&original.analysis.tree),y,x).id);}}
 let pasted=ids(cell(g,1,1));let source=ids(cell(g,0,0));assert!(pasted.is_disjoint(&source));
 let saved=t.snap();t.action(EditorAction::Undo);assert_eq!(contents(&t.snap()),contents(&before));
 t.action(EditorAction::Redo);assert_eq!(contents(&t.snap()),contents(&saved));
}
#[test]
fn explicit_cell_paste_from_caret_fits_without_growing_or_nesting_the_grid(){
 let mut t=Owned::new(SRC);let old=t.s.document_snapshot();let g=grid(&old.analysis.tree);
 t.action(EditorAction::SetSelection{selection:Selection::caret(Position::row(cell(g,1,1).id,0))});
 let r=t.edit(EditCommand::PasteGridCells{latex:r"\begin{matrix}p&q\\r&s\end{matrix}".into()});
 assert_eq!(r.editor.document.analysis.facts.grids,1);assert!(matches!(r.editor.selection,Selection::GridCells{..}));
 assert_eq!(model::latex(cell(grid(&r.editor.document.analysis.tree),2,2),LatexMode::Lossless),"s");
}
#[test]
fn invalid_shapes_unknown_source_and_raw_cross_grid_endpoints_are_atomic(){
 let mut t=Owned::new(SRC);t.select((0,0),(1,1));
 for src in ["x",r"\begin{matrix}x&y\end{matrix}",r"\begin{matrix}x",r"\future{x}"]{
  t.fail(EditorAction::Edit{command:EditCommand::PasteGridCells{latex:src.into()},group_id:None});
 }
 t.fail(EditorAction::Edit{command:EditCommand::InsertText{text:"x".into(),mode:InputMode::Math},group_id:None});
 let old=t.s.document_snapshot();let g=grid(&old.analysis.tree);
 for bad in [NodeId(0),old.analysis.tree.id,g.id]{t.fail(EditorAction::SetSelection{selection:Selection::GridCells{
  grid_id:g.id,anchor_cell:cell(g,0,0).id,focus_cell:bad}});}
 t.action(EditorAction::SetSelection{selection:Selection::caret(Position::row(cell(g,2,2).id,0))});
 t.fail(EditorAction::Edit{command:EditCommand::PasteGridCells{latex:r"\begin{matrix}p&q\end{matrix}".into()},group_id:None});
}
#[test]
fn keyboard_rectangle_extension_and_pointer_selection_are_not_content_edits(){
 let mut t=Owned::new(SRC);let old=t.s.document_snapshot();let g=grid(&old.analysis.tree);
 let a=Position::row(cell(g,0,0).id,0);let b=Position::row(cell(g,1,1).id,0);
 t.action(EditorAction::PointerSelection{anchor:a,focus:b});let rect=t.s.editor_snapshot().unwrap();
 assert!(matches!(rect.selection,Selection::GridCells{..}));assert_eq!(rect.undo_depth,0);assert_eq!(rect.document.revision,old.revision);
 t.action(EditorAction::ExtendGridSelection{direction:GridDirection::Right});
 let sel=t.s.editor_snapshot().unwrap().selection;let Selection::GridCells{focus_cell,..}=sel else{panic!()};assert_eq!(focus_cell,cell(g,1,2).id);
 t.action(EditorAction::MoveHorizontal{direction:Horizontal::Left,extend:false});
 assert_eq!(t.s.editor_snapshot().unwrap().selection,Selection::caret(a));
 t.action(EditorAction::ExtendGridSelection{direction:GridDirection::Down});assert_eq!(t.s.editor_snapshot().unwrap().undo_depth,0);
}
#[test]
fn empty_rectangle_clear_is_noop_and_scalar_paste_fills_one_real_cell(){
 let mut t=Owned::new(r"\begin{matrix}&\\&\end{matrix}");t.select((0,0),(1,1));let before=t.snap();
 let r=t.edit(EditCommand::ClearGridCells);assert!(!r.content_changed);assert_eq!(t.snap(),before);
 t.select((1,1),(1,1));let r=t.edit(EditCommand::PasteGridCells{latex:r"\sqrt{z}".into()});
 assert_eq!(r.editor.document.analysis.facts.radicals,1);assert_eq!(r.editor.document.analysis.facts.grids,1);
}
#[test]
fn source_patch_and_undo_keep_or_safely_map_rectangle_identity(){
 let mut t=Owned::new(SRC);t.select((0,0),(1,1));let before=t.snap();let at=t.s.source().find("a&").unwrap()as u32;
 t.s.patch_source(t.s.revision(),SourceEdit{start_utf16:at,end_utf16:at+1,replacement:"z".into()}).unwrap();
 let snap=t.s.editor_snapshot().unwrap();TreeIndex::new(&snap.document.analysis.tree).unwrap().validate_selection(&snap.selection).unwrap();
 oracle(&snap.document.analysis.tree,&snap.document.source);
 t.action(EditorAction::Undo);assert_eq!(contents(&t.snap()),contents(&before));
}
#[test]
fn composition_and_duplicate_event_cannot_clear_or_paste_twice(){
 let mut t=Owned::new(SRC);t.select((0,0),(1,1));
 let rev=t.s.revision();let a=EditorAction::Edit{command:EditCommand::ClearGridCells,group_id:None};let r=t.action(a.clone());
 let replay=t.s.apply_action(rev,t.n,a).unwrap();assert!(replay.replayed);assert_eq!(replay.editor.document.source,r.editor.document.source);
 t.action(EditorAction::BeginComposition{mode:InputMode::Math});
 t.fail(EditorAction::Edit{command:EditCommand::ClearGridCells,group_id:None});
 t.fail(EditorAction::ExtendGridSelection{direction:GridDirection::Down});
}
#[test]
fn nested_cell_endpoints_cannot_impersonate_outer_matrix_cells(){
 let t=Owned::new(r"\begin{matrix}\begin{matrix}a&b\end{matrix}&c\end{matrix}");let snap=t.s.document_snapshot();
 let outer=grid(&snap.analysis.tree);let inner=grid(cell(outer,0,0));let idx=TreeIndex::new(&snap.analysis.tree).unwrap();
 assert!(idx.grid_rectangle(outer.id,cell(inner,0,0).id,cell(outer,0,1).id).is_err());
 assert!(idx.grid_rectangle(inner.id,cell(inner,0,0).id,cell(inner,0,1).id).is_ok());
}
#[test]
fn large_rectangles_have_fresh_source_oracles_and_single_grid_identity(){
 let source=format!("\\begin{{matrix}}{}\\end{{matrix}}",(0..16).map(|_|(0..16).map(|_|"x").collect::<Vec<_>>().join("&")).collect::<Vec<_>>().join("\\\\"));
 let mut t=Owned::new(&source);t.select((1,2),(14,13));let clip=t.s.copy_selection(t.s.revision()).unwrap().latex.unwrap();
 let before=t.snap();t.edit(EditCommand::ClearGridCells);let r=t.edit(EditCommand::PasteGridCells{latex:clip});assert_eq!(r.editor.document.analysis.facts.grids,1);
 t.action(EditorAction::Undo);t.action(EditorAction::Undo);assert_eq!(contents(&t.snap()),contents(&before));
}
#[test]
fn real_scene_has_one_rectangle_per_selected_cell_not_outer_grid(){
 let path=PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf");
 let bytes=std::fs::read(path).unwrap();let mut rt=Runtime::default();
 let send=|rt:&mut Runtime,c:Value|->Value{serde_json::from_str(&rt.dispatch_json(&json!({"protocol":1,"command":c}).to_string())).unwrap()};
 let font_id=rt.register_font(&bytes,0).unwrap().id;
 let created=send(&mut rt,json!({"op":"create","documentId":"geometry","source":SRC}));let h=&created["result"]["handle"];
 let doc=send(&mut rt,json!({"op":"inspectDocument","handle":h}));let g=&doc["result"]["snapshot"]["analysis"]["tree"]["entries"][0]["node"];
 let sel=json!({"kind":"gridCells","gridId":g["id"],"anchorCell":g["rows"][0]["cells"][0]["id"],"focusCell":g["rows"][1]["cells"][1]["id"]});
 let lay=send(&mut rt,json!({"op":"layout","handle":h,"expectedRevision":0,"options":{"mathFont":font_id,"fontSizePx":16,"style":"display","paddingEm":0.125},"includeSvg":true}));
 assert_eq!(lay["status"],"ok","{lay}");
 let q=send(&mut rt,json!({"op":"queryScene","handle":h,"expectedRevision":0,"sceneToken":lay["result"]["sceneToken"],"query":{"kind":"selection","selection":sel}}));
 assert_eq!(q["status"],"ok","{q}");assert_eq!(q["result"]["result"]["rectangles"].as_array().unwrap().len(),4);assert!(q["result"]["result"]["caret"].is_null());
 // Also validate the native scene independently of serialized host rectangles.
 let mut fonts=FontRegistry::default();let id=fonts.register(&bytes,0).unwrap().id;let s=EngineSession::from_source("native-scene".into(),SRC.into()).unwrap();
 let scene=LayoutEngine::default().layout(&s,&fonts,&LayoutOptions::new(id)).unwrap();scene.validate().unwrap();
}
