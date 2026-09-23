use std::{collections::BTreeSet,path::PathBuf};
use serde_json::Value;
use visualtex_engine_core::{model::{Node,NodeKind,LatexMode,IdAllocator},parser::{self,ParseResult},source::SourceText,
 session::{EngineSession,NodeId},editing::{EditorAction,EditorActionResult},positions::{Position,Selection,TreeIndex},
 transactions::{EditCommand,InputMode,StructureKind},fonts::FontRegistry,layout::{LayoutEngine,LayoutOptions},source_edit::SourceEdit};
fn parse(s:&str)->ParseResult{parser::parse(SourceText::new(1,s).unwrap(),NodeId(1),&mut IdAllocator::new(2)).unwrap()}
fn semantic(n:&Node)->Value{
 fn trim(v:&mut Value){match v{Value::Object(m)=>{m.retain(|k,_|!matches!(k.as_str(),"id"|"origin"|"leading"|"trailing"|"grouping")&&!k.starts_with("before"));for v in m.values_mut(){trim(v);}},Value::Array(a)=>for v in a{trim(v);},_=>{}}}
 let mut v=serde_json::to_value(n).unwrap();trim(&mut v);v
}
fn ids(n:&Node)->BTreeSet<NodeId>{let mut s=BTreeSet::new();n.visit(&mut|n|{assert!(s.insert(n.id));});s}
struct Owned{s:EngineSession,seq:u32}
impl Owned{
 fn new(source:&str)->Self{Self{s:EngineSession::from_source("fence-test".into(),source.into()).unwrap(),seq:0}}
 fn act(&mut self,a:EditorAction)->EditorActionResult{self.seq+=1;let before=self.s.document_snapshot();let r=self.s.apply_action(self.s.revision(),self.seq,a).unwrap();
  let index=TreeIndex::new(&r.editor.document.analysis.tree).unwrap();index.validate_selection(&r.editor.selection).unwrap();
  if let Some(map)=&r.position_map{for p in TreeIndex::new(&before.analysis.tree).unwrap().carets{map.apply_position(p,&index).unwrap();}}
  let reopened=parse(self.s.source());assert!(reopened.complete&&reopened.supported,"{}: {:?}",self.s.source(),reopened.diagnostics);
  assert_eq!(semantic(&r.editor.document.analysis.tree),semantic(&reopened.tree),"{}",self.s.source());r
 }
 fn edit(&mut self,c:EditCommand)->EditorActionResult{self.act(EditorAction::Edit{command:c,group_id:None})}
 fn text(&mut self,text:&str){self.edit(EditCommand::InsertText{text:text.into(),mode:InputMode::Math});}
 fn bracket(&mut self,text:&str)->EditorActionResult{self.edit(EditCommand::TypeDelimiter{text:text.into()})}
 fn select(&mut self,s:Selection){self.act(EditorAction::SetSelection{selection:s});}
}
fn first_fence(n:&Node)->&Node{if matches!(n.kind,NodeKind::Fence{..}){return n;}n.children().into_iter().find_map(|c|{let mut yes=false;c.visit(&mut|n|yes|=matches!(n.kind,NodeKind::Fence{..}));yes.then(||first_fence(c))}).unwrap()}
fn font()->(FontRegistry,LayoutOptions){let path=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));let mut f=FontRegistry::default();let id=f.register(&std::fs::read(path).unwrap(),0).unwrap().id;(f,LayoutOptions::new(id))}
#[test]
fn pairs_nested_single_sided_and_hidden_delimiters_roundtrip(){
 for s in [r"\left(x\right)",r"\left[\sqrt{x}+\left(\frac{a}{b}\right)\right]",r"\left. x \right)",r"\left\{x\right.",r"\left. x\right.",r"\left\langle x\right\rangle",r"\left(x\right]","\\left %left\n ( x %tail\n \\right %right\n )"]{
  let p=parse(s);assert!(p.complete&&p.supported,"{s}: {:?}",p.diagnostics);assert_eq!(visualtex_engine_core::model::latex(&p.tree,LatexMode::Lossless),s);
  let canonical=visualtex_engine_core::model::latex(&p.tree,LatexMode::Canonical);let q=parse(&canonical);assert!(q.complete&&q.supported,"{canonical}");
  assert_eq!(semantic(&p.tree),semantic(&q.tree));assert_eq!(visualtex_engine_core::model::latex(&q.tree,LatexMode::Canonical),canonical);
  assert_eq!(ids(&p.tree).len(),p.facts.nodes);
 }
}
#[test]
fn missing_unknown_and_unmatched_delimiters_preserve_the_entire_draft(){
 for s in [r"\left",r"\left(x",r"\left(x\right",r"\right)",r"{\left(x}",r"\left{ x \right}"]{
  let p=parse(s);assert!(!p.complete&&!p.exportable,"{s}");assert_eq!(visualtex_engine_core::model::latex(&p.tree,LatexMode::Lossless),s);
 }
 let p=parse(r"\left\unknown x\right)");assert!(!p.supported&&!p.exportable);
 let p=parse(r"\left(\right)");assert!(p.complete&&p.supported&&!p.exportable);
}
#[test]
fn visual_pairing_and_matching_exit_do_not_insert_a_second_closer(){
 let mut t=Owned::new("");t.bracket("(");t.text("x");let before=t.s.editor_snapshot().unwrap();let r=t.bracket(")");
 assert!(!r.content_changed);assert_eq!(r.editor.document.revision,before.document.revision);assert_eq!(r.editor.undo_depth,before.undo_depth);
 assert_eq!(t.s.source(),r"\left(x\right)");assert_eq!(r.editor.selection,Selection::caret(Position::row(NodeId(1),1)));
 t.bracket(")");assert_eq!(t.s.source(),r"\left(x\right))");
}
#[test]
fn mismatched_inner_closer_is_literal_and_does_not_jump_to_matching_outer_pair(){
 let mut t=Owned::new("");t.bracket("(");t.bracket("[");t.text("x");t.bracket(")");
 assert_eq!(t.s.source(),r"\left(\left[x)\right]\right)");let before=t.s.source().to_owned();t.bracket("]");t.bracket(")");assert_eq!(t.s.source(),before);
 assert_eq!(t.s.editor_snapshot().unwrap().selection,Selection::caret(Position::row(NodeId(1),1)));
}
#[test]
fn matching_closer_exits_a_completed_nested_script_without_deleting_it(){
 let mut t=Owned::new("");t.bracket("(");t.text("x");t.edit(EditCommand::InsertStructure{structure:StructureKind::Superscript});t.text("2");
 t.bracket(")");assert_eq!(t.s.source(),r"\left(x^{2}\right)");assert_eq!(t.s.editor_snapshot().unwrap().selection,Selection::caret(Position::row(NodeId(1),1)));
}
#[test]
fn wrapping_selection_preserves_nodes_and_duplicate_subtrees_are_independent(){
 let mut t=Owned::new("x+y");let original=ids(&t.s.document_snapshot().analysis.tree);
 t.select(Selection::Range{anchor:Position::row(NodeId(1),0),focus:Position::row(NodeId(1),3)});
 t.edit(EditCommand::InsertFence{left:"[".into(),right:"]".into()});assert!(original.is_subset(&ids(&t.s.document_snapshot().analysis.tree)));
 t.select(Selection::Node{node_id:first_fence(&t.s.document_snapshot().analysis.tree).id});t.edit(EditCommand::DuplicateSelection);
 let p=t.s.document_snapshot();let NodeKind::Row{entries,..}=&p.analysis.tree.kind else{panic!()};assert!(ids(&entries[0].node).is_disjoint(&ids(&entries[1].node)));
 let all=ids(&p.analysis.tree);t.edit(EditCommand::DeleteBackward);t.act(EditorAction::Undo);assert_eq!(ids(&t.s.document_snapshot().analysis.tree),all);
}
#[test]
fn named_delimiter_does_not_merge_with_first_symbol_after_visual_edit(){
 let mut t=Owned::new(r"\left\langle x\right\rangle");let p=t.s.document_snapshot();let NodeKind::Fence{body,..}=&first_fence(&p.analysis.tree).kind else{panic!()};
 t.select(Selection::caret(Position::row(body.id,0)));t.text("y");assert_eq!(t.s.source(),r"\left\langle y x\right\rangle");
 let canonical=t.s.document_snapshot().canonical_latex.unwrap();assert_eq!(canonical,r"\left\langle yx\right\rangle");
}
#[test]
fn square_fence_in_root_degree_is_brace_protected_and_reopenable(){
 let mut t=Owned::new("");t.edit(EditCommand::InsertStructure{structure:StructureKind::IndexedRoot});t.bracket("[");t.text("n");t.bracket("]");
 assert_eq!(t.s.source(),r"\sqrt[{\left[n\right]}]{}");
}
#[test]
fn source_patch_keeps_other_fence_identity_and_plain_source_parentheses_stay_plain(){
 assert_eq!(parse("(x)").facts.fences,0);
 let mut t=Owned::new(r"\left(x\right)+\left(x\right)");let before=t.s.document_snapshot();let NodeKind::Row{entries,..}=&before.analysis.tree.kind else{panic!()};let other=ids(&entries[2].node);
 t.s.patch_source(t.s.revision(),SourceEdit{start_utf16:6,end_utf16:7,replacement:"z".into()}).unwrap();let after=t.s.document_snapshot();
 let NodeKind::Row{entries,..}=&after.analysis.tree.kind else{panic!()};assert_eq!(ids(&entries[2].node),other);assert_eq!(semantic(&after.analysis.tree),semantic(&parse(t.s.source()).tree));
}
#[test]
fn font_driven_fences_center_on_axis_and_hidden_fences_add_no_ink(){
 let(f,o)=font();let mut engine=LayoutEngine::default();
 for (left,right) in [("(",")"),("[","]"),(r"\{",r"\}"),("|","|"),(r"\Vert",r"\Vert"),(r"\langle",r"\rangle"),(r"\lfloor",r"\rceil")]{
  let s=format!("\\left{left}\\frac{{1}}{{\\frac{{2}}{{3}}}}\\right{right}");let t=Owned::new(&s);let scene=engine.layout(&t.s,&f,&o).unwrap();scene.validate().unwrap();
  let p=t.s.document_snapshot();let fence=first_fence(&p.analysis.tree);let NodeKind::Fence{body,..}=&fence.kind else{panic!()};let body=scene.regions.iter().find(|r|r.node_id==body.id).unwrap();
  let axis=f.get(&o.math_font).unwrap().math.as_ref().unwrap().axis;let mut top=f64::INFINITY;let mut bottom=f64::NEG_INFINITY;
  for g in scene.glyphs.iter().filter(|g|g.node_id==fence.id&&g.x<body.logical.x_min){let ink=scene.resources[g.resource].shape.ink.unwrap();top=top.min(ink[1]*g.scale+g.y);bottom=bottom.max(ink[3]*g.scale+g.y);}
  assert!(((top+bottom)/2.0-(scene.baseline_em-axis)).abs()<1e-9);assert!(scene.svg().is_ok());
 }
 let t=Owned::new(r"\left.x\right.");let scene=engine.layout(&t.s,&f,&o).unwrap();let id=first_fence(&t.s.document_snapshot().analysis.tree).id;
 assert!(scene.glyphs.iter().all(|g|g.node_id!=id));
}
