//! Editable mathop bodies, authored tails, common operator placement and identity.
use std::{collections::BTreeSet,path::PathBuf,sync::Arc};
use serde_json::{Value,json};
use visualtex_engine_core::{editing::{EditorAction,EditorActionResult},fonts::FontRegistry,
    layout::{LayoutEngine,LayoutOptions,MathStyle},model::{self,Node,NodeKind,IdAllocator,LatexMode,ScriptKind},
    operators::{self,LimitMode},parser::{self,ParseResult},positions::{Position,Selection,Horizontal,TreeIndex},
    session::{EngineSession,NodeId},source::SourceText,source_edit::SourceEdit,
    transactions::{EditCommand,InputMode,StructureKind},ErrorCode,Runtime};
fn parse(source:&str)->ParseResult{parser::parse(SourceText::new(1,source).unwrap(),NodeId(1),&mut IdAllocator::new(2)).unwrap()}
fn semantic(node:&Node)->Value{
 fn clean(value:Value)->Value{match value{
  Value::Object(m)=>Value::Object(m.into_iter().filter(|(k,_)|!matches!(k.as_str(),"id"|"origin"|"grouping"|"leading"|"trailing"|"before")&&!k.starts_with("before")).map(|(k,v)|(k,clean(v))).collect()),
  Value::Array(a)=>Value::Array(a.into_iter().map(clean).collect()),v=>v,
 }}clean(serde_json::to_value(node).unwrap())
}
fn ids(node:&Node)->BTreeSet<NodeId>{let mut result=BTreeSet::new();node.visit(&mut|n|assert!(result.insert(n.id)));result}
fn general(node:&Node)->&Node{if matches!(node.kind,NodeKind::GeneralOperator{..}){return node}
 for child in node.children(){let mut yes=false;child.visit(&mut|n|yes|=matches!(n.kind,NodeKind::GeneralOperator{..}));if yes{return general(child)}}panic!("Missing general operator")}
fn body(node:&Node)->&Node{let NodeKind::GeneralOperator{body,..}=&general(node).kind else{panic!()};body}
fn bound(node:&Node,kind:ScriptKind)->&Node{operators::slots(operators::parts(general(node)).unwrap()).find(|s|s.script==kind).unwrap().argument.as_ref()}
fn resources()->(FontRegistry,LayoutOptions){
 let root=PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../..");let mut f=FontRegistry::default();
 let id=f.register(&std::fs::read(root.join("apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf")).unwrap(),0).unwrap().id;
 let mut options=LayoutOptions::new(id);options.text_font=Some(f.register(&std::fs::read("C:/Windows/Fonts/NotoSansSC-VF.ttf").unwrap(),0).unwrap().id);(f,options)
}
struct Owned{session:EngineSession,sequence:u32}
impl Owned{
 fn new(source:&str)->Self{Self{session:EngineSession::from_source("general-test".into(),source.into()).unwrap(),sequence:0}}
 fn act(&mut self,action:EditorAction)->EditorActionResult{
  let before=self.session.document_snapshot();self.sequence+=1;let r=self.session.apply_action(self.session.revision(),self.sequence,action).unwrap();
  let tree=&r.editor.document.analysis.tree;let index=TreeIndex::new(tree).unwrap();index.validate_selection(&r.editor.selection).unwrap();
  if let Some(map)=&r.position_map{for p in TreeIndex::new(&before.analysis.tree).unwrap().carets{map.apply_position(p,&index).unwrap();}}
  assert_eq!(ids(tree).len(),r.editor.document.analysis.facts.nodes);let q=parse(self.session.source());assert_eq!(semantic(tree),semantic(&q.tree),"{}",self.session.source());r
 }
 fn edit(&mut self,command:EditCommand)->EditorActionResult{self.act(EditorAction::Edit{command,group_id:None})}
 fn text(&mut self,text:&str){self.edit(EditCommand::InsertText{text:text.into(),mode:InputMode::Math});}
 fn select(&mut self,selection:Selection){self.act(EditorAction::SetSelection{selection});}
 fn right(&mut self){self.act(EditorAction::MoveHorizontal{direction:Horizontal::Right,extend:false});}
}
#[test]
fn general_operators_keep_original_bodies_bounds_directives_and_trivia(){
 for b in ["f(x)",r"\mathrm{rank}",r"\sum",r"\frac{a}{b}",r"\text{算子}",r"\mathop{g}_i"] {
  for tail in ["",r"_i^n",r"^n_i\limits",r"_i\nolimits^n\displaylimits", " % before\r\n _{i} % middle\n \\limits ^{n} "] {
   let source=format!("\\mathop % body\n {{{b}}}{tail} x");let p=parse(&source);
   assert!(p.complete&&p.supported&&p.exportable,"{source}: {:?}",p.diagnostics);assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
   let canonical=model::latex(&p.tree,LatexMode::Canonical);let q=parse(&canonical);assert!(q.complete&&q.supported);
   assert_eq!(semantic(&p.tree),semantic(&q.tree),"{source}");assert_eq!(model::latex(&q.tree,LatexMode::Canonical),canonical);
  }
 }
 for source in [r"\mathop f_a^b",r"\mathop\sum_i^n",r"\sqrt[{\mathop]}_i]{x}"] {
  let p=parse(source);assert!(p.complete&&p.supported,"{source}: {:?}",p.diagnostics);assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
  assert_eq!(semantic(&p.tree),semantic(&parse(&model::latex(&p.tree,LatexMode::Canonical)).tree));
 }
}
#[test]
fn incomplete_operator_bodies_and_duplicate_bounds_stay_drafts(){
 for source in [r"\mathop",r"\mathop{",r"\mathop{f}_",r"\mathop{f}^n^m",r"\mathop{f}_i_j",r"\mathop{f}\limits_{"] {
  let p=parse(source);assert!(!p.complete&&!p.exportable,"{source}");assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
 }
 for source in [r"\mathop{}",r"\mathop{}_i",r"\mathop{f}_{}",r"\mathop{f}^{}"] {
  let p=parse(source);assert!(p.complete&&p.supported&&!p.exportable,"{source}");
 }
 for source in [r"\unknownop{x}_i",r"\mathop{\unknown{x}}"]{let p=parse(source);assert!(!p.supported&&!p.exportable);assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);}
 let source=format!("\\mathop{{f}}{}",r"\limits".repeat(129));
 let error=parser::parse(SourceText::new(1,source).unwrap(),NodeId(1),&mut IdAllocator::new(2)).unwrap_err();assert_eq!(error.code,ErrorCode::ResourceLimit);
}
#[test]
fn body_and_optional_bounds_are_built_without_source_reparsing(){
 let mut t=Owned::new("");let empty=t.edit(EditCommand::InsertGeneralOperator{with_lower:true,with_upper:true});
 assert!(!empty.editor.document.analysis.exportable);t.text("0");t.right();t.text("1");t.right();t.text("f");t.right();t.text("x");
 assert_eq!(t.session.source(),r"\mathop{f}_{0}^{1}x");
 let prior=t.session.editor_snapshot().unwrap();t.act(EditorAction::Undo);t.act(EditorAction::Redo);let after=t.session.editor_snapshot().unwrap();
 assert_eq!(after.document.source,prior.document.source);assert_eq!(after.selection,prior.selection);assert_eq!(ids(&after.document.analysis.tree),ids(&prior.document.analysis.tree));
 let mut t=Owned::new("");t.edit(EditCommand::InsertGeneralOperator{with_lower:false,with_upper:false});t.text("f(x)");
 assert_eq!(t.session.source(),r"\mathop{f(x)}");
}
#[test]
fn selected_expression_becomes_operator_body_without_losing_node_identity(){
 let mut t=Owned::new(r"\hat{f}(x)");let initial=t.session.document_snapshot();let root=initial.analysis.tree.id;let old=ids(&initial.analysis.tree);
 let NodeKind::Row{entries,..}=&initial.analysis.tree.kind else{panic!()};let count=entries.len();
 t.select(Selection::Range{anchor:Position::row(root,0),focus:Position::row(root,count)});
 let result=t.edit(EditCommand::InsertGeneralOperator{with_lower:true,with_upper:false});
 let b=body(&result.editor.document.analysis.tree);assert!(old.into_iter().filter(|id|*id!=root).all(|id|b.find(id).is_some()));
 assert_eq!(result.editor.selection,Selection::caret(Position::row(bound(&result.editor.document.analysis.tree,ScriptKind::Subscript).id,0)));
 t.text("i");assert_eq!(t.session.source(),r"\mathop{\hat{f}(x)}_{i}");
}
#[test]
fn copied_general_operator_body_and_bounds_are_independent_and_delete_undoable(){
 let mut t=Owned::new(r"\mathop{f}_a^b");let initial=t.session.document_snapshot();let id=general(&initial.analysis.tree).id;
 t.select(Selection::Node{node_id:id});let copied=t.edit(EditCommand::DuplicateSelection);
 let NodeKind::Row{entries,..}=&copied.editor.document.analysis.tree.kind else{panic!()};let a=&entries[0].node;let b=&entries[1].node;
 assert!(ids(a).is_disjoint(&ids(b)));let copy_id=b.id;let copy_body=body(b).id;
 t.select(Selection::caret(Position::row(copy_body,1)));t.text("+g");
 assert_eq!(t.session.source(),r"\mathop{f}_a^b\mathop{f+g}_a^b");
 t.select(Selection::Node{node_id:copy_id});t.edit(EditCommand::DeleteForward);assert_eq!(t.session.source(),initial.source);
 t.act(EditorAction::Undo);assert!(t.session.document_snapshot().analysis.tree.find(copy_id).is_some());
}
#[test]
fn limits_changes_from_body_retain_all_rows_and_restore_authored_source_on_undo(){
 let mut t=Owned::new("\\mathop % op\n {f} % lower\n _a % upper\n ^b+x");let initial=t.session.document_snapshot();let op=general(&initial.analysis.tree);
 let old_ids=ids(op);let p=Position::row(body(op).id,1);t.select(Selection::caret(p));
 let changed=t.edit(EditCommand::SetOperatorLimits{mode:LimitMode::NoLimits});assert_eq!(ids(general(&changed.editor.document.analysis.tree)),old_ids);
 assert!(t.session.source().contains(r"\nolimits"));let before=t.session.editor_snapshot().unwrap();
 let no_op=t.edit(EditCommand::SetOperatorLimits{mode:LimitMode::NoLimits});assert!(!no_op.content_changed);assert_eq!(no_op.editor.undo_depth,before.undo_depth);
 t.act(EditorAction::Undo);assert_eq!(t.session.source(),initial.source);t.act(EditorAction::Redo);
 t.edit(EditCommand::SetOperatorLimits{mode:LimitMode::Auto});assert_eq!(t.session.source(),initial.source);
}
#[test]
fn adding_and_reentering_bounds_keeps_the_general_operator_body(){
 let mut t=Owned::new(r"\mathop{f}");let document=t.session.document_snapshot();let id=general(&document.analysis.tree).id;let body_id=body(&document.analysis.tree).id;
 t.select(Selection::caret(Position::row(document.analysis.tree.id,1)));
 let inserted=t.edit(EditCommand::InsertStructure{structure:StructureKind::Subscript});let lower=bound(&inserted.editor.document.analysis.tree,ScriptKind::Subscript).id;t.text("i");
 t.select(Selection::caret(Position::row(document.analysis.tree.id,1)));t.edit(EditCommand::InsertStructure{structure:StructureKind::Superscript});t.text("n");
 let after=t.session.document_snapshot();assert_eq!(general(&after.analysis.tree).id,id);assert_eq!(body(&after.analysis.tree).id,body_id);assert_eq!(bound(&after.analysis.tree,ScriptKind::Subscript).id,lower);
 assert_eq!(t.session.source(),r"\mathop{f}_{i}^{n}");t.select(Selection::caret(Position::row(document.analysis.tree.id,1)));
 let reentered=t.edit(EditCommand::InsertStructure{structure:StructureKind::Subscript});assert!(!reentered.content_changed);assert_eq!(reentered.editor.selection,Selection::caret(Position::row(lower,0)));
}
#[test]
fn a_power_of_the_entire_selected_operator_is_grouped_not_a_duplicate_upper_limit(){
 let mut t=Owned::new(r"\mathop{f}_a^b");let d=t.session.document_snapshot();t.select(Selection::Node{node_id:general(&d.analysis.tree).id});
 t.edit(EditCommand::InsertStructure{structure:StructureKind::Superscript});t.text("2");assert_eq!(t.session.source(),r"{\mathop{f}_a^b}^{2}");
}
#[test]
fn local_body_and_bound_patches_change_content_not_unrelated_operator_identities(){
 for (needle,replacement) in [("{f}","g"),("{a}","0"),("{b}","1")] {
  let source=r"\mathop{f}_{a}^{b}+\mathop{f}_{a}^{b}";let mut t=EngineSession::from_source("patch".into(),source.into()).unwrap();
  let initial=t.document_snapshot();let NodeKind::Row{entries,..}=&initial.analysis.tree.kind else{panic!()};let other=ids(&entries[2].node);let op_id=entries[0].node.id;
  let at=source.find(needle).unwrap()+1;let result=t.patch_source(t.revision(),SourceEdit{start_utf16:at as u32,end_utf16:(at+1)as u32,replacement:replacement.into()}).unwrap().0;
  let NodeKind::Row{entries,..}=&result.analysis.tree.kind else{panic!()};assert_eq!(entries[0].node.id,op_id);assert_eq!(ids(&entries[2].node),other);
  assert_eq!(semantic(&result.analysis.tree),semantic(&parse(&result.source).tree));assert!(result.source.contains(replacement));
 }
}
#[test]
fn general_operator_layout_obeys_limits_policy_and_preserves_body_typography(){
 let (fonts,mut options)=resources();let math=fonts.get(&options.math_font).unwrap().math.as_ref().unwrap();
 for style in [MathStyle::Display,MathStyle::Text,MathStyle::Script,MathStyle::ScriptScript]{options.style=style;
  for directive in ["",r"\limits",r"\nolimits",r"\displaylimits"] {
   let source=format!("\\mathop{{f}}{directive}_a^b");let s=EngineSession::from_source("layout".into(),source).unwrap();let d=s.document_snapshot();
   let op=general(&d.analysis.tree);let b=body(op);let lower=bound(op,ScriptKind::Subscript);let upper=bound(op,ScriptKind::Superscript);
   let scene=LayoutEngine::default().layout(&s,&fonts,&options).unwrap();scene.validate().unwrap();scene.svg().unwrap();
   let br=scene.regions.iter().find(|r|r.node_id==b.id).unwrap();let lr=scene.regions.iter().find(|r|r.node_id==lower.id).unwrap();let ur=scene.regions.iter().find(|r|r.node_id==upper.id).unwrap();
   let scale=match style{MathStyle::Display|MathStyle::Text=>1.0,MathStyle::Script=>math.script_scale,_=>math.scriptscript_scale};
   let stacked=directive==r"\limits"||directive!=r"\nolimits"&&style==MathStyle::Display;
   if stacked{assert!(br.logical.y_min-ur.logical.y_max+1e-9>=math.upper_limit_gap*scale);assert!(lr.logical.y_min-br.logical.y_max+1e-9>=math.lower_limit_gap*scale);}
   else{assert!(lr.logical.x_min+1e-9>=br.logical.x_max);assert!(ur.logical.x_min+1e-9>=br.logical.x_max);}
   let plain=EngineSession::from_source("plain".into(),"f".into()).unwrap();let p=LayoutEngine::default().layout(&plain,&fonts,&options).unwrap();
   let own=scene.glyphs.iter().find(|g|ids(b).contains(&g.node_id)).unwrap();let pg=&p.glyphs[0];
   assert_eq!(scene.resources[own.resource].shape.glyph_id,p.resources[pg.resource].shape.glyph_id);assert_eq!(own.scale,pg.scale);
  }
 }
}
#[test]
fn vertical_queries_land_in_the_same_operator_and_body_is_a_real_middle_stop(){
 let (fonts,options)=resources();let s=EngineSession::from_source("vertical".into(),r"\mathop{f}_{a}^{b}+\frac{x}{y}".into()).unwrap();let d=s.document_snapshot();let op=general(&d.analysis.tree);
 let b=body(op);let lo=bound(op,ScriptKind::Subscript);let up=bound(op,ScriptKind::Superscript);let scene=LayoutEngine::default().layout(&s,&fonts,&options).unwrap();
 assert_eq!(scene.vertical(Position::row(lo.id,0),false,None).unwrap().node_id(),b.id);
 assert_eq!(scene.vertical(Position::row(b.id,0),false,None).unwrap().node_id(),up.id);
 assert_eq!(scene.vertical(Position::row(up.id,0),true,None).unwrap().node_id(),b.id);
 assert_eq!(scene.vertical(Position::row(b.id,0),true,None).unwrap().node_id(),lo.id);
}
#[test]
fn body_structure_colors_hidden_paint_and_empty_slots_share_layout_rules(){
 let (fonts,options)=resources();
 for source in [r"\mathop{\binom{n}{k}}_0^1",r"\mathop{\sum\sum}_i^n",r"\mathop{\text{算子}}_a^b",r"\textcolor{red}{\mathop{f}_a^b}",r"\phantom{\mathop{f}_a^b}+x",r"\mathop{}",r"\mathop{f}_{}"]{
  let s=EngineSession::from_source("nested".into(),source.into()).unwrap();let scene=LayoutEngine::default().layout(&s,&fonts,&options).unwrap();scene.validate().unwrap();
  if scene.export_ready{scene.svg().unwrap();}else{assert_eq!(scene.svg().unwrap_err().code,ErrorCode::IncompleteSource);}
  if source.starts_with(r"\textcolor"){assert!(scene.glyphs.iter().all(|g|g.color.is_some()));}
  if source.starts_with(r"\phantom"){assert_eq!(scene.glyphs.len(),2);}
 }
}
#[test]
fn forged_general_operator_and_malformed_wire_actions_fail_atomically(){
 let mut p=parse(r"\mathop{f}_a");let root=Arc::make_mut(&mut p.tree);let NodeKind::Row{entries,..}=&mut root.kind else{panic!()};
 let node=Arc::make_mut(&mut entries[0].node);let NodeKind::GeneralOperator{command,..}=&mut node.kind else{panic!()};*command=r"\unknownop".into();
 assert_eq!(model::validate_tree(&p.tree).unwrap_err().code,ErrorCode::InvalidInput);
 let mut r=Runtime::default();let send=|r:&mut Runtime,c:Value|serde_json::from_str::<Value>(&r.dispatch_json(&json!({"protocol":1,"command":c}).to_string())).unwrap();
 send(&mut r,json!({"op":"create","documentId":"wire","source":"x"}));let before=send(&mut r,json!({"op":"inspectEditor","handle":1}));
 for action in [json!({"kind":"insertGeneralOperator","withLower":true}),json!({"kind":"insertGeneralOperator","withLower":true,"withUpper":false,"source":"evil"})]{
  let error=send(&mut r,json!({"op":"editorAction","handle":1,"expectedRevision":0,"eventSequence":1,"action":{"kind":"edit","command":action}}));assert_eq!(error["error"]["code"],"invalidInput");
  assert_eq!(send(&mut r,json!({"op":"inspectEditor","handle":1})),before);
 }
}
#[test]
fn seeded_general_operator_source_edits_match_a_fresh_parse_without_data_loss(){
 let seed=0x0F3A2026u32;let mut random=seed;
 for i in 0..600 {random=random.wrapping_mul(1664525).wrapping_add(1013904223);
  let source=r"\mathop{f+\mathop{g}_i}\limits_{a}^{b}+\mathop{f}_a^b % keep";
  let at=random as usize%(source.len()+1);let end=if i%3==0{(at+1).min(source.len())}else{at};let addition=["x","{","}","%","_","^",r"\mathop{x}_0",r"\nolimits"][i%8];
  let expected=format!("{}{}{}",&source[..at],addition,&source[end..]);let mut s=EngineSession::from_source("fuzz".into(),source.into()).unwrap();
  let d=s.patch_source(s.revision(),SourceEdit{start_utf16:at as u32,end_utf16:end as u32,replacement:addition.into()}).unwrap().0;let q=parse(&expected);
  assert_eq!(d.source,expected);assert_eq!(semantic(&d.analysis.tree),semantic(&q.tree),"{seed:x}/{i}: {expected}");
  assert_eq!((d.analysis.complete,d.analysis.supported),(q.complete,q.supported));
 }
}
