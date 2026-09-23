//! Remaining product decoration commands: no copied renderer or opaque output.
use std::{collections::BTreeSet,path::PathBuf};
use serde_json::Value;
use visualtex_engine_core::{accents::{self,Mark},editing::{EditorAction,EditorActionResult},fonts::FontRegistry,
    layout::{LayoutEngine,LayoutOptions,MathStyle},model::{self,Node,NodeKind,IdAllocator,LatexMode},
    parser::{self,ParseResult},positions::{Selection,Position,TreeIndex,Horizontal},session::{EngineSession,NodeId},
    source::SourceText,source_edit::SourceEdit,transactions::{EditCommand,InputMode},stretch,ErrorCode};
const COMMANDS:[&str;6]=[r"\overgroup",r"\undergroup",r"\overleftharpoon",r"\overrightharpoon",r"\overlinesegment",r"\underlinesegment"];
fn parse(s:&str)->ParseResult{parser::parse(SourceText::new(1,s).unwrap(),NodeId(1),&mut IdAllocator::new(2)).unwrap()}
fn semantic(node:&Node)->Value{
 fn clean(v:Value)->Value{match v{
  Value::Object(m)=>Value::Object(m.into_iter().filter(|(k,_)|!matches!(k.as_str(),"id"|"origin"|"grouping"|"leading"|"trailing")&&!k.starts_with("before")).map(|(k,v)|(k,clean(v))).collect()),
  Value::Array(a)=>Value::Array(a.into_iter().map(clean).collect()),v=>v,
 }}clean(serde_json::to_value(node).unwrap())
}
fn ids(node:&Node)->BTreeSet<NodeId>{let mut result=BTreeSet::new();node.visit(&mut|n|assert!(result.insert(n.id)));result}
fn first(node:&Node)->(&Node,&Node){if let NodeKind::Accent{body,..}=&node.kind{return(node,body)}
 for child in node.children(){let mut found=false;child.visit(&mut|n|found|=matches!(n.kind,NodeKind::Accent{..}));if found{return first(child)}}panic!("Missing accent")}
fn resources()->(FontRegistry,LayoutOptions){
 let root=PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../..");let mut fonts=FontRegistry::default();
 let math=root.join("apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf");
 let mut options=LayoutOptions::new(fonts.register(&std::fs::read(math).unwrap(),0).unwrap().id);
 options.math_fallbacks.push(fonts.register(&std::fs::read(root.join("engine/target/test-resources/stix-two-2.13/STIXTwoMath-Regular.otf")).unwrap(),0).unwrap().id);
 (fonts,options)
}
struct Owned{s:EngineSession,seq:u32}
impl Owned{
 fn new(source:&str)->Self{Self{s:EngineSession::from_source("mark".into(),source.into()).unwrap(),seq:0}}
 fn action(&mut self,action:EditorAction)->EditorActionResult{let old=self.s.document_snapshot();self.seq+=1;let result=self.s.apply_action(self.s.revision(),self.seq,action).unwrap();
  let index=TreeIndex::new(&result.editor.document.analysis.tree).unwrap();index.validate_selection(&result.editor.selection).unwrap();
  if let Some(map)=&result.position_map{for p in TreeIndex::new(&old.analysis.tree).unwrap().carets{map.apply_position(p,&index).unwrap();}}
  assert_eq!(semantic(&result.editor.document.analysis.tree),semantic(&parse(self.s.source()).tree),"{}",self.s.source());result}
 fn edit(&mut self,c:EditCommand)->EditorActionResult{self.action(EditorAction::Edit{command:c,group_id:None})}
 fn select(&mut self,s:Selection){self.action(EditorAction::SetSelection{selection:s});}
}
#[test]
fn decoration_registry_preserves_distinct_roles_and_authored_roundtrips(){
 for (c,g,below) in [(r"\overgroup",'⏠',false),(r"\undergroup",'⏡',true),(r"\overleftharpoon",'↼',false),(r"\overrightharpoon",'⇀',false)]{
  let s=accents::spec(c).unwrap();assert_eq!(s.mark,Mark::Glyph{codepoint:g,stretch:true});assert_eq!(s.below,below);
 }
 for c in COMMANDS{for body in ["{x}"," x"," % keep\r\n {a+b}","{\\binom{n}{k}}"]{
  let source=format!("{c}{body}+z");let p=parse(&source);assert!(p.complete&&p.supported&&p.exportable,"{source}");
  assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);let canonical=model::latex(&p.tree,LatexMode::Canonical);let q=parse(&canonical);
  assert!(q.complete&&q.supported);assert_eq!(semantic(&p.tree),semantic(&q.tree));assert_eq!(model::latex(&q.tree,LatexMode::Canonical),canonical);
 }}
}
#[test]
fn empty_draft_and_unknown_decoration_source_cannot_be_exported_or_erased(){
 for c in COMMANDS{for tail in ["","{"]{let source=format!("{c}{tail}");let p=parse(&source);assert!(!p.complete&&!p.exportable);assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);}
  let p=parse(&format!("{c}{{}}"));assert!(p.complete&&p.supported&&!p.exportable);}
 let p=parse(r"\futureovergroup{x}");assert!(!p.supported&&!p.exportable);assert_eq!(model::latex(&p.tree,LatexMode::Lossless),r"\futureovergroup{x}");
}
#[test]
fn all_new_decorations_are_editable_copied_independently_and_undoable(){
 for c in COMMANDS{
  let mut t=Owned::new("a+b");let before=t.s.document_snapshot();let old=ids(&before.analysis.tree);
  t.select(Selection::Range{anchor:Position::row(before.analysis.tree.id,0),focus:Position::row(before.analysis.tree.id,3)});
  let wrapped=t.edit(EditCommand::InsertAccent{command:c.into()});let (owner,body)=first(&wrapped.editor.document.analysis.tree);
  assert_eq!(t.s.source(),format!("{c}{{a+b}}"));assert!(old.into_iter().filter(|id|*id!=before.analysis.tree.id).all(|id|body.find(id).is_some()));
  let owner_id=owner.id;t.select(Selection::Node{node_id:owner_id});let copy=t.edit(EditCommand::DuplicateSelection);
  let NodeKind::Row{entries,..}=&copy.editor.document.analysis.tree.kind else{panic!()};assert!(ids(&entries[0].node).is_disjoint(&ids(&entries[1].node)));
  let copied=entries[1].node.id;let (_,inside)=first(&entries[1].node);t.select(Selection::caret(Position::row(inside.id,3)));
  t.edit(EditCommand::InsertText{text:"+c".into(),mode:InputMode::Math});assert_eq!(t.s.source(),format!("{c}{{a+b}}{c}{{a+b+c}}"));
  t.select(Selection::Node{node_id:copied});t.edit(EditCommand::DeleteBackward);assert_eq!(t.s.source(),format!("{c}{{a+b}}"));
  t.action(EditorAction::Undo);assert!(t.s.document_snapshot().analysis.tree.find(copied).is_some());
  t.action(EditorAction::Redo);assert_eq!(t.s.source(),format!("{c}{{a+b}}"));
 }
}
#[test]
fn group_harpoon_and_segment_geometry_uses_real_marks_and_consistent_vertical_gaps(){
 let (fonts,mut options)=resources();
 for c in COMMANDS{for style in [MathStyle::Display,MathStyle::Text,MathStyle::Script,MathStyle::ScriptScript]{
  options.style=style;for body in ["x".to_owned(),"x+y".into(),format!("{}x","x+".repeat(24))]{
   let session=EngineSession::from_source("geometry".into(),format!("{c}{{{body}}}")).unwrap();let doc=session.document_snapshot();let (owner,child)=first(&doc.analysis.tree);
   let scene=LayoutEngine::default().layout(&session,&fonts,&options).unwrap_or_else(|e|panic!("{c}/{style:?}: {e:?}"));scene.validate().unwrap();scene.svg().unwrap();
   let spec=accents::spec(c).unwrap();let region=scene.regions.iter().find(|r|r.node_id==child.id).unwrap();let ink=region.ink.unwrap_or(region.logical);
   match spec.mark{
    Mark::Glyph{codepoint,..}=>{
     let font=if fonts.get(&options.math_font).unwrap().glyph_index(codepoint).unwrap().is_some(){fonts.get(&options.math_font).unwrap()}else{fonts.get(&options.math_fallbacks[0]).unwrap()};
     let own:Vec<_>=scene.glyphs.iter().filter(|g|g.node_id==owner.id).collect();assert!(!own.is_empty());assert!(own.iter().all(|g|scene.resources[g.resource].font_id==font.identity.id));
     for g in own{let b=scene.resources[g.resource].shape.ink.unwrap();if spec.below{assert!(g.y+b[1]*g.scale>=ink.y_max-1e-9);}else{assert!(g.y+b[3]*g.scale<=ink.y_min+1e-9);}}
     if body.len()>20{let width=region.logical.width();let scale=scene.glyphs.iter().find(|g|g.node_id==owner.id).unwrap().scale;
      let mark=stretch::horizontal_accent(font,font.glyph_id(codepoint).unwrap(),width/scale).unwrap();assert!(mark.assembled&&mark.target_met);}
    },
    Mark::Segment=>{let rules:Vec<_>=scene.rules.iter().filter(|r|r.node_id==owner.id).collect();assert_eq!(rules.len(),3);
     let line=rules[0].rect;let left=rules[1].rect;let right=rules[2].rect;
     assert!(left.y_min<line.y_min&&left.y_max>line.y_max);assert_eq!(left.height(),right.height());
     assert!((left.width()-line.height()).abs()<1e-9);assert!((right.width()-line.height()).abs()<1e-9);
     assert!((left.x_min-line.x_min).abs()<1e-9&&(right.x_max-line.x_max).abs()<1e-9);
     for r in rules{if spec.below{assert!(r.rect.y_min>=ink.y_max);}else{assert!(r.rect.y_max<=ink.y_min);}}
    },Mark::Rule=>panic!("Expected marked segment not plain rule"),
   }
  }
 }}
}
#[test]
fn nested_new_decoration_exit_and_font_color_phantom_inheritance_remain_local(){
 let (fonts,options)=resources();let mut t=Owned::new("");
 t.edit(EditCommand::InsertAccent{command:r"\overgroup".into()});t.edit(EditCommand::InsertAccent{command:r"\overleftharpoon".into()});
 t.edit(EditCommand::InsertText{text:"x".into(),mode:InputMode::Math});let before=t.s.editor_snapshot().unwrap();
 t.action(EditorAction::MoveHorizontal{direction:Horizontal::Right,extend:false});let after=t.s.editor_snapshot().unwrap();
 assert_eq!(before.document.source,after.document.source);assert_eq!(before.document.revision,after.document.revision);assert_eq!(before.undo_depth,after.undo_depth);
 assert_eq!(after.selection,Selection::caret(Position::row(after.document.analysis.tree.id,1)));
 for source in [r"\textcolor{red}{\overlinesegment{x+y}}",r"\textcolor{red}{\overgroup{x+y}}",r"\phantom{\underlinesegment{x}}+y",r"\binom{\overgroup{n}}{\undergroup{k}}",r"\overlinesegment{\!}"]{
  let session=EngineSession::from_source("nested".into(),source.into()).unwrap();let scene=LayoutEngine::default().layout(&session,&fonts,&options).unwrap();scene.validate().unwrap();if scene.export_ready{scene.svg().unwrap();}
  if source.starts_with(r"\textcolor"){assert!(scene.glyphs.iter().all(|g|g.color.is_some()));assert!(scene.rules.iter().all(|r|r.color.is_some()));}
  if source.starts_with(r"\phantom"){assert_eq!(scene.glyphs.len(),2);assert!(scene.rules.is_empty());}
 }
}
#[test]
fn scaling_and_undo_restore_identical_decorated_svg_without_stale_paint(){
 let (fonts,mut options)=resources();
 for c in COMMANDS{
  let mut t=Owned::new(&format!("{c}{{x+y}}"));let mut layout=LayoutEngine::default();let initial=layout.layout(&t.s,&fonts,&options).unwrap();
  let document=t.s.document_snapshot();let (owner,_)=first(&document.analysis.tree);let owner_id=owner.id;
  t.select(Selection::Node{node_id:owner_id});t.edit(EditCommand::DeleteForward);t.action(EditorAction::Undo);
  assert_eq!(layout.layout(&t.s,&fonts,&options).unwrap().svg().unwrap(),initial.svg().unwrap());
  options.font_size_px*=2.0;let large=layout.layout(&t.s,&fonts,&options).unwrap();assert_eq!(large.rebuilt_fragments,0);assert_eq!(large.width,initial.width*2.0);
  assert_eq!(serde_json::to_value(&large.glyphs).unwrap(),serde_json::to_value(&initial.glyphs).unwrap());options.font_size_px/=2.0;
 }
}
#[test]
fn seeded_decoration_edits_match_fresh_parse_and_keep_all_user_source(){
 let seed=0xaCce2026u32;let mut random=seed;
 for i in 0..360{
  random=random.wrapping_mul(1664525).wrapping_add(1013904223);let c=COMMANDS[i%6];let source=format!("{c}{{x+\\binom{{n}}{{k}}}}+{c}{{y}} % keep");
  let at=random as usize%(source.len()+1);let end=if i%3==0{(at+1).min(source.len())}else{at};let insert=["x","}","{","%","\\",r"\overgroup{z}"][i%6];
  let expected=format!("{}{}{}",&source[..at],insert,&source[end..]);let mut s=EngineSession::from_source("fuzz".into(),source).unwrap();
  let result=s.patch_source(s.revision(),SourceEdit{start_utf16:at as u32,end_utf16:end as u32,replacement:insert.into()}).unwrap().0;let fresh=parse(&expected);
  assert_eq!(result.source,expected);assert_eq!(semantic(&result.analysis.tree),semantic(&fresh.tree),"{seed:x}/{i}: {expected}");
  assert_eq!((result.analysis.complete,result.analysis.supported),(fresh.complete,fresh.supported));
 }
}
#[test]
fn unknown_accent_command_does_not_consume_document_or_history_state(){
 let mut s=EngineSession::from_source("atomic".into(),"x".into()).unwrap();let before=serde_json::to_value(s.editor_snapshot().unwrap()).unwrap();
 let e=s.apply_action(s.revision(),1,EditorAction::Edit{command:EditCommand::InsertAccent{command:r"\unknownsegment".into()},group_id:None}).unwrap_err();
 assert_eq!(e.code,ErrorCode::UnsupportedSyntax);assert_eq!(serde_json::to_value(s.editor_snapshot().unwrap()).unwrap(),before);
}
