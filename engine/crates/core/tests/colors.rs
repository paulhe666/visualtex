//! P07 paint: author scopes, real editing, safe bounded RGB and exact source retention.
use std::{collections::BTreeSet,path::PathBuf};
use visualtex_engine_core::{ErrorCode,Runtime,editing::{EditorAction,EditorActionResult},fonts::FontRegistry,
 layout::{LayoutEngine,LayoutOptions,MathStyle},model::{self,IdAllocator,LatexMode,Node,NodeKind},
 paint::{self,RgbColor,ColorSpec},parser::{self,ParseResult},positions::{Position,Selection,TreeIndex,Horizontal},
 scene::LayoutSnapshot,session::{EngineSession,NodeId},source::SourceText,source_edit::SourceEdit,
 transactions::{EditCommand,InputMode}};
const RED:RgbColor=RgbColor{r:255,g:0,b:0};const BLUE:RgbColor=RgbColor{r:0,g:0,b:255};
fn parse(s:&str)->ParseResult{parser::parse(SourceText::new(1,s).unwrap(),NodeId(1),&mut IdAllocator::new(2)).unwrap()}
fn ids(n:&Node)->BTreeSet<NodeId>{let mut set=BTreeSet::new();n.visit(&mut |n|assert!(set.insert(n.id)));set}
fn fonts()->(FontRegistry,LayoutOptions){
 let path=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||
 PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
 let mut fonts=FontRegistry::default();let id=fonts.register(&std::fs::read(path).unwrap(),0).unwrap().id;(fonts,LayoutOptions::new(id))
}
fn draw(s:&str)->LayoutSnapshot{let(f,o)=fonts();LayoutEngine::default().layout(&EngineSession::from_source("paint-layout".into(),s.into()).unwrap(),&f,&o).unwrap()}
fn close(a:f64,b:f64){assert!((a-b).abs()<1e-9,"{a} != {b}");}
fn check(doc:&visualtex_engine_core::session::DocumentSnapshot){
 let p=parse(&doc.source);assert_eq!(model::latex(&p.tree,LatexMode::Canonical),model::latex(&doc.analysis.tree,LatexMode::Canonical),"{}",doc.source);
 assert_eq!(p.facts,doc.analysis.facts);assert_eq!((p.complete,p.supported,p.exportable),(doc.analysis.complete,doc.analysis.supported,doc.analysis.exportable));
 doc.analysis.tree.visit(&mut |n|{let o=n.origin.as_ref().unwrap();assert_eq!(o.span.source_id,doc.source_id);
 assert_eq!(o.text(),&doc.source[o.span.byte_start as usize..o.span.byte_end as usize]);});
}
struct Owned{session:EngineSession,seq:u32}
impl Owned{
 fn new(s:&str)->Self{Self{session:EngineSession::from_source("paint-editor".into(),s.into()).unwrap(),seq:0}}
 fn action(&mut self,a:EditorAction)->EditorActionResult{let before=self.session.document_snapshot();self.seq+=1;
 let r=self.session.apply_action(self.session.revision(),self.seq,a).unwrap();let idx=TreeIndex::new(&r.editor.document.analysis.tree).unwrap();
 idx.validate_selection(&r.editor.selection).unwrap();if let Some(map)=&r.position_map{for p in TreeIndex::new(&before.analysis.tree).unwrap().carets{map.apply_position(p,&idx).unwrap();}}
 check(&r.editor.document);r}
 fn edit(&mut self,c:EditCommand)->EditorActionResult{self.action(EditorAction::Edit{command:c,group_id:None})}
 fn all(&mut self){let d=self.session.document_snapshot();let NodeKind::Row{entries,..}=&d.analysis.tree.kind else{panic!()};
 self.action(EditorAction::SetSelection{selection:Selection::Range{anchor:Position::row(d.analysis.tree.id,0),focus:Position::row(d.analysis.tree.id,entries.len())}});}
}
#[test]fn accepted_models_spelling_comments_and_mixing_roundtrip_without_normalizing_source(){
 for header in ["{red}","[RGB]{255, 0, 0}","[rgb]{1,.5,0}","[HTML]{aAbB09}","[gray]{.3}","{red!20!blue!50}",
 " % before\n [HTML] % gap\r\n {Ab%inner\n 1234}"]{
  // Whitespace inside HTML digits is unsupported; test an actual comment between complete values below.
  if header.contains("Ab%inner"){let p=parse(&format!("\\textcolor{header}{{x}}"));assert!(!p.supported);continue;}
  for command in ["\\textcolor","\\color"]{let source=format!("{command}{header}{{x+y}}+z");let p=parse(&source);
   assert!(p.exportable,"{source}: {:?}",p.diagnostics);assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
   let canonical=model::latex(&p.tree,LatexMode::Canonical);let q=parse(&canonical);assert!(q.exportable);assert_eq!(p.facts,q.facts);
   assert_eq!(model::latex(&q.tree,LatexMode::Canonical),canonical);
  }
 }
 let source="\\textcolor % first\n [rgb] % second\n {1,0,0 % keep\n } % body\n {x}";
 let p=parse(source);assert!(p.exportable);assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
}
#[test]fn invalid_unsafe_unknown_and_unfinished_color_is_preserved_and_not_exported(){
 for source in [r"\textcolor{url(https://example.invalid)}{x}",r"\textcolor[RGB]{256,0,0}{x}",r"\color{currentColor}x",
 r"\textcolor[HTML]{#123456}{x}",r"\color[rgb]{NaN,0,0}x",r"\textcolor{red;fill:blue}{x}",r"\color[cmyk]{0,1,1,0}x"]{
  let p=parse(source);assert!(!p.supported&&!p.exportable,"{source}");assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
  let(f,o)=fonts();let s=EngineSession::from_source("bad-color".into(),source.into()).unwrap();assert_eq!(LayoutEngine::default().layout(&s,&f,&o).unwrap_err().code,ErrorCode::UnsupportedSyntax);
 }
 for source in [r"\color",r"\textcolor",r"\textcolor{red}",r"\textcolor{red}{",r"\color[HTML",r"\color[HTML{ffffff}"]{
  let p=parse(source);assert!(!p.complete&&!p.exportable,"{source}");assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
 }
 for source in [r"\color{red}",r"\textcolor{red}{}"]{let p=parse(source);assert!(p.complete&&p.supported&&!p.exportable);}
}
#[test]fn nested_scopes_restore_outer_paint_and_declarations_affect_following_siblings_only(){
 let s=draw(r"a\color{red}b{\color{blue}c}d\textcolor{green}{e}f");
 assert_eq!(s.glyphs.iter().map(|g|g.color).collect::<Vec<_>>(),vec![None,Some(RED),Some(BLUE),Some(RED),Some(RgbColor{r:0,g:255,b:0}),Some(RED)]);
 let s=draw(r"\textcolor{red}{a\color{blue}b}c");
 assert_eq!(s.glyphs.iter().map(|g|g.color).collect::<Vec<_>>(),vec![Some(RED),Some(BLUE),None]);
}
#[test]fn coloring_does_not_change_atom_spacing_font_glyphs_or_natural_size(){
 for (plain,colored) in [("x+y",r"x\color{red}+y"),("x+y",r"x\textcolor{red}{+}y"),
 (r"\frac{x}{y}",r"\textcolor{red}{\frac{x}{y}}"),(r"\sqrt{x}",r"\textcolor{blue}{\sqrt{x}}")]{
  let a=draw(plain);let b=draw(colored);for(x,y)in[(a.width,b.width),(a.height,b.height),(a.baseline,b.baseline),(a.advance_em,b.advance_em)]{close(x,y);}
  assert_eq!(a.glyphs.len(),b.glyphs.len());
  for(x,y)in a.glyphs.iter().zip(&b.glyphs){close(x.x,y.x);close(x.y,y.y);close(x.scale,y.scale);assert_eq!(a.resources[x.resource].shape.glyph_id,b.resources[y.resource].shape.glyph_id);}
 }
}
#[test]fn fraction_radical_accent_and_border_rules_inherit_color_and_phantoms_never_paint(){
 let s=draw(r"\textcolor{red}{\boxed{\frac{\sqrt{x}}{\overline{y}}}}\phantom{\textcolor{blue}{\boxed{z}}}");
 assert!(s.rules.len()>=7);assert!(!s.glyphs.is_empty());
 assert!(s.glyphs.iter().all(|g|g.color==Some(RED)));assert!(s.rules.iter().all(|r|r.color==Some(RED)));
 let svg=s.svg().unwrap();assert!(!svg.contains("#0000ff"));assert_eq!(svg.matches("fill=\"#ff0000\"").count(),s.glyphs.len()+s.rules.len());
 assert!(!svg.contains("style=")&&!svg.contains("url(")&&!svg.contains("<text"));
}
#[test]fn selected_formula_color_is_one_undoable_edit_and_copy_retains_color_with_fresh_ids(){
 let mut t=Owned::new("x+y");let original=ids(&t.session.document_snapshot().analysis.tree);t.all();let colored=t.edit(EditCommand::SetColor{color:RED});
 assert_eq!(t.session.source(),r"\textcolor[HTML]{FF0000}{x+y}");assert!(original.is_subset(&ids(&colored.editor.document.analysis.tree)));
 let snapshot=t.session.document_snapshot();t.action(EditorAction::Undo);assert_eq!(t.session.source(),"x+y");
 t.action(EditorAction::Redo);assert_eq!(ids(&t.session.document_snapshot().analysis.tree),ids(&snapshot.analysis.tree));
 t.all();t.edit(EditCommand::DuplicateSelection);let duplicate=t.session.document_snapshot();let NodeKind::Row{entries,..}=&duplicate.analysis.tree.kind else{panic!()};
 assert_eq!(entries.len(),2);assert!(ids(&entries[0].node).is_disjoint(&ids(&entries[1].node)));
 t.edit(EditCommand::DeleteBackward);assert_eq!(t.session.source(),snapshot.source);t.action(EditorAction::Undo);assert_eq!(ids(&t.session.document_snapshot().analysis.tree),ids(&duplicate.analysis.tree));
}
#[test]fn empty_color_slot_accepts_input_and_new_color_does_not_leak_after_exit(){
 let mut t=Owned::new("");t.edit(EditCommand::SetColor{color:BLUE});assert!(!t.session.document_snapshot().analysis.exportable);
 t.edit(EditCommand::InsertText{text:"x".into(),mode:InputMode::Math});let revision=t.session.revision();
 t.action(EditorAction::MoveHorizontal{direction:Horizontal::Right,extend:false});assert_eq!(revision,t.session.revision());
 t.edit(EditCommand::InsertText{text:"y".into(),mode:InputMode::Math});let s=draw(t.session.source());assert_eq!(s.glyphs.iter().map(|g|g.color).collect::<Vec<_>>(),vec![Some(BLUE),None]);
}
#[test]fn header_and_nested_body_source_edits_keep_actual_math_ranges_and_refresh_paint(){
 let mut s=EngineSession::from_source("color-patch".into(),r"\textcolor{red}{\mathbf{x}}+y".into()).unwrap();
 let(f,o)=fonts();let mut layout=LayoutEngine::default();let initial=layout.layout(&s,&f,&o).unwrap();
 let at=s.source().find("red").unwrap()as u32;
 s.patch_source(s.revision(),SourceEdit{start_utf16:at,end_utf16:at+3,replacement:"blue".into()}).unwrap();check(&s.document_snapshot());
 let next=layout.layout(&s,&f,&o).unwrap();close(next.width,initial.width);assert_eq!(next.glyphs[0].color,Some(BLUE));assert_ne!(initial.svg().unwrap(),next.svg().unwrap());
 let at=s.source().find('x').unwrap()as u32; // first x occurs in textcolor: choose the argument explicitly.
 let argument=s.source().find("{x}").unwrap()as u32+1;assert!(argument>at);
 s.patch_source(s.revision(),SourceEdit{start_utf16:argument,end_utf16:argument+1,replacement:"z".into()}).unwrap();check(&s.document_snapshot());
 assert_ne!(layout.layout(&s,&f,&o).unwrap().svg().unwrap(),next.svg().unwrap());
}
#[test]fn color_optional_headers_inside_root_degree_are_brace_protected_on_structural_edit(){
 let mut t=Owned::new(r"\sqrt[n]{x}");let d=t.session.document_snapshot();let NodeKind::Row{entries,..}=&d.analysis.tree.kind else{panic!()};
 let NodeKind::Radical{degree:Some(degree),..}=&entries[0].node.kind else{panic!()};
 t.action(EditorAction::SetSelection{selection:Selection::Range{anchor:Position::row(degree.id,0),focus:Position::row(degree.id,1)}});
 t.edit(EditCommand::SetColor{color:RED});assert!(t.session.source().starts_with(r"\sqrt[{\textcolor[HTML]"));assert!(parse(t.session.source()).exportable);
}
#[test]fn scene_cache_reuses_geometry_on_same_paint_and_never_keeps_retired_colors(){
 let(f,mut o)=fonts();let mut layout=LayoutEngine::default();let mut t=Owned::new(r"\textcolor{red}{x}");
 let a=layout.layout(&t.session,&f,&o).unwrap();let b=layout.layout(&t.session,&f,&o).unwrap();assert_eq!(b.rebuilt_fragments,0);assert_eq!(a.svg().unwrap(),b.svg().unwrap());
 t.session.replace_source(t.session.revision(),r"\textcolor{blue}{x}".into()).unwrap();let c=layout.layout(&t.session,&f,&o).unwrap();assert_eq!(c.glyphs[0].color,Some(BLUE));
 o.font_size_px*=2.;let d=layout.layout(&t.session,&f,&o).unwrap();assert_eq!(d.rebuilt_fragments,0);close(c.width*2.,d.width);
 o.style=MathStyle::Script;let e=layout.layout(&t.session,&f,&o).unwrap();assert_eq!(e.glyphs[0].color,Some(BLUE));
}
#[test]fn malformed_wire_color_fields_cannot_mutate_the_session_or_replay_receipt(){
 let mut rt=Runtime::default();let request=|command|serde_json::json!({"protocol":1,"command":command}).to_string();
 rt.dispatch_json(&request(serde_json::json!({"op":"create","documentId":"bad-paint-wire"})));
 let read=request(serde_json::json!({"op":"inspectEditor","handle":1}));let before=rt.dispatch_json(&read);
 for color in [serde_json::json!({"r":256,"g":0,"b":0}),serde_json::json!({"r":-1,"g":0,"b":0}),serde_json::json!("url(x)"),
  serde_json::json!({"r":1.5,"g":0,"b":0}),serde_json::json!({"r":0,"g":0,"b":0,"fill":"red"})]{
  let r:serde_json::Value=serde_json::from_str(&rt.dispatch_json(&request(serde_json::json!({"op":"editorAction","handle":1,"expectedRevision":0,"eventSequence":1,
   "action":{"kind":"edit","command":{"kind":"setColor","color":color}}})))).unwrap();
  assert_eq!(r["status"],"error");assert_eq!(rt.dispatch_json(&read),before);
 }
}
#[test]fn forged_color_metadata_cannot_be_projected_as_valid_math(){
 let mut p=parse(r"\textcolor{red}{x}");let tree=std::sync::Arc::make_mut(&mut p.tree);
 let NodeKind::Row{entries,..}=&mut tree.kind else{panic!()};let n=std::sync::Arc::make_mut(&mut entries[0].node);
 let NodeKind::Color{spec,..}=&mut n.kind else{panic!()};spec.color=BLUE;
 assert_eq!(model::validate_tree(tree).unwrap_err().code,ErrorCode::InvalidInput);
 assert!(ColorSpec::authored("\\evil".into(),Some("HTML".into()),String::new(),"ff0000".into()).is_err());
 assert_eq!(paint::resolve_color(None,&"r".repeat(1025)).unwrap_err().code,ErrorCode::ResourceLimit);
}
#[test]fn seeded_nested_color_source_edits_match_fresh_structure_and_keep_all_input(){
 let mut seed=0xc010_2026u32;let initial=r"\color{red}x+{\color[RGB]{0,0,255}\boxed{\textcolor{green}{y}}}+\phantom{\textcolor{blue}{z}}";
 let mut s=EngineSession::from_source("paint-fuzz".into(),initial.into()).unwrap();let replacements=["", "q", "{", "}", "%\n", "中", "[", "]", r"\color{red}"];
 for i in 0..640{if i%16==0{s.replace_source(s.revision(),initial.into()).unwrap();}
  seed=seed.wrapping_mul(1664525).wrapping_add(1013904223);let boundaries:Vec<_>=s.source().char_indices().map(|(b,_)|b).chain(std::iter::once(s.source().len())).collect();
  let index=seed as usize%boundaries.len();let a=boundaries[index];let b=boundaries[(index+usize::from(seed&1==0)).min(boundaries.len()-1)];
  let patch=SourceEdit{start_utf16:s.source()[..a].encode_utf16().count()as u32,end_utf16:s.source()[..b].encode_utf16().count()as u32,replacement:replacements[(seed>>8)as usize%replacements.len()].into()};
  let before=s.document_snapshot();match s.patch_source(s.revision(),patch){Ok(_)=>check(&s.document_snapshot()),Err(_)=>assert_eq!(serde_json::to_value(s.document_snapshot()).unwrap(),serde_json::to_value(before).unwrap(),"seed={seed} i={i}")}
 }
}
#[test]fn declaration_deletion_is_one_action_and_undo_restores_paint_not_just_source(){
 let mut t=Owned::new(r"\color{red}x");
 let root=t.session.document_snapshot().analysis.tree.id;
 t.action(EditorAction::SetSelection{selection:Selection::caret(Position::row(root,0))});
 t.edit(EditCommand::DeleteForward);assert_eq!(t.session.source(),"x");assert_eq!(draw(t.session.source()).glyphs[0].color,None);
 t.action(EditorAction::Undo);assert_eq!(draw(t.session.source()).glyphs[0].color,Some(RED));
}
