use std::path::PathBuf;
use visualtex_engine_core::{editing::{EditorAction,EditorActionResult},fonts::FontRegistry,layout::{LayoutEngine,LayoutOptions},model::{self,LatexMode,Node,NodeKind},positions::{Position,Selection},session::EngineSession,transactions::{EditCommand,InputMode}};

fn middle<'a>(node:&'a Node)->&'a Node{if matches!(node.kind,NodeKind::MiddleDelimiter{..}){return node;}for child in node.children(){if let Some(found)=find_middle(child){return found;}}panic!("middle delimiter missing")}
fn find_middle(node:&Node)->Option<&Node>{if matches!(node.kind,NodeKind::MiddleDelimiter{..}){return Some(node);}node.children().into_iter().find_map(find_middle)}
fn fence<'a>(node:&'a Node)->&'a Node{if matches!(node.kind,NodeKind::Fence{..}){return node;}node.children().into_iter().find_map(|n|if matches!(n.kind,NodeKind::Fence{..}){Some(n)}else{None}).expect("fence missing")}
fn font()->(FontRegistry,LayoutOptions){let path=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));let mut fonts=FontRegistry::default();let id=fonts.register(&std::fs::read(path).unwrap(),0).unwrap().id;(fonts,LayoutOptions::new(id))}
fn layout(source:&str)->(visualtex_engine_core::scene::LayoutSnapshot,EngineSession){let session=EngineSession::from_source("middle-layout".into(),source.into()).unwrap();let (fonts,options)=font();let mut engine=LayoutEngine::default();let scene=engine.layout(&session,&fonts,&options).unwrap();(scene,session)}

#[test]
fn middle_delimiters_are_real_fence_children_and_roundtrip(){
 for source in [r"\left\langle a\middle|b\right\rangle",r"\left\langle a\middle\vert b\right\rangle",r"\left\lVert a\middle\Vert b\right\rVert",r"\left(\frac{a}{b}\middle|\sqrt{x}\right)"]{
  let session=EngineSession::from_source("middle".into(),source.into()).unwrap();let doc=session.document_snapshot();assert!(doc.analysis.complete&&doc.analysis.supported&&doc.analysis.exportable,"{:?}",doc.analysis.diagnostics);assert_eq!(model::latex(&doc.analysis.tree,LatexMode::Lossless),source);let m=middle(&doc.analysis.tree);let NodeKind::MiddleDelimiter{delimiter,..}=&m.kind else{panic!()};assert!(visualtex_engine_core::delimiters::glyph(delimiter).is_some());
 }
}

#[test]
fn middle_outside_fence_fails_closed_and_keeps_source(){
 let source=r"\middle|x";let session=EngineSession::from_source("outside".into(),source.into()).unwrap();let d=session.document_snapshot();assert_eq!(d.source,source);assert!(!d.analysis.supported&&!d.analysis.exportable);assert!(d.analysis.diagnostics.iter().any(|x|format!("{:?}",x.code)=="UnexpectedMiddleDelimiter"));
}

#[test]
fn tall_middle_uses_the_same_stretch_target_as_outer_fence(){
 let (simple,s1)=layout(r"\left\langle a\middle|b\right\rangle");let (tall,s2)=layout(r"\left\langle\frac{a}{b}\middle|\frac{c}{d}\right\rangle");let m1=middle(&s1.document_snapshot().analysis.tree).id;let m2=middle(&s2.document_snapshot().analysis.tree).id;let i1=simple.regions.iter().find(|r|r.node_id==m1).unwrap().ink.unwrap();let i2=tall.regions.iter().find(|r|r.node_id==m2).unwrap().ink.unwrap();assert!(i2.height()>i1.height()*1.25,"simple={} tall={}",i1.height(),i2.height());let f2=fence(&s2.document_snapshot().analysis.tree).id;assert!(tall.glyphs.iter().any(|g|g.node_id==m2));assert!(tall.glyphs.iter().any(|g|g.node_id==f2));
}

struct Owned{session:EngineSession,sequence:u32}
impl Owned{fn new(source:&str)->Self{Self{session:EngineSession::from_source("edit-middle".into(),source.into()).unwrap(),sequence:0}}fn act(&mut self,action:EditorAction)->EditorActionResult{self.sequence+=1;self.session.apply_action(self.session.revision(),self.sequence,action).unwrap()}fn edit(&mut self,command:EditCommand)->EditorActionResult{self.act(EditorAction::Edit{command,group_id:None})}}

#[test]
fn visual_edit_around_middle_preserves_middle_identity_and_history(){
 let mut owned=Owned::new(r"\left\langle a\middle|b\right\rangle");let before=owned.session.editor_snapshot().unwrap();let f=fence(&before.document.analysis.tree);let NodeKind::Fence{body,..}=&f.kind else{panic!()};let middle_id=middle(body).id;let NodeKind::Row{entries,..}=&body.kind else{panic!()};let b_id=entries.last().unwrap().node.id;owned.act(EditorAction::SetSelection{selection:Selection::caret(Position::row(body.id,2))});let changed=owned.edit(EditCommand::InsertText{text:"q".into(),mode:InputMode::Math}).editor;assert!(changed.document.source.contains(r"\middle|q"));assert_eq!(middle(&changed.document.analysis.tree).id,middle_id);assert!(changed.document.analysis.tree.find(b_id).is_some());let changed_source=changed.document.source.clone();owned.act(EditorAction::Undo);assert_eq!(owned.session.source(),before.document.source);assert_eq!(middle(&owned.session.document_snapshot().analysis.tree).id,middle_id);owned.act(EditorAction::Redo);assert_eq!(owned.session.source(),changed_source);assert_eq!(middle(&owned.session.document_snapshot().analysis.tree).id,middle_id);
}

#[test]
fn nested_fences_assign_middle_to_the_nearest_fence(){
 let source=r"\left( a+\left\langle\frac{x}{y}\middle|z\right\rangle\middle|b\right)";let session=EngineSession::from_source("nested-middle".into(),source.into()).unwrap();let d=session.document_snapshot();assert!(d.analysis.complete&&d.analysis.supported);let mut middles=Vec::new();d.analysis.tree.visit(&mut|n|if matches!(n.kind,NodeKind::MiddleDelimiter{..}){middles.push(n.id)});assert_eq!(middles.len(),2);let (fonts,options)=font();let mut engine=LayoutEngine::default();let scene=engine.layout(&session,&fonts,&options).unwrap();for id in middles{assert!(scene.glyphs.iter().any(|g|g.node_id==id));}
}
