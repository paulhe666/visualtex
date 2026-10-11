//! P09 bounded compatibility macros: authored spelling stays in the model while
//! fixed semantics are rendered/editable without MathLive/MathJax expansion.
use std::{collections::BTreeSet,path::PathBuf};
use serde_json::Value;
use visualtex_engine_core::{editing::{EditorAction,EditorActionResult},fonts::FontRegistry,
 layout::{LayoutEngine,LayoutOptions},macros::{self,BuiltinMacroKind},model::{self,LatexMode,Node,NodeKind},
 positions::{Position,Selection,TreeIndex},session::EngineSession,source_edit::SourceEdit,
 transactions::{EditCommand,InputMode}};

struct Owned{s:EngineSession,seq:u32}
impl Owned{
 fn new(source:&str)->Self{Self{s:EngineSession::from_source("macro-review".into(),source.into()).unwrap(),seq:0}}
 fn action(&mut self,action:EditorAction)->EditorActionResult{self.seq+=1;self.s.apply_action(self.s.revision(),self.seq,action).unwrap()}
 fn select(&mut self,p:Position){self.action(EditorAction::SetSelection{selection:Selection::caret(p)});}
 fn edit(&mut self,c:EditCommand)->EditorActionResult{self.action(EditorAction::Edit{command:c,group_id:None})}
}
fn macro_node(root:&Node)->&Node{if matches!(root.kind,NodeKind::MacroCall{..}){return root;}root.children().into_iter().find_map(|child|{let mut found=None;if matches!(child.kind,NodeKind::MacroCall{..}){found=Some(child);}else{for nested in child.children(){if let Some(node)=find_macro(nested){found=Some(node);break;}}}found}).expect("macro call")}
fn find_macro(root:&Node)->Option<&Node>{if matches!(root.kind,NodeKind::MacroCall{..}){return Some(root);}root.children().into_iter().find_map(find_macro)}
fn ids(node:&Node)->BTreeSet<u32>{let mut result=BTreeSet::new();node.visit(&mut|n|{assert!(result.insert(n.id.0));});result}
fn semantic(v:Value)->Value{match v{Value::Object(m)=>Value::Object(m.into_iter().filter(|(k,_)|!matches!(k.as_str(),"id"|"origin"|"leading"|"trailing"|"grouping")&&!k.starts_with("before"))
 .map(|(k,v)|(k,semantic(v))).collect()),Value::Array(a)=>Value::Array(a.into_iter().map(semantic).collect()),v=>v}}
fn font()->(FontRegistry,LayoutOptions){let path=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));let mut fonts=FontRegistry::default();let id=fonts.register(&std::fs::read(path).unwrap(),0).unwrap().id;(fonts,LayoutOptions::new(id))}

#[test]
fn all_bounded_macros_preserve_authored_source_and_real_arguments(){
 let samples=[
  r"\abs{x+1}",r"\norm{\frac{x}{y}}",r"\dd{x}",r"\dv{f}{x}",r"\pdv{f}{x}",
  r"\comm{A}{B}",r"\acomm{A}{B}",r"\pb{f}{g}",r"\bra{\psi}",r"\ket{\phi}",
  r"\braket{\phi}{\psi}",r"\expval{A}",r"\mel{\phi}{A}{\psi}",r"\ketbra{\psi}{\phi}",
  r"\vb{v}",r"\va{v}",r"\vu{e}",r"\mathrel{⫽}",r"\mathbin{=}",r"\xrightarrow{f}",r"\xleftarrow{n\to\infty}",r"\pmod{p}"];
 assert_eq!(samples.len(),macros::all().len());
 for source in samples {let owned=Owned::new(source);let d=owned.s.document_snapshot();
  assert!(d.analysis.complete&&d.analysis.supported&&d.analysis.exportable,"{source}: {:?}",d.analysis.diagnostics);
  assert_eq!(model::latex(&d.analysis.tree,LatexMode::Lossless),source);
  let m=macro_node(&d.analysis.tree);let NodeKind::MacroCall{command,macro_kind,before_arguments,arguments}= &m.kind else{panic!()};
  let spec=macros::spec(command).unwrap();assert_eq!(*macro_kind,spec.kind);assert_eq!(arguments.len(),spec.arity);assert_eq!(before_arguments.len(),spec.arity);
  assert!(arguments.iter().all(|a|a.is_row()));let index=TreeIndex::new(&d.analysis.tree).unwrap();
  for arg in arguments {index.validate_position(Position::row(arg.id,0)).unwrap();}
 }
}

#[test]
fn comments_and_spacing_roundtrip_without_expanding_macro_source(){
 let source="\\pdv %a\n {f+1} %b\n { x }+\\abs  {y}";let owned=Owned::new(source);let d=owned.s.document_snapshot();
 assert!(d.analysis.supported);assert_eq!(model::latex(&d.analysis.tree,LatexMode::Lossless),source);
 assert_eq!(model::latex(&d.analysis.tree,LatexMode::Canonical),r"\pdv{f+1}{x}+\abs{y}");
}

#[test]
fn visual_edit_inside_macro_argument_preserves_command_other_argument_and_identity(){
 let mut owned=Owned::new(r"\pdv{f}{x}+z");let before=owned.s.document_snapshot();let m=macro_node(&before.analysis.tree);
 let NodeKind::MacroCall{arguments,..}=&m.kind else{panic!()};let macro_id=m.id;let denominator_id=arguments[1].id;
 let numerator_id=arguments[0].id;owned.select(Position::row(numerator_id,1));let edited=owned.edit(EditCommand::InsertText{text:"g".into(),mode:InputMode::Math}).editor;
 assert_eq!(edited.document.source,r"\pdv{fg}{x}+z");let after=macro_node(&edited.document.analysis.tree);assert_eq!(after.id,macro_id);
 let NodeKind::MacroCall{command,arguments,..}=&after.kind else{panic!()};assert_eq!(command,"\\pdv");assert_eq!(arguments[1].id,denominator_id);
 let fresh=Owned::new(&edited.document.source).s.document_snapshot();assert_eq!(semantic(serde_json::to_value(&edited.document.analysis.tree).unwrap()),semantic(serde_json::to_value(&fresh.analysis.tree).unwrap()));
 let undone=owned.action(EditorAction::Undo).editor;assert_eq!(undone.document.source,before.source);assert_eq!(macro_node(&undone.document.analysis.tree).id,macro_id);
 let redone=owned.action(EditorAction::Redo).editor;assert_eq!(redone.document.source,r"\pdv{fg}{x}+z");assert_eq!(macro_node(&redone.document.analysis.tree).id,macro_id);
}

#[test]
fn source_patch_inside_macro_argument_reuses_unaffected_macro_and_sibling_ids(){
 let mut owned=Owned::new(r"\mel{\phi}{A}{\psi}+q");let before=owned.s.document_snapshot();let m=macro_node(&before.analysis.tree);
 let NodeKind::MacroCall{arguments,..}=&m.kind else{panic!()};let old=(m.id,arguments[0].id,arguments[2].id);
 let at=before.source.find('A').unwrap();let (after,_)=owned.s.patch_source(owned.s.revision(),SourceEdit{start_utf16:at as u32,end_utf16:(at+1) as u32,replacement:"B".into()}).unwrap();
 assert_eq!(after.source,r"\mel{\phi}{B}{\psi}+q");let m=macro_node(&after.analysis.tree);let NodeKind::MacroCall{arguments,..}=&m.kind else{panic!()};
 assert_eq!(m.id,old.0);assert_eq!(arguments[0].id,old.1);assert_eq!(arguments[2].id,old.2);
}

#[test]
fn copying_macro_allocates_disjoint_argument_identities_and_undo_restores_original(){
 let mut owned=Owned::new(r"\braket{a}{b}");let before=owned.s.editor_snapshot().unwrap();let root=&before.document.analysis.tree;
 let NodeKind::Row{entries,..}=&root.kind else{panic!()};let original=ids(&entries[0].node);
 owned.action(EditorAction::SetSelection{selection:Selection::Range{anchor:Position::row(root.id,0),focus:Position::row(root.id,1)}});
 let duplicated=owned.edit(EditCommand::DuplicateSelection).editor;let NodeKind::Row{entries,..}=&duplicated.document.analysis.tree.kind else{panic!()};assert_eq!(entries.len(),2);
 let second=ids(&entries[1].node);assert!(original.is_disjoint(&second));assert_eq!(duplicated.document.source,r"\braket{a}{b}\braket{a}{b}");
 let undone=owned.action(EditorAction::Undo).editor;assert_eq!(undone.document.source,r"\braket{a}{b}");let NodeKind::Row{entries,..}=&undone.document.analysis.tree.kind else{panic!()};assert_eq!(ids(&entries[0].node),original);
}

#[test]
fn native_layout_uses_real_geometry_for_every_macro_kind(){
 let source=r"\abs{\frac{x}{y}}+\norm{x}+\dd{x}+\dv{f}{x}+\pdv{f}{x}+\comm{A}{B}+\acomm{A}{B}+\pb{f}{g}+\bra{\psi}+\ket{\phi}+\braket{\phi}{\psi}+\expval{A}+\mel{\phi}{A}{\psi}+\ketbra{\psi}{\phi}+\vb{v}+\va{v}+\vu{e}";
 let owned=Owned::new(source);let d=owned.s.document_snapshot();assert!(d.analysis.exportable,"{:?}",d.analysis.diagnostics);let(fonts,options)=font();let mut layout=LayoutEngine::default();
 let scene=layout.layout(&owned.s,&fonts,&options).unwrap();assert!(scene.width>0.0&&scene.height>0.0);assert!(scene.glyphs.len()>20);assert!(scene.rules.len()>=3);
 let svg=scene.svg().unwrap();assert!(svg.starts_with("<svg")&&!svg.contains("<text")&&!svg.contains("foreignObject"));
 for spec in macros::all(){let source=match spec.kind{BuiltinMacroKind::Derivative=>r"\dv{f}{x}",BuiltinMacroKind::PartialDerivative=>r"\pdv{f}{x}",BuiltinMacroKind::MatrixElement=>r"\mel{a}{b}{c}",BuiltinMacroKind::BraKet=>r"\braket{a}{b}",BuiltinMacroKind::KetBra=>r"\ketbra{a}{b}",BuiltinMacroKind::Commutator=>r"\comm{a}{b}",BuiltinMacroKind::Anticommutator=>r"\acomm{a}{b}",BuiltinMacroKind::PoissonBracket=>r"\pb{a}{b}",BuiltinMacroKind::Absolute=>r"\abs{x}",BuiltinMacroKind::Norm=>r"\norm{x}",BuiltinMacroKind::Differential=>r"\dd{x}",BuiltinMacroKind::Bra=>r"\bra{x}",BuiltinMacroKind::Ket=>r"\ket{x}",BuiltinMacroKind::Expectation=>r"\expval{x}",BuiltinMacroKind::BoldVector=>r"\vb{x}",BuiltinMacroKind::ArrowVector=>r"\va{x}",BuiltinMacroKind::UnitVector=>r"\vu{x}",BuiltinMacroKind::MathRelation=>r"\mathrel{x}",BuiltinMacroKind::MathBinary=>r"\mathbin{x}",BuiltinMacroKind::ExtensibleRightArrow=>r"\xrightarrow{f}",BuiltinMacroKind::ExtensibleLeftArrow=>r"\xleftarrow{f}",BuiltinMacroKind::ParenthesizedMod=>r"\pmod{p}"};
  let s=Owned::new(source);layout.layout(&s.s,&fonts,&options).unwrap();}
}

#[test]
fn explicit_atom_class_macros_keep_enum_source_editing_and_spacing_contracts(){
 for (source,kind,label) in [
  (r"\mathrel{⫽}",BuiltinMacroKind::MathRelation,"mathRelation"),
  (r"\mathbin{=}",BuiltinMacroKind::MathBinary,"mathBinary"),
 ] {
  let serialized=serde_json::to_value(kind).unwrap();assert_eq!(serialized,label);
  assert_eq!(serde_json::from_value::<BuiltinMacroKind>(serialized).unwrap(),kind);
  let mut owned=Owned::new(source);let before=owned.s.document_snapshot();let m=macro_node(&before.analysis.tree);
  let NodeKind::MacroCall{macro_kind,arguments,..}=&m.kind else{panic!()};assert_eq!(*macro_kind,kind);
  owned.select(Position::row(arguments[0].id,1));
  let edited=owned.edit(EditCommand::InsertText{text:"x".into(),mode:InputMode::Math}).editor;
  let fresh=Owned::new(&edited.document.source).s.document_snapshot();
  assert_eq!(semantic(serde_json::to_value(&edited.document.analysis.tree).unwrap()),semantic(serde_json::to_value(&fresh.analysis.tree).unwrap()));
  assert_eq!(model::latex(&fresh.analysis.tree,LatexMode::Lossless),edited.document.source);
  assert_eq!(owned.action(EditorAction::Undo).editor.document.source,source);
 }
 let(fonts,options)=font();let mut layout=LayoutEngine::default();
 let rel=layout.layout(&Owned::new(r"a\mathrel{+}b").s,&fonts,&options).unwrap();
 let bin=layout.layout(&Owned::new(r"a\mathbin{+}b").s,&fonts,&options).unwrap();
 assert!(rel.width>bin.width,"Explicit relation spacing must exceed binary spacing");
 for source in [r"\mathrel",r"\mathbin"]{assert!(!Owned::new(source).s.document_snapshot().analysis.exportable);}
}

#[test]
fn missing_arguments_and_unknown_user_macros_remain_nonexportable_without_execution(){
 for source in [r"\pdv{x}",r"\mel{a}{b}",r"\abs",r"\userMacro{rm -rf /}"] {let owned=Owned::new(source);let d=owned.s.document_snapshot();assert_eq!(model::latex(&d.analysis.tree,LatexMode::Lossless),source);assert!(!d.analysis.exportable, "{source}");}
}
