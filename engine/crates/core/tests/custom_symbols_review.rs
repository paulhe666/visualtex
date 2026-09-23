//! P09 immutable legacy-symbol snapshots, native structure and vector geometry.
//! Browser interaction and Native/WASM parity are independent gates.
use std::{collections::BTreeSet,path::PathBuf,sync::Arc};
use serde_json::{json,Value};
use visualtex_engine_core::{custom_symbols::{ArtworkFontStyle,SymbolAsset,SymbolCatalog},editing::{EditorAction,EditorActionResult},fonts::{FontRegistry,PathOp},layout::{CustomSymbolTextFontBinding,LayoutEngine,LayoutOptions,MathStyle},model::{self,Node,NodeKind,LatexMode},positions::{Position,Selection,TreeIndex},session::EngineSession,source_edit::SourceEdit,transactions::{EditCommand,InputMode},ErrorCode};
fn library()->Value{serde_json::from_str(include_str!("../../../fixtures/custom-symbol-foundation.json")).unwrap()}
fn catalog()->Arc<SymbolCatalog>{SymbolCatalog::from_library(&library()).unwrap()}
fn definition()->Value{library()["symbols"][0].clone()}
fn first(tree:&Node)->&Node{tree.children().into_iter().find(|n|matches!(n.kind,NodeKind::CustomSymbol{..})).expect("custom symbol")}
fn ids(tree:&Node)->BTreeSet<u32>{let mut ids=BTreeSet::new();tree.visit(&mut|n|{assert!(ids.insert(n.id.0));if let NodeKind::Grid{columns,..}=&n.kind{for c in columns{assert!(ids.insert(c.0));}}});ids}
fn semantic(v:Value)->Value{match v{Value::Array(a)=>Value::Array(a.into_iter().map(semantic).collect()),Value::Object(m)=>Value::Object(m.into_iter().filter(|(k,_)|!matches!(k.as_str(),"id"|"origin"|"grouping"|"leading"|"trailing")&&!k.starts_with("before")).map(|(k,v)|{let v=if k=="columns"{json!(v.as_array().unwrap().len())}else{semantic(v)};(k,v)}).collect()),v=>v}}
fn state(e:&Value)->Value{json!({"source":e["document"]["source"],"tree":e["document"]["analysis"]["tree"],"assets":e["document"]["analysis"]["symbolAssets"],"selection":e["selection"]})}
fn fonts()->(FontRegistry,LayoutOptions){let p=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));let mut fonts=FontRegistry::default();let id=fonts.register(&std::fs::read(p).unwrap(),0).unwrap().id;(fonts,LayoutOptions::new(id))}
struct Owned{s:EngineSession,seq:u32}
impl Owned{
 fn new(source:&str)->Self{Self{s:EngineSession::from_source_with_symbols("symbols-test".into(),source.into(),catalog()).unwrap(),seq:0}}
 fn empty()->Self{Self{s:EngineSession::from_source("plain-target".into(),String::new()).unwrap(),seq:0}}
 fn snap(&self)->Value{serde_json::to_value(self.s.editor_snapshot().unwrap()).unwrap()}
 fn act(&mut self,action:EditorAction)->EditorActionResult{let before=self.s.document_snapshot();self.seq+=1;let r=self.s.apply_action(self.s.revision(),self.seq,action).unwrap();let index=TreeIndex::new(&r.editor.document.analysis.tree).unwrap();index.validate_selection(&r.editor.selection).unwrap();if let Some(map)=&r.position_map{for p in TreeIndex::new(&before.analysis.tree).unwrap().carets{index.validate_position(map.apply_position(p,&index).unwrap()).unwrap();}}
  self.oracle();r}
 fn edit(&mut self,command:EditCommand)->EditorActionResult{self.act(EditorAction::Edit{command,group_id:None})}
 fn select(&mut self,selection:Selection){self.act(EditorAction::SetSelection{selection});}
 fn reject(&mut self,command:EditCommand,code:ErrorCode){let before=self.snap();self.seq+=1;let error=self.s.apply_action(self.s.revision(),self.seq,EditorAction::Edit{command,group_id:None}).unwrap_err();assert_eq!(error.code,code,"{}",error.message);assert_eq!(self.snap(),before);}
 fn oracle(&self){let d=self.s.document_snapshot();let new=EngineSession::from_source_with_symbols("oracle".into(),d.source.clone(),Arc::clone(&d.analysis.symbol_assets)).unwrap().document_snapshot();assert_eq!(semantic(serde_json::to_value(&d.analysis.tree).unwrap()),semantic(serde_json::to_value(&new.analysis.tree).unwrap()),"{}",d.source);assert_eq!(serde_json::to_value(&d.analysis.facts).unwrap(),serde_json::to_value(&new.analysis.facts).unwrap());assert_eq!(model::latex(&d.analysis.tree,LatexMode::Lossless),d.source);}
}
#[test]
fn legacy_library_freezes_exact_definition_without_following_later_same_name_changes(){
 let mut def=definition();let original=def.clone();let a=SymbolAsset::freeze(def.clone()).unwrap();assert_eq!(a.definition(),&original);assert_eq!(a.sha256(),SymbolAsset::freeze(original.clone()).unwrap().sha256());assert!(a.supported());assert_eq!(a.geometry().len(),3);
 def["artwork"]["shapes"][0]["d"]=json!("M0 0L100 0L0 100Z");let b=SymbolAsset::freeze(def).unwrap();assert_ne!(a.sha256(),b.sha256());assert_eq!(a.definition(),&original);
 let old=catalog();let incoming=SymbolCatalog::freeze(vec![b.definition().clone()]).unwrap();assert_eq!(old.merge(&incoming).unwrap_err().code,ErrorCode::Conflict);
 let data=library();assert_eq!(SymbolCatalog::from_library(&data).unwrap().definitions().len(),3);assert_eq!(data,library());
}
#[test]
fn invalid_commands_metrics_transforms_and_path_input_are_rejected_not_normalized(){
 for command in ["frac","partial","sqrt","begin","input","mathop","hphantom","overset","bad-name",""]{let mut d=definition();d["command"]=json!(command);assert!(SymbolAsset::freeze(d).is_err(),"{command}");}
 for path in ["MNaN 0L0 1","M0 0L1e999 0","M0 0<script>","L0 0","M1","M0 0Z1 2"]{let mut d=definition();d["artwork"]["shapes"]=json!([{"kind":"path","d":path}]);assert!(SymbolAsset::freeze(d).is_err(),"{path}");}
 for value in [json!(-1),json!(65),json!("1")]{let mut d=definition();d["metrics"]["widthEm"]=value;assert!(SymbolAsset::freeze(d).is_err());}
 for transform in [json!({"scaleX":0}),json!({"matrix":[1,1,1,1,0,0]}),json!({"matrix":[1,2]}),json!({"translateX":"oops"})]{let mut d=definition();d["artwork"]["shapes"][0]["transform"]=transform;assert!(SymbolAsset::freeze(d).is_err());}
 assert!(SymbolCatalog::freeze(vec![definition(),definition()]).is_err());
}
#[test]
fn text_artwork_requires_exact_registered_font_binding_and_becomes_outline_only(){
 let mut d=definition();d["artwork"]["shapes"]=json!([{"kind":"text","text":"Ax","fontFamily":"VisualTeX Verified Test","fontSize":900,"fontStyle":"normal","fontWeight":400,"x":40,"y":760,"fill":true,"transform":{"rotateDeg":4,"originX":400,"originY":400}}]);
 let a=SymbolAsset::freeze(d.clone()).unwrap();assert!(a.supported());assert_eq!(a.geometry().len(),1);let unresolved=&a.geometry()[0];assert!(unresolved.path.is_empty());assert!(unresolved.text_source.is_some());assert!(unresolved.outline_font_id.is_none());assert_eq!(a.definition(),&d);
 let c=SymbolCatalog::freeze(vec![d]).unwrap();let s=EngineSession::from_source_with_symbols("bound-text".into(),r"\userglyph+x".into(),c).unwrap();let snap=s.document_snapshot();assert!(snap.analysis.complete&&snap.analysis.supported&&snap.analysis.exportable);
 let(f,mut o)=fonts();let mut layout=LayoutEngine::default();let missing=layout.layout(&s,&f,&o).unwrap_err();assert_eq!(missing.code,ErrorCode::InvalidInput);
 o.custom_symbol_text_fonts.push(CustomSymbolTextFontBinding{font_family:"VisualTeX Verified Test".into(),font_style:ArtworkFontStyle::Italic,font_weight:400,font_id:o.math_font.clone()});assert_eq!(layout.layout(&s,&f,&o).unwrap_err().code,ErrorCode::InvalidInput);
 o.custom_symbol_text_fonts[0].font_style=ArtworkFontStyle::Normal;let scene=layout.layout(&s,&f,&o).unwrap();scene.validate().unwrap();assert_eq!(scene.artworks.len(),1);let shape=&scene.artworks[0].shape;assert!(!shape.path.is_empty());assert!(shape.text_source.is_none());assert_eq!(shape.outline_font_id.as_deref(),Some(o.math_font.as_str()));assert!(scene.fonts.iter().any(|font|font.id==o.math_font));let svg=scene.svg().unwrap();assert!(!svg.contains("<text")&&!svg.contains("font-family")&&!svg.contains("VisualTeX Verified Test"));
 let mut forged=scene.clone();Arc::make_mut(&mut forged.artworks[0].shape).outline_font_id=Some("0".repeat(64)+":0");assert!(forged.validate().is_err());
}
#[test]
fn erase_is_ordered_inside_one_custom_asset_and_repaint_can_restore_local_ink(){
 let mut d=definition();d["artwork"]["shapes"]=json!([
  {"kind":"rect","x":80,"y":80,"width":820,"height":760,"fill":true},
  {"kind":"circle","cx":450,"cy":430,"r":230,"fill":true,"operation":"erase","transform":{"rotateDeg":7,"originX":450,"originY":430}},
  {"kind":"circle","cx":450,"cy":430,"r":70,"fill":true}
 ]);
 let a=SymbolAsset::freeze(d.clone()).unwrap();assert!(a.supported());assert_eq!(a.geometry().len(),3);assert_eq!(a.geometry()[0].operation,visualtex_engine_core::custom_symbols::ArtworkOperation::Paint);assert_eq!(a.geometry()[1].operation,visualtex_engine_core::custom_symbols::ArtworkOperation::Erase);assert_eq!(a.geometry()[2].operation,visualtex_engine_core::custom_symbols::ArtworkOperation::Paint);
 let c=SymbolCatalog::freeze(vec![d]).unwrap();let s=EngineSession::from_source_with_symbols("erase".into(),r"\userglyph+x".into(),c).unwrap();let(f,o)=fonts();let scene=LayoutEngine::default().layout(&s,&f,&o).unwrap();scene.validate().unwrap();assert_eq!(scene.artworks.len(),3);let owner=scene.artworks[0].node_id;assert!(scene.artworks.iter().all(|x|x.node_id==owner));
 let svg=scene.svg().unwrap();assert_eq!(svg.matches("<mask ").count(),1);let start=svg.find("<mask ").unwrap();let end=start+svg[start..].find("</mask>").unwrap();let mask=&svg[start..end];let white1=mask.find("#ffffff").unwrap();let black=mask[white1+1..].find("#000000").map(|i|i+white1+1).unwrap();let white2=mask[black+1..].find("#ffffff").map(|i|i+black+1).unwrap();assert!(white1<black&&black<white2);assert!(svg.contains("mask=\"url(#vt-art-mask-0)\""));
 let positions:Vec<_>=scene.paint_order.iter().enumerate().filter_map(|(i,p)|matches!(p,visualtex_engine_core::artwork_scene::PaintItem::Artwork{..}).then_some(i)).collect();assert_eq!(positions.len(),3);assert_eq!(positions,[positions[0],positions[0]+1,positions[0]+2]);
}
#[test]
fn erase_group_validation_rejects_cross_glyph_reentry_or_mismatched_instance(){
 let mut d=definition();d["artwork"]["shapes"]=json!([
  {"kind":"rect","x":80,"y":80,"width":820,"height":760,"fill":true},
  {"kind":"circle","cx":450,"cy":430,"r":230,"fill":true,"operation":"erase"},
  {"kind":"circle","cx":450,"cy":430,"r":70,"fill":true}
 ]);let c=SymbolCatalog::freeze(vec![d]).unwrap();let s=EngineSession::from_source_with_symbols("erase-isolation".into(),r"\userglyph+x".into(),c).unwrap();let(f,o)=fonts();let scene=LayoutEngine::default().layout(&s,&f,&o).unwrap();scene.validate().unwrap();
 let art:Vec<_>=scene.paint_order.iter().enumerate().filter_map(|(i,p)|matches!(p,visualtex_engine_core::artwork_scene::PaintItem::Artwork{..}).then_some(i)).collect();let glyph=scene.paint_order.iter().enumerate().find_map(|(i,p)|matches!(p,visualtex_engine_core::artwork_scene::PaintItem::Glyph{..}).then_some(i)).unwrap();assert!(glyph>art[2]);
 let mut crossed=scene.clone();crossed.paint_order.swap(art[1],glyph);assert!(crossed.validate().is_err());
 let mut shifted=scene.clone();shifted.artworks[1].x+=0.01;assert!(shifted.validate().is_err());
 let mut wrong_index=scene.clone();wrong_index.artworks[1].path_index=9;assert!(wrong_index.validate().is_err());
}
#[test]
fn clip_rect_sits_between_outer_user_transform_and_inner_import_matrix(){
 let mut d=definition();d["artwork"]["shapes"]=json!([{"kind":"path","d":"M0 0H1000V1000H0Z","fill":true,"clipRect":{"x":200,"y":100,"width":400,"height":500},"transform":{"translateX":40,"translateY":20,"scaleX":1.1,"scaleY":0.9,"matrix":[1,0,0,1,100,0]}}]);
 let a=SymbolAsset::freeze(d.clone()).unwrap();assert!(a.supported());assert_eq!(a.geometry().len(),1);let g=&a.geometry()[0];let clip=g.clip.expect("clip");
 assert!((clip.x_min-0.2).abs()<1e-9&&(clip.x_max-0.6).abs()<1e-9);assert!((g.inner_transform[4]-0.1).abs()<1e-9);assert!((g.transform[0]-1.1).abs()<1e-9);assert!(g.ink.x_min>=0.25&&g.ink.x_max<=0.71);
 let c=SymbolCatalog::freeze(vec![d]).unwrap();let s=EngineSession::from_source_with_symbols("clip".into(),r"\userglyph".into(),c).unwrap();let(f,o)=fonts();let scene=LayoutEngine::default().layout(&s,&f,&o).unwrap();scene.validate().unwrap();let svg=scene.svg().unwrap();
 let outer=svg.find("<g transform=\"").unwrap();let clipped=svg.find("<g clip-path=\"").unwrap();let path=svg[clipped..].find("<path d=\"").map(|i|i+clipped).unwrap();assert!(outer<clipped&&clipped<path);assert!(svg.contains("<rect x=\"0.2\"")&&svg.contains("transform=\"matrix(1 0 0 1 0.1 0)\""));
}
#[test]
fn general_strokes_keep_svg_stroke_semantics_through_affine_transform(){
 for shape in [
  json!({"kind":"path","d":"M80 700Q450 40 820 700","fill":false,"strokeWidth":36,"lineCap":"round","lineJoin":"bevel","transform":{"scaleX":1.3,"scaleY":0.7,"skewXDeg":9,"rotateDeg":12,"originX":450,"originY":400}}),
  json!({"kind":"ellipse","cx":450,"cy":420,"rx":320,"ry":140,"fill":false,"strokeWidth":28,"lineJoin":"round","transform":{"scaleX":0.8,"scaleY":1.25}}),
  json!({"kind":"polygon","points":[[100,700],[450,120],[820,700]],"fill":false,"strokeWidth":30,"lineJoin":"bevel"}),
  json!({"kind":"rect","x":100,"y":120,"width":700,"height":620,"rx":120,"ry":80,"fill":false,"strokeWidth":24,"lineJoin":"round"})
 ]{
  let mut d=definition();d["artwork"]["shapes"]=json!([shape]);let a=SymbolAsset::freeze(d.clone()).unwrap();assert!(a.supported());assert_eq!(a.geometry().len(),1);let g=&a.geometry()[0];assert!(!g.fill&&g.stroke_width>0.);assert!(g.ink.valid());assert_eq!(a.definition(),&d);
  let c=SymbolCatalog::freeze(vec![d]).unwrap();let s=EngineSession::from_source_with_symbols("stroke".into(),r"\userglyph".into(),c).unwrap();let(f,o)=fonts();let scene=LayoutEngine::default().layout(&s,&f,&o).unwrap();let svg=scene.svg().unwrap();assert!(svg.contains("stroke-width=")&&svg.contains("matrix(")&&svg.contains("fill=\"none\""));assert!(!svg.contains("<text")&&!svg.contains("foreignObject"));
 }
}
#[test]
fn custom_insert_edit_duplicate_source_patch_and_history_share_real_node_ownership(){
 let mut t=Owned::new("x");t.edit(EditCommand::InsertCustomSymbol{command:r"\userglyph".into()});let inserted=t.snap();let custom=inserted["document"]["analysis"]["tree"]["entries"][1]["node"]["id"].as_u64().unwrap() as u32;
 t.select(Selection::Node{node_id:visualtex_engine_core::session::NodeId(custom)});let before=t.snap();t.edit(EditCommand::DuplicateSelection);let d=t.s.document_snapshot();let children=d.analysis.tree.children();assert!(ids(children[1]).is_disjoint(&ids(children[2])));let after=t.snap();
 t.act(EditorAction::Undo);assert_eq!(state(&t.snap()),state(&before));t.act(EditorAction::Redo);assert_eq!(state(&t.snap()),state(&after));
 let asset=t.s.document_snapshot().analysis.symbol_assets.get("userglyph").unwrap();let old_sha=asset.sha256().to_owned();t.s.patch_source(t.s.revision(),SourceEdit{start_utf16:0,end_utf16:1,replacement:"q".into()}).unwrap();t.oracle();assert_eq!(t.s.document_snapshot().analysis.symbol_assets.get("userglyph").unwrap().sha256(),old_sha);
 t.s.replace_source(t.s.revision(),r"\frac{\userglyph}{\userrel}".into()).unwrap();t.oracle();let d=t.s.document_snapshot();let NodeKind::Fraction{numerator,..}=&d.analysis.tree.children()[0].kind else{panic!()};t.select(Selection::caret(Position::row(numerator.id,1)));t.edit(EditCommand::InsertText{text:"z".into(),mode:InputMode::Math});assert_eq!(t.s.source(),r"\frac{\userglyph z}{\userrel}");
}
#[test]
fn custom_operator_limits_are_editable_and_override_frozen_policy_without_mutating_asset(){
 let mut t=Owned::new(r"\userop_i^n x");let d=t.s.document_snapshot();let op=first(&d.analysis.tree);let id=op.id;let NodeKind::CustomSymbol{asset,parts,..}=&op.kind else{panic!()};let sha=asset.sha256().to_owned();assert_eq!(parts.len(),2);
 t.select(Selection::Node{node_id:id});let before=t.snap();t.edit(EditCommand::SetOperatorLimits{mode:visualtex_engine_core::operators::LimitMode::NoLimits});assert!(t.s.source().contains(r"\nolimits"));t.act(EditorAction::Undo);assert_eq!(state(&before),state(&t.snap()));
 let (f,mut o)=fonts();let mut l=LayoutEngine::default();let display=l.layout(&t.s,&f,&o).unwrap();o.style=MathStyle::Text;let inline=l.layout(&t.s,&f,&o).unwrap();assert!(display.height>inline.height);assert_eq!(t.s.document_snapshot().analysis.symbol_assets.get("userop").unwrap().sha256(),sha);
 t.select(Selection::Node{node_id:id});t.edit(EditCommand::InsertStructure{structure:visualtex_engine_core::transactions::StructureKind::Superscript});t.edit(EditCommand::InsertText{text:"2".into(),mode:InputMode::Math});t.oracle();
}
#[test]
fn custom_svg_uses_nonfont_resources_ordered_paint_and_same_natural_geometry(){
 let mut t=Owned::new(r"\textcolor{red}{\userglyph}\!\textcolor{blue}{\userrel}+\userop_i^n");let (f,mut o)=fonts();let mut l=LayoutEngine::default();let a=l.layout(&t.s,&f,&o).unwrap();assert!(!a.artworks.is_empty());assert_eq!(a.paint_order.len(),a.artworks.len()+a.glyphs.len()+a.rules.len());assert!(a.resources.iter().all(|r|r.shape.glyph_id>0));assert!(a.artworks.iter().any(|a|a.color.is_some()));let svg=a.svg().unwrap();for bad in ["<text","<image","<mask","foreignObject","href=","userglyph","assetSha256"]{assert!(!svg.contains(bad),"{bad}");}
 o.font_size_px=32.;let b=l.layout(&t.s,&f,&o).unwrap();assert_eq!(b.width,a.width*2.);assert_eq!(b.height,a.height*2.);assert_eq!(serde_json::to_value(&a.artworks).unwrap(),serde_json::to_value(&b.artworks).unwrap());
 t.s.replace_source(t.s.revision(),r"\phantom{\userglyph}x".into()).unwrap();let hidden=l.layout(&t.s,&f,&o).unwrap();assert!(hidden.artworks.is_empty());assert!(hidden.paint_order.is_empty());
}
#[test]
fn rich_clipboard_transfers_frozen_symbol_and_undo_restores_catalog_and_ids(){
 let mut source=Owned::new(r"\userglyph+\userop_i^n");let root=Arc::clone(&source.s.document_snapshot().analysis.tree);source.select(Selection::Range{anchor:Position::row(root.id,0),focus:Position::row(root.id,3)});let before=source.snap();let clip=source.s.copy_selection(source.s.revision()).unwrap();assert_eq!(source.snap(),before);let structure=clip.structure.unwrap();assert_eq!(structure.assets.len(),2);assert!(structure.assets.iter().all(|a|a.definition.is_some()));
 let mut target=Owned::empty();let blank=target.snap();target.edit(EditCommand::PasteStructured{latex:clip.latex.clone().unwrap(),fragment:structure.clone()});let pasted=target.snap();assert_eq!(pasted["document"]["source"],source.s.source());assert_eq!(target.s.document_snapshot().analysis.facts.custom_symbols,2);assert_eq!(target.s.document_snapshot().analysis.symbol_assets.definitions().len(),2);
 target.act(EditorAction::Undo);assert_eq!(state(&target.snap()),state(&blank));target.act(EditorAction::Redo);assert_eq!(state(&target.snap()),state(&pasted));
 let mut changed=definition();changed["name"]=json!("different snapshot");let conflict=SymbolCatalog::freeze(vec![changed]).unwrap();let mut other=Owned{s:EngineSession::from_source_with_symbols("conflict".into(),"x".into(),conflict).unwrap(),seq:0};other.reject(EditCommand::PasteStructured{latex:clip.latex.clone().unwrap(),fragment:structure.clone()},ErrorCode::Conflict);
 let mut invalid=structure;invalid.assets[0].sha256="0".repeat(64);target.reject(EditCommand::PasteStructured{latex:clip.latex.unwrap(),fragment:invalid},ErrorCode::InvalidInput);
}
#[test]
fn missing_symbol_assets_cannot_render_or_be_silently_invented(){
 let plain=EngineSession::from_source("plain".into(),r"\userglyph".into()).unwrap();assert!(!plain.document_snapshot().analysis.supported);
 let mut target=Owned::empty();target.reject(EditCommand::PasteLatex{latex:r"\userglyph".into()},ErrorCode::UnsupportedSyntax);
 target.reject(EditCommand::InsertCustomSymbol{command:r"\userglyph".into()},ErrorCode::UnsupportedSyntax);
 let a=Owned::new(r"\userglyph");let tree=Arc::clone(&a.s.document_snapshot().analysis.tree);assert!(visualtex_engine_core::projection::project_tree(tree,20).is_err());
}
#[test]
fn artwork_scene_validation_rejects_bad_identity_order_nonfinite_and_clipped_geometry(){
 let t=Owned::new(r"\userglyph+x");let (f,o)=fonts();let a=LayoutEngine::default().layout(&t.s,&f,&o).unwrap();a.validate().unwrap();
 let mut b=a.clone();b.artworks[0].asset_sha256="bad".into();assert!(b.validate().is_err());
 let mut b=a.clone();b.paint_order.pop();assert!(b.validate().is_err());
 let mut b=a.clone();b.artworks[0].x=f64::NAN;assert!(b.validate().is_err());
 let mut b=a.clone();b.artworks[0].x+=100.;assert!(b.validate().is_err());
 let mut b=a.clone();b.artworks[0].node_id=b.glyphs[0].node_id;assert!(b.validate().is_err());
 let mut b=a.clone();Arc::make_mut(&mut b.artworks[0].shape).path[0]=PathOp::Move([f64::INFINITY,0.]);assert!(b.validate().is_err());
}
#[test]
fn every_supported_path_command_and_transform_keeps_finite_identity_free_geometry(){
 for path in ["M20 20h100v100h-100z","m20 20 100 0 0 100z","M10 30C20 0 40 0 50 30S80 60 90 30Z","M1e1 30Q30 0 50 30T90 30Z","M100 400A300 160 25 0 1 800 400Z","M100 400a300 160 25 1 0 700 0Z"]{
  let mut d=definition();d["artwork"]["shapes"]=json!([{"kind":"path","d":path,"transform":{"translateX":200,"translateY":100,"scaleX":1.2,"scaleY":0.8,"skewXDeg":5,"rotateDeg":15,"originX":50,"originY":30,"matrix":[1,0,0,1,20,30]}}]);let a=SymbolAsset::freeze(d.clone()).unwrap();assert!(a.supported());assert_eq!(a.definition(),&d);assert!(!a.geometry().is_empty());assert!(a.geometry()[0].ink.valid());
 }
}
#[test]
fn rounded_rect_and_arc_flags_compile_as_finite_cubic_geometry(){
 for shape in [json!({"kind":"rect","x":80,"y":100,"width":760,"height":650,"rx":120,"ry":80,"fill":true}),json!({"kind":"rect","x":80,"y":100,"width":760,"height":650,"rx":100,"fill":true})]{let mut d=definition();d["artwork"]["shapes"]=json!([shape]);let a=SymbolAsset::freeze(d).unwrap();assert!(a.supported());assert!(a.geometry()[0].path.iter().any(|op|matches!(op,PathOp::Cubic(_))));assert!(a.geometry()[0].ink.valid());}
 for path in ["M0 0A20 20 0 2 0 40 0","M0 0A20 20 0 0 3 40 0"]{let mut d=definition();d["artwork"]["shapes"]=json!([{"kind":"path","d":path}]);assert!(SymbolAsset::freeze(d).is_err());}
}
