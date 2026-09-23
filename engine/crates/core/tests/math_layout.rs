use std::path::PathBuf;
use visualtex_engine_core::fonts::FontRegistry;
use visualtex_engine_core::layout::{LayoutEngine,LayoutOptions,MathStyle};
use visualtex_engine_core::model::{NodeKind,ScriptKind};
use visualtex_engine_core::positions::{Position,TreeIndex};
use visualtex_engine_core::session::{EngineSession,Revision};
use visualtex_engine_core::source_edit::SourceEdit;
use visualtex_engine_core::ErrorCode;

fn setup()->(FontRegistry,LayoutOptions) {
    let p=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    let mut fonts=FontRegistry::default();let id=fonts.register(&std::fs::read(p).unwrap(),0).unwrap();
    (fonts,LayoutOptions::new(id.id))
}
fn session(source:&str)->EngineSession {EngineSession::from_source("layout-test".into(),source.into()).unwrap()}
fn close(a:f64,b:f64) {assert!((a-b).abs()<1e-9,"{a} != {b}");}

#[test]
fn a_scene_is_nonmutating_and_shares_the_document_identity_and_revision() {
    let (fonts,options)=setup();let s=session(r"\frac{x_1+1}{y^2}");let before=serde_json::to_value(s.editor_snapshot().unwrap()).unwrap();
    let scene=LayoutEngine::default().layout(&s,&fonts,&options).unwrap();
    scene.validate().unwrap();assert_eq!(scene.revision,s.revision());assert_eq!(scene.document_id,s.document_id());
    assert_eq!(serde_json::to_value(s.editor_snapshot().unwrap()).unwrap(),before);
    let doc=s.document_snapshot();let index=TreeIndex::new(&doc.analysis.tree).unwrap();
    for c in &scene.carets {index.validate_position(c.position).unwrap();}
    assert!(!scene.glyphs.is_empty() && scene.rules.len()==1);
}
#[test]
fn fraction_rule_axis_and_gaps_are_font_metric_constraints() {
    let (fonts,options)=setup();let s=session(r"\frac{x}{y}");let snap=s.document_snapshot();
    let frac=&snap.analysis.tree.children()[0];
    let NodeKind::Fraction {numerator,denominator,..}=&frac.kind else {panic!()};
    let scene=LayoutEngine::default().layout(&s,&fonts,&options).unwrap();
    let math=fonts.get(&options.math_font).unwrap().math.as_ref().unwrap();
    let rule=&scene.rules[0].rect;
    close(rule.height(),math.fraction_rule);close((rule.y_min+rule.y_max)/2.0,scene.baseline_em-math.axis);
    let n=scene.regions.iter().find(|r|r.node_id==numerator.id).unwrap().ink.unwrap();
    let d=scene.regions.iter().find(|r|r.node_id==denominator.id).unwrap().ink.unwrap();
    assert!(rule.y_min-n.y_max+1e-9>=math.numerator_display_gap);
    assert!(d.y_min-rule.y_max+1e-9>=math.denominator_display_gap);
}
#[test]
fn scripts_shrink_at_most_to_scriptscript_and_do_not_overlap() {
    let (fonts,options)=setup();let s=session(r"x_{y}^{z}");let snap=s.document_snapshot();
    let script=&snap.analysis.tree.children()[0];let NodeKind::Scripts {slots,..}=&script.kind else {panic!()};
    let sup=slots.iter().find(|s|s.script==ScriptKind::Superscript).unwrap().argument.id;
    let sub=slots.iter().find(|s|s.script==ScriptKind::Subscript).unwrap().argument.id;
    let scene=LayoutEngine::default().layout(&s,&fonts,&options).unwrap();
    let math=fonts.get(&options.math_font).unwrap().math.as_ref().unwrap();
    let top=scene.regions.iter().find(|r|r.node_id==sup).unwrap().ink.unwrap();
    let bottom=scene.regions.iter().find(|r|r.node_id==sub).unwrap().ink.unwrap();
    assert!(bottom.y_min-top.y_max+1e-9>=math.sub_sup_gap);
    let nested=LayoutEngine::default().layout(&session(r"x^{y^{z^2}}"),&fonts,&options).unwrap();
    for g in &nested.glyphs {assert!(g.scale+1e-9>=math.scriptscript_scale);}
    assert!(nested.glyphs.iter().any(|g|(g.scale-math.scriptscript_scale).abs()<1e-9));
}
#[test]
fn font_size_changes_only_natural_dimensions_and_reuses_em_geometry() {
    let (fonts,mut options)=setup();let s=session(r"\frac{x^2}{y_1}");let mut layout=LayoutEngine::default();
    let a=layout.layout(&s,&fonts,&options).unwrap();options.font_size_px*=2.0;
    let b=layout.layout(&s,&fonts,&options).unwrap();
    close(a.width*2.0,b.width);close(a.height*2.0,b.height);close(a.baseline*2.0,b.baseline);
    close(a.width_em,b.width_em);assert_eq!(b.rebuilt_fragments,0);assert!(b.reused_fragments>0);
    assert_eq!(serde_json::to_value(&a.glyphs).unwrap(),serde_json::to_value(&b.glyphs).unwrap());
}
#[test]
fn source_edits_rebuild_ancestors_but_reuse_unrelated_geometry() {
    let (fonts,options)=setup();let mut s=session(r"\frac{x}{2}+\frac{x}{2}");let mut layout=LayoutEngine::default();
    let a=layout.layout(&s,&fonts,&options).unwrap();
    s.patch_source(Revision(0),SourceEdit {start_utf16:6,end_utf16:7,replacement:"z".into()}).unwrap();
    let b=layout.layout(&s,&fonts,&options).unwrap();
    assert!(b.reused_fragments>0);assert!(b.rebuilt_fragments<a.rebuilt_fragments);
    assert_eq!(b.revision,Revision(1));assert_ne!(a.svg().unwrap(),b.svg().unwrap());
    assert_eq!(a.svg().unwrap(),LayoutEngine::default().layout(&session(r"\frac{x}{2}+\frac{x}{2}"),&fonts,&options).unwrap().svg().unwrap());
}
#[test]
fn svg_has_only_vectors_and_contains_all_glyph_ink() {
    let (fonts,options)=setup();let scene=LayoutEngine::default().layout(&session(r"\frac{f^{g_y}}{j_2}"),&fonts,&options).unwrap();
    let svg=scene.svg().unwrap();assert!(svg.starts_with("<svg ") && svg.ends_with("</svg>"));
    for disallowed in ["<text","<image","<script","foreignObject","href=","NaN","Infinity"] {assert!(!svg.contains(disallowed),"{disallowed}");}
    for g in &scene.glyphs {if let Some([xmin,ymin,xmax,ymax])=scene.resources[g.resource].shape.ink {
        assert!(xmin*g.scale+g.x>=-1e-9);assert!(ymin*g.scale+g.y>=-1e-9);
        assert!(xmax*g.scale+g.x<=scene.width_em+1e-9);assert!(ymax*g.scale+g.y<=scene.height_em+1e-9);
    }}
}
#[test]
fn empty_slots_are_editable_but_never_exported_as_finished_formulas() {
    let (fonts,options)=setup();let scene=LayoutEngine::default().layout(&session(r"\frac{}{}"),&fonts,&options).unwrap();
    assert!(!scene.export_ready);assert!(scene.carets.len()>=4);
    assert_eq!(scene.svg().unwrap_err().code,ErrorCode::IncompleteSource);
}
#[test]
fn unsupported_drafts_and_missing_text_fonts_fail_without_data_loss() {
    let (fonts,options)=setup();let mut layout=LayoutEngine::default();
    // The former accent rejection fixture is now supported and kept below.
    // Unknown commands, incomplete drafts and missing text fonts still fail.
    for source in [r"\frac{x}{",r"\futureaccent{x}",r"\text{速度}"] {
        let s=session(source);let before=s.snapshot();assert!(layout.layout(&s,&fonts,&options).is_err());assert_eq!(s.snapshot(),before);
    }
    assert!(layout.layout(&session("x"),&fonts,&options).is_ok());
    let accent=layout.layout(&session(r"\hat{x}"),&fonts,&options).unwrap();
    assert!(accent.export_ready);accent.validate().unwrap();
    assert!(accent.glyphs.len()>=2);
}
#[test]
fn invalid_sizes_and_hit_coordinates_are_rejected() {
    let (fonts,mut options)=setup();let s=session("x");let mut layout=LayoutEngine::default();
    for size in [f64::NAN,f64::INFINITY,0.0,-1.0,1e20] {options.font_size_px=size;assert!(layout.layout(&s,&fonts,&options).is_err());}
    options.font_size_px=16.0;let scene=layout.layout(&s,&fonts,&options).unwrap();
    assert!(scene.hit_test(f64::NAN,0.0).is_err());assert!(scene.hit_test(0.0,f64::INFINITY).is_err());
}
#[test]
fn text_carets_are_grapheme_based_and_come_from_shaping() {
    let (fonts,options)=setup();let s=session("\\text{Cafe\u{301}}");let doc=s.document_snapshot();let text=doc.analysis.tree.children()[0].id;
    let scene=LayoutEngine::default().layout(&s,&fonts,&options).unwrap();
    let carets:Vec<_>=scene.carets.iter().filter(|c|matches!(c.position,Position::Text {node_id,..} if node_id==text)).collect();
    assert_eq!(carets.len(),5);assert!(carets.windows(2).all(|p|p[0].x<=p[1].x));
    assert!(scene.caret(Position::text(text,4)).is_some());assert!(scene.caret(Position::text(text,5)).is_none());
}
#[test]
fn display_and_text_fractions_obey_distinct_math_styles() {
    let (fonts,mut options)=setup();let mut layout=LayoutEngine::default();
    let a=layout.layout(&session(r"\frac{x}{y}"),&fonts,&options).unwrap();options.style=MathStyle::Text;
    let b=layout.layout(&session(r"\frac{x}{y}"),&fonts,&options).unwrap();
    assert!(a.height>b.height);assert!(a.glyphs.iter().all(|g|(g.scale-1.0).abs()<1e-9));
    assert!(b.glyphs.iter().all(|g|g.scale<1.0));
    let forced=layout.layout(&session(r"\dfrac{x}{y}"),&fonts,&options).unwrap();close(forced.height,a.height);
}
#[test]
fn seeded_nested_formulas_produce_valid_geometry_and_reusable_snapshots() {
    let (fonts,options)=setup();let mut layout=LayoutEngine::default();let mut seed=0x50452026u32;
    let mut source="x".to_owned();
    for i in 0..100 {seed=seed.wrapping_mul(1664525).wrapping_add(1013904223);if i%8==0 {source="x".into();}
        source=match seed%4 {0=>format!("\\frac{{{source}}}{{y_2}}"),1=>format!("{{{source}}}^{{z}}"),2=>format!("{source}+\\text{{AV}}"),_=>format!("{source}=y")};
        let s=session(&source);let a=layout.layout(&s,&fonts,&options).unwrap();a.validate().unwrap();assert!(a.svg().unwrap().len()>50);
        let b=layout.layout(&s,&fonts,&options).unwrap();assert_eq!(a.svg().unwrap(),b.svg().unwrap());assert_eq!(b.rebuilt_fragments,0);
    }
}
