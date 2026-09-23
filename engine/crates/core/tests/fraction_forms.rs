//! P07 fraction-family vertical slice, independent of browser/Office execution.
use std::{collections::BTreeSet, path::PathBuf, sync::Arc};
use serde_json::{json, Value};
use visualtex_engine_core::{
    editing::{EditorAction, EditorActionResult}, error::ErrorCode, fonts::FontRegistry,
    layout::{LayoutEngine, LayoutOptions, MathStyle},
    model::{self, IdAllocator, LatexMode, Node, NodeKind}, parser::{self, ParseResult},
    positions::{Horizontal, Position, Selection, TreeIndex}, projection::project_tree,
    session::{EngineSession, NodeId}, source::SourceText, source_edit::SourceEdit,
    transactions::{EditCommand, InputMode, StructureKind},
};
const FORMS: &[&str] = &[r"\binom", r"\dbinom", r"\tbinom", r"\nicefrac"];
fn parse(source: &str) -> ParseResult {
    parser::parse(SourceText::new(1, source).unwrap(), NodeId(1), &mut IdAllocator::new(2)).unwrap()
}
fn semantic(node: &Node) -> Value {
    fn strip(value: Value) -> Value {
        match value {
            Value::Object(map) => Value::Object(map.into_iter()
                .filter(|(k,_)| !matches!(k.as_str(), "id"|"origin"|"grouping"|"leading"|"trailing") && !k.starts_with("before"))
                .map(|(k,v)|(k,strip(v))).collect()),
            Value::Array(a) => Value::Array(a.into_iter().map(strip).collect()),
            other => other,
        }
    }
    strip(serde_json::to_value(node).unwrap())
}
fn ids(node: &Node) -> BTreeSet<NodeId> {
    let mut result = BTreeSet::new(); node.visit(&mut |n| assert!(result.insert(n.id))); result
}
fn fraction(tree: &Node) -> (&Node, &Node, &Node) {
    if let NodeKind::Fraction {numerator,denominator,..} = &tree.kind { return (tree,numerator,denominator); }
    for child in tree.children() {
        let mut has=false; child.visit(&mut |n|has |= matches!(n.kind,NodeKind::Fraction{..}));
        if has {return fraction(child);}
    }
    panic!("No fraction in tree")
}
fn resources() -> (FontRegistry, LayoutOptions) {
    let root=PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../..");
    let path=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from)
        .unwrap_or_else(||root.join("apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    let mut fonts=FontRegistry::default();
    let mut options=LayoutOptions::new(fonts.register(&std::fs::read(path).unwrap(),0).unwrap().id);
    let text=std::env::var_os("VISUALTEX_TEST_TEXT_FONT").map(PathBuf::from)
        .unwrap_or_else(||PathBuf::from("C:/Windows/Fonts/NotoSansSC-VF.ttf"));
    options.text_font=Some(fonts.register(&std::fs::read(text).unwrap(),0).unwrap().id);
    (fonts,options)
}
struct Owned { session: EngineSession, sequence: u32 }
impl Owned {
    fn new(source:&str)->Self {Self{session:EngineSession::from_source("fraction-test".into(),source.into()).unwrap(),sequence:0}}
    fn act(&mut self,action:EditorAction)->EditorActionResult {
        let before=self.session.document_snapshot(); self.sequence+=1;
        let result=self.session.apply_action(self.session.revision(),self.sequence,action).unwrap();
        let tree=&result.editor.document.analysis.tree;
        let index=TreeIndex::new(tree).unwrap();index.validate_selection(&result.editor.selection).unwrap();
        if let Some(map)=&result.position_map {for p in TreeIndex::new(&before.analysis.tree).unwrap().carets {map.apply_position(p,&index).unwrap();}}
        assert_eq!(ids(tree).len(),result.editor.document.analysis.facts.nodes);
        assert_eq!(semantic(tree),semantic(&parse(self.session.source()).tree),"{}",self.session.source());
        result
    }
    fn edit(&mut self,command:EditCommand)->EditorActionResult {self.act(EditorAction::Edit{command,group_id:None})}
    fn text(&mut self,text:&str) {self.edit(EditCommand::InsertText{text:text.into(),mode:InputMode::Math});}
    fn right(&mut self) {self.act(EditorAction::MoveHorizontal{direction:Horizontal::Right,extend:false});}
    fn select(&mut self,selection:Selection) {self.act(EditorAction::SetSelection{selection});}
}

#[test]
fn authored_fraction_forms_keep_commands_arguments_comments_and_roundtrip() {
    for command in FORMS {for args in ["{n}{k}"," nk", " % numerator\r\n {n_1} % denominator\n {k+1}","{\\frac{n}{2}}{\\sqrt{k}}"] {
        let source=format!("{command}{args}+\\alpha x");let p=parse(&source);
        assert!(p.complete&&p.supported&&p.exportable,"{source}: {:?}",p.diagnostics);
        assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
        let canonical=model::latex(&p.tree,LatexMode::Canonical);let q=parse(&canonical);
        assert!(q.complete&&q.supported);assert_eq!(semantic(&p.tree),semantic(&q.tree),"{canonical}");
        assert_eq!(model::latex(&q.tree,LatexMode::Canonical),canonical);
    }}
}
#[test]
fn incomplete_unknown_and_unfilled_fraction_source_is_not_repaired_or_exported() {
    for command in FORMS {for tail in ["","{","{n}","{n}{"] {
        let source=format!("{command}{tail}");let p=parse(&source);
        assert!(!p.complete&&!p.exportable,"{source}");assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
    }}
    for command in FORMS {for tail in ["{}{k}","{n}{}","{{}}{k}"] {
        let p=parse(&format!("{command}{tail}"));assert!(p.complete&&p.supported&&!p.exportable);
    }}
    for source in [r"\futurefrac{a}{b}",r"\genfrac{(}{)}{0pt}{}{n}{k}",r"{n\choose k}"] {
        let p=parse(source);assert!(!p.supported&&!p.exportable);assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
    }
    let p=parse(r"\frac\binom nk");assert!(!p.complete);
}
#[test]
fn binomial_and_skewed_fraction_build_by_transactions_and_navigate_real_slots() {
    for (structure,command) in [(StructureKind::Binomial,r"\binom"),(StructureKind::SkewedFraction,r"\nicefrac")] {
        let mut t=Owned::new("");t.edit(EditCommand::InsertStructure{structure});t.text("n");t.right();t.text("k");t.right();t.text("+x");
        assert_eq!(t.session.source(),format!("{command}{{n}}{{k}}+x"));
        let before=t.session.editor_snapshot().unwrap();t.act(EditorAction::Undo);t.act(EditorAction::Redo);
        let after=t.session.editor_snapshot().unwrap();assert_eq!(before.document.source,after.document.source);
        assert_eq!(before.selection,after.selection);assert_eq!(ids(&before.document.analysis.tree),ids(&after.document.analysis.tree));
    }
}
#[test]
fn wrapping_an_existing_selection_preserves_numerator_ids_and_targets_denominator() {
    for structure in [StructureKind::Binomial,StructureKind::SkewedFraction] {
        let mut t=Owned::new("n+1");let before=t.session.document_snapshot();
        let children=ids(&before.analysis.tree).into_iter().filter(|id|*id!=before.analysis.tree.id).collect::<BTreeSet<_>>();
        t.select(Selection::Range{anchor:Position::row(before.analysis.tree.id,0),focus:Position::row(before.analysis.tree.id,3)});
        let result=t.edit(EditCommand::InsertStructure{structure});let (_,num,den)=fraction(&result.editor.document.analysis.tree);
        assert!(children.is_subset(&ids(num)));assert_eq!(result.editor.selection,Selection::caret(Position::row(den.id,0)));
        t.text("k");assert!(t.session.document_snapshot().analysis.exportable);
    }
}
#[test]
fn duplicate_edit_delete_and_undo_keep_two_fractions_independent() {
    for command in FORMS {
        let mut t=Owned::new(&format!("{command}{{n}}{{k}}"));let initial=t.session.document_snapshot();
        let original=fraction(&initial.analysis.tree).0.id;t.select(Selection::Node{node_id:original});
        let duplicated=t.edit(EditCommand::DuplicateSelection);
        let NodeKind::Row{entries,..}=&duplicated.editor.document.analysis.tree.kind else{panic!()};
        let a=&entries[0].node;let b=&entries[1].node;
        assert!(ids(a).is_disjoint(&ids(b)));let copy_id=b.id;let (_,num,_)=fraction(b);
        t.select(Selection::caret(Position::row(num.id,1)));t.text("+1");
        assert_eq!(t.session.source(),format!("{command}{{n}}{{k}}{command}{{n+1}}{{k}}"));
        t.select(Selection::Node{node_id:copy_id});let removed=t.edit(EditCommand::DeleteForward);
        assert_eq!(removed.editor.document.source,initial.source);assert_eq!(fraction(&removed.editor.document.analysis.tree).0.id,original);
        t.act(EditorAction::Undo);assert!(t.session.document_snapshot().analysis.tree.find(copy_id).is_some());
    }
}
#[test]
fn source_patch_updates_actual_numerator_but_retains_unaffected_duplicate() {
    let source=r"\binom{x}{k}+\binom{x}{k} % keep";
    let mut session=EngineSession::from_source("patch".into(),source.into()).unwrap();
    let before=session.document_snapshot();let NodeKind::Row{entries,..}=&before.analysis.tree.kind else{panic!()};
    let other=ids(&entries[2].node);let at=source.find('x').unwrap() as u32;
    let (after,_)=session.patch_source(session.revision(),SourceEdit{start_utf16:at,end_utf16:at+1,replacement:"y+1".into()}).unwrap();
    assert_eq!(after.source,r"\binom{y+1}{k}+\binom{x}{k} % keep");
    assert_eq!(semantic(&after.analysis.tree),semantic(&parse(&after.source).tree));
    let NodeKind::Row{entries,..}=&after.analysis.tree.kind else{panic!()};assert_eq!(ids(&entries[2].node),other);
    assert_eq!(model::latex(fraction(&entries[0].node).1,LatexMode::Canonical),"{y+1}");
}
#[test]
fn binomial_uses_stack_metrics_without_painting_a_fraction_rule() {
    let (fonts,mut options)=resources();let metrics=fonts.get(&options.math_font).unwrap().math.as_ref().unwrap();
    for style in [MathStyle::Display,MathStyle::Text,MathStyle::Script,MathStyle::ScriptScript] {
        options.style=style;let session=EngineSession::from_source("geometry".into(),r"\binom{n+1}{k}".into()).unwrap();
        let document=session.document_snapshot();let (owner,num,den)=fraction(&document.analysis.tree);
        let scene=LayoutEngine::default().layout(&session,&fonts,&options).unwrap();scene.validate().unwrap();
        let n=scene.regions.iter().find(|r|r.node_id==num.id).unwrap();let d=scene.regions.iter().find(|r|r.node_id==den.id).unwrap();
        let scale=match style {MathStyle::Display|MathStyle::Text=>1.0,MathStyle::Script=>metrics.script_scale,_=>metrics.scriptscript_scale};
        let gap=if style==MathStyle::Display {metrics.stack_display_gap}else{metrics.stack_gap};
        assert!(d.logical.y_min-n.logical.y_max+1e-9>=gap*scale);
        assert!(scene.rules.iter().all(|r|r.node_id!=owner.id));
        assert!(scene.glyphs.iter().filter(|g|g.node_id==owner.id).count()>=2);
        let own:Vec<_>=scene.glyphs.iter().filter(|g|g.node_id==owner.id).collect();
        let left_x=own.iter().map(|g|g.x).min_by(f64::total_cmp).unwrap();
        let mut top=f64::INFINITY;let mut bottom=f64::NEG_INFINITY;
        for g in own.iter().filter(|g|(g.x-left_x).abs()<1e-9) {
            let ink=scene.resources[g.resource].shape.ink.unwrap();top=top.min(g.y+ink[1]*g.scale);bottom=bottom.max(g.y+ink[3]*g.scale);
        }
        assert!(((top+bottom)/2.0-(scene.baseline_em-metrics.axis*scale)).abs()<1e-9);
        assert!(scene.svg().unwrap().contains("<path"));
    }
}
#[test]
fn forced_binomial_style_overrides_ambient_size_without_source_or_history_edits() {
    let (fonts,mut options)=resources();
    for command in [r"\dbinom",r"\tbinom"] {
        let session=EngineSession::from_source("forced".into(),format!("{command}{{n}}{{k}}" )).unwrap();
        let before=serde_json::to_value(session.editor_snapshot().unwrap()).unwrap();let mut engine=LayoutEngine::default();let mut reference=None;
        for style in [MathStyle::Display,MathStyle::Text,MathStyle::Script,MathStyle::ScriptScript] {
            options.style=style;let scene=engine.layout(&session,&fonts,&options).unwrap();
            let value=json!([scene.glyphs,scene.rules,scene.width,scene.height,scene.baseline]);
            if let Some(prior)=&reference {assert_eq!(&value,prior);}else{reference=Some(value);}
            assert_eq!(serde_json::to_value(session.editor_snapshot().unwrap()).unwrap(),before);
        }
    }
}
#[test]
fn bevelled_fraction_has_real_slash_and_two_separate_reduced_slots() {
    let (fonts,mut options)=resources();let metrics=fonts.get(&options.math_font).unwrap().math.as_ref().unwrap();
    for style in [MathStyle::Display,MathStyle::Text,MathStyle::Script,MathStyle::ScriptScript] {
        options.style=style;let session=EngineSession::from_source("bevelled".into(),r"\nicefrac{a+1}{b}".into()).unwrap();
        let document=session.document_snapshot();let (owner,num,den)=fraction(&document.analysis.tree);
        let scene=LayoutEngine::default().layout(&session,&fonts,&options).unwrap();
        let nr=scene.regions.iter().find(|r|r.node_id==num.id).unwrap();let dr=scene.regions.iter().find(|r|r.node_id==den.id).unwrap();
        let scale=match style {MathStyle::Display|MathStyle::Text=>1.0,MathStyle::Script=>metrics.script_scale,_=>metrics.scriptscript_scale};
        assert!(dr.logical.x_min-nr.logical.x_max+1e-9>=metrics.skewed_horizontal_gap*scale);
        assert!(dr.logical.y_min-nr.logical.y_max+1e-9>=metrics.skewed_vertical_gap*scale);
        let nc=scene.caret(Position::row(num.id,0)).unwrap();let dc=scene.caret(Position::row(den.id,0)).unwrap();
        assert!(nc.baseline<dc.baseline);assert!(scene.rules.iter().all(|r|r.node_id!=owner.id));
        assert!(scene.glyphs.iter().any(|g|g.node_id==owner.id));scene.svg().unwrap();
        assert_eq!(scene.vertical(Position::row(num.id,0),true,None).unwrap().node_id(),den.id);
    }
}
#[test]
fn nested_forms_colors_phantom_and_negative_glue_keep_valid_natural_geometry() {
    let (fonts,options)=resources();
    for source in [r"\binom{\frac{n}{2}}{\sqrt{k}}",r"\sqrt{\binom{n}{k}}",r"\nicefrac{\text{速度}}{y}",
        r"\textcolor{red}{\binom{n}{k}}",r"\phantom{\binom{n}{k}}+x",r"\binom{\!}{\!}",r"\nicefrac{\!}{b}",
        r"x^{\tbinom{n}{k}}",r"\binom{\nicefrac{a}{b}}{\binom{x}{y}}"] {
        let session=EngineSession::from_source("nested".into(),source.into()).unwrap();
        let scene=LayoutEngine::default().layout(&session,&fonts,&options).unwrap();scene.validate().unwrap();
        if scene.export_ready {scene.svg().unwrap();}
        if source.starts_with(r"\phantom") {
            let doc=session.document_snapshot();let NodeKind::Row{entries,..}=&doc.analysis.tree.kind else{panic!()};
            let hidden=ids(&entries[0].node);
            assert_eq!(scene.glyphs.len(),2,"Only the visible + and x should paint");
            assert!(scene.glyphs.iter().all(|g|!hidden.contains(&g.node_id)));
            assert!(scene.rules.iter().all(|r|!hidden.contains(&r.node_id)));
        }
        if source.starts_with(r"\textcolor") {assert!(scene.glyphs.iter().all(|g|g.color.is_some()));}
    }
}
#[test]
fn new_fraction_empty_slots_render_for_editing_but_cannot_be_exported() {
    let (fonts,options)=resources();for command in FORMS {
        let session=EngineSession::from_source("empty".into(),format!("{command}{{}}{{}}" )).unwrap();
        let scene=LayoutEngine::default().layout(&session,&fonts,&options).unwrap();assert!(!scene.export_ready);
        assert_eq!(scene.svg().unwrap_err().code,ErrorCode::IncompleteSource);
        let p=session.document_snapshot();let (_,num,den)=fraction(&p.analysis.tree);
        assert!(scene.caret(Position::row(num.id,0)).is_some());assert!(scene.caret(Position::row(den.id,0)).is_some());
        assert_eq!(scene.regions.iter().filter(|region|region.role=="emptyRow").count(),2,
            "Empty fraction arguments must remain explicit editor affordances in the scene");
    }
}
#[test]
fn authored_fraction_kind_changes_invalidate_geometry_and_undo_restores_the_same_svg() {
    let (fonts,options)=resources();let mut t=Owned::new(r"\binom{n}{k}");let mut layout=LayoutEngine::default();
    let first=layout.layout(&t.session,&fonts,&options).unwrap();let vector=first.svg().unwrap();
    t.session.patch_source(t.session.revision(),SourceEdit{start_utf16:0,end_utf16:6,replacement:r"\nicefrac".into()}).unwrap();
    let changed=layout.layout(&t.session,&fonts,&options).unwrap();assert_ne!(changed.svg().unwrap(),vector);
    t.act(EditorAction::Undo);assert_eq!(layout.layout(&t.session,&fonts,&options).unwrap().svg().unwrap(),vector);
    let cached=layout.layout(&t.session,&fonts,&options).unwrap();assert_eq!(cached.rebuilt_fragments,0);
}
#[test]
fn unknown_forged_fraction_presentation_is_rejected_before_source_projection() {
    let mut p=parse(r"\frac{x}{y}");let tree=Arc::make_mut(&mut p.tree);
    let NodeKind::Row{entries,..}=&mut tree.kind else{panic!()};
    let fraction=Arc::make_mut(&mut entries[0].node);fraction.origin=None;
    let NodeKind::Fraction{command,..}=&mut fraction.kind else{panic!()};*command=r"\notAnImplementedFraction".into();tree.origin=None;
    assert_eq!(model::validate_tree(&p.tree).unwrap_err().code,ErrorCode::InvalidInput);
    assert!(project_tree(p.tree,2).is_err());
}
#[test]
fn seeded_fraction_source_edits_preserve_exact_input_and_fresh_parse_semantics() {
    let seed=0xB10F2026u32;let mut rng=seed;
    for n in 0..480 {
        rng=rng.wrapping_mul(1664525).wrapping_add(1013904223);
        let command=FORMS[n%FORMS.len()];let source=format!("{command}{{x_1+\\frac{{a}}{{b}}}}{{y}}+{command}{{x}}{{y}} % keep");
        let mut session=EngineSession::from_source("fuzz".into(),source.clone()).unwrap();
        let at=(rng as usize)%(source.len()+1);let end=if n%3==0 {(at+1).min(source.len())}else{at};
        let replacement=["x","{","}","%","\\","^","","+","\\binom{n}{k}"][n%9];
        let expected=format!("{}{}{}",&source[..at],replacement,&source[end..]);
        let result=session.patch_source(session.revision(),SourceEdit{start_utf16:at as u32,end_utf16:end as u32,replacement:replacement.into()}).unwrap().0;
        assert_eq!(result.source,expected,"seed={seed:x}, n={n}");assert_eq!(model::latex(&result.analysis.tree,LatexMode::Lossless),expected);
        let fresh=parse(&expected);assert_eq!(semantic(&result.analysis.tree),semantic(&fresh.tree),"seed={seed:x}, n={n}, {expected}");
        assert_eq!((result.analysis.complete,result.analysis.supported),(fresh.complete,fresh.supported));
    }
}
