//! P07.3 vertical slice: parser, source preservation, identities, editing and
//! font-driven geometry. Browser/OS IME verification is a separate test layer.
use std::{collections::BTreeSet,path::PathBuf};
use serde_json::{json,Value};
use visualtex_engine_core::{accents,editing::{EditorAction,EditorActionResult},fonts::FontRegistry,
    layout::{LayoutEngine,LayoutOptions,MathStyle},model::{self,Node,NodeKind,IdAllocator,LatexMode},
    parser::{self,ParseResult},positions::{Position,Selection,Horizontal,TreeIndex},
    session::{EngineSession,NodeId},source::SourceText,source_edit::SourceEdit,
    transactions::{EditCommand,InputMode},ErrorCode};
const MARKS:&[&str]=&[r"\hat",r"\widehat",r"\tilde",r"\widetilde",r"\vec",r"\bar",r"\dot",r"\ddot",
    r"\dddot",r"\ddddot",r"\breve",r"\check",r"\acute",r"\grave",r"\mathring",r"\overline",r"\underline",
    r"\overrightarrow",r"\overleftarrow",r"\overleftrightarrow",r"\underrightarrow",r"\underleftarrow",
    r"\underleftrightarrow",r"\overbrace",r"\underbrace",r"\overparen",r"\underparen",r"\overarc",r"\underarc"];
const LABELS:&[&str]=&[r"\overset",r"\underset",r"\stackrel",r"\stackbin"];
fn parse(s:&str)->ParseResult{parser::parse(SourceText::new(1,s).unwrap(),NodeId(1),&mut IdAllocator::new(2)).unwrap()}
fn semantic(n:&Node)->Value{match &n.kind{
    NodeKind::Row{entries,..}=>json!(["row",entries.iter().map(|e|semantic(&e.node)).collect::<Vec<_>>()]),
    NodeKind::Symbol{value,class,..}=>json!(["symbol",value,class]),NodeKind::Text{text,..}=>json!(["text",text]),
    NodeKind::Fraction{command,numerator,denominator,..}=>json!([command,semantic(numerator),semantic(denominator)]),
    NodeKind::Scripts{base,slots}=>json!(["scripts",semantic(base),slots.iter().map(|s|json!([s.script,semantic(&s.argument)])).collect::<Vec<_>>()]),
    NodeKind::Radical{degree,radicand,..}=>json!(["root",degree.as_ref().map(|d|semantic(d)),semantic(radicand)]),
    NodeKind::Fence{left,right,body,..}=>json!(["fence",left,right,semantic(body)]),
    NodeKind::MiddleDelimiter{delimiter,..}=>json!(["middle",delimiter]),
    NodeKind::LargeOperator{command,parts}=>json!(["operator",command,visualtex_engine_core::operators::effective_mode(parts),visualtex_engine_core::operators::slots(parts).map(|s|json!([s.script,semantic(&s.argument)])).collect::<Vec<_>>()]),
    NodeKind::Style{command,variant,body,..}=>json!(["style",command,variant,semantic(body)]),
    NodeKind::Space{command,width_mu}=>json!(["space",command,width_mu]),
    NodeKind::Phantom{command,axis,body,..}=>json!(["phantom",command,axis,semantic(body)]),
    NodeKind::Boxed{body,..}=>json!(["boxed",semantic(body)]),
    NodeKind::Color{spec,body,..}=>json!(["color",spec.color,semantic(body)]),
    NodeKind::ColorChange{spec}=>json!(["colorChange",spec.color]),
    NodeKind::StyleChange{command,..}=>json!(["styleChange",command]),
    NodeKind::GeneralOperator{command,body,parts,..}=>json!(["generalOperator",command,semantic(body),visualtex_engine_core::operators::effective_mode(parts),visualtex_engine_core::operators::slots(parts).map(|s|json!([s.script,semantic(&s.argument)])).collect::<Vec<_>>()]),
    NodeKind::MacroCall{command,macro_kind,arguments,..}=>json!(["macro",command,macro_kind,arguments.iter().map(|a|semantic(a)).collect::<Vec<_>>()]),
    NodeKind::Grid{environment,columns,rows,..}=>json!(["grid",environment,columns.len(),rows.iter().map(|r|semantic(r)).collect::<Vec<_>>()]),
    NodeKind::GridRow{cells,..}=>json!(["gridRow",cells.iter().map(|c|semantic(c)).collect::<Vec<_>>()]),
    NodeKind::Accent{command,body,..}=>json!(["accent",command,semantic(body)]),
    NodeKind::Annotation{command,above,annotation,base,..}=>json!(["annotation",command,above,semantic(annotation),semantic(base)]),
    NodeKind::CustomSymbol{command,asset_sha256,parts,..}=>json!(["customSymbol",command,asset_sha256,visualtex_engine_core::operators::effective_mode(parts),visualtex_engine_core::operators::slots(parts).map(|s|json!([s.script,semantic(&s.argument)])).collect::<Vec<_>>()]),
    NodeKind::EmptySlot{label}=>json!(["slot",label]),NodeKind::UnparsedSource{raw}=>json!(["opaque",raw]),
}}
fn ids(n:&Node)->BTreeSet<NodeId>{let mut ids=BTreeSet::new();n.visit(&mut |n|{assert!(ids.insert(n.id));});ids}
struct Owned{session:EngineSession,sequence:u32}
impl Owned{
    fn new(source:&str)->Self{Self{session:EngineSession::from_source("accent-test".into(),source.into()).unwrap(),sequence:0}}
    fn action(&mut self,action:EditorAction)->EditorActionResult{
        let before=self.session.document_snapshot();self.sequence+=1;
        let result=self.session.apply_action(self.session.revision(),self.sequence,action).unwrap();
        let tree=&result.editor.document.analysis.tree;let index=TreeIndex::new(tree).unwrap();
        index.validate_selection(&result.editor.selection).unwrap();
        if let Some(map)=&result.position_map{for p in TreeIndex::new(&before.analysis.tree).unwrap().carets{map.apply_position(p,&index).unwrap();}}
        let reopened=parse(self.session.source());
        assert_eq!(semantic(tree),semantic(&reopened.tree),"{}",self.session.source());
        assert_eq!(ids(tree).len(),result.editor.document.analysis.facts.nodes);
        result
    }
    fn edit(&mut self,command:EditCommand)->EditorActionResult{self.action(EditorAction::Edit{command,group_id:None})}
    fn text(&mut self,text:&str){self.edit(EditCommand::InsertText{text:text.into(),mode:InputMode::Math});}
    fn select(&mut self,selection:Selection){self.action(EditorAction::SetSelection{selection});}
}
fn fonts()->(FontRegistry,LayoutOptions){
    let path=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    let mut fonts=FontRegistry::default();let id=fonts.register(&std::fs::read(path).unwrap(),0).unwrap().id;
    (fonts,LayoutOptions::new(id))
}
#[test]
fn all_registered_accents_and_annotations_preserve_spelling_trivia_and_semantics(){
    let mut sources:Vec<String>=MARKS.iter().map(|cmd|format!("{cmd} % keep\r\n {{x_1+\\text{{速度}}}} ")).collect();
    sources.extend(LABELS.iter().map(|cmd|format!("{cmd} % label\n {{a}} % base\n {{x+y}}")));
    sources.extend([r"\hat x",r"\overset a=",r"\hat{\vec{x}}",r"\overset{a}{\underset{b}{x}}",
        r"\overbrace{x+y}^{n}",r"\underbrace{x+y}_{n}",r"\sqrt[{\widehat{n}}]{\hat{x}}"].map(str::to_owned));
    for source in sources{
        let p=parse(&source);assert!(p.complete&&p.supported&&p.exportable,"{source}: {:?}",p.diagnostics);
        assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
        let canonical=model::latex(&p.tree,LatexMode::Canonical);let q=parse(&canonical);
        assert!(q.complete&&q.supported);assert_eq!(semantic(&p.tree),semantic(&q.tree),"{source}");
        assert_eq!(model::latex(&q.tree,LatexMode::Canonical),canonical);
    }
}
#[test]
fn incomplete_empty_and_unknown_marks_never_become_export_ready(){
    for source in [r"\hat",r"\hat{",r"\overset{a}",r"\underset",r"\overset{a}{"]{
        let p=parse(source);assert!(!p.complete&&!p.exportable,"{source}");assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
    }
    for source in [r"\hat{}",r"\overset{}{x}",r"\underset{a}{}"]{let p=parse(source);assert!(p.complete&&p.supported);assert!(!p.exportable);}
    let p=parse(r"\futureaccent{\hat{x}}");assert!(!p.supported&&!p.exportable);
    assert_eq!(model::latex(&p.tree,LatexMode::Lossless),r"\futureaccent{\hat{x}}");
}
#[test]
fn nested_accents_build_by_transactions_and_exit_without_a_content_edit(){
    let mut t=Owned::new("");
    t.edit(EditCommand::InsertAccent{command:r"\hat".into()});
    t.edit(EditCommand::InsertAccent{command:r"\vec".into()});t.text("x");
    assert_eq!(t.session.source(),r"\hat{\vec{x}}");
    let revision=t.session.revision();let result=t.edit(EditCommand::ExitStructure);
    assert!(!result.content_changed);assert_eq!(t.session.revision(),revision);
    assert!(matches!(result.editor.selection,Selection::Range{focus:Position::Row{node_id:NodeId(1),offset:1,..},..}));
    t.text("+y");assert_eq!(t.session.source(),r"\hat{\vec{x}}+y");
    t.action(EditorAction::Undo);assert_eq!(t.session.source(),r"\hat{\vec{x}}");
}
#[test]
fn annotation_insertion_wraps_selection_and_navigates_label_to_base(){
    for cmd in LABELS{
        let mut t=Owned::new("x+y");let before=t.session.document_snapshot();let original=ids(&before.analysis.tree);
        t.select(Selection::Range{anchor:Position::row(NodeId(1),0),focus:Position::row(NodeId(1),3)});
        t.edit(EditCommand::InsertAnnotation{command:(*cmd).into()});t.text("a");
        let s=t.session.document_snapshot();let NodeKind::Annotation{base,annotation,..}=&s.analysis.tree.children()[0].kind else{panic!()};
        assert!(ids(base).is_superset(&original.difference(&BTreeSet::from([NodeId(1)])).copied().collect()));
        t.select(Selection::caret(Position::row(annotation.id,1)));
        let result=t.action(EditorAction::MoveHorizontal{direction:Horizontal::Right,extend:false});
        assert!(matches!(result.editor.selection,Selection::Range{focus:Position::Row{node_id,offset:0,..},..} if node_id==base.id));
        t.text("z");assert_eq!(t.session.source(),format!("{cmd}{{a}}{{zx+y}}"));
    }
}
#[test]
fn right_arrow_exits_nested_accents_but_does_not_skip_siblings_or_other_slots(){
    for source in [r"\hat{\vec{x}}",r"\hat{\vec{x}+y}",r"\frac{\hat{\vec{x}}}{y}",r"\hat{\vec{x^{2}}}"] {
        let mut t=Owned::new(source);let snapshot=t.session.document_snapshot();let tree=&snapshot.analysis.tree;
        let mut target=None;
        tree.visit(&mut |n| if let NodeKind::Symbol{value,..}=&n.kind{if value==if source.contains("^{2}"){ "2" }else{"x"}{target=Some(n.id);}});
        let index=TreeIndex::new(tree).unwrap();let (row,offset)=index.row_owner(target.unwrap()).unwrap();
        t.select(Selection::caret(Position::row(row,offset+1)));let before=t.session.editor_snapshot().unwrap();
        let result=t.action(EditorAction::MoveHorizontal{direction:Horizontal::Right,extend:false});
        assert!(!result.content_changed);assert_eq!(result.editor.document.revision,before.document.revision);assert_eq!(result.editor.undo_depth,before.undo_depth);
        let expected=if source==r"\hat{\vec{x}+y}" {
            let NodeKind::Accent{body,..}=&tree.children()[0].kind else{panic!()};Position::row(body.id,1)
        } else if source.starts_with(r"\frac") {
            let NodeKind::Fraction{numerator,..}=&tree.children()[0].kind else{panic!()};Position::row(numerator.id,1)
        } else {Position::row(tree.id,1)};
        assert!(matches!(result.editor.selection,Selection::Range{focus,..} if focus.same_location(expected)),"{source}: {:?}",result.editor.selection);
    }
}
#[test]
fn shift_right_retains_the_anchor_and_does_not_auto_exit_accent_chain(){
    let mut t=Owned::new(r"\hat{\vec{x}}");let s=t.session.document_snapshot();
    let NodeKind::Accent{body,..}=&s.analysis.tree.children()[0].kind else{panic!()};
    let NodeKind::Accent{body:inner,..}=&body.children()[0].kind else{panic!()};
    let start=Position::row(inner.id,1);t.select(Selection::caret(start));
    let result=t.action(EditorAction::MoveHorizontal{direction:Horizontal::Right,extend:true});
    let Selection::Range{anchor,focus}=result.editor.selection else{panic!()};
    assert!(anchor.same_location(start));assert_ne!(focus.node_id(),s.analysis.tree.id);
    assert!(!result.content_changed);assert_eq!(result.editor.document.revision,s.revision);
}
#[test]
fn copied_decorations_have_disjoint_ids_and_delete_undo_restores_them(){
    for source in [r"\hat{\vec{x}}",r"\overset{a}{x}",r"\underbrace{x+y}_{n}"]{
        let mut t=Owned::new(source);t.select(Selection::Range{anchor:Position::row(NodeId(1),0),focus:Position::row(NodeId(1),1)});
        t.edit(EditCommand::DuplicateSelection);let p=t.session.document_snapshot();let children=p.analysis.tree.children();
        let first=ids(children[0]);let second=ids(children[1]);assert!(first.is_disjoint(&second));
        let mut editable=None;children[1].visit(&mut |n|{if matches!(n.kind,NodeKind::Row{..}){editable=Some(n.id);}});
        t.select(Selection::caret(Position::row(editable.unwrap(),0)));t.text("q");
        let after=t.session.document_snapshot();assert_eq!(ids(after.analysis.tree.children()[0]),first);
        assert_eq!(model::latex(after.analysis.tree.children()[0],LatexMode::Lossless),source);
        t.action(EditorAction::Undo);assert_eq!(ids(&t.session.document_snapshot().analysis.tree),ids(&p.analysis.tree));
        t.select(Selection::caret(Position::row(NodeId(1),2)));t.edit(EditCommand::DeleteBackward);
        assert!(matches!(t.session.editor_snapshot().unwrap().selection,Selection::Node{..}));
        t.edit(EditCommand::DeleteBackward);assert_eq!(t.session.source(),source);
        t.action(EditorAction::Undo);assert_eq!(ids(&t.session.document_snapshot().analysis.tree),ids(&p.analysis.tree));
    }
}
#[test]
fn local_source_edit_retains_other_accent_and_annotation_identities(){
    let mut t=Owned::new(r"\overset{a}{\hat{x}}+\overset{a}{\hat{x}}");let before=t.session.document_snapshot();
    let other=ids(before.analysis.tree.children()[2]);let at=t.session.source().find("{x}").unwrap()+1;
    t.session.patch_source(t.session.revision(),SourceEdit{start_utf16:at as u32,end_utf16:at as u32+1,replacement:"z+1".into()}).unwrap();
    let after=t.session.document_snapshot();assert_eq!(ids(after.analysis.tree.children()[2]),other);
    assert_eq!(semantic(&after.analysis.tree),semantic(&parse(t.session.source()).tree));
}
#[test]
fn unrecognized_decoration_command_is_atomic(){
    let mut t=Owned::new(r"\hat{x}");let before=serde_json::to_value(t.session.editor_snapshot().unwrap()).unwrap();
    for command in [EditCommand::InsertAccent{command:r"\futureaccent".into()},EditCommand::InsertAnnotation{command:r"\futureannotation".into()}]{
        t.sequence+=1;let error=t.session.apply_action(t.session.revision(),t.sequence,EditorAction::Edit{command,group_id:None}).unwrap_err();
        assert_eq!(error.code,ErrorCode::UnsupportedSyntax);assert_eq!(serde_json::to_value(t.session.editor_snapshot().unwrap()).unwrap(),before);
    }
}
#[test]
fn all_mark_variants_have_valid_font_geometry_in_every_math_style(){
    let (fonts,mut options)=fonts();let mut layout=LayoutEngine::default();
    let mut sources:Vec<String>=MARKS.iter().map(|cmd|format!("{cmd}{{x+y}}" )).collect();
    sources.extend(LABELS.iter().map(|cmd|format!("{cmd}{{a+1}}{{x}}")));
    sources.extend([r"\hat{\vec{x}}",r"\overbrace{x+y+z}^{n}",r"\underbrace{x+y+z}_{n}",r"\overset{a}{\underset{b}{x}}"].map(str::to_owned));
    for style in [MathStyle::Display,MathStyle::Text,MathStyle::Script,MathStyle::ScriptScript]{
        options.style=style;
        for source in &sources{
            let t=Owned::new(source);let scene=layout.layout(&t.session,&fonts,&options).unwrap_or_else(|e|panic!("{source}, {style:?}: {e:?}"));
            scene.validate().unwrap();assert!(scene.export_ready);let svg=scene.svg().unwrap();
            for forbidden in ["<text","<image","<script","foreignObject","NaN","Infinity"]{assert!(!svg.contains(forbidden));}
            let doc=t.session.document_snapshot();let index=TreeIndex::new(&doc.analysis.tree).unwrap();for c in &scene.carets{index.validate_position(c.position).unwrap();}
            for g in &scene.glyphs{if let Some([x0,y0,x1,y1])=scene.resources[g.resource].shape.ink{
                assert!(g.x+x0*g.scale>=-1e-9&&g.y+y0*g.scale>=-1e-9,"{source}");
                assert!(g.x+x1*g.scale<=scene.width_em+1e-9&&g.y+y1*g.scale<=scene.height_em+1e-9,"{source}");
            }}
        }
    }
}
#[test]
fn overline_and_underline_use_font_gap_and_rule_thickness(){
    let (fonts,options)=fonts();let metrics=fonts.get(&options.math_font).unwrap().math.as_ref().unwrap();
    for (command,below) in [(r"\overline",false),(r"\underline",true)]{
        let t=Owned::new(&format!("{command}{{x+y}}"));let snapshot=t.session.document_snapshot();
        let accent=snapshot.analysis.tree.children()[0];let NodeKind::Accent{body,..}=&accent.kind else{panic!()};
        let scene=LayoutEngine::default().layout(&t.session,&fonts,&options).unwrap();let rule=scene.rules.iter().find(|r|r.node_id==accent.id).unwrap();
        let body=scene.regions.iter().find(|r|r.node_id==body.id).unwrap().ink.unwrap();
        let (actual_gap,expected_gap,expected_rule)=if below{(rule.rect.y_min-body.y_max,metrics.underbar_gap,metrics.underbar_rule)}else{(body.y_min-rule.rect.y_max,metrics.overbar_gap,metrics.overbar_rule)};
        assert!((actual_gap-expected_gap).abs()<1e-9);assert!((rule.rect.height()-expected_rule).abs()<1e-9);
    }
}
#[test]
fn annotation_and_brace_labels_are_above_or_below_instead_of_right_scripts(){
    let (fonts,options)=fonts();let mut layout=LayoutEngine::default();
    for source in [r"\overset{a}{x}",r"\underset{a}{x}",r"\overbrace{x+y}^{a}",r"\underbrace{x+y}_{a}"]{
        let t=Owned::new(source);let snapshot=t.session.document_snapshot();let n=snapshot.analysis.tree.children()[0];
        let (base,label,above)=match &n.kind{
            NodeKind::Annotation{base,annotation,above,..}=>(base,annotation,*above),
            NodeKind::Scripts{base,slots}=>(base,&slots[0].argument,source.contains("overbrace")),_=>panic!(),
        };
        let scene=layout.layout(&t.session,&fonts,&options).unwrap();
        let b=scene.regions.iter().find(|r|r.node_id==base.id).unwrap().ink.unwrap();
        let l=scene.regions.iter().find(|r|r.node_id==label.id).unwrap().ink.unwrap();
        assert!(if above{l.y_max<=b.y_min+1e-9}else{l.y_min>=b.y_max-1e-9},"{source}");
        assert!(l.x_min<b.x_max&&l.x_max>b.x_min,"label was placed like a right-hand script: {source}");
    }
}
#[test]
fn long_marks_use_genuine_font_pieces_and_scaling_does_not_relayout(){
    let (fonts,mut options)=fonts();let mut layout=LayoutEngine::default();
    for command in [r"\overbrace",r"\overrightarrow",r"\widehat"]{
        let t=Owned::new(&format!("{command}{{{}}}","x+".repeat(40)+"y"));
        options.font_size_px=16.0;let a=layout.layout(&t.session,&fonts,&options).unwrap();
        options.font_size_px=32.0;let b=layout.layout(&t.session,&fonts,&options).unwrap();
        assert_eq!(b.rebuilt_fragments,0);assert!((b.width-a.width*2.0).abs()<1e-8);
        assert_eq!(serde_json::to_value(&a.glyphs).unwrap(),serde_json::to_value(&b.glyphs).unwrap());
        let root=t.session.document_snapshot().analysis.tree.children()[0].id;
        let pieces:Vec<_>=a.glyphs.iter().filter(|g|g.node_id==root).collect();assert!(!pieces.is_empty());
        assert!(pieces.iter().all(|g|(g.scale-1.0).abs()<1e-9));
        if command!=r"\widehat"{assert!(pieces.len()>1,"{command} should assemble font pieces for this length");}
        assert!(accents::spec(command).is_some());
    }
}
#[test]
fn seeded_nested_decoration_source_edits_match_fresh_parse(){
    let mut seed=0xacc307u32;fn next(s:&mut u32)->u32{*s=s.wrapping_mul(1664525).wrapping_add(1013904223);*s}
    for case in 0..100{
        let mut source="x".to_owned();for _ in 0..3{source=match next(&mut seed)%3{
            0=>format!("\\hat{{{source}}}"),1=>format!("\\overset{{a}}{{{source}}}"),_=>format!("\\frac{{{source}}}{{\\vec{{y}}}}"),
        };}
        let mut t=Owned::new(&source);
        for iteration in 0..6{
            let at=next(&mut seed) as usize%(t.session.source().len()+1);
            let replacement=["x","}","{","%","^"," "][next(&mut seed) as usize%6];
            t.session.patch_source(t.session.revision(),SourceEdit{start_utf16:at as u32,end_utf16:at as u32,replacement:replacement.into()}).unwrap();
            let p=t.session.document_snapshot();let q=parse(t.session.source());
            assert_eq!(semantic(&p.analysis.tree),semantic(&q.tree),"seed 0xacc307 case {case}, edit {iteration}");
            assert_eq!((p.analysis.complete,p.analysis.supported),(q.complete,q.supported));
            assert_eq!(model::latex(&p.analysis.tree,LatexMode::Lossless),t.session.source());
        }
    }
}
