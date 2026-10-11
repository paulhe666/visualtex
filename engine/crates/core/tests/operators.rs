//! P07.4: authored operator tails, editor ownership, navigation and real-font layout.
use std::{collections::BTreeSet,path::PathBuf};
use serde_json::{json,Value};
use visualtex_engine_core::{editing::{EditorAction,EditorActionResult},fonts::FontRegistry,
    layout::{LayoutEngine,LayoutOptions,MathStyle},model::{self,Node,NodeKind,IdAllocator,LatexMode,ScriptKind},
    operators::{self,LimitMode,OperatorPart},parser::{self,ParseResult},positions::{Position,Selection,Horizontal,TreeIndex},
    session::{EngineSession,NodeId},source::SourceText,source_edit::SourceEdit,
    transactions::{EditCommand,InputMode,StructureKind},ErrorCode};
const COMMANDS:&[&str]=&[r"\sum",r"\prod",r"\coprod",r"\bigcup",r"\bigcap",r"\bigsqcup",r"\biguplus",r"\bigvee",r"\bigwedge",
    r"\bigodot",r"\bigoplus",r"\bigotimes",r"\int",r"\iint",r"\iiint",r"\oint",r"\oiint",r"\oiiint",
    r"\lim",r"\limsup",r"\liminf",r"\min",r"\max",r"\inf",r"\sup"];
fn parse(s:&str)->ParseResult{parser::parse(SourceText::new(1,s).unwrap(),NodeId(1),&mut IdAllocator::new(2)).unwrap()}
fn semantic(n:&Node)->Value{match &n.kind{
    NodeKind::Row{entries,..}=>json!(["row",entries.iter().map(|e|semantic(&e.node)).collect::<Vec<_>>()]),
    NodeKind::Symbol{value,class,..}=>json!(["symbol",value,class]),NodeKind::Text{text,..}=>json!(["text",text]),
    NodeKind::Fraction{command,numerator,denominator,..}=>json!([command,semantic(numerator),semantic(denominator)]),
    NodeKind::Scripts{base,slots}=>json!(["scripts",semantic(base),slots.iter().map(|s|json!([s.script,semantic(&s.argument)])).collect::<Vec<_>>()]),
    NodeKind::Radical{degree,radicand,..}=>json!(["root",degree.as_ref().map(|d|semantic(d)),semantic(radicand)]),
    NodeKind::Fence{left,right,body,..}=>json!(["fence",left,right,semantic(body)]),
    NodeKind::MiddleDelimiter{delimiter,..}=>json!(["middle",delimiter]),
    NodeKind::GeneralOperator{command,body,parts,..}=>json!(["generalOperator",command,semantic(body),operators::effective_mode(parts),operators::slots(parts).map(|s|json!([s.script,semantic(&s.argument)])).collect::<Vec<_>>()]),
    NodeKind::MacroCall{command,macro_kind,arguments,..}=>json!(["macro",command,macro_kind,arguments.iter().map(|a|semantic(a)).collect::<Vec<_>>()]),
    NodeKind::Grid{environment,columns,rows,..}=>json!(["grid",environment,columns.len(),rows.iter().map(|r|semantic(r)).collect::<Vec<_>>()]),
    NodeKind::GridRow{cells,..}=>json!(["gridRow",cells.iter().map(|c|semantic(c)).collect::<Vec<_>>()]),
    NodeKind::Accent{command,body,..}=>json!(["accent",command,semantic(body)]),
    NodeKind::Annotation{command,above,annotation,base,..}=>json!(["annotation",command,above,semantic(annotation),semantic(base)]),
    NodeKind::LargeOperator{command,parts}=>json!(["operator",command,operators::effective_mode(parts),operators::slots(parts).map(|s|json!([s.script,semantic(&s.argument)])).collect::<Vec<_>>()]),
    NodeKind::Style{command,variant,body,..}=>json!(["style",command,variant,semantic(body)]),
    NodeKind::Space{command,width_mu}=>json!(["space",command,width_mu]),
    NodeKind::Phantom{command,axis,body,..}=>json!(["phantom",command,axis,semantic(body)]),
    NodeKind::Boxed{body,..}=>json!(["boxed",semantic(body)]),
    NodeKind::Color{spec,body,..}=>json!(["color",spec.color,semantic(body)]),
    NodeKind::ColorChange{spec}=>json!(["colorChange",spec.color]),
    NodeKind::StyleChange{command,..}=>json!(["styleChange",command]),
    NodeKind::CustomSymbol{command,asset_sha256,parts,..}=>json!(["customSymbol",command,asset_sha256,operators::effective_mode(parts),operators::slots(parts).map(|s|json!([s.script,semantic(&s.argument)])).collect::<Vec<_>>()]),
    NodeKind::EmptySlot{label}=>json!(["slot",label]),NodeKind::UnparsedSource{raw}=>json!(["opaque",raw]),
}}
fn ids(n:&Node)->BTreeSet<NodeId>{let mut all=BTreeSet::new();n.visit(&mut |n|{assert!(all.insert(n.id));});all}
fn first_operator(tree:&Node)->&Node{if matches!(tree.kind,NodeKind::LargeOperator{..}){return tree;}tree.children().into_iter().find(|n|matches!(n.kind,NodeKind::LargeOperator{..})).unwrap()}
fn bound(tree:&Node,kind:ScriptKind)->&Node{let NodeKind::LargeOperator{parts,..}=&first_operator(tree).kind else{panic!()};operators::slots(parts).find(|s|s.script==kind).unwrap().argument.as_ref()}
struct Owned{session:EngineSession,sequence:u32}
impl Owned{
    fn new(source:&str)->Self{Self{session:EngineSession::from_source("operator-test".into(),source.into()).unwrap(),sequence:0}}
    fn action(&mut self,action:EditorAction)->EditorActionResult{
        let before=self.session.document_snapshot();self.sequence+=1;
        let result=self.session.apply_action(self.session.revision(),self.sequence,action).unwrap();
        let tree=&result.editor.document.analysis.tree;let index=TreeIndex::new(tree).unwrap();index.validate_selection(&result.editor.selection).unwrap();
        if let Some(map)=&result.position_map{for p in TreeIndex::new(&before.analysis.tree).unwrap().carets{map.apply_position(p,&index).unwrap();}}
        let q=parse(self.session.source());assert_eq!(semantic(tree),semantic(&q.tree),"{}",self.session.source());
        assert_eq!(ids(tree).len(),result.editor.document.analysis.facts.nodes);result
    }
    fn edit(&mut self,command:EditCommand)->EditorActionResult{self.action(EditorAction::Edit{command,group_id:None})}
    fn text(&mut self,text:&str){self.edit(EditCommand::InsertText{text:text.into(),mode:InputMode::Math});}
    fn select(&mut self,selection:Selection){self.action(EditorAction::SetSelection{selection});}
    fn right(&mut self)->EditorActionResult{self.action(EditorAction::MoveHorizontal{direction:Horizontal::Right,extend:false})}
}
fn fonts()->(FontRegistry,LayoutOptions){
    let path=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    let mut fonts=FontRegistry::default();let id=fonts.register(&std::fs::read(path).unwrap(),0).unwrap().id;(fonts,LayoutOptions::new(id))
}
#[test]
fn operators_preserve_bounds_directive_order_comments_and_command_boundaries(){
    for command in COMMANDS{for tail in ["",r"_i^n",r"^{n}_{i=1}",r"\limits_i^n",r"_i\nolimits^n",r"^n_i\limits",r"\nolimits_i^n\limits\displaylimits",
        " % before\r\n ^{n} % middle\n \\limits _{i=1} % end\n "]{
        let source=format!("{command}{tail} x_i");let p=parse(&source);
        assert!(p.complete&&p.supported&&p.exportable,"{source}: {:?}",p.diagnostics);
        assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
        let canonical=model::latex(&p.tree,LatexMode::Canonical);let q=parse(&canonical);
        assert!(q.complete&&q.supported);assert_eq!(semantic(&p.tree),semantic(&q.tree),"{source}");
        assert_eq!(model::latex(&q.tree,LatexMode::Canonical),canonical);
    }}
    for source in [r"\sum^\sum_0",r"\sqrt[{\sum_]}]{x}",r"\hat{\sum_{i=1}^{n}}",r"\sum^{\sum_i^n}_j x_j",r"\lim_{x\to0}x",r"{\sum_i^n}^2"]{
        let p=parse(source);assert!(p.complete&&p.supported,"{source}");let q=parse(&model::latex(&p.tree,LatexMode::Canonical));
        assert_eq!(semantic(&p.tree),semantic(&q.tree),"{source}");
    }
}
#[test]
fn duplicate_missing_and_unknown_operator_tails_preserve_drafts(){
    for source in [r"\sum_1_2",r"\int^a^b",r"\sum_",r"\int^{",r"\prod_{i=1}^{"]{
        let p=parse(source);assert!(!p.complete&&!p.exportable,"{source}");assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
    }
    for source in [r"\sum_{}",r"\sum_{}^{}"]{let p=parse(source);assert!(p.complete&&p.supported&&!p.exportable);}
    for source in [r"\unknownop_i^n",r"x\limits"]{let p=parse(source);assert!(!p.supported&&!p.exportable);assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);}
}
#[test]
fn sum_constructs_lower_upper_then_body_without_reparsing_typing(){
    let mut t=Owned::new("");t.edit(EditCommand::InsertOperator{command:r"\sum".into(),with_lower:true,with_upper:true});
    t.text("i=1");t.right();t.text("n");t.right();t.text("x");t.edit(EditCommand::InsertStructure{structure:StructureKind::Subscript});t.text("i");
    assert_eq!(t.session.source(),r"\sum_{i=1}^{n}x_{i}");let before=t.session.document_snapshot();
    t.action(EditorAction::Undo);t.action(EditorAction::Redo);assert_eq!(semantic(&before.analysis.tree),semantic(&t.session.document_snapshot().analysis.tree));
    assert_eq!(ids(&before.analysis.tree),ids(&t.session.document_snapshot().analysis.tree));
}
#[test]
fn source_order_does_not_change_lower_upper_navigation(){
    let mut t=Owned::new(r"\sum^{n}_{i}x");let s=t.session.document_snapshot();
    t.select(Selection::caret(Position::row(NodeId(1),0)));let entered=t.right();
    assert!(matches!(entered.editor.selection,Selection::Range{focus,..} if focus.node_id()==bound(&s.analysis.tree,ScriptKind::Subscript).id));
    t.select(Selection::caret(Position::row(bound(&s.analysis.tree,ScriptKind::Subscript).id,1)));let upper=t.right();
    assert!(matches!(upper.editor.selection,Selection::Range{focus,..} if focus.node_id()==bound(&s.analysis.tree,ScriptKind::Superscript).id));
    t.select(Selection::caret(Position::row(bound(&s.analysis.tree,ScriptKind::Superscript).id,1)));let exit=t.right();
    assert!(matches!(exit.editor.selection,Selection::Range{focus:Position::Row{node_id:NodeId(1),offset:1,..},..}));
    assert_eq!(t.session.source(),s.source);assert_eq!(t.session.revision(),s.revision);
}
#[test]
fn insertion_preserves_selected_operand_identity_instead_of_turning_it_into_a_bound(){
    let mut t=Owned::new("x+y");let before=t.session.document_snapshot();let original:Vec<_>=before.analysis.tree.children().iter().map(|n|n.id).collect();
    t.select(Selection::Range{anchor:Position::row(NodeId(1),0),focus:Position::row(NodeId(1),3)});
    t.edit(EditCommand::InsertOperator{command:r"\int".into(),with_lower:false,with_upper:false});
    assert_eq!(t.session.source(),r"\int x+y");let after=t.session.document_snapshot();
    assert_eq!(after.analysis.tree.children()[1..].iter().map(|n|n.id).collect::<Vec<_>>(),original);
    assert_eq!(after.analysis.facts.large_operators,1);
}
#[test]
fn adding_or_reentering_operator_bounds_retains_the_operator_identity(){
    let mut t=Owned::new(r"\sum");let op=first_operator(&t.session.document_snapshot().analysis.tree).id;
    t.select(Selection::caret(Position::row(NodeId(1),1)));t.edit(EditCommand::InsertStructure{structure:StructureKind::Subscript});t.text("i");t.right();
    t.edit(EditCommand::InsertStructure{structure:StructureKind::Superscript});t.text("n");t.right();
    let before=t.session.document_snapshot();assert_eq!(first_operator(&before.analysis.tree).id,op);assert_eq!(t.session.source(),r"\sum_{i}^{n}");
    let result=t.edit(EditCommand::InsertStructure{structure:StructureKind::Subscript});assert!(!result.content_changed);
    assert_eq!(result.editor.document.revision,before.revision);
}
#[test]
fn applying_a_power_to_a_selected_whole_operator_preserves_its_existing_bounds(){
    let mut t=Owned::new(r"\sum_i^n");let old=t.session.document_snapshot();
    t.select(Selection::Node{node_id:first_operator(&old.analysis.tree).id});
    t.edit(EditCommand::InsertStructure{structure:StructureKind::Superscript});t.text("2");
    assert_eq!(t.session.source(),r"{\sum_i^n}^{2}");
    assert!(parse(t.session.source()).complete);
    t.action(EditorAction::Undo);t.action(EditorAction::Undo);assert_eq!(t.session.source(),old.source);
    assert_eq!(ids(&t.session.document_snapshot().analysis.tree),ids(&old.analysis.tree));
}
#[test]
fn unicode_operator_input_becomes_the_same_real_editable_node(){
    let mut t=Owned::new("");t.text("∑");t.edit(EditCommand::InsertStructure{structure:StructureKind::Subscript});t.text("i");
    assert_eq!(t.session.source(),"∑_{i}");assert_eq!(t.session.document_snapshot().analysis.facts.large_operators,1);
}
#[test]
fn limits_mode_changes_are_undoable_and_auto_restoration_preserves_comments(){
    for source in ["\\sum % first\n \\limits_{i} % second\n ^{n} % last\n \\nolimits x",r"\int_i^n"]{
        let mut t=Owned::new(source);let old=t.session.document_snapshot();let op=first_operator(&old.analysis.tree).id;
        t.select(Selection::Node{node_id:op});t.edit(EditCommand::SetOperatorLimits{mode:LimitMode::Limits});
        let limits=t.session.document_snapshot();assert_eq!(ids(&old.analysis.tree),ids(&limits.analysis.tree));
        let NodeKind::LargeOperator{parts,..}=&first_operator(&limits.analysis.tree).kind else{panic!()};assert_eq!(operators::effective_mode(parts),LimitMode::Limits);
        let no_op=t.edit(EditCommand::SetOperatorLimits{mode:LimitMode::Limits});assert!(!no_op.content_changed);
        t.edit(EditCommand::SetOperatorLimits{mode:LimitMode::Auto});
        let auto=t.session.document_snapshot();let NodeKind::LargeOperator{parts,..}=&first_operator(&auto.analysis.tree).kind else{panic!()};
        assert_eq!(operators::effective_mode(parts),LimitMode::Auto);assert!(parts.iter().all(|p|matches!(p,OperatorPart::Script{..})));
        if source.contains('%'){for comment in ["% first","% second","% last"]{assert!(auto.source.contains(comment));}}
        t.action(EditorAction::Undo);assert_eq!(t.session.source(),limits.source);
        t.action(EditorAction::Undo);assert_eq!(t.session.source(),source);assert_eq!(ids(&old.analysis.tree),ids(&t.session.document_snapshot().analysis.tree));
    }
}
#[test]
fn copied_operator_bounds_remain_independent_and_deleted_ids_return_on_undo(){
    let mut t=Owned::new(r"\sum_{i=1}^{n}\limits");t.select(Selection::Node{node_id:first_operator(&t.session.document_snapshot().analysis.tree).id});
    t.edit(EditCommand::DuplicateSelection);let copied=t.session.document_snapshot();let children=copied.analysis.tree.children();
    assert!(ids(children[0]).is_disjoint(&ids(children[1])));let second=bound(children[1],ScriptKind::Subscript).id;
    t.select(Selection::caret(Position::row(second,0)));t.text("j+");let after=t.session.document_snapshot();
    assert_eq!(model::latex(after.analysis.tree.children()[0],LatexMode::Lossless),r"\sum_{i=1}^{n}\limits");
    t.action(EditorAction::Undo);assert_eq!(ids(&t.session.document_snapshot().analysis.tree),ids(&copied.analysis.tree));
    t.select(Selection::caret(Position::row(NodeId(1),2)));t.edit(EditCommand::DeleteBackward);t.edit(EditCommand::DeleteBackward);
    assert_eq!(t.session.document_snapshot().analysis.facts.large_operators,1);t.action(EditorAction::Undo);
    assert_eq!(ids(&t.session.document_snapshot().analysis.tree),ids(&copied.analysis.tree));
}
#[test]
fn local_source_bound_edit_preserves_the_other_bound_and_repeated_operator(){
    let mut t=Owned::new(r"\sum_{i=1}^{n}+\sum_{i=1}^{n}");let old=t.session.document_snapshot();
    let upper=ids(bound(&old.analysis.tree,ScriptKind::Superscript));let other=ids(old.analysis.tree.children()[2]);
    let at=t.session.source().find("i=1").unwrap();
    t.session.patch_source(t.session.revision(),SourceEdit{start_utf16:at as u32,end_utf16:at as u32+1,replacement:"j".into()}).unwrap();
    let new=t.session.document_snapshot();assert_eq!(ids(bound(&new.analysis.tree,ScriptKind::Superscript)),upper);
    assert_eq!(ids(new.analysis.tree.children()[2]),other);assert_eq!(semantic(&new.analysis.tree),semantic(&parse(&new.source).tree));
}
#[test]
fn all_operator_fonts_and_limits_policies_have_valid_unclipped_geometry(){
    let (fonts,mut options)=fonts();let mut layout=LayoutEngine::default();
    for command in COMMANDS{for directive in ["",r"\limits",r"\nolimits",r"\displaylimits"]{for style in [MathStyle::Display,MathStyle::Text,MathStyle::Script,MathStyle::ScriptScript]{
        options.style=style;let source=format!("{command}{directive}_{{i=1}}^{{n}}x");let t=Owned::new(&source);
        let scene=layout.layout(&t.session,&fonts,&options).unwrap_or_else(|e|panic!("{source}, {style:?}: {e:?}"));scene.validate().unwrap();assert!(scene.export_ready);
        let svg=scene.svg().unwrap();for bad in ["<text","<image","NaN","Infinity"]{assert!(!svg.contains(bad));}
        let s=t.session.document_snapshot();let index=TreeIndex::new(&s.analysis.tree).unwrap();for c in &scene.carets{index.validate_position(c.position).unwrap();}
        for g in &scene.glyphs{if let Some([a,b,c,d])=scene.resources[g.resource].shape.ink{
            assert!(g.x+a*g.scale>=-1e-9&&g.y+b*g.scale>=-1e-9,"{source}");
            assert!(g.x+c*g.scale<=scene.width_em+1e-9&&g.y+d*g.scale<=scene.height_em+1e-9,"{source}");
        }}
    }}}
}
#[test]
fn display_operator_variant_is_real_font_data_and_not_operand_stretch(){
    let (fonts,mut options)=fonts();let mut layout=LayoutEngine::default();
    for command in [r"\sum",r"\int",r"\prod"]{
        let a=Owned::new(command);let b=Owned::new(&format!("{command}\\dfrac{{1}}{{\\dfrac{{2}}{{3}}}}"));
        options.style=MathStyle::Text;let text=layout.layout(&a.session,&fonts,&options).unwrap();
        options.style=MathStyle::Display;let display=layout.layout(&a.session,&fonts,&options).unwrap();let operand=layout.layout(&b.session,&fonts,&options).unwrap();
        let glyph=|s:&visualtex_engine_core::scene::LayoutSnapshot,id:NodeId|{let g=s.glyphs.iter().find(|g|g.node_id==id).unwrap();s.resources[g.resource].shape.clone()};
        let tg=glyph(&text,first_operator(&a.session.document_snapshot().analysis.tree).id);
        let dg=glyph(&display,first_operator(&a.session.document_snapshot().analysis.tree).id);
        let bg=glyph(&operand,first_operator(&b.session.document_snapshot().analysis.tree).id);
        assert_ne!(tg.glyph_id,dg.glyph_id);assert_eq!(dg.glyph_id,bg.glyph_id);
        let font=fonts.get(&options.math_font).unwrap();let ink=dg.ink.unwrap();assert!(ink[3]-ink[1]+1e-9>=font.math.as_ref().unwrap().display_operator_min_height);
        assert!(display.glyphs.iter().all(|g|(g.scale-1.0).abs()<1e-9));
    }
}
#[test]
fn limits_and_nolimits_geometry_matches_the_policy_and_font_gaps(){
    let (fonts,mut options)=fonts();let mut layout=LayoutEngine::default();let metrics=fonts.get(&options.math_font).unwrap().math.as_ref().unwrap();
    for (source,style,stacked) in [(r"\sum_i^n",MathStyle::Display,true),(r"\sum_i^n",MathStyle::Text,false),
        (r"\int_i^n",MathStyle::Display,false),(r"\int\limits_i^n",MathStyle::Text,true),
        (r"\sum\nolimits_i^n",MathStyle::Display,false),(r"\int\displaylimits_i^n",MathStyle::Display,true)]{
        options.style=style;let t=Owned::new(source);let p=t.session.document_snapshot();let op=first_operator(&p.analysis.tree);
        let scene=layout.layout(&t.session,&fonts,&options).unwrap();let glyph=scene.glyphs.iter().find(|g|g.node_id==op.id).unwrap();
        let shape=&scene.resources[glyph.resource].shape;let [a,b,c,d]=shape.ink.unwrap();
        let (left,top,right,bottom)=(glyph.x+a*glyph.scale,glyph.y+b*glyph.scale,glyph.x+c*glyph.scale,glyph.y+d*glyph.scale);
        let upper=scene.regions.iter().find(|r|r.node_id==bound(op,ScriptKind::Superscript).id).unwrap().ink.unwrap();
        let lower=scene.regions.iter().find(|r|r.node_id==bound(op,ScriptKind::Subscript).id).unwrap().ink.unwrap();
        if stacked {
            assert!(upper.y_max+metrics.upper_limit_gap<=top+1e-9,"{source}");assert!(lower.y_min>=bottom+metrics.lower_limit_gap-1e-9,"{source}");
            assert!(upper.x_min<right&&upper.x_max>left,"{source}");
        } else {
            // Advance is measured from the glyph origin, not its left ink
            // bearing. Assert the exact logical placement of both arguments.
            let ur=scene.regions.iter().find(|r|r.node_id==bound(op,ScriptKind::Superscript).id).unwrap().logical;
            let lr=scene.regions.iter().find(|r|r.node_id==bound(op,ScriptKind::Subscript).id).unwrap().logical;
            assert!((ur.x_min-(glyph.x+(shape.advance+shape.italic_correction)*glyph.scale)).abs()<1e-9,"{source}");
            assert!((lr.x_min-(glyph.x+shape.advance*glyph.scale)).abs()<1e-9,"{source}");
        }
    }
}
#[test]
fn changing_limits_invalidates_geometry_without_losing_bound_ids_and_undo_restores_svg(){
    let (fonts,options)=fonts();let mut layout=LayoutEngine::default();let mut t=Owned::new(r"\sum_{i=1}^{n}x");
    let old=t.session.document_snapshot();let initial=layout.layout(&t.session,&fonts,&options).unwrap();
    t.select(Selection::caret(Position::row(bound(&old.analysis.tree,ScriptKind::Subscript).id,0)));
    t.edit(EditCommand::SetOperatorLimits{mode:LimitMode::NoLimits});let side=layout.layout(&t.session,&fonts,&options).unwrap();
    assert_ne!(side.svg().unwrap(),initial.svg().unwrap());assert!(side.reused_fragments>0);
    assert_eq!(ids(&old.analysis.tree),ids(&t.session.document_snapshot().analysis.tree));
    t.action(EditorAction::Undo);let restored=layout.layout(&t.session,&fonts,&options).unwrap();assert_eq!(initial.svg().unwrap(),restored.svg().unwrap());
}
#[test]
fn unsupported_operator_and_excessive_tail_fail_without_mutating_the_session(){
    let mut t=Owned::new("x");let before=serde_json::to_value(t.session.editor_snapshot().unwrap()).unwrap();t.sequence+=1;
    let e=t.session.apply_action(t.session.revision(),t.sequence,EditorAction::Edit{command:EditCommand::InsertOperator{command:r"\futureop".into(),with_lower:true,with_upper:true},group_id:None}).unwrap_err();
    assert_eq!(e.code,ErrorCode::UnsupportedSyntax);assert_eq!(serde_json::to_value(t.session.editor_snapshot().unwrap()).unwrap(),before);
    assert_eq!(t.session.replace_source(t.session.revision(),r"\sum".to_owned()+&r"\limits".repeat(129)).unwrap_err().code,ErrorCode::ResourceLimit);
    assert_eq!(serde_json::to_value(t.session.editor_snapshot().unwrap()).unwrap(),before);
}
#[test]
fn seeded_operator_bound_edits_match_fresh_parse(){
    let mut seed=0x0b0734u32;fn next(s:&mut u32)->u32{*s=s.wrapping_mul(1664525).wrapping_add(1013904223);*s}
    for case in 0..100{
        let source=format!("{}^{{n}}_{{i=1}}\\limits+\\hat{{\\sum_j^n}}",COMMANDS[next(&mut seed) as usize%COMMANDS.len()]);
        let mut t=Owned::new(&source);
        for iteration in 0..6{let at=next(&mut seed) as usize%(t.session.source().len()+1);let insert=["{","}","%","x","_","^",r"\limits"][next(&mut seed) as usize%7];
            t.session.patch_source(t.session.revision(),SourceEdit{start_utf16:at as u32,end_utf16:at as u32,replacement:insert.into()}).unwrap();
            let p=t.session.document_snapshot();let q=parse(t.session.source());assert_eq!(semantic(&p.analysis.tree),semantic(&q.tree),"seed 0x0b0734 case {case} iteration {iteration}");
            assert_eq!((p.analysis.complete,p.analysis.supported,p.analysis.exportable),(q.complete,q.supported,q.exportable));
            assert_eq!(model::latex(&p.analysis.tree,LatexMode::Lossless),t.session.source());
        }
    }
}
