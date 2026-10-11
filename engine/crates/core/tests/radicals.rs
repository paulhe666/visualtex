use std::{collections::BTreeSet,path::PathBuf};
use serde_json::{Value,json};
use visualtex_engine_core::{editing::{EditorAction,EditorActionResult},fonts::FontRegistry,
    layout::{LayoutEngine,LayoutOptions,MathStyle},model::{self,Node,NodeKind,LatexMode,IdAllocator},
    parser::{self,ParseResult},positions::{Selection,Position,Horizontal,TreeIndex},
    session::{EngineSession,NodeId},source::SourceText,source_edit::SourceEdit,
    transactions::{EditCommand,InputMode,StructureKind}};
fn parse(s:&str)->ParseResult {parser::parse(SourceText::new(1,s).unwrap(),NodeId(1),&mut IdAllocator::new(2)).unwrap()}
fn semantic(n:&Node)->Value {
    match &n.kind {
        NodeKind::Row{entries,..}=>json!(["row",entries.iter().map(|e|semantic(&e.node)).collect::<Vec<_>>()]),
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
        NodeKind::Fraction{command,numerator,denominator,..}=>json!([command,semantic(numerator),semantic(denominator)]),
        NodeKind::Scripts{base,slots}=>json!(["scripts",semantic(base),slots.iter().map(|s|json!([s.script,semantic(&s.argument)])).collect::<Vec<_>>()]),
        NodeKind::Symbol{value,class,..}=>json!(["symbol",value,class]),
        NodeKind::Text{text,..}=>json!(["text",text]),
        NodeKind::CustomSymbol{command,asset_sha256,parts,..}=>json!(["customSymbol",command,asset_sha256,visualtex_engine_core::operators::effective_mode(parts),visualtex_engine_core::operators::slots(parts).map(|s|json!([s.script,semantic(&s.argument)])).collect::<Vec<_>>()]),
        NodeKind::EmptySlot{label}=>json!(["slot",label]),NodeKind::UnparsedSource{raw}=>json!(["opaque",raw]),
    }
}
struct Owned {s:EngineSession,seq:u32}
impl Owned {
    fn new(source:&str)->Self {Self{s:EngineSession::from_source("root-test".into(),source.into()).unwrap(),seq:0}}
    fn act(&mut self,a:EditorAction)->EditorActionResult {
        self.seq+=1; let before=self.s.document_snapshot();
        let r=self.s.apply_action(self.s.revision(),self.seq,a).unwrap();
        let index=TreeIndex::new(&r.editor.document.analysis.tree).unwrap();index.validate_selection(&r.editor.selection).unwrap();
        if let Some(map)=&r.position_map {for p in TreeIndex::new(&before.analysis.tree).unwrap().carets {map.apply_position(p,&index).unwrap();}}
        if r.editor.document.analysis.complete&&r.editor.document.analysis.supported {
            let reopened=parse(self.s.source());assert!(reopened.complete&&reopened.supported);
            assert_eq!(semantic(&r.editor.document.analysis.tree),semantic(&reopened.tree),"{}",self.s.source());
        } r
    }
    fn edit(&mut self,c:EditCommand)->EditorActionResult {self.act(EditorAction::Edit{command:c,group_id:None})}
    fn text(&mut self,t:&str) {self.edit(EditCommand::InsertText{text:t.into(),mode:InputMode::Math});}
    fn select(&mut self,p:Position) {self.act(EditorAction::SetSelection{selection:Selection::caret(p)});}
}
fn first_root(n:&Node)->&Node {if matches!(n.kind,NodeKind::Radical{..}) {return n;} n.children().into_iter().find_map(|c| {
    let mut has=false;c.visit(&mut |n|has|=matches!(n.kind,NodeKind::Radical{..}));has.then(||first_root(c))}).unwrap()}
fn all_ids(n:&Node)->BTreeSet<NodeId> {let mut s=BTreeSet::new();n.visit(&mut |n|{assert!(s.insert(n.id));});s}
fn fonts()->(FontRegistry,LayoutOptions) {
    let p=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    let mut f=FontRegistry::default();let id=f.register(&std::fs::read(p).unwrap(),0).unwrap().id;
    (f,LayoutOptions::new(id))
}
#[test]
fn square_and_indexed_roots_preserve_source_and_canonical_semantics() {
    for source in [r"\sqrt{x}",r"\sqrt x",r"\sqrt[3]{x_1}",r"\sqrt[{n+1}]{\frac{x}{y}}",r"\sqrt[{\sqrt[3]{x}}]{\sqrt{y}}",r"\sqrt[{]}]{x}","\\sqrt % keep\n [n % index\n +1] % body\n {x}"] {
        let p=parse(source);assert!(p.complete&&p.supported,"{source}: {:?}",p.diagnostics);
        assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);assert!(p.facts.radicals>0);
        let canonical=model::latex(&p.tree,LatexMode::Canonical);let q=parse(&canonical);
        assert!(q.complete&&q.supported,"{canonical}");assert_eq!(semantic(&p.tree),semantic(&q.tree));
        assert_eq!(model::latex(&q.tree,LatexMode::Canonical),canonical);
        assert_eq!(all_ids(&p.tree).len(),p.facts.nodes);
    }
}
#[test]
fn root_inside_an_index_protects_its_own_optional_closing_bracket() {
    let mut t=Owned::new(r"\sqrt[n]{x}");
    let p=t.s.document_snapshot();let NodeKind::Radical{degree:Some(d),..}=&first_root(&p.analysis.tree).kind else{panic!()};
    t.select(Position::row(d.id,1));
    t.edit(EditCommand::InsertStructure{structure:StructureKind::IndexedRoot});t.text("3");
    t.act(EditorAction::MoveHorizontal{direction:Horizontal::Right,extend:false});t.text("z");
    assert_eq!(t.s.source(),r"\sqrt[{n\sqrt[3]{z}}]{x}");
}
#[test]
fn incomplete_indices_empty_slots_and_unknown_commands_remain_distinct() {
    for source in ["\\sqrt", "\\sqrt[", "\\sqrt[3", "\\sqrt[3]", "\\sqrt{", "{\\sqrt[n}"] {
        let p=parse(source);assert!(!p.complete&&!p.exportable,"{source}");assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
    }
    for source in [r"\sqrt{}",r"\sqrt[]{x}"] {let p=parse(source);assert!(p.complete&&p.supported);assert!(!p.exportable);}
    let p=parse(r"\sqrt{\future{x}}");assert!(p.complete&&!p.supported&&!p.exportable);
    for source in [r"\sqrt[n\sqrt[3]{z}]{x}",r"\sqrt[\frac]2]{x}"] {
        let p=parse(source);assert!(!p.complete&&!p.exportable,"Unprotected optional delimiter was accepted: {source}");
        assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
    }
}
#[test]
fn indexed_root_constructs_editable_degree_and_radicand_with_navigation() {
    let mut t=Owned::new("");t.edit(EditCommand::InsertStructure{structure:StructureKind::IndexedRoot});t.text("3");
    t.act(EditorAction::MoveHorizontal{direction:Horizontal::Right,extend:false});t.text("x");
    t.edit(EditCommand::InsertStructure{structure:StructureKind::Superscript});t.text("2");
    assert_eq!(t.s.source(),r"\sqrt[3]{x^{2}}");
    let p=t.s.document_snapshot();assert_eq!(p.analysis.facts.radicals,1);assert!(p.analysis.exportable);
}
#[test]
fn editing_degree_right_bracket_adds_only_required_group_and_survives_reopen() {
    let mut t=Owned::new(r"\sqrt[n]{x}");
    let p=t.s.document_snapshot();let NodeKind::Radical{degree:Some(d),..}=&first_root(&p.analysis.tree).kind else{panic!()};
    let id=d.id;t.select(Position::row(id,1));t.text("]");
    assert_eq!(t.s.source(),r"\sqrt[{n]}]{x}");
    t.text("+1");assert_eq!(t.s.source(),r"\sqrt[{n]+1}]{x}");
    let p=t.s.document_snapshot();let NodeKind::Radical{degree:Some(d),..}=&first_root(&p.analysis.tree).kind else{panic!()};assert_eq!(d.id,id);
    t.act(EditorAction::Undo);t.act(EditorAction::Undo);assert_eq!(t.s.source(),r"\sqrt[n]{x}");
}
#[test]
fn source_patch_preserves_untouched_degree_and_other_root_identity() {
    let mut t=Owned::new(r"\sqrt[3]{x}+\sqrt[3]{x}");let before=t.s.document_snapshot();
    let NodeKind::Row{entries,..}=&before.analysis.tree.kind else{panic!()};let other=all_ids(&entries[2].node);
    let NodeKind::Radical{degree:Some(d),..}=&entries[0].node.kind else{panic!()};let d=all_ids(d);
    let at=t.s.source().find("{x}").unwrap()+1;
    t.s.patch_source(t.s.revision(),SourceEdit{start_utf16:at as u32,end_utf16:at as u32+1,replacement:"q+1".into()}).unwrap();
    let after=t.s.document_snapshot();let NodeKind::Row{entries,..}=&after.analysis.tree.kind else{panic!()};assert_eq!(all_ids(&entries[2].node),other);
    let NodeKind::Radical{degree:Some(now),..}=&entries[0].node.kind else{panic!()};assert_eq!(all_ids(now),d);
    assert_eq!(semantic(&after.analysis.tree),semantic(&parse(t.s.source()).tree));
}
#[test]
fn root_copy_delete_and_undo_preserve_independent_node_identity() {
    let mut t=Owned::new(r"\sqrt[3]{x}");let root=t.s.document_snapshot().analysis.tree.id;
    t.act(EditorAction::SetSelection{selection:Selection::Range{anchor:Position::row(root,0),focus:Position::row(root,1)}});
    t.edit(EditCommand::DuplicateSelection);
    let p=t.s.document_snapshot();let NodeKind::Row{entries,..}=&p.analysis.tree.kind else{panic!()};
    let original=all_ids(&entries[0].node);assert!(original.is_disjoint(&all_ids(&entries[1].node)));
    t.select(Position::row(root,2));t.edit(EditCommand::DeleteBackward);
    assert!(matches!(t.s.editor_snapshot().unwrap().selection,Selection::Node{..}));
    t.edit(EditCommand::DeleteBackward);assert_eq!(t.s.source(),r"\sqrt[3]{x}");
    t.act(EditorAction::Undo);let restored=t.s.document_snapshot();assert_eq!(all_ids(&restored.analysis.tree),all_ids(&p.analysis.tree));
}
#[test]
fn root_uses_font_rule_gaps_and_cramped_base_without_clipping() {
    let (f,o)=fonts();let mut l=LayoutEngine::default();
    for source in [r"\sqrt{x}",r"\sqrt[n+1]{x^2}",r"\sqrt{\frac{1}{\frac{2}{\frac{3}{4}}}}"] {
        let t=Owned::new(source);let scene=l.layout(&t.s,&f,&o).unwrap();scene.validate().unwrap();assert!(scene.export_ready);
        let p=t.s.document_snapshot();let root=first_root(&p.analysis.tree);let NodeKind::Radical{radicand,..}=&root.kind else{panic!()};
        let bar=scene.rules.iter().find(|r|r.node_id==root.id).unwrap();
        let body=scene.regions.iter().find(|r|r.node_id==radicand.id).unwrap();let metrics=f.get(&o.math_font).unwrap().math.as_ref().unwrap();
        assert!((bar.rect.height()-metrics.radical_rule).abs()<1e-9);
        assert!(body.ink.unwrap().y_min-bar.rect.y_max+1e-9>=metrics.radical_display_gap);
        let glyphs:Vec<_>=scene.glyphs.iter().filter(|g|g.node_id==root.id).collect();assert!(!glyphs.is_empty());
        assert!(glyphs.iter().all(|g|(g.scale-1.0).abs()<1e-9));
        assert!(scene.svg().unwrap().contains("<svg"));
    }
}
#[test]
fn seeded_nested_root_corpus_and_source_edits_match_fresh_parser() {
    let mut seed=0x7a0711u32;
    fn next(seed:&mut u32)->u32{*seed=seed.wrapping_mul(1664525).wrapping_add(1013904223);*seed}
    fn formula(seed:&mut u32,depth:u32)->String {
        if depth==0{return ["x","n","3","y"][next(seed) as usize%4].into();}
        let a=formula(seed,depth-1);let b=formula(seed,depth-1);
        match next(seed)%4 {
            0=>format!("\\sqrt{{{a}}}"),1=>format!("\\sqrt[{{{a}}}]{{{b}}}"),
            2=>format!("\\frac{{{a}}}{{{b}}}"),_=>format!("{{{a}}}^{{{b}}}"),
        }
    }
    for case in 0..200 {
        let source=formula(&mut seed,3);let p=parse(&source);assert!(p.complete&&p.supported,"seed case {case}");
        let canonical=model::latex(&p.tree,LatexMode::Canonical);
        assert_eq!(semantic(&p.tree),semantic(&parse(&canonical).tree),"case {case}");
        let mut t=Owned::new(&source);
        for iteration in 0..6 {
            let at=next(&mut seed) as usize%(t.s.source().len()+1);
            let replacement=["x","}","[","]","%"," "][next(&mut seed) as usize%6];
            t.s.patch_source(t.s.revision(),SourceEdit{start_utf16:at as u32,end_utf16:at as u32,replacement:replacement.into()}).unwrap();
            let p=t.s.document_snapshot();let q=parse(t.s.source());
            assert_eq!(semantic(&p.analysis.tree),semantic(&q.tree),"case={case}, iteration={iteration}, source={}",t.s.source());
            assert_eq!((p.analysis.complete,p.analysis.supported),(q.complete,q.supported));
            assert_eq!(model::latex(&p.analysis.tree,LatexMode::Lossless),t.s.source());
            assert_eq!(all_ids(&p.analysis.tree).len(),p.analysis.facts.nodes);
        }
    }
}
#[test]
fn display_and_script_roots_reuse_fixed_font_geometry_at_changed_pixel_size() {
    let (f,mut o)=fonts();let t=Owned::new(r"\sqrt[3]{x^2}");let mut l=LayoutEngine::default();
    for style in [MathStyle::Display,MathStyle::Text,MathStyle::Script,MathStyle::ScriptScript] {
        o.style=style;o.font_size_px=16.0;let a=l.layout(&t.s,&f,&o).unwrap();o.font_size_px=32.0;let b=l.layout(&t.s,&f,&o).unwrap();
        assert!((b.width-a.width*2.0).abs()<1e-8);assert_eq!(b.rebuilt_fragments,0);
        assert_eq!(serde_json::to_value(&a.glyphs).unwrap(),serde_json::to_value(&b.glyphs).unwrap());
    }
}
