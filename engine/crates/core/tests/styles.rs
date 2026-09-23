//! P07.5 mathematical alphabets and upright function operators.
use std::{collections::BTreeSet, path::PathBuf};
use visualtex_engine_core::{
    editing::{EditorAction, EditorActionResult},
    fonts::FontRegistry,
    layout::{LayoutEngine, LayoutOptions, MathStyle},
    model::{self, AtomClass, IdAllocator, LatexMode, Node, NodeKind},
    parser::{self, ParseResult},
    positions::{Position, Selection},
    session::{EngineSession, NodeId},
    source::SourceText,
    source_edit::SourceEdit,
    styles::{self, MathVariant},
    transactions::{EditCommand, InputMode},
};

fn parse(source: &str) -> ParseResult {
    parser::parse(SourceText::new(1, source).unwrap(), NodeId(1), &mut IdAllocator::new(2)).unwrap()
}
fn fonts() -> (FontRegistry, LayoutOptions) {
    let path=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    let mut fonts=FontRegistry::default();let id=fonts.register(&std::fs::read(path).unwrap(),0).unwrap().id;
    (fonts,LayoutOptions::new(id))
}
fn first_style(node: &Node) -> &Node {
    if matches!(node.kind, NodeKind::Style { .. }) { return node; }
    node.children().into_iter().find_map(|child| {
        let mut found=false;child.visit(&mut |n|found |= matches!(n.kind,NodeKind::Style{..}));found.then(||first_style(child))
    }).unwrap()
}
fn ids(node:&Node)->BTreeSet<NodeId>{let mut all=BTreeSet::new();node.visit(&mut |n|{assert!(all.insert(n.id));});all}

#[test]
fn all_product_style_commands_preserve_authored_spelling_and_roundtrip() {
    for (command, variant) in [
        (r"\mathnormal",MathVariant::Default),(r"\mathrm",MathVariant::Upright),(r"\mathbf",MathVariant::Bold),
        (r"\mathit",MathVariant::Italic),(r"\mathbfit",MathVariant::BoldItalic),(r"\boldsymbol",MathVariant::BoldItalic),
        (r"\bm",MathVariant::BoldItalic),(r"\mathsf",MathVariant::Sans),(r"\mathtt",MathVariant::Monospace),
        (r"\mathbb",MathVariant::DoubleStruck),(r"\mathcal",MathVariant::Script),(r"\mathscr",MathVariant::Script),
        (r"\mathfrak",MathVariant::Fraktur),
    ] {
        let source=format!("{command} % style\r\n {{A1x+\\alpha}}+z");
        let parsed=parse(&source);assert!(parsed.complete&&parsed.supported&&parsed.exportable,"{source}: {:?}",parsed.diagnostics);
        assert_eq!(model::latex(&parsed.tree,LatexMode::Lossless),source);
        assert_eq!(parsed.facts.styles,1);
        let NodeKind::Style{command:seen,variant:seen_variant,body,..}=&first_style(&parsed.tree).kind else{panic!()};
        assert_eq!(seen,command);assert_eq!(*seen_variant,variant);assert!(body.is_row());
        let canonical=model::latex(&parsed.tree,LatexMode::Canonical);let reopened=parse(&canonical);
        assert!(reopened.complete&&reopened.supported);assert_eq!(model::latex(&reopened.tree,LatexMode::Canonical),canonical);
        let NodeKind::Style{variant:again,..}=&first_style(&reopened.tree).kind else{panic!()};assert_eq!(*again,variant);
    }
    for source in [r"\mathbf",r"\mathrm{",r"\mathbb{}"] {
        let parsed=parse(source);assert!(!parsed.exportable);assert_eq!(model::latex(&parsed.tree,LatexMode::Lossless),source);
    }
}

#[test]
fn common_named_functions_are_upright_operator_atoms_not_implicit_text() {
    for (command,name) in [(r"\sin","sin"),(r"\cos","cos"),(r"\tan","tan"),(r"\log","log"),(r"\ln","ln"),(r"\exp","exp"),
        (r"\det","det"),(r"\gcd","gcd"),(r"\arg","arg"),(r"\ker","ker"),(r"\dim","dim"),(r"\hom","hom")] {
        let source=format!("{command} x");let p=parse(&source);assert!(p.complete&&p.supported&&p.exportable,"{source}");
        let NodeKind::Row{entries,..}=&p.tree.kind else{panic!()};
        let NodeKind::Symbol{value,class,command:Some(spelling)}=&entries[0].node.kind else{panic!("{source}")};
        assert_eq!(value,name);assert_eq!(*class,AtomClass::Operator);assert_eq!(spelling,command);
        assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
    }
}

#[test]
fn operatorname_is_editable_upright_math_text_and_preserves_source() {
    for source in [r"\operatorname{tr} A",r"\operatorname{rank}(A)",r"\operatorname{foo}_i"] {
        let p=parse(source);assert!(p.complete&&p.supported&&p.exportable,"{source}: {:?}",p.diagnostics);
        assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
        let mut operator_name=None;p.tree.visit(&mut |node|if matches!(&node.kind,NodeKind::Text{command,..} if command==r"\operatorname"){operator_name=Some(node.id);});
        let node=p.tree.find(operator_name.unwrap()).unwrap();let NodeKind::Text{command,text,..}=&node.kind else{panic!()};
        assert_eq!(command,r"\operatorname");assert!(!text.is_empty());
        if source.ends_with("_i") {assert_eq!(p.facts.scripts,1);}
    }
    for source in [r"\operatorname",r"\operatorname{"] {let p=parse(source);assert!(!p.complete&&!p.exportable);assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);}
}

#[cfg(target_os="windows")]
#[test]
fn product_font_faces_own_normal_symbols_and_common_math_styles() {
    let math_path=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    let windows=PathBuf::from(std::env::var_os("WINDIR").unwrap_or_else(||"C:\\Windows".into())).join("Fonts");
    let mut fonts=FontRegistry::default();
    let math=fonts.register(&std::fs::read(math_path).unwrap(),0).unwrap();
    let regular=fonts.register(&std::fs::read(windows.join("times.ttf")).unwrap(),0).unwrap();
    let italic=fonts.register(&std::fs::read(windows.join("timesi.ttf")).unwrap(),0).unwrap();
    let bold=fonts.register(&std::fs::read(windows.join("timesbd.ttf")).unwrap(),0).unwrap();
    let bold_italic=fonts.register(&std::fs::read(windows.join("timesbi.ttf")).unwrap(),0).unwrap();
    let mut options=LayoutOptions::new(math.id.clone());options.style=MathStyle::Text;
    options.alphabet_upright_font=Some(regular.id.clone());
    options.alphabet_italic_font=Some(italic.id.clone());
    options.alphabet_bold_font=Some(bold.id.clone());
    options.alphabet_bold_italic_font=Some(bold_italic.id.clone());
    let source=r"x+1=\alpha+\partial+\mathbf{A}+\mathit{B}+\mathbfit{C}+\mathcal{D}";
    let session=EngineSession::from_source("product-font-faces".into(),source.into()).unwrap();
    let scene=LayoutEngine::default().layout(&session,&fonts,&options).unwrap();
    let used=scene.resources.iter().map(|r|r.font_id.as_str()).collect::<BTreeSet<_>>();
    for id in [&regular.id,&italic.id,&bold.id,&bold_italic.id,&math.id] {
        assert!(used.contains(id.as_str()),"expected font resource {id} was not used");
    }
    let has=|font_id:&str,ch:char|{
        let gid=fonts.get(font_id).unwrap().glyph_id(ch).unwrap();
        scene.resources.iter().any(|resource|resource.font_id==font_id&&resource.shape.glyph_id==gid)
    };
    assert!(has(&italic.id,'x'));
    assert!(has(&italic.id,'α'));
    assert!(has(&regular.id,'+'));
    assert!(has(&regular.id,'='));
    assert!(has(&regular.id,'∂'));
    assert!(has(&regular.id,'1'));
    assert!(has(&bold.id,'A'));
    assert!(has(&italic.id,'B'));
    assert!(has(&bold_italic.id,'C'));
    let script_d=styles::map_character('D',MathVariant::Script);
    assert!(has(&math.id,script_d),"semantic \\mathcal must retain its dedicated mathematical alphabet");

    let fenced=EngineSession::from_source("product-font-fence".into(),r"\left(x\right)".into()).unwrap();
    let fenced_scene=LayoutEngine::default().layout(&fenced,&fonts,&options).unwrap();
    for ch in ['(',')'] {
        let gid=fonts.get(&regular.id).unwrap().glyph_id(ch).unwrap();
        assert!(fenced_scene.resources.iter().any(|resource|resource.font_id==regular.id&&resource.shape.glyph_id==gid),
            "a normal-size authored fence should use the selected product family for {ch}");
    }
}

#[cfg(target_os="windows")]
#[test]
fn every_installed_product_family_uses_its_own_four_faces_for_common_math() {
    let math_path=std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    let windows=PathBuf::from(std::env::var_os("WINDIR").unwrap_or_else(||"C:\\Windows".into())).join("Fonts");
    for (family,files) in [
        ("times",["times.ttf","timesi.ttf","timesbd.ttf","timesbi.ttf"]),
        ("palatino",["pala.ttf","palai.ttf","palab.ttf","palabi.ttf"]),
        ("arial",["arial.ttf","ariali.ttf","arialbd.ttf","arialbi.ttf"]),
    ] {
        let mut fonts=FontRegistry::default();
        let math=fonts.register(&std::fs::read(&math_path).unwrap(),0).unwrap();
        let regular=fonts.register(&std::fs::read(windows.join(files[0])).unwrap(),0).unwrap();
        let italic=fonts.register(&std::fs::read(windows.join(files[1])).unwrap(),0).unwrap();
        let bold=fonts.register(&std::fs::read(windows.join(files[2])).unwrap(),0).unwrap();
        let bold_italic=fonts.register(&std::fs::read(windows.join(files[3])).unwrap(),0).unwrap();
        let mut options=LayoutOptions::new(math.id.clone());options.style=MathStyle::Text;
        options.alphabet_upright_font=Some(regular.id.clone());options.alphabet_italic_font=Some(italic.id.clone());
        options.alphabet_bold_font=Some(bold.id.clone());options.alphabet_bold_italic_font=Some(bold_italic.id.clone());
        let session=EngineSession::from_source(format!("family-{family}"),r"x+1=\alpha+\partial+\mathbf{A}+\mathbfit{B}+\left(x\right)+\sqrt{x}".into()).unwrap();
        let scene=LayoutEngine::default().layout(&session,&fonts,&options).unwrap();
        let has=|font_id:&str,ch:char|{
            let gid=fonts.get(font_id).unwrap().glyph_id(ch).unwrap();
            scene.resources.iter().any(|resource|resource.font_id==font_id&&resource.shape.glyph_id==gid)
        };
        assert!(has(&italic.id,'x'),"{family}: variable x is not from italic face");
        assert!(has(&italic.id,'α'),"{family}: lowercase Greek is not from italic face");
        for ch in ['+','=','∂','1','(',')','√'] {assert!(has(&regular.id,ch),"{family}: {ch} is not from regular face");}
        assert!(has(&bold.id,'A'),"{family}: bold A is not from bold face");
        assert!(has(&bold_italic.id,'B'),"{family}: bold italic B is not from bold italic face");
    }
}

#[test]
fn fixed_math_font_uses_real_variant_glyphs_not_css_or_outline_scaling() {
    let (fonts,mut options)=fonts();options.style=MathStyle::Text;
    for (source, expected) in [(r"\mathbf{A}",'𝐀'),(r"\mathrm{A}",'A'),(r"\mathit{x}",'𝑥'),(r"\mathbb{R}",'ℝ'),
        (r"\mathcal{A}",'𝒜'),(r"\mathfrak{H}",'ℌ'),(r"\boldsymbol{\alpha}",'𝜶'),(r"\bm{\Gamma}",'𝜞')] {
        let session=EngineSession::from_source(format!("style-{source}"),source.into()).unwrap();
        let mut layout=LayoutEngine::default();let scene=layout.layout(&session,&fonts,&options).unwrap();
        let expected_id=fonts.get(&options.math_font).unwrap().glyph_id(expected).unwrap();
        assert!(scene.glyphs.iter().any(|g|scene.resources[g.resource].shape.glyph_id==expected_id),"{source} did not use U+{:04X}",u32::from(expected));
        assert!(scene.svg().unwrap().contains("<path"));
    }
}

struct Owned { session:EngineSession, sequence:u32 }
impl Owned {
    fn new(source:&str)->Self{Self{session:EngineSession::from_source("style-edit".into(),source.into()).unwrap(),sequence:0}}
    fn action(&mut self,action:EditorAction)->EditorActionResult{self.sequence+=1;self.session.apply_action(self.session.revision(),self.sequence,action).unwrap()}
    fn edit(&mut self,command:EditCommand)->EditorActionResult{self.action(EditorAction::Edit{command,group_id:None})}
    fn select(&mut self,selection:Selection){self.action(EditorAction::SetSelection{selection});}
}

#[test]
fn set_style_is_a_structural_edit_with_shared_history_and_fresh_wrapper_identity() {
    let mut t=Owned::new("x+y");let root=t.session.document_snapshot().analysis.tree.id;
    t.select(Selection::Range{anchor:Position::row(root,0),focus:Position::row(root,3)});
    let styled=t.edit(EditCommand::SetStyle{style:r"\mathbf".into()});
    assert_eq!(t.session.source(),r"\mathbf{x+y}");assert!(styled.content_changed);
    let style_id=first_style(&styled.editor.document.analysis.tree).id;
    t.action(EditorAction::Undo);assert_eq!(t.session.source(),"x+y");
    t.action(EditorAction::Redo);assert_eq!(t.session.source(),r"\mathbf{x+y}");
    assert_eq!(first_style(&t.session.document_snapshot().analysis.tree).id,style_id);
    assert_eq!(ids(&t.session.document_snapshot().analysis.tree).len(),t.session.document_snapshot().analysis.facts.nodes);
}

#[test]
fn collapsed_style_command_creates_an_editable_slot_and_typing_fills_it() {
    let mut t=Owned::new("");
    let styled=t.edit(EditCommand::SetStyle{style:r"\mathrm".into()});
    assert_eq!(t.session.source(),r"\mathrm{}");assert!(!styled.editor.document.analysis.exportable);
    t.edit(EditCommand::InsertText{text:"d".into(),mode:InputMode::Math});
    assert_eq!(t.session.source(),r"\mathrm{d}");assert!(t.session.document_snapshot().analysis.exportable);
}

#[test]
fn source_patch_inside_style_retains_wrapper_and_unaffected_sibling_identity() {
    let mut session=EngineSession::from_source("style-source".into(),r"\mathbf{x}+z".into()).unwrap();
    let before=session.document_snapshot();let style=first_style(&before.analysis.tree).id;
    let NodeKind::Row{entries,..}=&before.analysis.tree.kind else{panic!()};let z=entries.last().unwrap().node.id;
    let at=u32::try_from(session.source().find('x').unwrap()).unwrap();
    let (_,map)=session.patch_source(session.revision(),SourceEdit{start_utf16:at,end_utf16:at+1,replacement:"y".into()}).unwrap();
    let after=session.document_snapshot();assert_eq!(first_style(&after.analysis.tree).id,style);
    let NodeKind::Row{entries,..}=&after.analysis.tree.kind else{panic!()};assert_eq!(entries.last().unwrap().node.id,z);
    assert!(map.retained_ids.contains(&style)&&map.retained_ids.contains(&z));
    let NodeKind::Style{body,..}=&first_style(&after.analysis.tree).kind else{panic!()};
    let NodeKind::Row{entries,..}=&body.kind else{panic!()};
    assert!(matches!(&entries[0].node.kind,NodeKind::Symbol{value,..} if value=="y"),"Source changed but styled math still contains old x");
    after.analysis.tree.visit(&mut |n|{
        let origin=n.origin.as_ref().unwrap();assert_eq!(origin.span.source_id,after.source_id);
        assert_eq!(origin.text(),&after.source[origin.span.byte_start as usize..origin.span.byte_end as usize]);
    });
}

#[test]
fn explicit_variants_do_not_leak_outside_their_style_node_or_cache_key() {
    let (fonts,options)=fonts();let mut layout=LayoutEngine::default();
    let bold=EngineSession::from_source("cache-style".into(),r"\mathbf{x}+x".into()).unwrap();
    let scene=layout.layout(&bold,&fonts,&options).unwrap();
    let bold_id=fonts.get(&options.math_font).unwrap().glyph_id(styles::map_character('x',MathVariant::Bold)).unwrap();
    let italic_id=fonts.get(&options.math_font).unwrap().glyph_id(styles::default_math_character('x')).unwrap();
    let glyph_ids:Vec<_>=scene.glyphs.iter().map(|g|scene.resources[g.resource].shape.glyph_id).collect();
    assert!(glyph_ids.contains(&bold_id));assert!(glyph_ids.contains(&italic_id));assert_ne!(bold_id,italic_id);
}
