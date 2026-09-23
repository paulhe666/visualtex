//! Unicode-math aliases are real Style nodes, not old-runtime macro calls.
use std::{collections::BTreeSet, path::PathBuf};
use visualtex_engine_core::{
    editing::EditorAction, fonts::FontRegistry, layout::{LayoutEngine, LayoutOptions, MathStyle},
    model::{self, IdAllocator, LatexMode, Node, NodeKind}, parser,
    positions::{Position, Selection, TreeIndex}, session::{EngineSession, NodeId},
    source::SourceText, source_edit::SourceEdit,
    styles::{self, MathVariant, STYLE_ALIASES}, transactions::{EditCommand, InputMode},
};
fn fonts() -> (FontRegistry, LayoutOptions) {
    let path = std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    let mut fonts = FontRegistry::default(); let id = fonts.register(&std::fs::read(path).unwrap(), 0).unwrap().id;
    let fallback_path = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../target/test-resources/stix-two-2.13/STIXTwoMath-Regular.otf");
    let fallback=fonts.register(&std::fs::read(fallback_path).unwrap(),0).unwrap();
    assert_eq!(fallback.sha256,"f2076b9f1676438439dd41e23676f5ab99056e83d6b8f8c27841591ef2ccfa72");
    let mut options=LayoutOptions::new(id);options.math_fallbacks.push(fallback.id);
    (fonts, options)
}
fn ids(n: &Node) -> BTreeSet<NodeId> {
    let mut all = BTreeSet::new(); n.visit(&mut |n| { assert!(all.insert(n.id)); }); all
}
fn parse(text: &str) -> parser::ParseResult {
    parser::parse(SourceText::new(1, text).unwrap(), NodeId(1), &mut IdAllocator::new(2)).unwrap()
}
fn action(s: &mut EngineSession, seq: &mut u32, a: EditorAction) {
    *seq += 1; let r = s.apply_action(s.revision(), *seq, a).unwrap();
    TreeIndex::new(&r.editor.document.analysis.tree).unwrap().validate_selection(&r.editor.selection).unwrap();
    assert_eq!(ids(&r.editor.document.analysis.tree).len(), r.editor.document.analysis.facts.nodes);
}
#[test]
fn every_variant_has_a_canonical_command_that_preserves_the_variant() {
    use MathVariant::*;
    for variant in [Default,Upright,Bold,Italic,BoldItalic,Sans,SansBold,SansItalic,SansBoldItalic,
        Monospace,DoubleStruck,Script,BoldScript,Fraktur,BoldFraktur] {
        assert_eq!(styles::command_variant(styles::canonical_command(variant)), Some(variant));
    }
}
#[test]
fn supported_aliases_preserve_spelling_grouping_comments_and_actual_glyphs() {
    let (fonts, mut options) = fonts(); let mut names = BTreeSet::new();
    for &(command, variant) in STYLE_ALIASES {
        assert!(names.insert(command)); assert_eq!(styles::command_variant(command), Some(variant));
        let source = format!("{command} % retain\n {{Az09}}+x"); let p = parse(&source);
        assert!(p.complete && p.supported && p.exportable, "{source}");
        assert_eq!(model::latex(&p.tree, LatexMode::Lossless), source);
        let canonical = model::latex(&p.tree, LatexMode::Canonical);
        assert_eq!(model::latex(&parse(&canonical).tree, LatexMode::Canonical), canonical);
        let session = EngineSession::from_source("style-alias-font".into(), source).unwrap();
        for style in [MathStyle::Display, MathStyle::Text, MathStyle::Script, MathStyle::ScriptScript] {
            options.style = style; let scene = LayoutEngine::default().layout(&session, &fonts, &options).unwrap();
            let mapped = styles::map_character('A', variant);
            let expected = fonts.get(&options.math_font).unwrap().glyph_id(mapped).unwrap();
            assert_eq!(scene.resources[scene.glyphs[0].resource].shape.glyph_id, expected, "{command}");
            let last = scene.glyphs.last().unwrap();
            let default_x = fonts.get(&options.math_font).unwrap().glyph_id('𝑥').unwrap();
            assert_eq!(scene.resources[last.resource].shape.glyph_id, default_x, "{command} leaked outside its scope");
            scene.validate().unwrap(); scene.svg().unwrap();
        }
    }
    assert_eq!(names.len(), 17);
}
#[test]
fn greek_variant_symbols_keep_their_distinct_unicode_math_codepoints() {
    // Each offset is specified by the Unicode Mathematical Alphanumeric block.
    let symbols = [('ϴ',17),('∇',25),('∂',51),('ϵ',52),('ϑ',53),('ϰ',54),('ϕ',55),('ϱ',56),('ϖ',57)];
    let alphabets = [(MathVariant::Bold,0x1D6A8),(MathVariant::Italic,0x1D6E2),
        (MathVariant::BoldItalic,0x1D71C),(MathVariant::SansBold,0x1D756),(MathVariant::SansBoldItalic,0x1D790)];
    let (fonts, options) = fonts();
    for (variant, base) in alphabets {
        for (symbol, offset) in symbols {
            let mapped = styles::map_character(symbol, variant);
            assert_eq!(mapped as u32, base+offset);
            assert!(fonts.get(&options.math_font).unwrap().glyph_id(mapped).is_ok());
        }
        assert_ne!(styles::map_character('φ',variant), styles::map_character('ϕ',variant));
        assert_ne!(styles::map_character('κ',variant), styles::map_character('ϰ',variant));
    }
}
#[test]
fn aliases_participate_in_structure_copy_delete_and_single_history_transactions() {
    for &(command, variant) in STYLE_ALIASES {
        let mut s = EngineSession::from_source("alias-action".into(), "x+y".into()).unwrap(); let mut seq = 0;
        let root = s.document_snapshot().analysis.tree.id;
        let selected = Selection::Range {anchor:Position::row(root,0),focus:Position::row(root,3)};
        action(&mut s,&mut seq,EditorAction::SetSelection{selection:selected});
        let original = s.document_snapshot();
        action(&mut s,&mut seq,EditorAction::Edit{command:EditCommand::SetStyle{style:command.into()},group_id:None});
        let styled = s.document_snapshot(); assert_eq!(styled.source,format!("{command}{{x+y}}"));
        let NodeKind::Row { entries, .. } = &styled.analysis.tree.kind else {panic!()};
        assert!(matches!(entries[0].node.kind,NodeKind::Style{variant:v,..} if v==variant));
        let style_id = entries[0].node.id;
        action(&mut s,&mut seq,EditorAction::Undo); assert_eq!(s.source(),original.source);
        assert_eq!(ids(&s.document_snapshot().analysis.tree),ids(&original.analysis.tree));
        action(&mut s,&mut seq,EditorAction::Redo); assert_eq!(ids(&s.document_snapshot().analysis.tree),ids(&styled.analysis.tree));
        action(&mut s,&mut seq,EditorAction::SetSelection{selection:Selection::Node{node_id:style_id}});
        action(&mut s,&mut seq,EditorAction::Edit{command:EditCommand::DuplicateSelection,group_id:None});
        let copied = s.document_snapshot();let NodeKind::Row{entries,..}=&copied.analysis.tree.kind else {panic!()};
        assert!(ids(&entries[0].node).is_disjoint(&ids(&entries[1].node)));
        action(&mut s,&mut seq,EditorAction::Edit{command:EditCommand::DeleteBackward,group_id:None});
        assert_eq!(s.source(),styled.source);
        action(&mut s,&mut seq,EditorAction::Undo);assert_eq!(ids(&s.document_snapshot().analysis.tree),ids(&copied.analysis.tree));
    }
}
#[test]
fn blank_alias_body_is_editable_and_unknown_legacy_declarations_are_not_misparsed() {
    for &(command,_) in STYLE_ALIASES {
        let mut s=EngineSession::new("alias-empty".into()).unwrap();let mut seq=0;
        action(&mut s,&mut seq,EditorAction::Edit{command:EditCommand::SetStyle{style:command.into()},group_id:None});
        assert!(!s.document_snapshot().analysis.exportable);
        action(&mut s,&mut seq,EditorAction::Edit{command:EditCommand::InsertText{text:"x".into(),mode:InputMode::Math},group_id:None});
        assert_eq!(s.source(),format!("{command}{{x}}"));assert!(s.document_snapshot().analysis.exportable);
        for suffix in ["", "{"] {let source=format!("{command}{suffix}");let p=parse(&source);
            assert!(!p.complete && !p.exportable);assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);}
    }
    for source in [r"{\bf x+y}",r"\simbfit{x}",r"\symUnknown{x}"] {
        let p=parse(source);assert!(!p.supported && !p.exportable);assert_eq!(model::latex(&p.tree,LatexMode::Lossless),source);
    }
}
#[test]
fn seeded_alias_source_edits_preserve_fresh_parse_and_unrelated_nodes() {
    let seed=0x57a12026u32;let mut state=seed;
    for iteration in 0..340 {
        state=state.wrapping_mul(1664525).wrapping_add(1013904223);
        let (command,_)=STYLE_ALIASES[state as usize % STYLE_ALIASES.len()];
        let source=format!("\\frac{{{command}{{x}}}}{{{command}{{y}}}}+z");
        let mut s=EngineSession::from_source("alias-seeded".into(),source.clone()).unwrap();
        let before=s.document_snapshot();let NodeKind::Row{entries,..}=&before.analysis.tree.kind else{panic!()};
        let tail=entries.last().unwrap().node.id;
        let at=source.find("{x}").unwrap()+1;let replacement=char::from(b'a'+(state%26) as u8).to_string();
        let (_,map)=s.patch_source(s.revision(),SourceEdit{start_utf16:at as u32,end_utf16:(at+1) as u32,replacement}).unwrap();
        let after=s.document_snapshot();let fresh=parse(s.source());
        assert_eq!(model::latex(&after.analysis.tree,LatexMode::Canonical),model::latex(&fresh.tree,LatexMode::Canonical),"seed={seed:x} step={iteration}");
        assert!(map.retained_ids.contains(&tail));assert_eq!(ids(&after.analysis.tree).len(),after.analysis.facts.nodes);
    }
}
