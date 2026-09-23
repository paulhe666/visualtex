//! P07.8 authored atoms and Unicode integrals, independent of the old renderer.
use std::{collections::BTreeSet, path::PathBuf};
use serde_json::Value;
use visualtex_engine_core::{
    editing::{EditorAction, EditorActionResult}, fonts::FontRegistry,
    layout::{LayoutEngine, LayoutOptions, MathStyle},
    model::{self, AtomClass, IdAllocator, LatexMode, Node, NodeKind},
    operators::{self, LimitMode, OperatorMark}, parser::{self, ParseResult},
    positions::{Position, Selection, TreeIndex}, rare_operators, symbol_catalog,
    session::{EngineSession, NodeId}, source::SourceText, source_edit::SourceEdit,
    symbols::{literal_class, named_symbol}, transactions::{EditCommand, InputMode}, ErrorCode,
};
fn parse(s: &str) -> ParseResult {
    parser::parse(SourceText::new(1, s).unwrap(), NodeId(1), &mut IdAllocator::new(2)).unwrap()
}
fn atoms(root: &Node) -> Vec<(String, AtomClass)> {
    let mut out = Vec::new();
    root.visit(&mut |n| if let NodeKind::Symbol { value, class, .. } = &n.kind { out.push((value.clone(), *class)); });
    out
}
fn identities(root: &Node) -> BTreeSet<NodeId> {
    let mut ids = BTreeSet::new(); root.visit(&mut |n| { assert!(ids.insert(n.id)); }); ids
}
fn fonts() -> (FontRegistry, LayoutOptions) {
    let path = std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from).unwrap_or_else(||
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    let mut f = FontRegistry::default(); let id = f.register(&std::fs::read(path).unwrap(), 0).unwrap().id;
    let fallback_path = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../target/test-resources/stix-two-2.13/STIXTwoMath-Regular.otf");
    let bytes = std::fs::read(fallback_path).expect("Run node tools/prepare-symbol-test-font.mjs before catalog font tests");
    let fallback = f.register(&bytes, 0).unwrap();
    assert_eq!(fallback.sha256, "f2076b9f1676438439dd41e23676f5ab99056e83d6b8f8c27841591ef2ccfa72");
    let mut options = LayoutOptions::new(id); options.math_fallbacks.push(fallback.id);
    (f, options)
}
fn action(s: &mut EngineSession, seq: &mut u32, a: EditorAction) -> EditorActionResult {
    *seq += 1; let before = s.document_snapshot();
    let result = s.apply_action(s.revision(), *seq, a).unwrap();
    let index = TreeIndex::new(&result.editor.document.analysis.tree).unwrap();
    index.validate_selection(&result.editor.selection).unwrap();
    if let Some(map) = &result.position_map {
        for p in TreeIndex::new(&before.analysis.tree).unwrap().carets { map.apply_position(p, &index).unwrap(); }
    }
    let reopened = parse(s.source());
    assert_eq!(atoms(&reopened.tree), atoms(&result.editor.document.analysis.tree));
    assert_eq!(identities(&result.editor.document.analysis.tree).len(), result.editor.document.analysis.facts.nodes);
    result
}
fn edit(s: &mut EngineSession, seq: &mut u32, c: EditCommand) -> EditorActionResult {
    action(s, seq, EditorAction::Edit { command: c, group_id: None })
}

#[test]
fn explicit_font_fallback_is_required_and_never_replaces_existing_primary_glyphs() {
    let (fonts, options) = fonts(); let mut layout = LayoutEngine::default();
    let s = EngineSession::from_source("fallback-policy".into(), r"x+\digamma+\sumint_0^1".into()).unwrap();
    let strict = LayoutOptions::new(options.math_font.clone());
    assert!(layout.layout(&s, &fonts, &strict).is_err(), "Missing glyphs must not get an implicit fallback");
    let scene = layout.layout(&s, &fonts, &options).unwrap();
    let owners: BTreeSet<_> = scene.resources.iter().map(|r|r.font_id.as_str()).collect();
    assert!(owners.contains(options.math_font.as_str()));
    assert!(owners.contains(options.math_fallbacks[0].as_str()));
    let plain = EngineSession::from_source("plain-primary".into(), "x+y".into()).unwrap();
    assert_eq!(layout.layout(&plain, &fonts, &options).unwrap().svg().unwrap(),
        layout.layout(&plain, &fonts, &strict).unwrap().svg().unwrap());
    let mut invalid = options.clone(); invalid.math_fallbacks.push("not-a-font".into());
    assert!(layout.layout(&s, &fonts, &invalid).is_err());
    invalid = options.clone(); invalid.math_fallbacks.push(options.math_fallbacks[0].clone());
    assert!(layout.layout(&s, &fonts, &invalid).is_err());
    // Removing explicit fallback must not reuse its cached geometry.
    layout.layout(&s, &fonts, &options).unwrap();
    assert!(layout.layout(&s, &fonts, &strict).is_err());
}

#[test]
fn catalog_inventory_is_unique_and_lookup_preserves_semantics() {
    let mut commands = BTreeSet::new();
    for spec in symbol_catalog::SYMBOLS {
        assert!(commands.insert(spec.command), "duplicate {}", spec.command);
        assert_eq!(named_symbol(spec.command), Some((spec.value, spec.class)));
        let source = format!("{} % spelling retained\r\n +x", spec.command);
        let p = parse(&source); assert!(p.complete && p.supported && p.exportable, "{}: {:?}", source, p.diagnostics);
        assert_eq!(model::latex(&p.tree, LatexMode::Lossless), source);
        assert_eq!(atoms(&p.tree)[0], (spec.value.into(), spec.class));
        let canonical = model::latex(&p.tree, LatexMode::Canonical); let q = parse(&canonical);
        assert!(q.complete && q.supported); assert_eq!(atoms(&p.tree), atoms(&q.tree));
        assert_eq!(model::latex(&q.tree, LatexMode::Canonical), canonical);
    }
    assert!(commands.len() > 140);
}

#[test]
fn newly_registered_atoms_use_actual_nonmissing_font_glyphs_in_all_math_styles() {
    let (fonts, mut options) = fonts(); let mut failures = Vec::new();
    for spec in symbol_catalog::SYMBOLS {
        let s = EngineSession::from_source("catalog-font".into(), spec.command.into()).unwrap();
        for style in [MathStyle::Display, MathStyle::Text, MathStyle::Script, MathStyle::ScriptScript] {
            options.style = style;
            match LayoutEngine::default().layout(&s, &fonts, &options) {
                Ok(scene) => { scene.validate().unwrap(); assert!(!scene.glyphs.is_empty(), "{}", spec.command);
                    assert!(scene.resources.iter().all(|r| r.shape.glyph_id != 0)); scene.svg().unwrap(); }
                Err(e) => failures.push(format!("{} {style:?}: {e:?}", spec.command)),
            }
        }
    }
    assert!(failures.is_empty(), "Unrendered registered atoms:\n{}", failures.join("\n"));
}

#[test]
fn direct_unicode_and_named_unambiguous_atoms_have_the_same_math_class_and_geometry() {
    let (fonts, options) = fonts();
    for name in [r"\subseteq", r"\cup", r"\land", r"\Rightarrow", r"\mapsto", r"\digamma", r"\otimes", r"\propto"] {
        let (value, class) = named_symbol(name).unwrap(); assert_eq!(literal_class(value), class, "{name}");
        let a = EngineSession::from_source("named".into(), format!("a{name} b")).unwrap();
        let b = EngineSession::from_source("unicode".into(), format!("a{value}b")).unwrap();
        let left = LayoutEngine::default().layout(&a, &fonts, &options).unwrap();
        let right = LayoutEngine::default().layout(&b, &fonts, &options).unwrap();
        assert_eq!(left.svg().unwrap(), right.svg().unwrap(), "{name}");
    }
    assert_eq!(named_symbol(r"\lvert").unwrap().1, AtomClass::Open);
    assert_eq!(named_symbol(r"\rvert").unwrap().1, AtomClass::Close);
    assert_eq!(literal_class("|"), AtomClass::Ordinary);
}

#[test]
fn symbol_transaction_is_one_atom_and_does_not_run_a_template_or_parser() {
    for name in [r"\partial", r"\digamma", r"\subseteq", r"\Rightarrow"] {
        let mut s = EngineSession::new("atom-action".into()).unwrap(); let mut seq = 0;
        let result = edit(&mut s, &mut seq, EditCommand::InsertSymbol { command: name.into() });
        assert_eq!(s.source(), name); assert_eq!(result.editor.document.analysis.facts.fractions, 0);
        assert_eq!(result.editor.document.analysis.facts.nodes, 2);
        let ids = identities(&result.editor.document.analysis.tree);
        edit(&mut s, &mut seq, EditCommand::InsertText { text: "x".into(), mode: InputMode::Math });
        assert_eq!(s.source(), format!("{name} x"), "Control-word boundary changed the command");
        action(&mut s, &mut seq, EditorAction::Undo); assert_eq!(s.source(), name);
        assert_eq!(identities(&s.document_snapshot().analysis.tree), ids);
        action(&mut s, &mut seq, EditorAction::Redo); assert_eq!(s.source(), format!("{name} x"));
    }
}

#[test]
fn every_catalog_atom_copies_independently_and_delete_undo_restores_its_identity() {
    for spec in symbol_catalog::SYMBOLS {
        let mut s = EngineSession::from_source("copy-atom".into(), spec.command.into()).unwrap(); let mut seq = 0;
        let root = s.document_snapshot().analysis.tree.id;
        let selection = Selection::Range { anchor: Position::row(root, 0), focus: Position::row(root, 1) };
        action(&mut s, &mut seq, EditorAction::SetSelection { selection });
        let copy = edit(&mut s, &mut seq, EditCommand::DuplicateSelection);
        let NodeKind::Row { entries, .. } = &copy.editor.document.analysis.tree.kind else { panic!() };
        assert_ne!(entries[0].node.id, entries[1].node.id); let copied_id = entries[1].node.id;
        edit(&mut s, &mut seq, EditCommand::InsertSymbol { command: r"\forall".into() });
        assert_eq!(atoms(&s.document_snapshot().analysis.tree)[0], (spec.value.into(), spec.class));
        action(&mut s, &mut seq, EditorAction::Undo);
        assert!(s.document_snapshot().analysis.tree.find(copied_id).is_some());
        edit(&mut s, &mut seq, EditCommand::DeleteBackward);
        assert_eq!(atoms(&s.document_snapshot().analysis.tree), vec![(spec.value.into(), spec.class)]);
        action(&mut s, &mut seq, EditorAction::Undo);
        assert!(s.document_snapshot().analysis.tree.find(copied_id).is_some());
    }
}

#[test]
fn unsupported_symbol_actions_leave_source_selection_and_history_unchanged() {
    let mut s = EngineSession::from_source("atomic-symbol".into(), "x".into()).unwrap();
    let before: Value = serde_json::to_value(s.editor_snapshot().unwrap()).unwrap();
    for (i, name) in [r"\unknownsymbol", r"\frac", r"\sum", r"\partial{x}", "<svg onload=x>"].iter().enumerate() {
        let error = s.apply_action(s.revision(), (i + 1) as u32, EditorAction::Edit {
            command: EditCommand::InsertSymbol { command: (*name).into() }, group_id: None }).unwrap_err();
        assert_eq!(error.code, ErrorCode::UnsupportedSyntax);
        assert_eq!(serde_json::to_value(s.editor_snapshot().unwrap()).unwrap(), before);
    }
}

#[test]
fn symbol_source_changes_update_content_without_swapping_identical_siblings() {
    let source = r"\frac{\subseteq}{\subseteq}+\subseteq";
    let mut s = EngineSession::from_source("source-atom".into(), source.into()).unwrap();
    let before = s.document_snapshot(); let mut ids = Vec::new();
    before.analysis.tree.visit(&mut |n| if matches!(n.kind, NodeKind::Symbol { .. }) { ids.push(n.id); });
    let at = source.find(r"\subseteq").unwrap();
    let (_, map) = s.patch_source(s.revision(), SourceEdit { start_utf16: at as u32, end_utf16: (at+9) as u32,
        replacement: r"\supseteq".into() }).unwrap();
    assert_eq!(atoms(&s.document_snapshot().analysis.tree)[0].0, "⊇");
    for id in &ids[1..] { assert!(map.retained_ids.contains(id)); }
    assert_eq!(atoms(&s.document_snapshot().analysis.tree), atoms(&parse(s.source()).tree));
}

#[test]
fn unicode_integrals_preserve_authored_source_and_have_real_bounds_and_glyphs() {
    let (fonts, mut options) = fonts(); let mut failures = Vec::new(); let mut commands = BTreeSet::new();
    for spec in rare_operators::OPERATORS {
        assert!(commands.insert(spec.command));
        assert_eq!(operators::spec(spec.command).unwrap().mark, OperatorMark::Glyph(spec.glyph));
        for mode in ["", r"\limits", r"\nolimits", r"\displaylimits"] {
            let source = format!("{}{mode}_{{a}}^{{b}}x", spec.command);
            let p = parse(&source); assert!(p.complete && p.supported && p.exportable);
            assert_eq!(model::latex(&p.tree, LatexMode::Lossless), source);
            let s = EngineSession::from_source("rare-glyph".into(), source.clone()).unwrap();
            for style in [MathStyle::Display, MathStyle::Text, MathStyle::Script, MathStyle::ScriptScript] {
                options.style = style;
                match LayoutEngine::default().layout(&s, &fonts, &options) {
                    Ok(scene) => { assert!(scene.regions.iter().any(|r| r.role == "largeOperator")); scene.svg().unwrap(); }
                    Err(e) => failures.push(format!("{source} {style:?}: {e:?}")),
                }
            }
        }
    }
    assert!(failures.is_empty(), "Unrendered integral glyphs:\n{}", failures.join("\n"));
}

#[test]
fn rare_integral_transactions_preserve_bound_ids_across_modes_copy_and_history() {
    for spec in rare_operators::OPERATORS {
        let mut s = EngineSession::new("rare-action".into()).unwrap(); let mut seq = 0;
        let inserted = edit(&mut s, &mut seq, EditCommand::InsertOperator {
            command: spec.command.into(), with_lower: true, with_upper: true });
        let root = inserted.editor.document.analysis.tree.id;
        edit(&mut s, &mut seq, EditCommand::InsertText { text: "0".into(), mode: InputMode::Math });
        action(&mut s, &mut seq, EditorAction::MoveHorizontal { direction: visualtex_engine_core::positions::Horizontal::Right, extend: false });
        edit(&mut s, &mut seq, EditCommand::InsertText { text: "1".into(), mode: InputMode::Math });
        let before = s.document_snapshot();
        edit(&mut s, &mut seq, EditCommand::SetOperatorLimits { mode: LimitMode::Limits });
        assert_eq!(identities(&before.analysis.tree), identities(&s.document_snapshot().analysis.tree));
        action(&mut s, &mut seq, EditorAction::Undo); assert_eq!(s.source(), before.source);
        action(&mut s, &mut seq, EditorAction::SetSelection { selection: Selection::Range {
            anchor: Position::row(root, 0), focus: Position::row(root, 1) } });
        let copied = edit(&mut s, &mut seq, EditCommand::DuplicateSelection);
        let NodeKind::Row { entries, .. } = &copied.editor.document.analysis.tree.kind else { panic!() };
        assert!(identities(&entries[0].node).is_disjoint(&identities(&entries[1].node)));
    }
}

#[test]
fn esint_outline_aliases_are_not_falsely_advertised_as_unicode_replacements() {
    for source in [r"\sqint_0^1", r"\sqiint", r"\idotsint", r"\dotsint", r"\ointclockwise"] {
        let p = parse(source); assert!(!p.supported && !p.exportable, "{source}");
        assert_eq!(model::latex(&p.tree, LatexMode::Lossless), source);
    }
}
