//! Independent P04 tree-command tests against the exported core APIs.
use visualtex_engine_core::{model, parser, positions, session, source, transactions};
use std::collections::BTreeSet;
use std::sync::Arc;
use model::{IdAllocator, LatexMode, Node, NodeKind};
use parser::ParseResult;
use positions::{Horizontal, Position, Selection, TreeIndex};
use session::NodeId;
use source::SourceText;
use transactions::{EditCommand, InputMode, StructureKind, PreparedEdit};
use visualtex_engine_core::ErrorCode;

fn seed(text: &str) -> (Arc<ParseResult>, IdAllocator) {
    let mut ids = IdAllocator::new(2);
    let document = Arc::new(parser::parse(SourceText::new(1, text).unwrap(), NodeId(1), &mut ids).unwrap());
    (document, ids)
}
fn caret(id: NodeId, offset: usize) -> Selection { Selection::caret(Position::row(id, offset)) }
fn text_id(root: &Node) -> NodeId {
    let mut result = None;
    root.visit(&mut |node| if matches!(node.kind, NodeKind::Text { .. }) { result = Some(node.id); });
    result.unwrap()
}
fn node_ids(root: &Node) -> BTreeSet<NodeId> {
    let mut result = BTreeSet::new(); root.visit(&mut |node| { result.insert(node.id); }); result
}
fn valid_after(before: &Arc<ParseResult>, prepared: &PreparedEdit) {
    let after = &prepared.document;
    let index = TreeIndex::new(&after.tree).unwrap();
    index.validate_selection(&prepared.selection).unwrap();
    assert_eq!(model::latex(&after.tree, LatexMode::Lossless), after.source.text());
    let (reparsed, _) = seed(after.source.text());
    assert!(reparsed.complete && reparsed.supported, "saved syntax failed to reopen: {:?}", after.source.text());
    assert_eq!(model::latex(&after.tree, LatexMode::Canonical), model::latex(&reparsed.tree, LatexMode::Canonical),
        "saved source means something else than the edited tree: {:?}", after.source.text());
    let old_index = TreeIndex::new(&before.tree).unwrap();
    for position in &old_index.carets {
        let mapped = prepared.position_map.apply_position(*position, &index).unwrap();
        index.validate_position(mapped).unwrap();
    }
    let old = node_ids(&before.tree); let new = node_ids(&after.tree);
    assert_eq!(prepared.position_map.retained_ids.iter().copied().collect::<BTreeSet<_>>(), old.intersection(&new).copied().collect());
    assert_eq!(prepared.position_map.removed_ids.iter().copied().collect::<BTreeSet<_>>(), old.difference(&new).copied().collect());
    assert_eq!(prepared.position_map.inserted_ids.iter().copied().collect::<BTreeSet<_>>(), new.difference(&old).copied().collect());
}
fn run(doc: &Arc<ParseResult>, selection: &Selection, command: EditCommand, ids: &mut IdAllocator) -> PreparedEdit {
    let result = transactions::prepare(doc, selection, &command, doc.source.id() + 1, ids).unwrap();
    valid_after(doc, &result); result
}

#[test]
fn the_minimum_k0_formula_is_built_by_real_structural_commands() {
    let (mut doc, mut ids) = seed("");
    let first = run(&doc, &caret(doc.tree.id, 0), EditCommand::InsertStructure { structure: StructureKind::Fraction }, &mut ids);
    let mut selection = first.selection; doc = first.document;
    let NodeKind::Row { entries, .. } = &doc.tree.kind else { panic!() };
    let NodeKind::Fraction { numerator, denominator, .. } = &entries[0].node.kind else { panic!() };
    let (num_id, den_id) = (numerator.id, denominator.id);
    for command in [EditCommand::InsertText { text: "x".into(), mode: InputMode::Math },
        EditCommand::InsertStructure { structure: StructureKind::Subscript },
        EditCommand::InsertText { text: "1".into(), mode: InputMode::Math }] {
        let next = run(&doc, &selection, command, &mut ids); selection = next.selection; doc = next.document;
    }
    selection = caret(num_id, 1);
    for command in [EditCommand::InsertText { text: "+".into(), mode: InputMode::Math },
        EditCommand::InsertText { text: "速度".into(), mode: InputMode::Text }] {
        let next = run(&doc, &selection, command, &mut ids); selection = next.selection; doc = next.document;
    }
    selection = caret(den_id, 0);
    for command in [EditCommand::InsertText { text: "y".into(), mode: InputMode::Math },
        EditCommand::InsertStructure { structure: StructureKind::Superscript },
        EditCommand::InsertText { text: "2".into(), mode: InputMode::Math }] {
        let next = run(&doc, &selection, command, &mut ids); selection = next.selection; doc = next.document;
    }
    assert_eq!(doc.source.text(), "\\frac{x_{1}+\\text{速度}}{y^{2}}");
    assert!(doc.exportable);
    assert_eq!(doc.facts.fractions, 1); assert_eq!(doc.facts.scripts, 2);
}

#[test]
fn backspace_at_fraction_boundary_selects_then_deletes_without_destroying_neighbor() {
    let (doc, mut ids) = seed("a+\\frac{x}{y}+b");
    let next_id = ids.next_unused();
    let first = run(&doc, &caret(doc.tree.id, 3), EditCommand::DeleteBackward, &mut ids);
    assert!(!first.changed); assert!(Arc::ptr_eq(&doc, &first.document));
    assert!(matches!(first.selection, Selection::Node { .. })); assert_eq!(ids.next_unused(), next_id);
    let second = run(&doc, &first.selection, EditCommand::DeleteBackward, &mut ids);
    assert_eq!(second.document.source.text(), "a++b");
    assert!(second.selection.focus().unwrap().same_location(Position::row(doc.tree.id, 2)));
}

#[test]
fn text_deletion_removes_complete_graphemes_not_surrogates_or_combining_marks() {
    let (mut doc, mut ids) = seed("\\text{e\u{301}👩🏽‍🚀中}");
    let id = text_id(&doc.tree);
    let mut selection = Selection::caret(Position::text(id, 3));
    for expected in ["\\text{e\u{301}👩🏽‍🚀}", "\\text{e\u{301}}", "\\text{}"] {
        let next = run(&doc, &selection, EditCommand::DeleteBackward, &mut ids);
        assert_eq!(next.document.source.text(), expected);
        selection = next.selection; doc = next.document;
    }
}

#[test]
fn shift_selection_deletes_only_the_first_text_character() {
    let (doc, mut ids) = seed("\\text{abc}+x");
    let selection = TreeIndex::new(&doc.tree).unwrap().move_horizontal(&caret(doc.tree.id, 0), Horizontal::Right, true).unwrap();
    let next = run(&doc, &selection, EditCommand::DeleteForward, &mut ids);
    assert_eq!(next.document.source.text(), "\\text{bc}+x");
}

#[test]
fn duplicate_formula_has_fresh_identity_and_edits_independently() {
    let (doc, mut ids) = seed("\\frac{x}{y}");
    let NodeKind::Row { entries, .. } = &doc.tree.kind else { panic!() };
    let first_id = entries[0].node.id;
    let copied = run(&doc, &Selection::Node { node_id: first_id }, EditCommand::DuplicateSelection, &mut ids);
    let NodeKind::Row { entries, .. } = &copied.document.tree.kind else { panic!() };
    assert_eq!(entries.len(), 2);
    assert!(node_ids(&entries[0].node).is_disjoint(&node_ids(&entries[1].node)));
    let NodeKind::Fraction { numerator, .. } = &entries[1].node.kind else { panic!() };
    let changed = run(&copied.document, &caret(numerator.id, 1), EditCommand::InsertText { text: "+1".into(), mode: InputMode::Math }, &mut ids);
    assert_eq!(changed.document.source.text(), "\\frac{x}{y}\\frac{x+1}{y}");
    assert_eq!(model::latex(changed.document.tree.find(first_id).unwrap(), LatexMode::Lossless), "\\frac{x}{y}");
}

#[test]
fn invalid_paste_and_allocator_exhaustion_are_atomic() {
    let (doc, mut ids) = seed("x+1");
    let next = ids.next_unused();
    for bad in ["\\frac{x}{", "\\unknown{secret}", "\\input{never-execute}", "x^1^2"] {
        assert!(transactions::prepare(&doc, &caret(doc.tree.id, 1), &EditCommand::PasteLatex { latex: bad.into() }, 2, &mut ids).is_err());
        assert_eq!(ids.next_unused(), next);
        assert_eq!(doc.source.text(), "x+1");
    }
    let mut exhausted = IdAllocator::new(u32::MAX - 1);
    assert!(transactions::prepare(&doc, &caret(doc.tree.id, 1), &EditCommand::InsertStructure { structure: StructureKind::Fraction }, 2, &mut exhausted).is_err());
    assert_eq!(exhausted.next_unused(), u32::MAX - 1);
}

#[test]
fn empty_paste_does_not_allocate_an_unused_root_identity() {
    let (doc, mut ids) = seed("x");
    let next = ids.next_unused();
    let result = run(&doc, &caret(doc.tree.id, 1), EditCommand::PasteLatex { latex: String::new() }, &mut ids);
    assert!(!result.changed); assert_eq!(ids.next_unused(), next);
}

#[test]
fn a_pasted_trailing_comment_cannot_comment_out_existing_math() {
    let (doc, mut ids) = seed("a+b");
    let result = run(&doc, &caret(doc.tree.id, 0), EditCommand::PasteLatex { latex: "x% keep".into() }, &mut ids);
    assert_eq!(result.document.source.text(), "x% keep\na+b");
}

#[test]
fn a_comment_only_paste_cannot_erase_following_source_semantics() {
    let (doc, mut ids) = seed("a+b");
    let result = run(&doc, &caret(doc.tree.id, 0), EditCommand::PasteLatex { latex: "% keep".into() }, &mut ids);
    assert_eq!(result.document.source.text(), "% keep\na+b");
}

#[test]
fn text_split_then_merge_restores_text_and_valid_selection() {
    let (doc, mut ids) = seed("\\text{ab中😀}");
    let id = text_id(&doc.tree);
    let split = run(&doc, &Selection::caret(Position::text(id, 2)), EditCommand::SplitText, &mut ids);
    assert_eq!(split.document.source.text(), "\\text{ab}\\text{中😀}");
    let merged = run(&split.document, &split.selection, EditCommand::MergeText { direction: Horizontal::Left }, &mut ids);
    assert_eq!(merged.document.source.text(), "\\text{ab中😀}");
    assert_eq!(text_id(&merged.document.tree), id);
}

#[test]
fn literal_math_metacharacters_are_escaped_instead_of_becoming_source_commands() {
    let (doc, mut ids) = seed("");
    let result = run(&doc, &caret(doc.tree.id, 0), EditCommand::InsertText { text: "{%&#$}".into(), mode: InputMode::Math }, &mut ids);
    assert_eq!(result.document.source.text(), "\\{\\%\\&\\#\\$\\}");
}

#[test]
fn replacing_part_of_text_maps_old_carets_past_deleted_temporary_split_nodes() {
    for (start, end) in [(1, 4), (1, 3)] {
        let (doc, mut ids) = seed("\\text{abcd}");
        let id = text_id(&doc.tree);
        let selection = Selection::Range { anchor: Position::text(id, end), focus: Position::text(id, start) };
        let result = run(&doc, &selection, EditCommand::PasteLatex { latex: "\\text{X}".into() }, &mut ids);
        let expected = if end == 4 { "\\text{a}\\text{X}" } else { "\\text{a}\\text{X}\\text{d}" };
        assert_eq!(result.document.source.text(), expected);
        // valid_after checks every old caret against the returned position map,
        // not merely the successful command's own newly chosen selection.
    }
}

#[test]
fn seeded_structural_command_sequences_preserve_reopenability_and_atomic_failures() {
    let initial = ["", "a+\\frac{x_1}{y^2}+b", "\\text{速度 e\u{301}😀}+x", "x^{a_b}+\\alpha y"];
    let mut random = 0x5040_2026_0913_a17eu64;
    let mut next = || { random ^= random << 13; random ^= random >> 7; random ^= random << 17; random as usize };
    let mut accepted = 0;
    for sequence in 0..20 {
        let (mut doc, mut ids) = seed(initial[sequence % initial.len()]);
        for action in 0..80 {
            let index = TreeIndex::new(&doc.tree).unwrap();
            let anchor = index.carets[next() % index.carets.len()];
            let selection = if next() % 3 == 0 {
                Selection::Range { anchor, focus: index.carets[next() % index.carets.len()] }
            } else { Selection::caret(anchor) };
            let command = match next() % 13 {
                0 | 1 => EditCommand::InsertText { text: ["a", "b", "+", "12", "%", "{", "}", "", "𝑥"][next() % 9].into(), mode: InputMode::Math },
                2 => EditCommand::InsertText { text: ["速度", "😀", "e\u{301}", " ", "%", "\n", "👩‍🔬", "_"][next() % 8].into(), mode: InputMode::Text },
                3 => EditCommand::InsertStructure { structure: [StructureKind::Fraction, StructureKind::Superscript, StructureKind::Subscript, StructureKind::Text][next() % 4] },
                4 => EditCommand::DeleteBackward,
                5 => EditCommand::DeleteForward,
                6 => EditCommand::PasteLatex { latex: ["\\frac{x_1}{y^2}", "\\alpha x", "\\text{速度}", "x% note", "\n", "%note", "\\frac{a}{", "\\bad"][next() % 8].into() },
                7 => EditCommand::DuplicateSelection,
                8 => EditCommand::SplitText,
                9 => EditCommand::MergeText { direction: Horizontal::Left },
                10 => EditCommand::MergeText { direction: Horizontal::Right },
                11 => EditCommand::ReplaceRange { text: String::new(), mode: InputMode::Math },
                _ => EditCommand::InsertStructure { structure: StructureKind::Fraction },
            };
            let before_source = doc.source.text().to_owned();
            let before_ids = ids.next_unused();
            let request = format!("seed=504020260913a17e sequence={sequence} action={action} source={before_source:?} selection={selection:?} command={command:?}");
            let prepared = transactions::prepare(&doc, &selection, &command, doc.source.id() + 1, &mut ids);
            match prepared {
                Ok(prepared) => {
                    let check = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| valid_after(&doc, &prepared)));
                    assert!(check.is_ok(), "{request}");
                    if !prepared.changed { assert_eq!(ids.next_unused(), before_ids, "noop consumed identity: {request}"); }
                    doc = prepared.document; accepted += 1;
                }
                Err(error) => {
                    assert!(matches!(error.code, ErrorCode::InvalidPosition | ErrorCode::UnsupportedSyntax | ErrorCode::IncompleteSource | ErrorCode::ResourceLimit), "unexpected failure {error}: {request}");
                    assert_eq!(ids.next_unused(), before_ids, "failed command consumed identity: {request}");
                    assert_eq!(doc.source.text(), before_source, "failed command mutated source: {request}");
                }
            }
        }
    }
    assert!(accepted > 700, "random regression did not exercise enough successful actions: {accepted}");
}

#[test]
fn unsupported_style_and_visual_draft_edit_are_explicit_failures() {
    let (doc, mut ids) = seed("x");
    assert_eq!(transactions::prepare(&doc, &caret(doc.tree.id, 0), &EditCommand::SetStyle { style: "bold".into() }, 2, &mut ids).err().unwrap().code, ErrorCode::UnsupportedSyntax);
    let (draft, mut ids) = seed("\\frac{x}{");
    assert_eq!(transactions::prepare(&draft, &caret(draft.tree.id, 0), &EditCommand::InsertText { text: "a".into(), mode: InputMode::Math }, 2, &mut ids).err().unwrap().code, ErrorCode::IncompleteSource);
}
