//! Structural edits must project to equivalent LaTeX without reparsing as an
//! implementation shortcut. Reparse here is an independent test oracle only.
use std::collections::BTreeSet;
use std::sync::Arc;
use visualtex_engine_core::model::{self, AtomClass, Grouping, IdAllocator, LatexMode, Node, NodeKind, RowEntry, ScriptKind, ScriptSlot};
use visualtex_engine_core::parser::{self, ParseResult};
use visualtex_engine_core::projection::project_tree;
use visualtex_engine_core::session::NodeId;
use visualtex_engine_core::source::SourceText;
use visualtex_engine_core::ErrorCode;

fn parse(source: &str) -> Arc<ParseResult> {
    Arc::new(parser::parse(SourceText::new(1, source).unwrap(), NodeId(1), &mut IdAllocator::new(2)).unwrap())
}
fn symbol(id: u32, value: &str) -> Arc<Node> {
    Arc::new(Node { id: NodeId(id), origin: None, kind: NodeKind::Symbol {
        value: value.into(), class: AtomClass::Ordinary, command: None,
    }})
}
fn row(id: u32, children: Vec<Arc<Node>>) -> Arc<Node> {
    Arc::new(Node { id: NodeId(id), origin: None, kind: NodeKind::Row {
        grouping: Grouping::Bare, entries: children.into_iter().map(|node| RowEntry { leading: String::new(), node }).collect(),
        trailing: String::new(),
    }})
}
fn ids(node: &Node) -> BTreeSet<NodeId> {
    let mut result = BTreeSet::new(); node.visit(&mut |node| { assert!(result.insert(node.id)); }); result
}
fn verify_projection(before: &Node, projected: &ParseResult, expected: &str) {
    assert_eq!(projected.source.text(), expected);
    assert_eq!(model::latex(&projected.tree, LatexMode::Lossless), expected);
    assert_eq!(ids(before), ids(&projected.tree), "projection changed edit identities");
    projected.tree.visit(&mut |node| {
        let origin = node.origin.as_ref().expect("every projected node has a fresh source span");
        assert_eq!(origin.span.source_id, projected.source.id());
        assert!(origin.span.byte_end as usize <= expected.len());
        assert_eq!(origin.text(), &expected[origin.span.byte_start as usize..origin.span.byte_end as usize]);
    });
    let reparsed = parse(expected);
    assert!(reparsed.complete && reparsed.supported, "projected syntax cannot be opened again");
    assert_eq!(model::latex(&projected.tree, LatexMode::Canonical), model::latex(&reparsed.tree, LatexMode::Canonical));
}

#[test]
fn untouched_formula_projection_keeps_every_byte_and_identity() {
    for source in ["", "\\frac12", " \\dfrac %comment\r\n{a}{b} ", "x^2_1", "\\text{速度 👩‍🔬}", "\\placeholder{分母}"] {
        let parsed = parse(source);
        let projected = project_tree(parsed.tree.clone(), 2).unwrap();
        verify_projection(&parsed.tree, &projected, source);
    }
}

#[test]
fn synthesized_command_separator_is_included_in_all_following_node_ranges() {
    let parsed = parse("\\alpha");
    let NodeKind::Row { entries, .. } = &parsed.tree.kind else { panic!() };
    let scripted = Arc::new(Node { id: NodeId(100), origin: None, kind: NodeKind::Scripts {
        base: symbol(101, "x"), slots: vec![ScriptSlot { script: ScriptKind::Superscript,
            before_marker: String::new(), before_argument: String::new(), argument: row(102, vec![symbol(103, "2")]) }],
    }});
    let root = row(1, vec![entries[0].node.clone(), scripted]);
    let projected = project_tree(root.clone(), 2).unwrap();
    verify_projection(&root, &projected, "\\alpha x^{2}");
    assert_eq!(projected.tree.find(NodeId(101)).unwrap().origin.as_ref().unwrap().text(), "x");
    assert_eq!(projected.tree.find(NodeId(103)).unwrap().origin.as_ref().unwrap().text(), "2");
}

#[test]
fn preserved_unbraced_multi_atom_base_is_grouped_before_attaching_a_script() {
    let base = parse("a+b").tree.clone();
    let script = Arc::new(Node { id: NodeId(100), origin: None, kind: NodeKind::Scripts {
        base, slots: vec![ScriptSlot { script: ScriptKind::Superscript, before_marker: String::new(),
            before_argument: String::new(), argument: row(101, vec![symbol(102, "2")]) }],
    }});
    let root = row(103, vec![script]);
    let projected = project_tree(root.clone(), 2).unwrap();
    verify_projection(&root, &projected, "{a+b}^{2}");
}

#[test]
fn preserved_scripted_base_is_grouped_before_a_second_superscript() {
    let base = parse("x^2").tree.clone();
    let script = Arc::new(Node { id: NodeId(100), origin: None, kind: NodeKind::Scripts {
        base, slots: vec![ScriptSlot { script: ScriptKind::Superscript, before_marker: String::new(),
            before_argument: String::new(), argument: row(101, vec![symbol(102, "3")]) }],
    }});
    let root = row(103, vec![script]);
    let projected = project_tree(root.clone(), 2).unwrap();
    verify_projection(&root, &projected, "{x^2}^{3}");
}

#[test]
fn moving_an_existing_bare_row_into_fraction_argument_adds_required_braces() {
    let numerator = parse("ab").tree.clone();
    let fraction = Arc::new(Node { id: NodeId(100), origin: None, kind: NodeKind::Fraction {
        command: "\\frac".into(), before_numerator: String::new(), numerator,
        before_denominator: String::new(), denominator: row(101, vec![symbol(102, "c")]),
    }});
    let root = row(103, vec![fraction]);
    let projected = project_tree(root.clone(), 2).unwrap();
    verify_projection(&root, &projected, "\\frac{ab}{c}");
}

#[test]
fn a_preserved_trailing_comment_does_not_swallow_new_grouping_or_later_edits() {
    let numerator = parse("ab% keep").tree.clone();
    let fraction = Arc::new(Node { id: NodeId(100), origin: None, kind: NodeKind::Fraction {
        command: "\\frac".into(), before_numerator: String::new(), numerator,
        before_denominator: String::new(), denominator: row(101, vec![symbol(102, "c")]),
    }});
    let root = row(103, vec![fraction]);
    let projected = project_tree(root.clone(), 2).unwrap();
    verify_projection(&root, &projected, "\\frac{ab% keep\n}{c}");
    // A second structural edit must not reintroduce the swallowed closing brace
    // after the now-braced numerator loses its whole-node origin fast path.
    let mut next_root = projected.tree.as_ref().clone(); next_root.origin = None;
    let NodeKind::Row { entries, .. } = &mut next_root.kind else { panic!() };
    let mut next_fraction = entries[0].node.as_ref().clone(); next_fraction.origin = None;
    let NodeKind::Fraction { numerator, .. } = &mut next_fraction.kind else { panic!() };
    let mut next_numerator = numerator.as_ref().clone(); next_numerator.origin = None;
    let NodeKind::Row { entries: atoms, .. } = &mut next_numerator.kind else { panic!() };
    atoms[0].node = symbol(104, "z");
    *numerator = Arc::new(next_numerator); entries[0].node = Arc::new(next_fraction);
    let next_root = Arc::new(next_root);
    let next = project_tree(next_root.clone(), 3).unwrap();
    verify_projection(&next_root, &next, "\\frac{zb% keep\n}{c}");
}

#[test]
fn an_escaped_percent_is_not_mistaken_for_a_trailing_comment() {
    let numerator = parse("a\\%b").tree.clone();
    let fraction = Arc::new(Node { id: NodeId(100), origin: None, kind: NodeKind::Fraction {
        command: "\\frac".into(), before_numerator: String::new(), numerator,
        before_denominator: String::new(), denominator: row(101, vec![symbol(102, "c")]),
    }});
    let root = row(103, vec![fraction]);
    let projected = project_tree(root.clone(), 2).unwrap();
    verify_projection(&root, &projected, "\\frac{a\\%b}{c}");
}

#[test]
fn local_change_preserves_untouched_sibling_source_spelling_after_rebasing() {
    let original = "a+\\dfrac % keep\r\n{x_1}{y^2}  ";
    let parsed = parse(original);
    let mut root = parsed.tree.as_ref().clone();
    root.origin = None;
    let NodeKind::Row { entries, .. } = &mut root.kind else { panic!() };
    let unchanged_id = entries[2].node.id;
    entries[0].node = symbol(100, "速度");
    let root = Arc::new(root);
    let projected = project_tree(root.clone(), 2).unwrap();
    verify_projection(&root, &projected, "速度+\\dfrac % keep\r\n{x_1}{y^2}  ");
    assert_eq!(projected.tree.find(unchanged_id).unwrap().origin.as_ref().unwrap().text(), "\\dfrac % keep\r\n{x_1}{y^2}");
}

#[test]
fn stale_ancestor_origin_is_rejected_instead_of_silently_losing_a_changed_child() {
    let parsed = parse("a+b");
    let mut root = parsed.tree.as_ref().clone();
    let NodeKind::Row { entries, .. } = &mut root.kind else { panic!() };
    entries[0].node = symbol(100, "c"); // deliberately keep stale root origin
    assert_eq!(project_tree(Arc::new(root), 2).unwrap_err().code, ErrorCode::InvalidInput);
    assert_eq!(model::latex(&parsed.tree, LatexMode::Lossless), "a+b");
}

#[test]
fn opaque_visual_projection_is_rejected_without_changing_the_old_source() {
    let parsed = parse("a+\\unknown{x}");
    assert_eq!(project_tree(parsed.tree.clone(), 2).unwrap_err().code, ErrorCode::UnsupportedSyntax);
    assert_eq!(model::latex(&parsed.tree, LatexMode::Lossless), "a+\\unknown{x}");
}

#[test]
fn duplicate_identity_cannot_be_hidden_by_projected_source_map() {
    let repeated = symbol(2, "x");
    assert_eq!(project_tree(row(1, vec![repeated.clone(), repeated]), 2).unwrap_err().code, ErrorCode::InvalidInput);
}
