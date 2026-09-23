//! Independent structural serialization/identity regressions for the K0 model.
use std::sync::Arc;
use visualtex_engine_core::model::{AtomClass, Grouping, IdAllocator, LatexMode, Node, NodeKind, RowEntry, ScriptKind, ScriptSlot, latex, validate_tree};
use visualtex_engine_core::session::NodeId;
use visualtex_engine_core::source::SourceText;

fn symbol(id: u32, value: &str, command: Option<&str>) -> Arc<Node> {
    Arc::new(Node { id: NodeId(id), origin: None, kind: NodeKind::Symbol {
        value: value.into(), class: AtomClass::Ordinary, command: command.map(str::to_owned),
    }})
}
fn row(id: u32, children: Vec<Arc<Node>>) -> Arc<Node> {
    Arc::new(Node { id: NodeId(id), origin: None, kind: NodeKind::Row {
        grouping: Grouping::Bare,
        entries: children.into_iter().map(|node| RowEntry { leading: String::new(), node }).collect(),
        trailing: String::new(),
    }})
}

#[test]
fn inserted_letter_cannot_extend_a_preserved_control_word() {
    let source = SourceText::new(1, "\\alpha").unwrap();
    let original = Arc::new(Node { id: NodeId(2), origin: Some(source.slice(0, source.len()).unwrap()),
        kind: NodeKind::Symbol { value: "α".into(), class: AtomClass::Ordinary, command: Some("\\alpha".into()) } });
    let root = row(1, vec![original, symbol(3, "x", None)]);
    assert_eq!(latex(&root, LatexMode::Lossless), "\\alpha x",
        "local structural edit must not create the unrelated command \\alphax");
    assert_eq!(latex(&root, LatexMode::Canonical), "\\alpha x");
}

#[test]
fn control_word_boundary_handles_scripted_base_not_just_bare_symbols() {
    let scripted = Arc::new(Node { id: NodeId(3), origin: None, kind: NodeKind::Scripts {
        base: symbol(4, "x", None), slots: vec![ScriptSlot { script: ScriptKind::Superscript,
            before_marker: String::new(), before_argument: String::new(), argument: row(5, vec![symbol(6, "2", None)]) }],
    }});
    let root = row(1, vec![symbol(2, "α", Some("\\alpha")), scripted]);
    for mode in [LatexMode::Lossless, LatexMode::Canonical] {
        assert_eq!(latex(&root, mode), "\\alpha x^{2}");
    }
}

#[test]
fn untouched_root_preserves_comments_spaces_and_crlf_byte_for_byte() {
    let text = " \\dfrac %原注释\r\n{a}{b}  ";
    let source = SourceText::new(1, text).unwrap();
    let root = Node { id: NodeId(1), origin: Some(source.slice(0, source.len()).unwrap()),
        kind: NodeKind::Row { grouping: Grouping::Bare, entries: vec![], trailing: String::new() } };
    assert_eq!(latex(&root, LatexMode::Lossless), text);
}

#[test]
fn unchanged_sibling_spelling_and_trivia_survive_an_edited_row() {
    let old = "\\alpha % keep this\r\n";
    let source = SourceText::new(1, old).unwrap();
    let alpha = Arc::new(Node { id: NodeId(2), origin: Some(source.slice(0, 6).unwrap()),
        kind: NodeKind::Symbol { value: "α".into(), class: AtomClass::Ordinary, command: Some("\\alpha".into()) } });
    let root = Node { id: NodeId(1), origin: None, kind: NodeKind::Row {
        grouping: Grouping::Bare,
        entries: vec![RowEntry { leading: String::new(), node: alpha },
            RowEntry { leading: " % keep this\r\n".into(), node: symbol(3, "z", None) }],
        trailing: "  ".into(),
    }};
    assert_eq!(latex(&root, LatexMode::Lossless), "\\alpha % keep this\r\nz  ");
    assert_eq!(latex(&root, LatexMode::Canonical), "\\alpha z");
}

#[test]
fn identical_subexpressions_must_still_have_unique_node_ids() {
    let valid = row(1, vec![symbol(2, "x", None), symbol(3, "x", None)]);
    assert_eq!(validate_tree(&valid).unwrap().nodes, 3);
    let repeated = symbol(2, "x", None);
    let invalid = row(1, vec![repeated.clone(), repeated]);
    assert!(validate_tree(&invalid).is_err());
}

#[test]
fn fractions_require_editable_argument_rows() {
    let root = Node { id: NodeId(1), origin: None, kind: NodeKind::Fraction {
        command: "\\frac".into(), before_numerator: String::new(), before_denominator: String::new(),
        numerator: symbol(2, "a", None), denominator: row(3, vec![symbol(4, "b", None)]),
    }};
    assert!(validate_tree(&root).is_err());
}

#[test]
fn node_allocator_is_monotonic_and_errors_before_wrapping() {
    let mut ids = IdAllocator::new(10);
    assert_eq!(ids.allocate().unwrap(), NodeId(10));
    assert_eq!(ids.allocate().unwrap(), NodeId(11));
    assert_eq!(ids.next_unused(), 12);
    assert!(IdAllocator::new(0).allocate().is_err());
    let mut near_end = IdAllocator::new(u32::MAX - 1);
    assert_eq!(near_end.allocate().unwrap(), NodeId(u32::MAX - 1));
    assert!(near_end.allocate().is_err());
    assert_eq!(near_end.next_unused(), u32::MAX);
}

#[test]
fn source_text_escapes_tex_metacharacters_after_a_real_text_edit() {
    let text = Node { id: NodeId(1), origin: None, kind: NodeKind::Text {
        command: "\\text".into(), before_body: String::new(), text: "速度 50% {a}_#&$ \\ ^ ~".into(),
    }};
    let expected = "\\text{速度 50\\% \\{a\\}\\_\\#\\&\\$ \\textbackslash{} \\textasciicircum{} \\textasciitilde{}}";
    assert_eq!(latex(&text, LatexMode::Canonical), expected);
    assert_eq!(latex(&text, LatexMode::Lossless), expected);
}
