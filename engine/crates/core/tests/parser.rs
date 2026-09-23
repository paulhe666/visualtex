use std::sync::Arc;
use visualtex_engine_core::ErrorCode;
use visualtex_engine_core::model::{self, AtomClass, IdAllocator, LatexMode, Node, NodeKind, RowEntry};
use visualtex_engine_core::parser::{self, DiagnosticCode, ParseResult};
use visualtex_engine_core::session::NodeId;
use visualtex_engine_core::source::{SourceText, MAX_SOURCE_BYTES};

fn parse(input: &str) -> ParseResult {
    parser::parse(SourceText::new(1, input).unwrap(), NodeId(1), &mut IdAllocator::new(2)).unwrap()
}

#[test]
fn k0_fraction_scripts_and_chinese_text_are_real_structure() {
    let p = parse(r"\frac{x_1+\text{速度}}{y^2}");
    assert!(p.complete && p.supported && p.exportable);
    assert_eq!(p.facts.fractions, 1);
    assert_eq!(p.facts.scripts, 2);
    assert_eq!(p.facts.text_runs, 1);
    let mut text = Vec::new();
    p.tree.visit(&mut |node| if let NodeKind::Text { text: value, .. } = &node.kind { text.push(value.clone()); });
    assert_eq!(text, vec!["速度"]);
}

#[test]
fn lossless_source_includes_all_spelling_trivia_and_invalid_input() {
    for input in ["", "   ", "\n%comment\n", "x + y", r"\frac12", r"\dfrac {x}{y}",
        " \\frac %分子\n { x_1 } %分母\r\n{\\text{速度}} ", r"x^a_b", r"{x}^{2}",
        r"\frac{\frac{x_1}{2}}{y^{z_2}}", r"\alpha x^2+\beta a", r"\text{a\%b\{c\}}",
        "\\text{中😀e\u{301}}", r"\mystery*[x]{a{b}}{c}+z", r"\frac{x}{", r"x^", r"\", r"x}}", r"^2", r"x^2^3"] {
        let p = parse(input);
        assert_eq!(model::latex(&p.tree, LatexMode::Lossless), input, "{input}");
        assert_eq!(p.tokens.iter().map(|t| t.text(&p.source)).collect::<String>(), input);
        p.tree.visit(&mut |node| {
            if let Some(origin) = &node.origin {
                assert_eq!(p.source.byte_to_utf16(origin.span.byte_start as usize).unwrap(), origin.span.utf16_start);
                assert_eq!(p.source.byte_to_utf16(origin.span.byte_end as usize).unwrap(), origin.span.utf16_end);
            }
        });
    }
}

#[test]
fn incomplete_source_is_diagnosed_not_silently_completed() {
    for input in [r"\frac", r"\frac{x}", r"\frac{x}{", r"x^", r"^2", r"\text{速度", r"\", r"x}"] {
        let p = parse(input);
        assert!(!p.complete && !p.exportable, "{input}");
        assert!(!p.diagnostics.is_empty());
        assert_eq!(model::latex(&p.tree, LatexMode::Lossless), input);
    }
}

#[test]
fn unknown_macro_and_arguments_are_one_opaque_node() {
    let p = parse(r"a+\mystery*[x]{a{b}}{c}+z");
    assert!(p.complete && !p.supported && !p.exportable);
    let mut opaque = Vec::new();
    p.tree.visit(&mut |node| if let NodeKind::UnparsedSource { raw } = &node.kind { opaque.push(raw.clone()); });
    assert_eq!(opaque, vec![r"\mystery*[x]{a{b}}{c}"]);
    assert!(p.diagnostics.iter().any(|d| d.code == DiagnosticCode::UnknownCommand));
}

#[test]
fn structural_macro_cannot_be_an_unbraced_tex_argument() {
    let bad = parse(r"\frac\frac12{3}");
    assert!(!bad.complete);
    assert!(bad.diagnostics.iter().any(|d| d.code == DiagnosticCode::UnbracedStructure));
    let good = parse(r"\frac{\frac12}{3}");
    assert!(good.complete && good.supported);
    assert_eq!(good.facts.fractions, 2);
}

#[test]
fn duplicate_scripts_remain_source_errors() {
    assert!(!parse(r"x^2^3").complete);
    assert!(!parse(r"x_1_2").complete);
    for input in [r"x^2_3", r"x_3^2", r"x^{y_2}"] { assert!(parse(input).exportable); }
}

#[test]
fn empty_slots_are_editable_but_not_export_ready() {
    for input in [r"\placeholder{}", r"\frac{}{y}", r"\frac{x}{}", r"x^{}", ""] {
        let p = parse(input);
        assert!(p.complete && p.supported, "{input}");
        assert!(!p.exportable, "{input}");
    }
}

#[test]
fn canonical_control_word_boundaries_are_semantically_safe() {
    let p = parse(r"\alpha x^2 + \beta a");
    let canonical = model::latex(&p.tree, LatexMode::Canonical);
    assert_eq!(canonical, r"\alpha x^{2}+\beta a");
    let q = parse(&canonical);
    assert!(q.exportable);
    assert_eq!(p.facts, q.facts);
}

#[test]
fn local_rewrite_preserves_untouched_argument_and_neighbor_source() {
    let original = " \\dfrac  {x} % numerator done\n { y } + \\alpha z ";
    let mut p = parse(original);
    let root = Arc::make_mut(&mut p.tree);
    root.origin = None;
    let NodeKind::Row { entries, .. } = &mut root.kind else { panic!() };
    let fraction = Arc::make_mut(&mut entries[0].node);
    fraction.origin = None;
    let NodeKind::Fraction { numerator, .. } = &mut fraction.kind else { panic!() };
    let numerator = Arc::make_mut(numerator);
    numerator.origin = None;
    let NodeKind::Row { entries, .. } = &mut numerator.kind else { panic!() };
    let symbol = Arc::make_mut(&mut entries[0].node);
    symbol.origin = None;
    let NodeKind::Symbol { value, .. } = &mut symbol.kind else { panic!() };
    *value = "w".into();
    assert_eq!(model::latex(&p.tree, LatexMode::Lossless), " \\dfrac  {w} % numerator done\n { y } + \\alpha z ");
}

#[test]
fn inserted_letter_cannot_merge_into_a_neighbor_command() {
    let mut p = parse(r"\alpha");
    let root = Arc::make_mut(&mut p.tree);
    root.origin = None;
    let NodeKind::Row { entries, .. } = &mut root.kind else { panic!() };
    entries.push(RowEntry { leading: String::new(), node: Arc::new(Node { id: NodeId(100), origin: None,
        kind: NodeKind::Symbol { value: "x".into(), class: AtomClass::Ordinary, command: None } }) });
    assert_eq!(model::latex(&p.tree, LatexMode::Lossless), r"\alpha x");
}

#[test]
fn editing_an_unbraced_argument_adds_only_required_grouping() {
    let mut p = parse(r"\frac12");
    let root = Arc::make_mut(&mut p.tree);
    root.origin = None;
    let NodeKind::Row { entries, .. } = &mut root.kind else { panic!() };
    let fraction = Arc::make_mut(&mut entries[0].node);
    fraction.origin = None;
    let NodeKind::Fraction { numerator, .. } = &mut fraction.kind else { panic!() };
    let numerator = Arc::make_mut(numerator);
    numerator.origin = None;
    let NodeKind::Row { entries, .. } = &mut numerator.kind else { panic!() };
    entries.push(RowEntry { leading: String::new(), node: Arc::new(Node { id: NodeId(100), origin: None,
        kind: NodeKind::Symbol { value: "z".into(), class: AtomClass::Ordinary, command: None } }) });
    let result = model::latex(&p.tree, LatexMode::Lossless);
    assert_eq!(result, r"\frac{1z}2");
    assert!(parse(&result).exportable);
}

#[test]
fn text_escape_semantics_and_comments_are_preserved() {
    let p = parse("\\text{a%ignore\nb\\%\\{c\\}\\textbackslash{}}");
    assert!(p.exportable);
    let canonical = model::latex(&p.tree, LatexMode::Canonical);
    assert_eq!(canonical, r"\text{ab\%\{c\}\textbackslash{}}");
    assert!(parse(&canonical).exportable);
    assert!(!parse(r"\text{\unknown{x}}").supported);
}

#[test]
fn parser_has_explicit_depth_node_and_source_budgets() {
    let deep = format!("{}x{}", "{".repeat(70), "}".repeat(70));
    let result = parser::parse(SourceText::new(1, deep.as_str()).unwrap(), NodeId(1), &mut IdAllocator::new(2));
    assert_eq!(result.unwrap_err().code, ErrorCode::ResourceLimit);
    let many = "x".repeat(9000);
    let result = parser::parse(SourceText::new(1, many.as_str()).unwrap(), NodeId(1), &mut IdAllocator::new(2));
    assert_eq!(result.unwrap_err().code, ErrorCode::ResourceLimit);
    assert_eq!(SourceText::new(1, "x".repeat(MAX_SOURCE_BYTES + 1)).unwrap_err().code, ErrorCode::ResourceLimit);
}

#[test]
fn seeded_malformed_corpus_never_loses_source_or_panics() {
    let pieces = ["x", "y", "+", "_", "^", "{", "}", "\\frac", "\\text", "\\alpha", " ", "%c\n", "中", "\\mystery", "'"];
    let mut seed = 0x3155_2267_u64;
    for _ in 0..500 {
        let mut input = String::new();
        for _ in 0..48 {
            seed = seed.wrapping_mul(6364136223846793005).wrapping_add(1442695040888963407);
            input.push_str(pieces[(seed >> 32) as usize % pieces.len()]);
        }
        let result = parser::parse(SourceText::new(1, input.as_str()).unwrap(), NodeId(1), &mut IdAllocator::new(2));
        match result {
            Ok(p) => assert_eq!(model::latex(&p.tree, LatexMode::Lossless), input),
            Err(error) => assert_eq!(error.code, ErrorCode::ResourceLimit, "{input}: {error}"),
        }
    }
}

#[test]
fn seeded_valid_structures_have_idempotent_canonical_serialization() {
    fn generate(seed: &mut u64, depth: usize) -> String {
        *seed = seed.wrapping_mul(6364136223846793005).wrapping_add(1);
        let choice = (*seed >> 32) as usize;
        if depth == 0 {
            return ["x", "12", "\\alpha", "\\text{速度}", "\\text{e\u{301}😀}"][choice % 5].into();
        }
        match choice % 4 {
            0 => format!("\\frac{{{}}}{{{}}}", generate(seed, depth - 1), generate(seed, depth - 1)),
            1 => format!("{{{}}}_{{{}}}^{{{}}}", generate(seed, depth - 1), generate(seed, depth - 1), generate(seed, depth - 1)),
            2 => format!("({}) + {}", generate(seed, depth - 1), generate(seed, depth - 1)),
            _ => generate(seed, 0),
        }
    }
    let mut seed = 42;
    for _ in 0..300 {
        let input = generate(&mut seed, 3);
        let p = parse(&input);
        assert!(p.exportable, "{input}");
        let canonical = model::latex(&p.tree, LatexMode::Canonical);
        let q = parse(&canonical);
        assert!(q.exportable, "{canonical}");
        assert_eq!(canonical, model::latex(&q.tree, LatexMode::Canonical));
        assert_eq!(p.facts, q.facts);
    }
}
