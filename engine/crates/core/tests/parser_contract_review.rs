//! Independent P03 parser regression corpus and deterministic generative tests.
use std::collections::BTreeSet;
use visualtex_engine_core::model::{self, IdAllocator, LatexMode, Node, NodeKind, latex, validate_tree};
use visualtex_engine_core::parser::{self, ParseResult, Severity};
use visualtex_engine_core::session::NodeId;
use visualtex_engine_core::source::SourceText;
use visualtex_engine_core::ErrorCode;
use serde_json::Value;

fn parse(text: &str) -> ParseResult {
    parser::parse(SourceText::new(1, text).unwrap(), NodeId(1), &mut IdAllocator::new(2)).unwrap()
}
fn semantic_tokens(node: &Node, out: &mut Vec<String>) {
    match &node.kind {
        NodeKind::Symbol { value, class, .. } => out.push(format!("symbol:{class:?}:{value}")),
        NodeKind::Text { text, .. } => out.push(format!("text:{text}")),
        NodeKind::Fraction { command, .. } => out.push(format!("fraction:{command}")),
        NodeKind::Radical { degree, .. } => out.push(format!("radical:indexed={}", degree.is_some())),
        NodeKind::Fence { left, right, .. } => out.push(format!("fence:{left}:{right}")),
        NodeKind::MiddleDelimiter { delimiter, .. } => out.push(format!("middle:{delimiter}")),
        NodeKind::LargeOperator { command, parts } => out.push(format!("operator:{command}:{:?}:{:?}",visualtex_engine_core::operators::effective_mode(parts),visualtex_engine_core::operators::slots(parts).map(|s|s.script).collect::<Vec<_>>())),
        NodeKind::Style { command, variant, .. } => out.push(format!("style:{command}:{variant:?}")),
        NodeKind::Space { command, width_mu } => out.push(format!("space:{command}:{width_mu}")),
        NodeKind::Phantom { command, axis, .. } => out.push(format!("phantom:{command}:{axis:?}")),
        NodeKind::Boxed { .. } => out.push("boxed".into()),
        NodeKind::Color {spec,..}=>out.push(format!("color:{}",spec.color.hex())),
        NodeKind::ColorChange {spec}=>out.push(format!("colorChange:{}",spec.color.hex())),
        NodeKind::StyleChange {command,..}=>out.push(format!("styleChange:{command}")),
        NodeKind::GeneralOperator { command, parts, .. } => out.push(format!("generalOperator:{command}:{:?}:{:?}",visualtex_engine_core::operators::effective_mode(parts),visualtex_engine_core::operators::slots(parts).map(|s|s.script).collect::<Vec<_>>())),
        NodeKind::MacroCall { command, macro_kind, arguments, .. } => out.push(format!("macro:{command}:{macro_kind:?}:{}",arguments.len())),
        NodeKind::Grid { environment, columns, rows, .. } => out.push(format!("grid:{environment:?}:{}:{}",columns.len(),rows.len())),
        NodeKind::GridRow { cells, .. } => out.push(format!("gridRow:{}",cells.len())),
        NodeKind::Accent { command, .. } => out.push(format!("accent:{command}")),
        NodeKind::Annotation { command, above, .. } => out.push(format!("annotation:{command}:above={above}")),
        NodeKind::Scripts { slots, .. } => out.push(format!("scripts:{:?}", slots.iter().map(|s| s.script).collect::<Vec<_>>())),
        NodeKind::EmptySlot { label } => out.push(format!("slot:{label}")),
        NodeKind::UnparsedSource { raw } => out.push(format!("opaque:{raw}")),
        NodeKind::CustomSymbol { command, asset_sha256, parts, .. } => out.push(format!("customSymbol:{command}:{asset_sha256}:{:?}",visualtex_engine_core::operators::effective_mode(parts))),
        NodeKind::Row { .. } => {},
    }
    for child in node.children() { semantic_tokens(child, out); }
}
fn semantics(result: &ParseResult) -> Vec<String> {
    let mut out = Vec::new(); semantic_tokens(&result.tree, &mut out); out
}
fn verify_node_ranges(result: &ParseResult) {
    let mut seen = BTreeSet::new();
    result.tree.visit(&mut |node| {
        assert!(seen.insert(node.id), "duplicate identity");
        if let Some(origin) = &node.origin {
            let span = origin.span;
            assert_eq!(span.source_id, result.source.id());
            assert!(span.byte_start <= span.byte_end);
            assert!((span.byte_end as usize) <= result.source.len());
            assert_eq!(origin.text(), &result.source.text()[span.byte_start as usize..span.byte_end as usize]);
        }
    });
    assert_eq!(seen.len(), result.facts.nodes);
}

#[test]
fn independent_k0_corpus_lossless_classification_and_semantic_round_trip() {
    let corpus: Value = serde_json::from_str(include_str!("../../../fixtures/k0-review-corpus.json")).unwrap();
    for case in corpus["cases"].as_array().unwrap() {
        let id = case["id"].as_str().unwrap();
        let source = case["source"].as_str().unwrap();
        let parsed = parse(source);
        assert_eq!(latex(&parsed.tree, LatexMode::Lossless), source, "lossless case={id}");
        verify_node_ranges(&parsed);
        match case["expect"].as_str().unwrap() {
            "supported" => {
                assert!(parsed.complete && parsed.supported, "case={id}, diagnostics={:?}", parsed.diagnostics);
                if let Some(expected) = case["canonical"].as_str() {
                    assert_eq!(latex(&parsed.tree, LatexMode::Canonical), expected, "canonical case={id}");
                }
                assert_eq!(parsed.facts.fractions, case["fractions"].as_u64().unwrap() as usize, "case={id}");
                assert_eq!(parsed.facts.scripts, case["scripts"].as_u64().unwrap() as usize, "case={id}");
                let normalized = parse(&latex(&parsed.tree, LatexMode::Canonical));
                assert!(normalized.complete && normalized.supported, "canonical unsupported case={id}");
                assert_eq!(semantics(&parsed), semantics(&normalized), "semantic round trip case={id}");
                if case["unfilled"].as_bool() == Some(true) {
                    assert!(parsed.tree.has_unfilled_slots());
                    assert!(!parsed.exportable);
                }
            }
            "incomplete" | "invalid" => {
                assert!(!parsed.complete && !parsed.exportable, "case={id}");
                assert!(parsed.diagnostics.iter().any(|d| d.severity == Severity::Error), "case={id}");
            }
            "unsupported" => {
                assert!(!parsed.supported && !parsed.exportable, "case={id}");
                assert!(parsed.diagnostics.iter().any(|d| d.severity == Severity::Unsupported), "case={id}");
            }
            expectation => panic!("unknown expectation {expectation}"),
        }
    }
}

#[test]
fn parser_rejects_resource_exhaustion_without_losing_caller_source() {
    for text in [format!("{}x{}", "{".repeat(80), "}".repeat(80)),
        "x".repeat(model::MAX_NODES + 2), "\\unknown ".repeat(200)] {
        let source = SourceText::new(1, text.clone()).unwrap();
        let error = parser::parse(source.clone(), NodeId(1), &mut IdAllocator::new(2)).unwrap_err();
        assert_eq!(error.code, ErrorCode::ResourceLimit);
        assert_eq!(source.text(), text);
    }
}

#[test]
fn repeated_subexpressions_do_not_share_any_edit_identity() {
    let parsed = parse("\\frac{x_1}{y^2}+\\frac{x_1}{y^2}");
    let NodeKind::Row { entries, .. } = &parsed.tree.kind else { panic!("root must be row") };
    let mut left = BTreeSet::new();
    let mut right = BTreeSet::new();
    entries[0].node.visit(&mut |node| { left.insert(node.id); });
    entries[2].node.visit(&mut |node| { right.insert(node.id); });
    assert!(left.is_disjoint(&right));
    assert_eq!(validate_tree(&parsed.tree).unwrap(), parsed.facts);
}

#[test]
fn unparsed_command_keeps_nested_optional_arguments_and_star() {
    for source in ["\\unknown*[a{[b]}c]{x}{y}", "\\unknown{\\frac{a}{b}}", "a+\\unknown[速度😀]{é}+b"] {
        let parsed = parse(source);
        assert!(!parsed.supported && !parsed.exportable);
        assert_eq!(latex(&parsed.tree, LatexMode::Lossless), source);
        assert_eq!(latex(&parsed.tree, LatexMode::Canonical), source);
    }
}

#[test]
fn placeholder_label_is_not_discarded_by_canonical_serialization() {
    let parsed = parse("\\placeholder{numerator}");
    assert_eq!(latex(&parsed.tree, LatexMode::Canonical), "\\placeholder{numerator}");
    assert_eq!(semantics(&parsed), semantics(&parse(&latex(&parsed.tree, LatexMode::Canonical))));
}

fn next(seed: &mut u64) -> u64 {
    *seed ^= *seed << 13; *seed ^= *seed >> 7; *seed ^= *seed << 17; *seed
}
fn generated(seed: &mut u64, depth: usize) -> String {
    if depth == 0 {
        return ["x", "12", "\\alpha x", "\\text{速度 é}", "\\%", "(a+b)"][(next(seed) % 6) as usize].into();
    }
    match next(seed) % 5 {
        0 => format!("\\frac{{{}}}{{{}}}", generated(seed, depth - 1), generated(seed, depth - 1)),
        1 => format!("{{{}}}_{{{}}}^{{{}}}", generated(seed, depth - 1), generated(seed, depth - 1), generated(seed, depth - 1)),
        2 => format!("{{{}}}", generated(seed, depth - 1)),
        3 => format!("{} %keep\r\n+ {}", generated(seed, depth - 1), generated(seed, depth - 1)),
        _ => generated(seed, 0),
    }
}

#[test]
fn seeded_recursive_k0_round_trips_without_rendering_dependency() {
    const SEED: u64 = 0x20260913_4b305445;
    let mut seed = SEED;
    for index in 0..512 {
        let source = generated(&mut seed, 3);
        let parsed = parse(&source);
        assert!(parsed.complete && parsed.supported, "seed={SEED:x} case={index}: {source}");
        assert_eq!(latex(&parsed.tree, LatexMode::Lossless), source);
        let canonical = latex(&parsed.tree, LatexMode::Canonical);
        let again = parse(&canonical);
        assert!(again.complete && again.supported, "seed={SEED:x} case={index}: {canonical}");
        assert_eq!(canonical, latex(&again.tree, LatexMode::Canonical), "canonical not idempotent seed={SEED:x} case={index}");
        assert_eq!(semantics(&parsed), semantics(&again), "seed={SEED:x} case={index}");
        verify_node_ranges(&parsed);
    }
}
