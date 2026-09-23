//! Seeded P08 source disturbances and cross-feature presentation invariants.
//! Every editable state is compared with a fresh full parse; we do not merely
//! compare serialized source strings while leaving stale mathematical children.
use serde_json::Value;
use visualtex_engine_core::{
    model::{self, IdAllocator, LatexMode},
    parser::{self, ParseResult},
    session::{EngineSession, NodeId},
    source::SourceText,
    source_edit::SourceEdit,
    ErrorCode,
};
fn parse(source: &str) -> Result<ParseResult, visualtex_engine_core::EngineError> {
    parser::parse(SourceText::new(1, source)?, NodeId(1), &mut IdAllocator::new(2))
}
fn semantic(value: Value) -> Value {
    match value {
        Value::Array(values) => Value::Array(values.into_iter().map(semantic).collect()),
        Value::Object(map) => Value::Object(map.into_iter()
            .filter(|(k,_)| k != "id" && k != "origin" && k != "sourceId")
            .map(|(k,v)| {
                let v = if k == "columns" { Value::from(v.as_array().unwrap().len()) } else { semantic(v) };
                (k, v)
            }).collect()),
        value => value,
    }
}
fn verify_ranges(p: &ParseResult, expected: &str) {
    assert_eq!(model::latex(&p.tree, LatexMode::Lossless), expected);
    let mut at = 0;
    for token in &p.tokens { assert_eq!(token.span.byte_start as usize, at); at = token.span.byte_end as usize; }
    assert_eq!(at, expected.len());
    p.tree.visit(&mut |node| {
        let origin = node.origin.as_ref().expect("Every parsed/projection node has its own source range");
        let span = origin.span;
        assert_eq!(span.source_id, p.source.id());
        assert!(span.byte_start <= span.byte_end && span.byte_end as usize <= expected.len());
        assert_eq!(origin.text(), &expected[span.byte_start as usize..span.byte_end as usize]);
        assert_eq!(expected[..span.byte_start as usize].encode_utf16().count(), span.utf16_start as usize);
        assert_eq!(expected[..span.byte_end as usize].encode_utf16().count(), span.utf16_end as usize);
    });
}
fn next(state: &mut u32) -> u32 { *state = state.wrapping_mul(1_664_525).wrapping_add(1_013_904_223); *state }

#[test]
fn seeded_environment_boundary_patches_match_full_parser_without_losing_source() {
    const SEED: u32 = 0x0800_2026;
    let sources = [
        r"\begin{gathered}a+b\\c-d\end{gathered}",
        r"\begin{gathered}\frac{x_1}{y}\\\text{中文 é}\end{gathered}",
        r"\textcolor{red}{\begin{gathered}\mathbf{x}\\\sqrt{y}\end{gathered}}",
        r"\begin{gathered}\begin{gathered}a\\b\end{gathered}\\c\end{gathered}",
        "\\begin %name\n {gathered}a %not separator \\\\\n+b\\\\\n c\\end{gathered}",
    ];
    let replacements = ["", "x", "}", "{", "%", "\n", "\\", "\\\\", "&", "中", "é", "[", "*", "\\begin{gathered}", "\\end{gathered}", "\\future{x}"];
    let mut rng = SEED;
    for case in 0..640 {
        let source = sources[next(&mut rng) as usize % sources.len()];
        let boundaries: Vec<_> = source.char_indices().map(|(i,_)|i).chain(std::iter::once(source.len())).collect();
        let start_index = next(&mut rng) as usize % boundaries.len();
        let end_index = (start_index + next(&mut rng) as usize % 4).min(boundaries.len()-1);
        let start = boundaries[start_index]; let end = boundaries[end_index];
        let replacement = replacements[next(&mut rng) as usize % replacements.len()];
        let expected = format!("{}{}{}", &source[..start], replacement, &source[end..]);
        let mut session = EngineSession::from_source(format!("grid-seed-{case}"), source.into()).unwrap();
        let before = serde_json::to_value(session.editor_snapshot().unwrap()).unwrap();
        let fresh = parse(&expected);
        let patched = session.patch_source(session.revision(), SourceEdit {
            start_utf16: source[..start].encode_utf16().count() as u32,
            end_utf16: source[..end].encode_utf16().count() as u32,
            replacement: replacement.into(),
        });
        match (fresh, patched) {
            (Ok(fresh), Ok((actual,_))) => {
                verify_ranges(&actual.analysis, &expected);
                assert_eq!(semantic(serde_json::to_value(&actual.analysis.tree).unwrap()), semantic(serde_json::to_value(&fresh.tree).unwrap()),
                    "seed={SEED:#x} case={case} edit={start}..{end} input={expected:?}");
                assert_eq!(actual.analysis.facts, fresh.facts);
                assert_eq!(actual.analysis.complete, fresh.complete);
                assert_eq!(actual.analysis.supported, fresh.supported);
                assert_eq!(actual.analysis.exportable, fresh.exportable);
            }
            (Err(expected_error), Err(actual_error)) => {
                assert_eq!(actual_error.code, expected_error.code, "seed={SEED:#x} case={case}");
                assert_eq!(serde_json::to_value(session.editor_snapshot().unwrap()).unwrap(), before);
            }
            (fresh, patched) => panic!("seed={SEED:#x} case={case} input={expected:?}: fresh={} patched={}", fresh.is_ok(), patched.is_ok()),
        }
    }
}

#[test]
fn unknown_nested_environments_cannot_bypass_the_parser_depth_budget() {
    let source = format!("{}x{}", r"\begin{future}".repeat(80), r"\end{future}".repeat(80));
    assert!(matches!(parse(&source), Err(error) if error.code == ErrorCode::ResourceLimit),
        "Unknown nested environments bypassed MAX_DEPTH by skipping recursive descent");
}

#[test]
fn source_physical_newlines_do_not_create_extra_grid_rows() {
    for source in [
        "\\begin{gathered}\na+b\n+c\n\\end{gathered}",
        "\\begin{gathered}\r\na+b\r\n+c\r\n\\end{gathered}",
        "\\begin{gathered}\\text{line\n text}\n\\end{gathered}",
    ] {
        let p = parse(source).unwrap();
        assert!(p.complete && p.supported); assert_eq!(p.facts.grid_rows, 1);
        verify_ranges(&p, source);
    }
}

#[test]
fn row_gap_options_and_unimplemented_environments_are_explicit_not_fake_formulas() {
    for source in [
        r"\begin{gathered}[t]x\\y\end{gathered}",
        r"\begin{gathered}x\\[2pt]y\end{gathered}",
        r"\begin{gathered}x\\*y\end{gathered}",
        r"\begin{future}x\\y\end{future}",
    ] {
        let p = parse(source).unwrap(); assert!(!p.supported && !p.exportable, "{source}");
        verify_ranges(&p, source);
    }
}
