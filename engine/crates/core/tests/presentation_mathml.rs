use std::sync::Arc;
use serde_json::json;
use visualtex_engine_core::{
    ErrorCode,
    custom_symbols::SymbolCatalog,
    model::IdAllocator,
    parser,
    presentation_mathml,
    session::NodeId,
    source::SourceText,
};

fn parse(source: &str) -> parser::ParseResult {
    parser::parse(
        SourceText::new(1, source.to_owned()).unwrap(),
        NodeId(1),
        &mut IdAllocator::new(2),
    ).unwrap()
}

fn encode(source: &str, display: bool) -> presentation_mathml::PresentationMathMl {
    let parsed = parse(source);
    assert!(parsed.complete, "source diagnostics: {:?}", parsed.diagnostics);
    assert!(parsed.supported, "source unsupported: {:?}", parsed.diagnostics);
    presentation_mathml::encode(&parsed, display).unwrap()
}

#[test]
fn named_alternates_keep_scalar_annotation_styles_and_scripts() {
    for spec in visualtex_engine_core::symbol_catalog::ALTERNATE_SYMBOLS {
        for display in [false, true] {
            let token = format!("<mo data-mjx-alternate=\"1\">{}</mo>", spec.value);
            assert!(encode(spec.command, display).math_ml.contains(&token));
            for (wrapper, variant) in [(r"\mathit", "italic"), (r"\mathbf", "bold"), (r"\mathbfit", "bold-italic")] {
                let source = format!(r"x^{{{wrapper}{{{}}}}}_{{i}}", spec.command);
                let mml = encode(&source, display).math_ml;
                assert!(mml.contains(&token), "{mml}");
                assert!(mml.contains(&format!("<mstyle mathvariant=\"{variant}\">")), "{mml}");
                assert!(mml.contains("<msubsup>"), "{mml}");
                assert!(!mml.contains('\u{E007}'), "Font-local PUA leaked into semantics");
            }
        }
    }
    for source in [r"\parallel", r"\nparallel", "∥", "∦"] {
        assert!(!encode(source, false).math_ml.contains("data-mjx-alternate"));
    }
}

#[test]
fn common_structures_are_encoded_from_the_retained_tree() {
    let result = encode(r"\frac{x_1+\text{速度}}{\sqrt{y^2}}", false);
    assert!(result.math_ml.starts_with("<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"inline\">"));
    for token in ["<mfrac>", "<msub>", "<mtext>速度</mtext>", "<msqrt>", "<msup>"] {
        assert!(result.math_ml.contains(token), "missing {token}: {}", result.math_ml);
    }
    for legacy in ["MathJax", "MathLive", "data-mml-node"] {
        assert!(!result.math_ml.contains(legacy));
    }
    assert!(result.aligned_grid_ids.is_empty());
}

#[test]
fn operator_limit_modes_choose_native_mathml_script_containers() {
    let sum = encode(r"\sum\limits_{i=0}^{n}x_i", false).math_ml;
    assert!(sum.contains("<munderover><mo largeop=\"true\">∑</mo>"), "{sum}");
    let integral = encode(r"\int_0^1 x\,\mathrm{d}x", true).math_ml;
    assert!(integral.contains("<msubsup><mo largeop=\"true\">∫</mo>"), "{integral}");
    let display_sum = encode(r"\sum_{i=0}^{n}x_i", true).math_ml;
    assert!(display_sum.contains("<munderover><mo largeop=\"true\">∑</mo>"), "{display_sum}");
}

#[test]
fn aligned_grid_exposes_private_identity_and_math_type_shape_without_svg_dom() {
    let result = encode(
        r"\begin{aligned}a&=b&c&=d\\x&=y&u&=v\end{aligned}",
        true,
    );
    assert_eq!(result.aligned_grid_ids.len(), 1);
    let id = result.aligned_grid_ids[0];
    assert!(result.math_ml.contains("columnalign=\"right left right left\""), "{}", result.math_ml);
    assert!(result.math_ml.contains("columnspacing=\"0em 2em 0em\""), "{}", result.math_ml);
    assert!(result.math_ml.contains(&format!("data-visualtex-core-grid-id=\"{id}\"")));
    assert!(result.math_ml.contains("data-mtef-tabs=\"true\""));
    assert_eq!(result.math_ml.matches("<mtr>").count(), 2);
    assert_eq!(result.math_ml.matches("<mtd>").count(), 8);
}

#[test]
fn matrices_cases_and_arrays_keep_environment_semantics() {
    let matrix = encode(r"\begin{pmatrix}a&b\\c&d\end{pmatrix}", true).math_ml;
    assert!(matrix.contains("<mo fence=\"true\" stretchy=\"true\">(</mo><mtable"), "{matrix}");
    assert!(matrix.contains("</mtable><mo fence=\"true\" stretchy=\"true\">)</mo>"), "{matrix}");

    let cases = encode(r"\begin{cases}x&x>0\\-x&x\le0\end{cases}", true).math_ml;
    assert!(cases.contains("<mo fence=\"true\" stretchy=\"true\">{</mo><mtable columnalign=\"left left\""), "{cases}");

    let array = encode(r"\begin{array}{r|cl}a&b&c\\d&e&f\end{array}", true).math_ml;
    assert!(array.contains("columnalign=\"right center left\""), "{array}");
}

#[test]
fn style_color_accent_annotation_and_bounded_macros_remain_semantic() {
    let source = r"\mathbf{x}+\textcolor{red}{y}+\widehat{z}+\overset{a}{b}+\dv{q}{t}+\braket{u}{v}";
    let result = encode(source, true).math_ml;
    assert!(result.contains("<mstyle mathvariant=\"bold\">"), "{result}");
    assert!(result.contains("<mstyle mathcolor=\"#ff0000\">"), "{result}");
    assert!(result.contains("<mover accent=\"true\">"), "{result}");
    assert!(result.contains("<mover><mi>b</mi><mi>a</mi></mover>"), "{result}");
    assert!(result.contains("<mfrac>"), "{result}");
    assert!(result.contains("<mo>⟨</mo>"), "{result}");
}

fn custom_symbol(fallback: Option<&str>) -> serde_json::Value {
    json!({
        "id":"p11-symbol",
        "command":"pelevensymbol",
        "name":"P11 Symbol",
        "role":"ordinary",
        "limitsBehavior":"auto",
        "metrics":{"widthEm":0.8,"ascentEm":0.7,"descentEm":0.1},
        "artwork":{"shapes":[{"kind":"rect","x":100,"y":100,"width":600,"height":500,"fill":true}]},
        "ommlFallback":fallback,
        "createdAt":1,
        "updatedAt":1
    })
}

#[test]
fn custom_symbol_requires_and_uses_its_frozen_office_fallback() {
    let catalog = SymbolCatalog::freeze(vec![custom_symbol(Some(r"\star"))]).unwrap();
    let parsed = parser::parse_with_symbols(
        SourceText::new(1, r"1+\pelevensymbol".to_owned()).unwrap(),
        NodeId(1),
        &mut IdAllocator::new(2),
        Arc::clone(&catalog),
    ).unwrap();
    let result = presentation_mathml::encode(&parsed, false).unwrap();
    assert!(result.math_ml.contains("⋆") || result.math_ml.contains("★") || result.math_ml.contains("∗"), "{}", result.math_ml);
    assert!(!result.math_ml.contains("pelevensymbol"));

    let catalog = SymbolCatalog::freeze(vec![custom_symbol(None)]).unwrap();
    let parsed = parser::parse_with_symbols(
        SourceText::new(1, r"\pelevensymbol".to_owned()).unwrap(),
        NodeId(1),
        &mut IdAllocator::new(2),
        catalog,
    ).unwrap();
    let error = presentation_mathml::encode(&parsed, false).unwrap_err();
    assert_eq!(error.code, ErrorCode::UnsupportedSyntax);
}

#[test]
fn explicit_and_named_atom_classes_survive_presentation_mathml() {
    for display in [false, true] {
        for source in [r"\varparallel", r"\parallelslant", r"\mathrel{⫽}"] {
            let math_ml = encode(source, display).math_ml;
            assert!(math_ml.contains("<mrow data-mjx-texclass=\"REL\"><mo>⫽</mo></mrow>"), "{math_ml}");
        }
        for source in [r"\sslash", "⫽"] {
            let math_ml = encode(source, display).math_ml;
            assert!(math_ml.contains("<mo>⫽</mo>"), "{math_ml}");
            assert!(!math_ml.contains("texclass=\"REL\""), "{math_ml}");
        }
        let operator = encode(r"\mathop{x+y}_i^n", display).math_ml;
        assert!(operator.contains("data-mjx-texclass=\"OP\""), "{operator}");
        assert!(operator.contains(if display { "<munderover>" } else { "<msubsup>" }), "{operator}");
        let binary = encode(r"x^{a\mathbin{=}b}", display).math_ml;
        assert!(binary.contains("<mrow data-mjx-texclass=\"BIN\"><mo>=</mo></mrow>"), "{binary}");
        let composite = encode(r"\mathrel{a+b}", display).math_ml;
        assert!(composite.contains("data-mjx-texclass=\"REL\""), "{composite}");
        assert!(composite.contains("<mi>a</mi><mo>+</mo><mi>b</mi>"), "{composite}");
        let parallel = encode(r"\mathit{\parallel}", display).math_ml;
        assert!(parallel.contains("<mstyle mathvariant=\"italic\"><mo>∥</mo></mstyle>"), "{parallel}");
        assert!(!parallel.contains("⫽"), "{parallel}");
    }
}

#[test]
fn incomplete_and_opaque_sources_fail_closed() {
    let parsed = parse(r"\frac{x}{");
    let error = presentation_mathml::encode(&parsed, true).unwrap_err();
    assert_eq!(error.code, ErrorCode::IncompleteSource);

    let parsed = parse(r"\futurecommand{x}");
    let error = presentation_mathml::encode(&parsed, true).unwrap_err();
    assert_eq!(error.code, ErrorCode::UnsupportedSyntax);
}
