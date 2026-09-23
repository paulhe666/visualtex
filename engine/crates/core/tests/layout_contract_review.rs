//! Independent P05 checks against the public layout API. No renderer oracle,
//! production edits, screenshot constants or user-owned Office documents.
use std::path::PathBuf;
use visualtex_engine_core::fonts::FontRegistry;
use visualtex_engine_core::layout::{LayoutEngine, LayoutOptions, MathStyle};
use visualtex_engine_core::model::NodeKind;
use visualtex_engine_core::positions::Position;
use visualtex_engine_core::session::{EngineSession, Revision};
use visualtex_engine_core::source_edit::SourceEdit;
use visualtex_engine_core::ErrorCode;

fn setup() -> (FontRegistry, LayoutOptions) {
    let path = std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    let mut fonts = FontRegistry::default();
    let identity = fonts.register(&std::fs::read(path).unwrap(), 0).unwrap();
    (fonts, LayoutOptions::new(identity.id))
}
fn session(source: &str) -> EngineSession {
    EngineSession::from_source("p05-independent-layout".into(), source.into()).unwrap()
}
fn near(a: f64, b: f64) { assert!((a-b).abs() < 1e-9, "{a} != {b}"); }
fn geom(scene: &visualtex_engine_core::scene::LayoutSnapshot) -> serde_json::Value {
    serde_json::json!({"glyphs":scene.glyphs,"rules":scene.rules,"carets":scene.carets,
        "regions":scene.regions,"width":scene.width,"height":scene.height,"baseline":scene.baseline})
}

#[test]
fn explicit_fraction_style_overrides_all_ambient_styles_and_cached_results() {
    let (fonts, mut options) = setup();
    let mut layout = LayoutEngine::default();
    let display = layout.layout(&session(r"\frac{x}{y}"), &fonts, &options).unwrap();
    options.style = MathStyle::Text;
    let text = layout.layout(&session(r"\frac{x}{y}"), &fonts, &options).unwrap();
    assert!(display.height > text.height);
    for ambient in [MathStyle::Display, MathStyle::Text, MathStyle::Script, MathStyle::ScriptScript] {
        options.style = ambient;
        let forced_display = layout.layout(&session(r"\dfrac{x}{y}"), &fonts, &options).unwrap();
        let forced_text = layout.layout(&session(r"\tfrac{x}{y}"), &fonts, &options).unwrap();
        near(forced_display.height, display.height);
        near(forced_display.width, display.width);
        near(forced_text.height, text.height);
        near(forced_text.width, text.width);
        // A fresh cache is an independent check of invalidation correctness.
        assert_eq!(geom(&forced_display), geom(&LayoutEngine::default()
            .layout(&session(r"\dfrac{x}{y}"), &fonts, &options).unwrap()));
    }
}

#[test]
fn nested_forced_fractions_obey_the_same_rule_as_top_level_fractions() {
    let (fonts, options) = setup();
    let mut layout = LayoutEngine::default();
    let scene = layout.layout(&session(r"x^{\dfrac{a}{b}}+\frac{\tfrac{c}{d}}{y}"), &fonts, &options).unwrap();
    let m = fonts.get(&options.math_font).unwrap().math.as_ref().unwrap();
    // Outer, display-forced and text-forced fraction bars all retain full-size
    // metrics even though two are nested in otherwise smaller math styles.
    assert_eq!(scene.rules.len(), 3);
    for rule in &scene.rules { near(rule.rect.height(), m.fraction_rule); }
    for glyph in &scene.glyphs { assert!(glyph.scale >= m.scriptscript_scale - 1e-9); }
}

#[test]
fn blank_input_is_an_editable_draft_not_a_finished_vector_formula() {
    let (fonts, options) = setup();
    for source in ["", " ", "% not a formula", "{}", "{ }"] {
        let s = session(source);
        assert!(!s.document_snapshot().analysis.exportable);
        let scene = LayoutEngine::default().layout(&s, &fonts, &options).unwrap();
        assert!(!scene.carets.is_empty());
        assert!(!scene.export_ready, "blank source was labeled export-ready: {source:?}");
        assert_eq!(scene.svg().unwrap_err().code, ErrorCode::IncompleteSource);
    }
}

#[test]
fn vertical_navigation_inside_a_fraction_targets_the_other_argument() {
    let (fonts, options) = setup();
    let s = session(r"\frac{a+b}{c+d}");
    let doc = s.document_snapshot();
    let fraction = &doc.analysis.tree.children()[0];
    let NodeKind::Fraction { numerator, denominator, .. } = &fraction.kind else { panic!() };
    let scene = LayoutEngine::default().layout(&s, &fonts, &options).unwrap();
    for offset in 0..=3 {
        let from = Position::row(numerator.id, offset);
        let down = scene.vertical(from, true, None).unwrap();
        assert!(matches!(down, Position::Row {node_id,..} if node_id == denominator.id),
            "down from numerator {offset} escaped the fraction: {down:?}");
        let up = scene.vertical(Position::row(denominator.id, offset), false, None).unwrap();
        assert!(matches!(up, Position::Row {node_id,..} if node_id == numerator.id),
            "up from denominator {offset} escaped the fraction: {up:?}");
    }
}

#[test]
fn semantic_source_changes_never_reuse_stale_cached_geometry() {
    let (fonts, options) = setup();
    let mut layout = LayoutEngine::default();
    let mut s = session(r"\frac{x}{2}+\frac{x}{2}");
    for replacement in ["y", "\\text{AV}", "\\dfrac{a}{b}", "x_1", "\\frac{p}{q}"] {
        let old = s.source().to_owned();
        let start = old.find('{').unwrap()+1;
        // Replace the whole source explicitly here: local patch behavior is
        // tested separately. Identical document IDs must not make geometry stale.
        let next = format!("\\frac{{{replacement}}}{{2}}+\\frac{{x}}{{2}}");
        s.replace_source(s.revision(), next).unwrap();
        let cached = layout.layout(&s, &fonts, &options).unwrap();
        let fresh = LayoutEngine::default().layout(&s, &fonts, &options).unwrap();
        assert_eq!(geom(&cached), geom(&fresh), "source={}; old start={start}", s.source());
    }
}

#[test]
fn comment_only_patch_reuses_geometry_without_reusing_the_wrong_revision() {
    let (fonts, options) = setup();
    let mut s = session("x+1% old\n");
    let mut layout = LayoutEngine::default();
    let old = layout.layout(&s, &fonts, &options).unwrap();
    s.patch_source(Revision(0), SourceEdit { start_utf16: 5, end_utf16: 8, replacement: "new comment".into() }).unwrap();
    let next = layout.layout(&s, &fonts, &options).unwrap();
    assert_eq!(old.svg().unwrap(), next.svg().unwrap());
    assert_eq!(next.revision, Revision(1));
    assert_ne!(next.source_id, old.source_id);
    assert_eq!(next.rebuilt_fragments, 0);
}

#[test]
fn failed_font_or_draft_requests_cannot_poison_the_last_valid_scene() {
    let (fonts, options) = setup();
    let mut layout = LayoutEngine::default();
    let s = session(r"\frac{x_1}{y^2}");
    let old = layout.layout(&s, &fonts, &options).unwrap();
    let mut invalid = options.clone(); invalid.math_font = "unregistered-resource".into();
    assert!(layout.layout(&s, &fonts, &invalid).is_err());
    assert!(layout.layout(&session(r"\frac{x}{"), &fonts, &options).is_err());
    let next = layout.layout(&s, &fonts, &options).unwrap();
    assert_eq!(old.svg().unwrap(), next.svg().unwrap());
    assert_eq!(geom(&old), geom(&next));
}
