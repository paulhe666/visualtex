//! Test the scene/export trust boundary with copies, never production snapshots.
use std::path::PathBuf;
use std::sync::Arc;
use visualtex_engine_core::fonts::{FontRegistry, PathOp};
use visualtex_engine_core::layout::{LayoutEngine, LayoutOptions};
use visualtex_engine_core::scene::LayoutSnapshot;
use visualtex_engine_core::session::EngineSession;

fn scene() -> LayoutSnapshot {
    let p = std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    let mut fonts = FontRegistry::default();
    let id = fonts.register(&std::fs::read(p).unwrap(), 0).unwrap();
    let s = EngineSession::from_source("scene-review".into(), r"\frac{f^2}{y_1}".into()).unwrap();
    LayoutEngine::default().layout(&s, &fonts, &LayoutOptions::new(id.id)).unwrap()
}

#[test]
fn correct_scene_roundtrips_and_export_does_not_mutate_it() {
    let s = scene();
    let before = serde_json::to_value(&s).unwrap();
    s.validate().unwrap();
    assert_eq!(s.svg().unwrap(), s.svg().unwrap());
    assert_eq!(serde_json::to_value(&s).unwrap(), before);
    assert_eq!(s.glyphs.len(), 4);
}

#[test]
fn natural_css_dimensions_cannot_silently_diverge_from_em_geometry() {
    let s = scene();
    let cases: &[(&str, fn(&mut LayoutSnapshot))] = &[
        ("width unit drift", |s| s.width *= 0.75),
        ("height unit drift", |s| s.height *= 2.0),
        ("baseline unit drift", |s| s.baseline *= 0.75),
        ("negative font scale", |s| s.font_size_px = -16.0),
        ("zero font scale", |s| s.font_size_px = 0.0),
        ("negative width em", |s| s.width_em = -1.0),
        ("zero height em", |s| s.height_em = 0.0),
        ("negative advance", |s| s.advance_em = -1.0),
    ];
    let mut accepted = Vec::new();
    for (name, mutate) in cases {
        let mut broken = s.clone(); mutate(&mut broken);
        if broken.validate().is_ok() || broken.svg().is_ok() { accepted.push(*name); }
    }
    assert!(accepted.is_empty(), "inconsistent public scene accepted: {accepted:?}");
}

#[test]
fn glyph_path_nan_and_infinity_are_rejected_before_svg_or_json_output() {
    let s = scene();
    for bad in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
        let mut broken = s.clone();
        Arc::make_mut(&mut broken.resources[0].shape).path.push(PathOp::Line([bad, 0.0]));
        assert!(broken.validate().is_err(), "non-finite resource path was not validated");
        assert!(broken.svg().is_err(), "SVG must not encode a non-finite path");
    }
}

#[test]
fn glyph_resource_metrics_are_part_of_scene_validation() {
    let s = scene();
    for field in 0..3 {
        let mut broken = s.clone();
        let shape = Arc::make_mut(&mut broken.resources[0].shape);
        match field {
            0 => shape.advance = f64::INFINITY,
            1 => shape.italic_correction = f64::NAN,
            _ => shape.ink = Some([0.0, 0.0, -1.0, 1.0]),
        }
        assert!(broken.validate().is_err(), "resource metric case {field} accepted");
    }
}

#[test]
fn glyph_resource_reference_and_font_provenance_must_both_resolve() {
    let s = scene();
    let mut broken = s.clone(); broken.glyphs[0].resource = s.resources.len();
    assert!(broken.validate().is_err());
    let mut broken = s.clone(); broken.resources[0].font_id = "not-in-snapshot-fonts".into();
    assert!(broken.validate().is_err(), "unbound font resource accepted");
    let mut broken = s.clone(); broken.fonts.clear();
    assert!(broken.validate().is_err(), "resource provenance silently lost");
}

#[test]
fn out_of_frame_ink_cannot_be_reported_as_an_unclipped_export() {
    let s = scene();
    let mut broken = s.clone(); broken.glyphs[0].x += s.width_em + 10.0;
    assert!(broken.validate().is_err(), "glyph outside the natural output frame was accepted");
    let mut broken = s.clone(); broken.rules[0].rect.y_max += s.height_em + 10.0;
    assert!(broken.validate().is_err(), "fraction bar outside the frame was accepted");
}

#[test]
fn invalid_hit_and_navigation_queries_leave_the_valid_snapshot_unchanged() {
    let s = scene(); let before = serde_json::to_value(&s).unwrap();
    for invalid in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
        assert!(s.hit_test(invalid, 0.0).is_err());
        assert!(s.hit_test(0.0, invalid).is_err());
        assert!(s.vertical(s.carets[0].position, true, Some(invalid)).is_err());
    }
    assert_eq!(serde_json::to_value(&s).unwrap(), before);
}
