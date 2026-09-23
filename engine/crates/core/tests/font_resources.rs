use std::path::PathBuf;
use visualtex_engine_core::fonts::{FontAsset, FontRegistry};

fn math_bytes() -> Vec<u8> {
    let path = std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    std::fs::read(&path).unwrap_or_else(|e| panic!("Required fixed math font fixture {}: {e}", path.display()))
}
#[test]
fn font_identity_uses_bytes_and_face_not_family_name() {
    let bytes = math_bytes(); let mut registry = FontRegistry::default();
    let first = registry.register(&bytes, 0).unwrap();
    assert_eq!(first.sha256, "6075562b771f8b82f0c179e363389684f2dd09de30038269e2628e504bd7be0f");
    assert_eq!(first.id, registry.register(&bytes, 0).unwrap().id);
    assert_eq!(registry.len(), 1);
    assert!(registry.register(&bytes, 100).is_err());
    assert_eq!(registry.len(), 1);
    assert!(registry.get("Latin Modern Math").is_err());
}
#[test]
fn math_metrics_and_glyphs_come_from_the_actual_font() {
    let font = FontAsset::from_bytes(&math_bytes(), 0).unwrap();
    let math = font.math.as_ref().unwrap();
    assert!(math.axis > 0.0 && math.fraction_rule > 0.0);
    assert!(math.scriptscript_scale <= math.script_scale && math.script_scale < 1.0);
    let id = font.glyph_id('x').unwrap(); let x = font.glyph(id).unwrap();
    assert!(x.advance > 0.0 && x.ink.is_some() && !x.path.is_empty());
    assert!(std::sync::Arc::ptr_eq(&x, &font.glyph(id).unwrap()));
    let space = font.glyph(font.glyph_id(' ').unwrap()).unwrap();
    assert!(space.advance > 0.0 && space.path.is_empty());
}
#[test]
fn corrupt_fonts_and_missing_glyphs_are_not_silently_replaced() {
    assert!(FontAsset::from_bytes(&[0; 128], 0).is_err());
    let font = FontAsset::from_bytes(&math_bytes(), 0).unwrap();
    assert!(font.glyph_id('\u{10ffff}').is_err());
    assert!(font.glyph(0).is_err());
    assert!(font.shape_text("速度").is_err());
    assert!(font.shape_text("שלום").is_err());
}
#[test]
fn text_is_shaped_with_real_advances_and_clusters() {
    let font = FontAsset::from_bytes(&math_bytes(), 0).unwrap();
    let run = font.shape_text("office fi AV").unwrap();
    assert!(run.advance > 0.0);
    assert_eq!(run.glyphs.len(), 12);
    assert_eq!(run.glyphs.first().unwrap().cluster, 0);
    assert!(run.glyphs.windows(2).all(|p| p[0].cluster <= p[1].cluster));
    assert!(font.shape_text("").unwrap().glyphs.is_empty());
}
