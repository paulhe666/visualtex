//! App-private formula-font acceptance. These tests deliberately use only files
//! shipped with VisualTeX: no installed fonts, OS APIs, browser CSS or MathLive
//! runtime. MathLive 0.109.2 is a development oracle for the KaTeX role policy;
//! VisualTeX owns all syntax, layout, editing, selection and history semantics.
use std::{
    collections::{BTreeMap, BTreeSet},
    path::PathBuf,
};
use visualtex_engine_core::{
    fonts::FontRegistry,
    layout::{LayoutEngine, LayoutOptions, MathStyle},
    scene::LayoutSnapshot,
    session::EngineSession,
};

fn root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../../apps/windows/src-tauri/resources/formula-fonts")
}

fn load(fonts: &mut FontRegistry, file: &str) -> String {
    fonts
        .register(
            &std::fs::read(root().join(file)).expect("required bundled font missing"),
            0,
        )
        .expect("bundled font rejected")
        .id
}

fn resources() -> (FontRegistry, LayoutOptions, BTreeMap<String, String>) {
    let mut fonts = FontRegistry::default();
    let math = load(&mut fonts, "latinmodern-math.otf");
    let mut layout = LayoutOptions::new(math);
    let mut ids = BTreeMap::new();
    for (role, file) in [
        ("main", "KaTeX_Main-Regular.ttf"),
        ("mainItalic", "KaTeX_Main-Italic.ttf"),
        ("mainBold", "KaTeX_Main-Bold.ttf"),
        ("mainBoldItalic", "KaTeX_Main-BoldItalic.ttf"),
        ("mathItalic", "KaTeX_Math-Italic.ttf"),
        ("mathBoldItalic", "KaTeX_Math-BoldItalic.ttf"),
        ("ams", "KaTeX_AMS-Regular.ttf"),
        ("calligraphic", "KaTeX_Caligraphic-Regular.ttf"),
        ("calligraphicBold", "KaTeX_Caligraphic-Bold.ttf"),
        ("script", "KaTeX_Script-Regular.ttf"),
        ("fraktur", "KaTeX_Fraktur-Regular.ttf"),
        ("frakturBold", "KaTeX_Fraktur-Bold.ttf"),
        ("sans", "KaTeX_SansSerif-Regular.ttf"),
        ("sansBold", "KaTeX_SansSerif-Bold.ttf"),
        ("sansItalic", "KaTeX_SansSerif-Italic.ttf"),
        ("monospace", "KaTeX_Typewriter-Regular.ttf"),
        ("size1", "KaTeX_Size1-Regular.ttf"),
        ("size2", "KaTeX_Size2-Regular.ttf"),
        ("size3", "KaTeX_Size3-Regular.ttf"),
        ("size4", "KaTeX_Size4-Regular.ttf"),
    ] {
        ids.insert(role.to_string(), load(&mut fonts, file));
    }
    layout.tex_font_set = ids.clone();
    layout.alphabet_upright_font = Some(ids["main"].clone());
    layout.alphabet_italic_font = Some(ids["mathItalic"].clone());
    layout.alphabet_bold_font = Some(ids["mainBold"].clone());
    layout.alphabet_bold_italic_font = Some(ids["mathBoldItalic"].clone());
    layout
        .math_fallbacks
        .push(load(&mut fonts, "STIXTwoMath-Regular.otf"));
    layout.text_font = Some(load(&mut fonts, "NotoSansCJKsc-Regular.otf"));
    (fonts, layout, ids)
}

fn assert_glyph(scene: &LayoutSnapshot, fonts: &FontRegistry, font_id: &str, ch: char) {
    let font = fonts.get(font_id).unwrap();
    let gid = font.glyph_id(ch).unwrap();
    let expected = font.glyph(gid).unwrap();
    let used = scene
        .resources
        .iter()
        .find(|resource| resource.font_id == font_id && resource.shape.glyph_id == gid)
        .unwrap_or_else(|| {
            panic!(
                "U+{:04X} is not painted by requested true font {font_id}",
                u32::from(ch)
            )
        });
    assert_eq!(
        serde_json::to_string(&used.shape).unwrap(),
        serde_json::to_string(&expected).unwrap(),
        "font path/advance was synthesized"
    );
}

#[test]
fn mathlive_katex_role_policy_is_reproduced_without_mathlive_runtime() {
    let (fonts, mut options, ids) = resources();
    for (source, role, ch) in [
        (r"x", "mathItalic", 'x'),
        (r"\alpha", "mathItalic", '\u{03B1}'),
        (r"\Gamma", "main", '\u{0393}'),
        (r"1+2=3", "main", '+'),
        (r"1+2=3", "main", '='),
        (r"1+2=3", "main", '2'),
        (r"\mathrm{d}", "main", 'd'),
        (r"\mathit{A}", "mathItalic", 'A'),
        (r"\mathbf{A}", "mainBold", 'A'),
        // MathLive keeps lowercase Greek in KaTeX_Math when \mathbf is applied.
        (r"\mathbf{\alpha}", "mathBoldItalic", '\u{03B1}'),
        (r"\mathbfit{x}", "mathBoldItalic", 'x'),
        (r"\mathbb{R}", "ams", 'R'),
        (r"\mathcal{A}", "calligraphic", 'A'),
        (r"\mathscr{A}", "script", 'A'),
        (r"\mathfrak{H}", "fraktur", 'H'),
        (r"\mathsf{A}", "sans", 'A'),
        (r"\mathtt{x}", "monospace", 'x'),
    ] {
        let session = EngineSession::from_source(source.into(), source.into()).unwrap();
        let scene = LayoutEngine::default()
            .layout(&session, &fonts, &options)
            .unwrap();
        assert_glyph(&scene, &fonts, &ids[role], ch);
        scene.svg().unwrap();
    }

    for (style, role) in [(MathStyle::Text, "size1"), (MathStyle::Display, "size2")] {
        options.style = style;
        let session = EngineSession::from_source(
            "operators".into(),
            r"\sum_{i=1}^{n}x_i+\int_0^1x\mathrm{d}x".into(),
        )
        .unwrap();
        let scene = LayoutEngine::default()
            .layout(&session, &fonts, &options)
            .unwrap();
        for ch in ['\u{2211}', '\u{222B}'] {
            assert_glyph(&scene, &fonts, &ids[role], ch);
        }
    }
}

#[test]
fn all_shipped_faces_fit_one_bounded_registry_and_use_no_machine_fonts() {
    let manifest: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root().join("manifest.json")).unwrap(),
    )
    .unwrap();
    let mut registry = FontRegistry::default();
    let mut identities = BTreeSet::new();
    let records = manifest["fonts"].as_array().unwrap();
    assert_eq!(records.len(), 37);
    for record in records {
        let bytes = std::fs::read(root().join(record["file"].as_str().unwrap())).unwrap();
        let identity = registry.register(&bytes, 0).unwrap();
        assert_eq!(identity.sha256, record["sha256"].as_str().unwrap());
        assert_eq!(bytes.len() as u64, record["bytes"].as_u64().unwrap());
        assert!(identities.insert(identity.id.clone()));
        assert_eq!(
            registry.register(&bytes, 0).unwrap().id,
            identity.id,
            "cached registration changed identity"
        );
    }
}

#[test]
fn font_switch_changes_only_layout_resources_and_katex_switchback_is_exact() {
    let (mut fonts, mut options, _) = resources();
    let mut renderer = LayoutEngine::default();
    let source = r"\frac{1}{\psi_{1,2}}\frac{\mathrm{d}^{2}}{\mathrm{d}x^{2}}\psi_{1,2}+\frac{2m(E-u)}{\hbar^{2}}=0";
    let session = EngineSession::from_source("reference".into(), source.into()).unwrap();
    let before_document = serde_json::to_string(&session.document_snapshot()).unwrap();
    let before_revision = session.revision();
    let original_options = options.clone();
    let original = renderer
        .layout(&session, &fonts, &options)
        .unwrap()
        .svg()
        .unwrap();

    for family in ["termes", "pagella", "heros"] {
        options.tex_font_set.clear();
        options.alphabet_upright_font = Some(load(
            &mut fonts,
            &format!("texgyre{family}-regular.otf"),
        ));
        options.alphabet_italic_font = Some(load(
            &mut fonts,
            &format!("texgyre{family}-italic.otf"),
        ));
        options.alphabet_bold_font = Some(load(
            &mut fonts,
            &format!("texgyre{family}-bold.otf"),
        ));
        options.alphabet_bold_italic_font = Some(load(
            &mut fonts,
            &format!("texgyre{family}-bolditalic.otf"),
        ));
        let scene = renderer.layout(&session, &fonts, &options).unwrap();
        assert_ne!(scene.svg().unwrap(), original);
        assert_glyph(
            &scene,
            &fonts,
            options.alphabet_italic_font.as_deref().unwrap(),
            '\u{03C8}',
        );
        assert_glyph(
            &scene,
            &fonts,
            options.alphabet_upright_font.as_deref().unwrap(),
            '=',
        );
        assert_eq!(session.source(), source);
        assert_eq!(session.revision(), before_revision);
        assert_eq!(
            serde_json::to_string(&session.document_snapshot()).unwrap(),
            before_document,
            "changing fonts mutated retained syntax/selection state"
        );
    }
    assert_eq!(
        renderer
            .layout(&session, &fonts, &original_options)
            .unwrap()
            .svg()
            .unwrap(),
        original
    );
}

#[test]
fn stretch_capability_falls_back_when_primary_math_font_has_the_character_but_not_the_construction() {
    let (fonts, mut options, _) = resources();
    let latin_math = options.math_font.clone();
    let stix_math = options.math_fallbacks[0].clone();
    options.math_font = stix_math;
    options.math_fallbacks = vec![latin_math];
    options.tex_font_set.clear();
    options.alphabet_upright_font = None;
    options.alphabet_italic_font = None;
    options.alphabet_bold_font = None;
    options.alphabet_bold_italic_font = None;
    for source in [
        r"\overleftarrow{AB}",
        r"\overrightarrow{AB}",
        r"\overleftrightarrow{AB}",
        r"\overleftharpoon{AB}",
        r"\overrightharpoon{AB}",
        r"\underleftarrow{AB}",
        r"\underrightarrow{AB}",
        r"\underleftrightarrow{AB}",
        r"\begin{vmatrix}a&b\\c&d\end{vmatrix}",
        r"\left\lvert\psi\middle\rangle\!\middle\langle\phi\right\rvert",
    ] {
        let session = EngineSession::from_source(format!("fallback:{source}"), source.into()).unwrap();
        let scene = LayoutEngine::default().layout(&session, &fonts, &options).unwrap_or_else(|error| {
            panic!("capability fallback failed for {source}: {error}")
        });
        scene.validate().unwrap();
        scene.svg().unwrap();
    }
}

#[test]
fn scripts_multiline_cjk_and_draft_slots_are_real_geometry_offline() {
    let (mut fonts, mut options, _) = resources();
    for text in [
        "NotoSansCJKsc-Regular.otf",
        "NotoSerifCJKsc-Regular.otf",
        "LXGWWenKai-Regular.ttf",
    ] {
        options.text_font = Some(load(&mut fonts, text));
        for source in [
            r"x^{a_2}+\sqrt{\frac{a}{b}}",
            r"\begin{aligned}a&=b+c\\x&=y\end{aligned}",
            r"\begin{pmatrix}a&b\\c&d\end{pmatrix}",
            r"\text{中文公式}+\alpha_2",
            r"\frac{}{}",
        ] {
            let session = EngineSession::from_source(
                format!("{text}:{source}"),
                source.into(),
            )
            .unwrap();
            let scene = LayoutEngine::default()
                .layout(&session, &fonts, &options)
                .unwrap();
            scene.validate().unwrap();
            if source.contains("中文") {
                assert_glyph(
                    &scene,
                    &fonts,
                    options.text_font.as_deref().unwrap(),
                    '\u{4E2D}',
                );
            }
            if source == r"\frac{}{}" {
                assert!(!scene.export_ready);
                assert_eq!(
                    scene
                        .regions
                        .iter()
                        .filter(|region| region.role == "emptyRow")
                        .count(),
                    2
                );
            } else {
                scene.svg().unwrap();
            }
        }
    }
}
