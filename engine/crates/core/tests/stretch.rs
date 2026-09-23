use std::path::PathBuf;
use visualtex_engine_core::{fonts::FontAsset, stretch, ErrorCode};
fn font() -> FontAsset {
    let path = std::env::var_os("VISUALTEX_TEST_MATH_FONT").map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf"));
    FontAsset::from_bytes(&std::fs::read(path).unwrap(), 0).unwrap()
}
#[test]
fn designed_surd_variants_and_assemblies_reach_requested_ink_height() {
    let f=font(); let id=f.glyph_id('√').unwrap();
    let mut saw_assembly=false;
    for target in [0.1,0.8,1.0,1.2,1.8,2.5,3.0,4.0,6.0,10.0,25.0] {
        let result=stretch::vertical(&f,id,target).unwrap_or_else(|e|panic!("target={target}: {e:?}"));
        assert!(result.height+1e-9>=target); assert!(result.width>0.0);
        assert!(!result.pieces.is_empty());
        assert!(result.overlaps.iter().all(|o| *o+1e-9>=result.minimum_overlap));
        let mut top=f64::INFINITY; let mut bottom=f64::NEG_INFINITY;
        for piece in &result.pieces {
            let actual=f.glyph(piece.shape.glyph_id).unwrap();
            assert_eq!(serde_json::to_value(&piece.shape).unwrap(),serde_json::to_value(actual).unwrap());
            let ink=piece.shape.ink.unwrap(); top=top.min(piece.y+ink[1]);bottom=bottom.max(piece.y+ink[3]);
        }
        assert!(top.abs()<1e-9); assert!((bottom-result.height).abs()<1e-9);
        saw_assembly|=result.assembled;
    }
    assert!(saw_assembly,"Large roots must use real font assemblies, not scale a small glyph");
}
#[test]
fn invalid_targets_and_missing_constructions_are_explicit_errors() {
    let f=font(); let root=f.glyph_id('√').unwrap();
    for bad in [0.0,-1.0,f64::NAN,f64::INFINITY] { assert_eq!(stretch::vertical(&f,root,bad).unwrap_err().code,ErrorCode::InvalidInput); }
    assert_eq!(stretch::vertical(&f,root,2049.0).unwrap_err().code,ErrorCode::ResourceLimit);
    assert!(stretch::vertical(&f,f.glyph_id('x').unwrap(),30.0).is_err());
}
#[test]
fn horizontal_arrows_and_braces_are_real_outline_assemblies() {
    let f=font();
    for ch in ['→','←','↔','⏞','⏟'] {
        for target in [0.3,1.0,2.0,5.0,20.0] {
            let result=stretch::horizontal(&f,f.glyph_id(ch).unwrap(),target).unwrap_or_else(|e|panic!("{ch} at {target}: {e:?}"));
            assert!(result.width+1e-9>=target);assert!(result.ink_top<=result.ink_bottom);
            assert!(result.overlaps.iter().all(|v|*v+1e-9>=result.minimum_overlap));
            for piece in &result.pieces {assert_eq!(serde_json::to_value(&piece.shape).unwrap(),serde_json::to_value(f.glyph(piece.shape.glyph_id).unwrap()).unwrap());}
            if target==20.0 {assert!(result.assembled);}
        }
    }
}
#[test]
fn combining_accent_attachment_and_wide_variants_are_available() {
    let f=font();
    for ch in ['\u{302}','\u{303}','\u{20d7}','\u{307}','\u{308}'] {
        let id=f.glyph_id(ch).unwrap();assert!(f.accent_attachment(id).unwrap().is_finite());
        assert!(f.glyph(id).unwrap().ink.is_some());
    }
    for ch in ['\u{302}','\u{303}'] {
        for target in [0.5,1.0,2.0,4.0] {
            let id=f.glyph_id(ch).unwrap();
            let r=stretch::horizontal_accent(&f,id,target).unwrap();
            assert_eq!(r.target_met,r.width+1e-9>=target);
            if !r.target_met {
                // The actual fixed font ends near 1.9 em. Preserve the strict
                // failing assertion; only the explicit accent policy may cap.
                assert!(stretch::horizontal(&f,id,target).is_err());
                assert_eq!(r.pieces.len(),1);assert!(!r.assembled);
                assert_eq!(serde_json::to_value(&r.pieces[0].shape).unwrap(),serde_json::to_value(f.glyph(r.pieces[0].shape.glyph_id).unwrap()).unwrap());
            }
        }
    }
}
#[test]
fn radical_parameters_are_read_from_math_table() {
    let f=font(); let m=f.math.as_ref().unwrap();
    assert!(m.radical_rule>0.0);assert!(m.radical_display_gap>=m.radical_gap);
    assert!(m.radical_kern_before_degree>=0.0);assert!(m.radical_kern_after_degree<=0.0);
    assert!(m.radical_degree_raise>0.0&&m.radical_degree_raise<1.0);
}
