//! Behaviour the Windows corpus does not reach (it has no extended integrals
//! and no document-level whitespace), derived from WordOmmlConverter.cs.

use crate::office::omml_write::proj as port;

const MML: &str = "http://www.w3.org/1998/Math/MathML";

#[test]
fn extended_integrals_are_restored_after_the_projection() {
    let s11 = format!(
        "<math xmlns=\"{MML}\" display=\"block\"><msub><mo>\u{2230}</mo><mi>V</mi></msub><mrow><mi>f</mi></mrow></math>"
    );
    let (prepared, nary) = port::replace_extended_integrals_with_office_placeholders(&s11).unwrap();
    assert_eq!(nary, vec!["\u{2230}".to_string()]);
    assert_eq!(prepared, s11.replace('\u{2230}', "\u{222B}"));
    let projection = port::project(&prepared).unwrap();
    assert!(projection.contains("<m:chr m:val=\"\u{222B}\" />"));
    let restored = port::restore_extended_integral_characters(&port::extract_single_omath(&projection).unwrap(), &nary)
        .unwrap();
    assert_eq!(restored, projection.replace("<m:chr m:val=\"\u{222B}\" />", "<m:chr m:val=\"\u{2230}\" />"));
    let omml = port::transform_prepared_mathml_to_omml(&s11).unwrap();
    assert!(omml.contains("<m:chr m:val=\"\u{2230}\" />"));
    // A changed n-ary count is an error, as in the C#.
    let error = port::restore_extended_integral_characters(&projection, &[nary[0].clone(), nary[0].clone()]);
    assert!(error.unwrap_err().starts_with("InvalidDataException: Office changed the number of n-ary operators"));
}

#[test]
fn document_whitespace_is_kept_by_the_placeholder_pass() {
    let s11 = format!("\n<math xmlns=\"{MML}\"><mo>\u{222B}</mo></math>\r\n");
    let (prepared, nary) = port::replace_extended_integrals_with_office_placeholders(&s11).unwrap();
    assert_eq!(prepared, format!("\r\n<math xmlns=\"{MML}\"><mo>\u{222B}</mo></math>\r\n"));
    assert_eq!(nary, vec!["\u{222B}".to_string()]);
    assert!(!port::is_block_mathml(&prepared));
}

#[test]
fn numbered_host_with_bookmark_and_run_fonts() {
    let w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    let omml = format!(
        "<m:oMath xmlns:m=\"{m}\" xmlns:w=\"{w}\"><m:r><w:rPr><w:rFonts w:ascii=\"Cambria Math\" /><w:sz w:val=\"24\" /></w:rPr><m:t>x</m:t></m:r></m:oMath>"
    );
    let bookmark = "VTEqNum_0123456789abcdef0123456789ABCDEF";
    let numbered =
        port::build_word_native_numbered_omml(&omml, "Equation", Some(bookmark), 0, "", "1", None, None).unwrap();
    // _nsTop is 8 at bookmarkStart (oMath, eqArr, e, d, e and its own default
    // namespace), so its attributes share the generated prefix p6.
    assert!(numbered.contains(&format!(
        "<bookmarkStart p6:id=\"31901\" p6:name=\"{bookmark}\" xmlns:p6=\"{w}\" xmlns=\"{w}\" />"
    )));
    assert!(numbered.contains(&format!(
        "<eqArrPr><maxDist p4:val=\"1\" xmlns:p4=\"{m}\" /><ctrlPr><rPr xmlns=\"{w}\">\
         <rFonts p6:ascii=\"Cambria Math\" xmlns:p6=\"{w}\" /><sz p6:val=\"24\" xmlns:p6=\"{w}\" /><i /></rPr></ctrlPr></eqArrPr>"
    )));
    for invalid in ["VTEqNum_0123", "VTEqNum_0123456789abcdef0123456789abcdeg", "VTEq_0123456789abcdef0123456789abcdef0"] {
        assert!(port::build_word_native_numbered_omml(&omml, "Equation", Some(invalid), 0, "", "1", None, None).is_err());
    }
    assert!(port::build_word_native_numbered_omml(&omml, "Eq\\1", None, 0, "", "1", None, None).is_err());
    assert!(port::build_word_native_numbered_omml(&omml, "Eq", None, 0, "", "1", Some("2"), None).is_err());
    assert!(port::build_word_native_numbered_omml(&omml, "Eq", None, 1, "", "1", None, Some(0)).is_err());
}
