//! Word's paragraph typography aligns the picture box, not the TeX baseline.
//! Keep the TeX origin and measure the text/letter ink separately: scripts must
//! never move the main letter, and tall operators must share the text midline.
use crate::system_math_glyphs::extract_stable_macos_svg_glyph;
use serde::Deserialize;

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct MetricsRequest {
    body_font: String,
    body_sample: String,
    metadata: String,
}

fn scripted_letter(source: &str) -> Option<char> {
    let mut source = source.trim();
    loop {
        let previous = source;
        for command in ["\\displaystyle", "\\textstyle", "\\scriptstyle", "\\scriptscriptstyle"] {
            if let Some(rest) = source.strip_prefix(command) { source = rest.trim(); }
        }
        if source.starts_with('{') && source.ends_with('}') {
            let mut depth = 0;
            let mut encloses_all = true;
            for (index, character) in source.char_indices() {
                match character { '{' => depth += 1, '}' => depth -= 1, _ => {} }
                if depth == 0 && index + 1 != source.len() { encloses_all = false; break; }
            }
            if encloses_all && depth == 0 { source = source[1..source.len()-1].trim(); }
        }
        if source == previous { break; }
    }
    let mut chars = source.chars();
    let letter = chars.next()?;
    if !letter.is_ascii_alphabetic() { return None; }
    let tail: Vec<char> = chars.filter(|c| !c.is_whitespace()).collect();
    let mut index = 0;
    while index < tail.len() {
        if !matches!(tail[index], '^' | '_') { return None; }
        index += 1;
        if index == tail.len() { return None; }
        if tail[index] == '{' {
            let mut depth = 1;
            index += 1;
            while index < tail.len() && depth > 0 {
                match tail[index] { '{' => depth += 1, '}' => depth -= 1, _ => {} }
                index += 1;
            }
            if depth != 0 { return None; }
        } else {
            if matches!(tail[index], '^' | '_' | '}' | '\\') { return None; }
            index += 1;
        }
    }
    Some(letter)
}

// Painted centres from the MathJax TeX italic glyph metrics used by runtime.ts
// (mathjax-full/output/common/fonts/tex/normal, Apache-2.0). Computer Modern
// letters are paths, so querying an installed system font would be incorrect.
fn tex_letter_center(letter: char) -> f64 {
    const CENTRES: [f64; 52] = [
        0.3580, 0.3415, 0.3415, 0.3415, 0.3400, 0.3400, 0.3415, 0.3415, 0.3415, 0.3305, 0.3415, 0.3415, 0.3415,
        0.3415, 0.3410, 0.3415, 0.2550, 0.3310, 0.3415, 0.3385, 0.3305, 0.3305, 0.3305, 0.3415, 0.3415, 0.3415,
        0.2155, 0.3415, 0.2155, 0.3420, 0.2155, 0.2500, 0.1185, 0.3415, 0.3250, 0.2285, 0.3415, 0.3415, 0.2155,
        0.2155, 0.2150, 0.1240, 0.1240, 0.2155, 0.2160, 0.3075, 0.2155, 0.2160, 0.2160, 0.2155, 0.1185, 0.2155,
    ];
    let index = if letter.is_ascii_uppercase() { letter as usize - 'A' as usize }
        else { 26 + letter as usize - 'a' as usize };
    CENTRES[index]
}

#[cfg(target_os = "macos")]
fn register_word_fonts(body_font: &str) {
    use std::{ffi::c_void, ptr};
    #[link(name = "CoreFoundation", kind = "framework")]
    unsafe extern "C" {
        fn CFURLCreateFromFileSystemRepresentation(allocator: *const c_void, bytes: *const u8, length: isize, directory: u8) -> *const c_void;
        fn CFRelease(value: *const c_void);
    }
    #[link(name = "CoreText", kind = "framework")]
    unsafe extern "C" {
        fn CTFontManagerRegisterFontsForURL(url: *const c_void, scope: u32, error: *mut *const c_void) -> u8;
    }
    // Match the actual Word face, including private Chinese theme fonts that
    // CoreText cannot see in a separate CLI process. Process scope only.
    let body_files: &[&str] = match body_font {
        "DengXian" => &["Deng.ttf"],
        "DengXian Light" => &["Dengl.ttf"],
        "SimSun" | "NSimSun" => &["Simsun.ttc"],
        "SimHei" => &["SimHei.ttf"],
        "FangSong" => &["Fangsong.ttf"],
        _ => &[],
    };
    for name in std::iter::once(&"Cambria.ttc").chain(body_files.iter()) {
        let path = format!("/Applications/Microsoft Word.app/Contents/Resources/DFonts/{name}");
        if !std::path::Path::new(&path).is_file() { continue; }
        unsafe {
            let url = CFURLCreateFromFileSystemRepresentation(ptr::null(), path.as_ptr(), path.len() as isize, 0);
            if !url.is_null() {
                CTFontManagerRegisterFontsForURL(url, 1, ptr::null_mut());
                CFRelease(url);
            }
        }
    }
}

fn word_body_font(value: &str) -> &str {
    match value.trim() {
        "等线" => "DengXian",
        "等线 Light" | "等线 细体" => "DengXian Light",
        "宋体" => "SimSun",
        "新宋体" => "NSimSun",
        "黑体" => "SimHei",
        "仿宋" => "FangSong",
        other => other,
    }
}

fn measure(request: MetricsRequest) -> Result<(f64, f64, f64, f64, f64, bool), String> {
    if request.body_sample.is_empty() || request.body_sample.chars().count() > 16 {
        return Err("Word text sample must contain 1 to 16 characters.".into());
    }
    let body_font = word_body_font(&request.body_font);
    if body_font.starts_with('+') {
        return Err("Word theme font must be resolved to its effective face.".into());
    }
    #[cfg(target_os = "macos")]
    register_word_fonts(body_font);
    let mut top = f64::NEG_INFINITY;
    let mut bottom = f64::INFINITY;
    let mut ascent = 0.0_f64;
    let mut descent = 0.0_f64;
    for character in request.body_sample.chars() {
        let glyph = extract_stable_macos_svg_glyph(&[body_font.to_owned()], &character.to_string(), false, false)?;
        top = top.max(glyph.ink_top_em);
        bottom = bottom.min(glyph.ink_bottom_em);
        ascent = ascent.max(glyph.font_ascent_em);
        descent = descent.max(glyph.font_descent_em);
    }
    let metadata = super::macos_offline::decode_metadata(&request.metadata)?;
    let primary_letter = scripted_letter(&metadata.latex);
    let mut anchor_em = 0.25; // MathJax TeX math axis: equals/fraction/sum centre.
    if let Some(letter) = primary_letter {
        let font = match metadata.formula_letter_font.as_deref() {
            Some("times") => "Times New Roman",
            Some("cambria") => "Cambria Math",
            Some("palatino") => "Palatino",
            Some("helvetica") => "Helvetica Neue",
            _ => "STIX Two Math",
        };
        if metadata.formula_letter_font.as_deref().unwrap_or("katex") == "katex" {
            anchor_em = tex_letter_center(letter);
        } else if let Ok(glyph) = extract_stable_macos_svg_glyph(&[font.into()], &letter.to_string(), true, false) {
            anchor_em = (glyph.ink_top_em + glyph.ink_bottom_em) / 2.0;
        }
    }
    let render_height = metadata.render_height_px.ok_or("Word image metadata has no rendered height.")?;
    // Express the anchor relative to the unrounded SVG height; Word's actual
    // quantized Height then scales both the TeX descent and anchor consistently.
    Ok((top, bottom, anchor_em * (14.0 * 96.0 / 72.0) / render_height, ascent, descent, primary_letter.is_some()))
}

pub fn run_cli_if_requested() -> Option<i32> {
    let arguments: Vec<String> = std::env::args().collect();
    let index = arguments.iter().position(|value| value == "--word-image-alignment-metrics")?;
    let result = arguments.get(index + 1).ok_or_else(|| "Missing Word alignment request.".to_string())
        .and_then(|value| {
            if value.len() > 2_000_000 { return Err("Word alignment request is too large.".into()); }
            serde_json::from_str(value).map_err(|error| error.to_string())
        }).and_then(measure);
    match result {
        Ok((top, bottom, anchor, ascent, descent, primary_letter)) => { println!("{top:.9}|{bottom:.9}|{anchor:.9}|{ascent:.9}|{descent:.9}|{}", u8::from(primary_letter)); Some(0) }
        Err(error) => { eprintln!("{error}"); Some(2) }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn scripts_do_not_change_the_letter_anchor() {
        for source in ["L", "L_z", "L^2", "L_{ij}^{n+1}", "x^{\\frac{a}{b}}"] {
            assert_eq!(scripted_letter(source), source.chars().next());
        }
        assert_eq!(scripted_letter("{\\displaystyle L_z}"), Some('L'));
        assert_eq!(tex_letter_center('L'), 0.3415);
        assert_eq!(tex_letter_center('x'), 0.2155);
        for source in ["L+L", "L=2", "L_", "L^{2", "\\sum_i x_i", "(a+b)^n"] {
            assert_eq!(scripted_letter(source), None);
        }
    }

    #[test]
    #[ignore = "real Word fixture preparation; requires VISUALTEX_ALIGNMENT_SVG_DIRECTORY"]
    fn freeze_alignment_fixture_vectors() {
        let directory = std::env::var("VISUALTEX_ALIGNMENT_SVG_DIRECTORY").unwrap();
        let mut count = 0;
        for entry in std::fs::read_dir(directory).unwrap() {
            let path = entry.unwrap().path();
            if path.extension().and_then(|value| value.to_str()) != Some("svg") { continue; }
            let svg = std::fs::read(&path).unwrap();
            let frozen = crate::svg_font_stabilizer::stabilize_word_svg_font_outlines(&svg).unwrap();
            std::fs::write(path, frozen).unwrap();
            count += 1;
        }
        assert_eq!(count, 45);
    }
}
