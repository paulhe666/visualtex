#![cfg(target_os = "windows")]

//! Whether Windows in this user session (where Word and PowerPoint run) draws
//! MathType's pictures with MathType's own fonts.
//!
//! MathType's pictures name MT Extra and the Euclid fonts and draw 8-bit codes
//! from them. Office installs a 1992 MT Extra under the same name that lacks
//! most of those glyphs, and PCs without MathType have no Euclid fonts, so the
//! check compares what Windows resolves with the fonts bundled with the private
//! renderer. The Repair action of the Office integration installs them.

use std::ffi::c_void;
use std::path::Path;

/// Families MathType's default styles draw from. Fences and the Tiger fonts
/// serve optional MathType preference sets and are not checked.
fn is_required(family: &str) -> bool {
    family.eq_ignore_ascii_case("MT Extra") || family.starts_with("Euclid")
}

pub fn check(install_root: &Path) -> Result<(), String> {
    let directory = install_root.join("runtime").join("Fonts").join("TrueType");
    let entries = std::fs::read_dir(&directory)
        .map_err(|error| format!("MathType fonts are missing from the renderer ({error})"))?;
    let mut problems = Vec::new();
    let mut checked = 0;
    for entry in entries.flatten() {
        let path = entry.path();
        if !path
            .extension()
            .is_some_and(|extension| extension.eq_ignore_ascii_case("ttf"))
        {
            continue;
        }
        let Ok(data) = std::fs::read(&path) else { continue };
        let Ok(face) = ttf_parser::Face::parse(&data, 0) else { continue };
        let Some(family) = family_name(&face) else { continue };
        if !is_required(&family) {
            continue;
        }
        checked += 1;
        let codes = drawn_codes(&face);
        let style = match (face.is_bold(), face.is_italic()) {
            (true, true) => " Bold Italic",
            (true, false) => " Bold",
            (false, true) => " Italic",
            (false, false) => "",
        };
        match missing_codes(&family, face.weight().to_number(), face.is_italic(), &codes) {
            None => problems.push(format!("{family}{style} is not installed")),
            Some(0) => {}
            Some(missing) => problems.push(format!(
                "{family}{style} lacks {missing} of MathType's {} symbols",
                codes.len()
            )),
        }
    }
    if checked == 0 {
        return Err("The renderer carries no MathType fonts".to_string());
    }
    if problems.is_empty() {
        Ok(())
    } else {
        Err(problems.join("; "))
    }
}

fn family_name(face: &ttf_parser::Face) -> Option<String> {
    let names = face.names();
    let mut fallback = None;
    for index in 0..names.len() {
        let Some(name) = names.get(index) else { continue };
        if name.name_id != ttf_parser::name_id::FAMILY {
            continue;
        }
        if let Some(value) = name.to_string() {
            if name.platform_id == ttf_parser::PlatformId::Windows {
                return Some(value.trim().to_string());
            }
            fallback.get_or_insert(value.trim().to_string());
        }
    }
    fallback
}

/// The codes MathType's pictures can draw from this font: for a symbol font,
/// the 8-bit codes Windows maps to U+F020-U+F0FF; otherwise its BMP cmap.
fn drawn_codes(face: &ttf_parser::Face) -> Vec<u16> {
    let mut symbol = Vec::new();
    let mut other = Vec::new();
    if let Some(cmap) = face.tables().cmap {
        for subtable in cmap.subtables {
            // Windows draws through the Windows (and Unicode) cmaps only; a
            // Macintosh cmap's codes are Mac Roman bytes, not Unicode.
            let windows_symbol = subtable.platform_id == ttf_parser::PlatformId::Windows
                && subtable.encoding_id == 0;
            if !windows_symbol && !subtable.is_unicode() {
                continue;
            }
            subtable.codepoints(|code| {
                if (0xF020..=0xF0FF).contains(&code) {
                    symbol.push(code as u16);
                } else if code <= 0xFFFF {
                    other.push(code as u16);
                }
            });
        }
    }
    let mut codes = if symbol.is_empty() { other } else { symbol };
    codes.sort_unstable();
    codes.dedup();
    codes
}

#[link(name = "gdi32")]
extern "system" {
    fn CreateCompatibleDC(dc: *mut c_void) -> *mut c_void;
    fn DeleteDC(dc: *mut c_void) -> i32;
    fn CreateFontW(
        height: i32,
        width: i32,
        escapement: i32,
        orientation: i32,
        weight: i32,
        italic: u32,
        underline: u32,
        strike_out: u32,
        char_set: u32,
        out_precision: u32,
        clip_precision: u32,
        quality: u32,
        pitch_and_family: u32,
        face: *const u16,
    ) -> *mut c_void;
    fn SelectObject(dc: *mut c_void, object: *mut c_void) -> *mut c_void;
    fn DeleteObject(object: *mut c_void) -> i32;
    fn GetTextFaceW(dc: *mut c_void, count: i32, face: *mut u16) -> i32;
    fn GetGlyphIndicesW(
        dc: *mut c_void,
        text: *const u16,
        count: i32,
        indices: *mut u16,
        flags: u32,
    ) -> u32;
}

/// `None` when Windows resolves the family to a different font altogether;
/// otherwise how many of `codes` have no glyph in the font it picks.
fn missing_codes(family: &str, weight: u16, italic: bool, codes: &[u16]) -> Option<usize> {
    const DEFAULT_CHARSET: u32 = 1; // as MathType's pictures request the font
    const OUT_TT_ONLY_PRECIS: u32 = 7;
    const GGI_MARK_NONEXISTING_GLYPHS: u32 = 1;
    let face: Vec<u16> = family.encode_utf16().chain(std::iter::once(0)).collect();
    unsafe {
        let dc = CreateCompatibleDC(std::ptr::null_mut());
        if dc.is_null() {
            return None;
        }
        let font = CreateFontW(
            -1000,
            0,
            0,
            0,
            i32::from(weight),
            u32::from(italic),
            0,
            0,
            DEFAULT_CHARSET,
            OUT_TT_ONLY_PRECIS,
            0,
            0,
            0,
            face.as_ptr(),
        );
        if font.is_null() {
            DeleteDC(dc);
            return None;
        }
        let old = SelectObject(dc, font);
        let mut actual = [0u16; 64];
        let length = GetTextFaceW(dc, actual.len() as i32, actual.as_mut_ptr());
        let actual = String::from_utf16_lossy(&actual[..(length.max(1) as usize - 1).min(actual.len())]);
        let result = if !actual.trim_end_matches('\0').eq_ignore_ascii_case(family) {
            None
        } else if codes.is_empty() {
            Some(0)
        } else {
            let mut indices = vec![0u16; codes.len()];
            let status = GetGlyphIndicesW(
                dc,
                codes.as_ptr(),
                codes.len() as i32,
                indices.as_mut_ptr(),
                GGI_MARK_NONEXISTING_GLYPHS,
            );
            if status == u32::MAX {
                None
            } else {
                Some(indices.iter().filter(|index| **index == 0xFFFF).count())
            }
        };
        SelectObject(dc, old);
        DeleteObject(font);
        DeleteDC(dc);
        result
    }
}
