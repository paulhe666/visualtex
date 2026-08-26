use serde::Serialize;

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SystemMathFontProbe {
    requested_family: String,
    resolved_family: String,
    available: bool,
}

#[derive(Debug, Clone, Copy, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SystemMathGlyphMetrics {
    width_em: f64,
    ascent_em: f64,
    descent_em: f64,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SystemMathGlyphOutline {
    character: String,
    requested_family: String,
    resolved_family: String,
    fallback_used: bool,
    glyph_id: u16,
    path: String,
    metrics: SystemMathGlyphMetrics,
}

const DEFAULT_MATH_FONT_FALLBACKS: &[&str] = &[
    "STIX Two Math",
    "Apple Symbols",
    "Times New Roman",
    "Helvetica",
];

fn normalize_font_family(value: &str) -> String {
    value
        .chars()
        .filter(|character| !character.is_whitespace() && *character != '-' && *character != '_')
        .flat_map(char::to_lowercase)
        .collect()
}

fn validate_font_family(value: &str) -> Result<String, String> {
    let trimmed = value.trim();
    if trimmed.is_empty()
        || trimmed.chars().count() > 128
        || trimmed
            .chars()
            .any(|character| character.is_control() || matches!(character, '"' | '\'' | '<' | '>'))
    {
        return Err("The requested system math font family is invalid.".to_string());
    }
    Ok(trimmed.to_string())
}

fn validate_character(value: &str) -> Result<String, String> {
    let normalized = value.trim();
    let scalar_count = normalized.chars().count();
    if scalar_count != 1
        || normalized.chars().any(char::is_control)
        || normalized.encode_utf16().count() > 2
    {
        return Err("The requested system math glyph is invalid.".to_string());
    }
    Ok(normalized.to_string())
}

#[cfg(target_os = "macos")]
mod macos {
    use super::*;
    use std::ffi::c_void;
    use std::ptr;

    type CFIndex = isize;
    type CFStringRef = *const c_void;
    type CTFontRef = *const c_void;
    type CGPathRef = *const c_void;
    type CGGlyph = u16;
    type UniChar = u16;
    type Boolean = u8;
    type CGFloat = f64;

    const K_CF_STRING_ENCODING_UTF8: u32 = 0x0800_0100;
    const K_CT_FONT_ORIENTATION_HORIZONTAL: u32 = 0;

    #[repr(C)]
    #[derive(Debug, Clone, Copy, Default)]
    struct CGPoint {
        x: CGFloat,
        y: CGFloat,
    }

    #[repr(C)]
    #[derive(Debug, Clone, Copy, Default)]
    struct CGSize {
        width: CGFloat,
        height: CGFloat,
    }

    #[repr(C)]
    #[derive(Debug, Clone, Copy, Default)]
    struct CGRect {
        origin: CGPoint,
        size: CGSize,
    }

    #[repr(C)]
    #[derive(Debug, Clone, Copy)]
    struct CGPathElement {
        element_type: i32,
        points: *mut CGPoint,
    }

    type CGPathApplierFunction = unsafe extern "C" fn(*mut c_void, *const CGPathElement);

    #[link(name = "CoreFoundation", kind = "framework")]
    unsafe extern "C" {
        fn CFStringCreateWithBytes(
            allocator: *const c_void,
            bytes: *const u8,
            byte_count: CFIndex,
            encoding: u32,
            is_external_representation: Boolean,
        ) -> CFStringRef;
        fn CFStringGetLength(value: CFStringRef) -> CFIndex;
        fn CFStringGetMaximumSizeForEncoding(length: CFIndex, encoding: u32) -> CFIndex;
        fn CFStringGetCString(
            value: CFStringRef,
            buffer: *mut i8,
            buffer_size: CFIndex,
            encoding: u32,
        ) -> Boolean;
        fn CFRelease(value: *const c_void);
    }

    #[link(name = "CoreText", kind = "framework")]
    unsafe extern "C" {
        fn CTFontCreateWithName(
            name: CFStringRef,
            size: CGFloat,
            matrix: *const c_void,
        ) -> CTFontRef;
        fn CTFontCopyFamilyName(font: CTFontRef) -> CFStringRef;
        fn CTFontGetGlyphsForCharacters(
            font: CTFontRef,
            characters: *const UniChar,
            glyphs: *mut CGGlyph,
            count: CFIndex,
        ) -> Boolean;
        fn CTFontCreatePathForGlyph(
            font: CTFontRef,
            glyph: CGGlyph,
            transform: *const c_void,
        ) -> CGPathRef;
        fn CTFontGetAdvancesForGlyphs(
            font: CTFontRef,
            orientation: u32,
            glyphs: *const CGGlyph,
            advances: *mut CGSize,
            count: CFIndex,
        ) -> CGFloat;
    }

    #[link(name = "CoreGraphics", kind = "framework")]
    unsafe extern "C" {
        fn CGPathApply(
            path: CGPathRef,
            info: *mut c_void,
            function: Option<CGPathApplierFunction>,
        );
        fn CGPathGetPathBoundingBox(path: CGPathRef) -> CGRect;
        fn CGPathRelease(path: CGPathRef);
    }

    struct OwnedCf(*const c_void);

    impl Drop for OwnedCf {
        fn drop(&mut self) {
            if !self.0.is_null() {
                unsafe { CFRelease(self.0) };
            }
        }
    }

    struct OwnedPath(CGPathRef);

    impl Drop for OwnedPath {
        fn drop(&mut self) {
            if !self.0.is_null() {
                unsafe { CGPathRelease(self.0) };
            }
        }
    }

    fn create_cf_string(value: &str) -> Result<OwnedCf, String> {
        let bytes = value.as_bytes();
        let reference = unsafe {
            CFStringCreateWithBytes(
                ptr::null(),
                bytes.as_ptr(),
                bytes.len() as CFIndex,
                K_CF_STRING_ENCODING_UTF8,
                0,
            )
        };
        if reference.is_null() {
            return Err("CoreText could not create the requested font name.".to_string());
        }
        Ok(OwnedCf(reference))
    }

    fn cf_string_to_string(value: CFStringRef) -> Result<String, String> {
        if value.is_null() {
            return Err("CoreText returned an empty font family.".to_string());
        }
        let length = unsafe { CFStringGetLength(value) };
        let maximum = unsafe {
            CFStringGetMaximumSizeForEncoding(length, K_CF_STRING_ENCODING_UTF8)
        };
        if maximum < 0 || maximum > 32_768 {
            return Err("CoreText returned an invalid font family length.".to_string());
        }
        let mut buffer = vec![0_i8; maximum as usize + 1];
        let copied = unsafe {
            CFStringGetCString(
                value,
                buffer.as_mut_ptr(),
                buffer.len() as CFIndex,
                K_CF_STRING_ENCODING_UTF8,
            )
        };
        if copied == 0 {
            return Err("CoreText could not decode the resolved font family.".to_string());
        }
        let bytes: Vec<u8> = buffer
            .into_iter()
            .take_while(|byte| *byte != 0)
            .map(|byte| byte as u8)
            .collect();
        String::from_utf8(bytes).map_err(|error| error.to_string())
    }

    struct ResolvedFont {
        reference: OwnedCf,
        requested_family: String,
        resolved_family: String,
        exact_family: bool,
    }

    fn resolve_font(requested: &str) -> Result<ResolvedFont, String> {
        let name = create_cf_string(requested)?;
        let font = unsafe { CTFontCreateWithName(name.0, 1000.0, ptr::null()) };
        if font.is_null() {
            return Err(format!("CoreText could not open {requested}."));
        }
        let reference = OwnedCf(font);
        let family_ref = unsafe { CTFontCopyFamilyName(font) };
        let family_guard = OwnedCf(family_ref);
        let resolved_family = cf_string_to_string(family_guard.0)?;
        let exact_family = normalize_font_family(requested)
            == normalize_font_family(&resolved_family);
        Ok(ResolvedFont {
            reference,
            requested_family: requested.to_string(),
            resolved_family,
            exact_family,
        })
    }

    fn mapped_glyph(font: CTFontRef, character: &str) -> Option<CGGlyph> {
        let characters: Vec<UniChar> = character.encode_utf16().collect();
        let mut glyphs = vec![0; characters.len()];
        unsafe {
            CTFontGetGlyphsForCharacters(
                font,
                characters.as_ptr(),
                glyphs.as_mut_ptr(),
                characters.len() as CFIndex,
            );
        }
        glyphs.into_iter().find(|glyph| *glyph != 0)
    }

    fn format_number(value: f64) -> String {
        let normalized = if value.abs() < 0.000_000_5 { 0.0 } else { value };
        let mut output = format!("{normalized:.5}");
        while output.contains('.') && output.ends_with('0') {
            output.pop();
        }
        if output.ends_with('.') {
            output.pop();
        }
        output
    }

    struct PathWriter {
        shift_x: f64,
        ascent: f64,
        output: String,
    }

    impl PathWriter {
        fn point(&self, point: CGPoint) -> (String, String) {
            (
                format_number(point.x + self.shift_x),
                format_number(self.ascent - point.y),
            )
        }
    }

    unsafe extern "C" fn append_path_element(
        info: *mut c_void,
        element: *const CGPathElement,
    ) {
        if info.is_null() || element.is_null() {
            return;
        }
        let writer = unsafe { &mut *(info as *mut PathWriter) };
        let element = unsafe { &*element };
        let points = element.points;
        match element.element_type {
            0 => {
                let point = unsafe { *points };
                let (x, y) = writer.point(point);
                writer.output.push_str(&format!("M{x} {y}"));
            }
            1 => {
                let point = unsafe { *points };
                let (x, y) = writer.point(point);
                writer.output.push_str(&format!("L{x} {y}"));
            }
            2 => {
                let control = unsafe { *points };
                let destination = unsafe { *points.add(1) };
                let (control_x, control_y) = writer.point(control);
                let (destination_x, destination_y) = writer.point(destination);
                writer.output.push_str(&format!(
                    "Q{control_x} {control_y} {destination_x} {destination_y}"
                ));
            }
            3 => {
                let control_one = unsafe { *points };
                let control_two = unsafe { *points.add(1) };
                let destination = unsafe { *points.add(2) };
                let (control_one_x, control_one_y) = writer.point(control_one);
                let (control_two_x, control_two_y) = writer.point(control_two);
                let (destination_x, destination_y) = writer.point(destination);
                writer.output.push_str(&format!(
                    "C{control_one_x} {control_one_y} {control_two_x} {control_two_y} {destination_x} {destination_y}"
                ));
            }
            4 => writer.output.push('Z'),
            _ => {}
        }
    }

    fn outline_from_font(
        font: ResolvedFont,
        character: &str,
        primary_requested_family: &str,
    ) -> Result<SystemMathGlyphOutline, String> {
        let glyph = mapped_glyph(font.reference.0, character).ok_or_else(|| {
            format!(
                "{} does not contain the requested mathematical glyph {}.",
                font.resolved_family, character
            )
        })?;
        let path_reference = unsafe {
            CTFontCreatePathForGlyph(font.reference.0, glyph, ptr::null())
        };
        if path_reference.is_null() {
            return Err(format!(
                "{} could not expose a vector outline for {}.",
                font.resolved_family, character
            ));
        }
        let path = OwnedPath(path_reference);
        let bounds = unsafe { CGPathGetPathBoundingBox(path.0) };
        let min_x = bounds.origin.x;
        let min_y = bounds.origin.y;
        let max_x = bounds.origin.x + bounds.size.width;
        let max_y = bounds.origin.y + bounds.size.height;
        if ![min_x, min_y, max_x, max_y]
            .into_iter()
            .all(f64::is_finite)
            || bounds.size.width <= 0.0
            || bounds.size.height <= 0.0
        {
            return Err(format!(
                "{} returned an empty vector outline for {}.",
                font.resolved_family, character
            ));
        }

        let mut advance = CGSize::default();
        unsafe {
            CTFontGetAdvancesForGlyphs(
                font.reference.0,
                K_CT_FONT_ORIENTATION_HORIZONTAL,
                &glyph,
                &mut advance,
                1,
            );
        }
        let shift_x = -min_x.min(0.0);
        let ascent = max_y.max(20.0);
        let descent = (-min_y).max(0.0);
        let width = (advance.width + shift_x)
            .max(max_x + shift_x)
            .max(20.0);
        let mut writer = PathWriter {
            shift_x,
            ascent,
            output: String::with_capacity(2048),
        };
        unsafe {
            CGPathApply(
                path.0,
                &mut writer as *mut PathWriter as *mut c_void,
                Some(append_path_element),
            );
        }
        if writer.output.is_empty() || writer.output.len() > 240_000 {
            return Err("The system mathematical glyph outline is empty or too complex.".to_string());
        }

        Ok(SystemMathGlyphOutline {
            character: character.to_string(),
            requested_family: primary_requested_family.to_string(),
            resolved_family: font.resolved_family,
            fallback_used: normalize_font_family(primary_requested_family)
                != normalize_font_family(&font.requested_family),
            glyph_id: glyph,
            path: writer.output,
            metrics: SystemMathGlyphMetrics {
                width_em: width / 1000.0,
                ascent_em: ascent / 1000.0,
                descent_em: descent / 1000.0,
            },
        })
    }

    pub(super) fn probe_fonts(font_families: &[String]) -> Vec<SystemMathFontProbe> {
        font_families
            .iter()
            .filter_map(|requested| {
                let requested_family = validate_font_family(requested).ok()?;
                match resolve_font(&requested_family) {
                    Ok(font) => Some(SystemMathFontProbe {
                        requested_family,
                        resolved_family: font.resolved_family,
                        available: font.exact_family,
                    }),
                    Err(_) => Some(SystemMathFontProbe {
                        requested_family,
                        resolved_family: String::new(),
                        available: false,
                    }),
                }
            })
            .collect()
    }

    pub(super) fn extract_outline(
        requested_families: &[String],
        character: &str,
    ) -> Result<SystemMathGlyphOutline, String> {
        let character = validate_character(character)?;
        let mut candidates = Vec::new();
        for family in requested_families {
            let family = validate_font_family(family)?;
            if !candidates.iter().any(|candidate| candidate == &family) {
                candidates.push(family);
            }
        }
        for fallback in DEFAULT_MATH_FONT_FALLBACKS {
            if !candidates
                .iter()
                .any(|candidate| normalize_font_family(candidate) == normalize_font_family(fallback))
            {
                candidates.push((*fallback).to_string());
            }
        }
        let primary = candidates
            .first()
            .cloned()
            .ok_or_else(|| "No system math font family was requested.".to_string())?;
        let mut errors = Vec::new();
        for candidate in candidates {
            let font = match resolve_font(&candidate) {
                Ok(font) if font.exact_family => font,
                Ok(font) => {
                    errors.push(format!(
                        "{} resolved to {} instead of the requested family.",
                        candidate, font.resolved_family
                    ));
                    continue;
                }
                Err(error) => {
                    errors.push(error);
                    continue;
                }
            };
            match outline_from_font(font, &character, &primary) {
                Ok(outline) => return Ok(outline),
                Err(error) => errors.push(error),
            }
        }
        Err(format!(
            "VisualTeX could not find a system math font containing {}. {}",
            character,
            errors.join(" ")
        ))
    }
}

#[tauri::command]
pub(crate) fn probe_macos_math_fonts(
    font_families: Vec<String>,
) -> Result<Vec<SystemMathFontProbe>, String> {
    if font_families.len() > 32 {
        return Err("Too many system math font families were requested.".to_string());
    }
    #[cfg(target_os = "macos")]
    {
        return Ok(macos::probe_fonts(&font_families));
    }
    #[cfg(not(target_os = "macos"))]
    {
        let _ = font_families;
        Err("System mathematical glyph extraction is available only on macOS.".to_string())
    }
}

#[tauri::command]
pub(crate) fn extract_macos_math_glyph(
    font_families: Vec<String>,
    character: String,
) -> Result<SystemMathGlyphOutline, String> {
    if font_families.is_empty() || font_families.len() > 16 {
        return Err("A bounded system math font fallback list is required.".to_string());
    }
    #[cfg(target_os = "macos")]
    {
        return macos::extract_outline(&font_families, &character);
    }
    #[cfg(not(target_os = "macos"))]
    {
        let _ = (font_families, character);
        Err("System mathematical glyph extraction is available only on macOS.".to_string())
    }
}

#[cfg(all(test, target_os = "macos"))]
mod tests {
    use super::*;

    #[test]
    fn rejects_multiple_unicode_scalars_for_one_glyph_request() {
        assert!(validate_character("AB").is_err());
        assert!(validate_character("αβ").is_err());
        assert_eq!(validate_character("𝛼").as_deref(), Ok("𝛼"));
    }

    #[test]
    fn rejects_a_missing_font_without_accepting_coretext_substitution() {
        let probes = macos::probe_fonts(&["VisualTeX Definitely Missing Math".to_string()]);
        assert_eq!(probes.len(), 1);
        assert!(!probes[0].available);
    }

    #[test]
    fn extracts_a_crisp_stix_math_outline_when_available() {
        let probes = macos::probe_fonts(&["STIX Two Math".to_string()]);
        if !probes.first().is_some_and(|probe| probe.available) {
            return;
        }
        let outline = macos::extract_outline(&["STIX Two Math".to_string()], "𝛼")
            .expect("STIX Two Math should expose the mathematical italic alpha outline");
        assert_eq!(outline.resolved_family, "STIX Two Math");
        assert!(outline.path.starts_with('M'));
        assert!(!outline.path.contains("NaN"));
        assert!(outline.metrics.width_em > 0.1);
        assert!(outline.metrics.ascent_em > 0.1);
    }
}
