//! The formula's colour for fields that do not tell their text colour to
//! the input method (Chrome, Electron): the line around the caret is captured
//! from beneath VisualTeX's own panel and its most common colour is the
//! field's background; the formula is near-white on a dark one, near-black on
//! a light one, whatever the system appearance says. Only the background
//! decides: the text before the caret (or the field's own caret, which some
//! fields draw in a colour) does not change it.
//!
//! Only with the Screen Recording permission (asked for once at startup when
//! the input method is in use, see mod.rs); without it the page keeps the
//! system appearance's text colour.

use super::ScreenRect;
use std::ffi::{c_void, CString};

type CFTypeRef = *const c_void;

#[repr(C)]
#[derive(Clone, Copy)]
struct CGRect {
    x: f64,
    y: f64,
    width: f64,
    height: f64,
}

/// CGWindowListCreateImage, looked up at run time: deprecated, and a missing
/// symbol must not keep VisualTeX from starting on a later macOS.
type CreateImage = unsafe extern "C" fn(CGRect, u32, u32, u32) -> CFTypeRef;

const LIST_ON_SCREEN_BELOW_WINDOW: u32 = 1 << 2;
const IMAGE_NOMINAL_RESOLUTION: u32 = 1 << 4;
const ALPHA_INFO_MASK: u32 = 0x1f;
const BYTE_ORDER_MASK: u32 = 0x7000;
const BYTE_ORDER_32_LITTLE: u32 = 2 << 12;
/// kCGImageAlphaPremultipliedLast, kCGImageAlphaLast, kCGImageAlphaNoneSkipLast
const ALPHA_LAST: [u32; 3] = [1, 3, 5];

#[link(name = "CoreGraphics", kind = "framework")]
unsafe extern "C" {
    fn CGPreflightScreenCaptureAccess() -> bool;
    fn CGRequestScreenCaptureAccess() -> bool;
    fn CGImageGetWidth(image: CFTypeRef) -> usize;
    fn CGImageGetHeight(image: CFTypeRef) -> usize;
    fn CGImageGetBytesPerRow(image: CFTypeRef) -> usize;
    fn CGImageGetBitsPerPixel(image: CFTypeRef) -> usize;
    fn CGImageGetBitmapInfo(image: CFTypeRef) -> u32;
    fn CGImageGetDataProvider(image: CFTypeRef) -> CFTypeRef;
    fn CGDataProviderCopyData(provider: CFTypeRef) -> CFTypeRef;
}

#[link(name = "CoreFoundation", kind = "framework")]
unsafe extern "C" {
    fn CFRelease(value: CFTypeRef);
    fn CFDataGetBytePtr(data: CFTypeRef) -> *const u8;
    fn CFDataGetLength(data: CFTypeRef) -> isize;
}

pub(super) fn can_read_screen() -> bool {
    unsafe { CGPreflightScreenCaptureAccess() }
}

/// Shows macOS's Screen Recording request (once per app; it takes effect
/// when VisualTeX next starts). Main thread.
pub(super) fn request_access() -> bool {
    unsafe { CGRequestScreenCaptureAccess() }
}

fn create_image_function() -> Option<CreateImage> {
    let name = CString::new("CGWindowListCreateImage").ok()?;
    let symbol = unsafe { libc::dlsym(libc::RTLD_DEFAULT, name.as_ptr()) };
    (!symbol.is_null()).then(|| unsafe { std::mem::transmute::<*mut c_void, CreateImage>(symbol) })
}

/// The field's colours at `caret` (Quartz coordinates), as #rrggbbaa: the
/// formula's (near-white or near-black) and the background's (for the Image
/// output mode). None without the permission or a readable capture.
/// `panel_window` is VisualTeX's panel: what is beneath it is captured.
pub(super) struct FieldColors {
    pub(super) text: String,
    pub(super) background: String,
}

pub(super) fn field_colors(caret: ScreenRect, panel_window: u32) -> Option<FieldColors> {
    if !can_read_screen() {
        return None;
    }
    let create = create_image_function()?;
    let height = caret.height.max(10.0);
    // The line around the caret, mostly before it (after it the field may
    // still show a placeholder).
    let before = 160.0;
    let rect = CGRect {
        x: caret.x - before,
        y: caret.y,
        width: before + 80.0,
        height,
    };
    let image = unsafe { create(rect, LIST_ON_SCREEN_BELOW_WINDOW, panel_window, IMAGE_NOMINAL_RESOLUTION) };
    if image.is_null() {
        return None;
    }
    let pixels = read_pixels(image);
    unsafe { CFRelease(image) };
    let background = background(&pixels?);
    Some(FieldColors {
        text: hex(text_color_on(background)),
        background: hex([background[0], background[1], background[2], 1.0]),
    })
}

struct Pixels {
    rgb: Vec<[f64; 3]>,
}

fn read_pixels(image: CFTypeRef) -> Option<Pixels> {
    unsafe {
        if CGImageGetBitsPerPixel(image) != 32 {
            return None;
        }
        let width = CGImageGetWidth(image);
        let height = CGImageGetHeight(image);
        let row = CGImageGetBytesPerRow(image);
        let info = CGImageGetBitmapInfo(image);
        let provider = CGImageGetDataProvider(image);
        if provider.is_null() || width == 0 || height == 0 {
            return None;
        }
        let data = CGDataProviderCopyData(provider);
        if data.is_null() {
            return None;
        }
        let length = CFDataGetLength(data).max(0) as usize;
        let bytes = std::slice::from_raw_parts(CFDataGetBytePtr(data), length);
        let little = info & BYTE_ORDER_MASK == BYTE_ORDER_32_LITTLE;
        let alpha_last = ALPHA_LAST.contains(&(info & ALPHA_INFO_MASK));
        // Byte positions of red, green and blue in a pixel.
        let (r, g, b) = match (little, alpha_last) {
            (true, false) => (2, 1, 0),  // BGRA
            (true, true) => (3, 2, 1),   // ABGR
            (false, false) => (1, 2, 3), // ARGB
            (false, true) => (0, 1, 2),  // RGBA
        };
        let mut rgb = Vec::with_capacity(width * height);
        for y in 0..height {
            for x in 0..width {
                let at = y * row + x * 4;
                if at + 3 >= length {
                    break;
                }
                rgb.push([bytes[at + r] as f64 / 255.0, bytes[at + g] as f64 / 255.0, bytes[at + b] as f64 / 255.0]);
            }
        }
        CFRelease(data);
        (rgb.len() == width * height).then_some(Pixels { rgb })
    }
}

fn luminance(color: [f64; 3]) -> f64 {
    0.2126 * color[0] + 0.7152 * color[1] + 0.0722 * color[2]
}

/// The most common colour.
fn background(pixels: &Pixels) -> [f64; 3] {
    let mut buckets = std::collections::HashMap::<(u8, u8, u8), (usize, [f64; 3])>::new();
    for color in &pixels.rgb {
        let key = ((color[0] * 15.0) as u8, (color[1] * 15.0) as u8, (color[2] * 15.0) as u8);
        let entry = buckets.entry(key).or_insert((0, [0.0; 3]));
        entry.0 += 1;
        for channel in 0..3 {
            entry.1[channel] += color[channel];
        }
    }
    buckets
        .values()
        .max_by_key(|(count, _)| *count)
        .map(|(count, sum)| [sum[0] / *count as f64, sum[1] / *count as f64, sum[2] / *count as f64])
        .unwrap_or([1.0; 3])
}

/// Near-white on a dark background, near-black on a light one.
fn text_color_on(background: [f64; 3]) -> [f64; 4] {
    if luminance(background) < 0.5 {
        [1.0, 1.0, 1.0, 0.92]
    } else {
        [0.0, 0.0, 0.0, 0.9]
    }
}

fn hex(color: [f64; 4]) -> String {
    let channel = |value: f64| (value.clamp(0.0, 1.0) * 255.0).round() as u8;
    format!(
        "#{:02x}{:02x}{:02x}{:02x}",
        channel(color[0]),
        channel(color[1]),
        channel(color[2]),
        channel(color[3])
    )
}
