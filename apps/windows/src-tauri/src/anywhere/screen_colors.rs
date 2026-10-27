//! The formula's colour: Windows' input methods do not learn the field's text colour, so the
//! line around the caret is read off the screen (the formula's own window is left out of
//! captures) and its most common colour is the field's background; the formula is near-white on
//! a dark one, near-black on a light one, whatever the system's theme says. Only the background
//! decides, as on the Mac (screen_colors.rs there).

use super::ScreenRect;
use windows::Win32::Graphics::Gdi::{
    BitBlt, CreateCompatibleDC, CreateDIBSection, DeleteDC, DeleteObject, GetDC, ReleaseDC, SelectObject, BITMAPINFO,
    BITMAPINFOHEADER, BI_RGB, CAPTUREBLT, DIB_RGB_COLORS, HGDIOBJ, SRCCOPY,
};

pub(super) struct FieldColors {
    pub(super) text: String,
    pub(super) background: String,
}

/// The field's colours at `caret` (physical pixels), as #rrggbbaa.
pub(super) fn field_colors(caret: ScreenRect) -> Option<FieldColors> {
    let height = caret.height.max(10.0).round() as i32;
    let (scale, _) = super::panel::monitor_at(caret.x, caret.y);
    // The line around the caret, mostly before it (after it the field may still show a
    // placeholder).
    let before = (160.0 * scale).round() as i32;
    let width = before + (80.0 * scale).round() as i32;
    let left = caret.x.round() as i32 - before;
    let top = caret.y.round() as i32;
    let pixels = capture(left, top, width, height)?;
    let background = background(&pixels);
    Some(FieldColors {
        text: hex(text_color_on(background)),
        background: hex([background[0], background[1], background[2], 1.0]),
    })
}

fn capture(left: i32, top: i32, width: i32, height: i32) -> Option<Vec<[f64; 3]>> {
    unsafe {
        let screen = GetDC(None);
        if screen.is_invalid() {
            return None;
        }
        let memory = CreateCompatibleDC(Some(screen));
        let info = BITMAPINFO {
            bmiHeader: BITMAPINFOHEADER {
                biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                biWidth: width,
                biHeight: -height,
                biPlanes: 1,
                biBitCount: 32,
                biCompression: BI_RGB.0,
                ..Default::default()
            },
            ..Default::default()
        };
        let mut bits: *mut std::ffi::c_void = std::ptr::null_mut();
        let bitmap = CreateDIBSection(Some(memory), &info, DIB_RGB_COLORS, &mut bits, None, 0);
        let result = match bitmap {
            Ok(bitmap) if !bits.is_null() => {
                let previous = SelectObject(memory, HGDIOBJ(bitmap.0));
                let copied = BitBlt(memory, 0, 0, width, height, Some(screen), left, top, SRCCOPY | CAPTUREBLT).is_ok();
                let pixels = copied.then(|| {
                    let bytes = std::slice::from_raw_parts(bits as *const u8, (width * height * 4) as usize);
                    bytes
                        .chunks_exact(4)
                        .map(|pixel| [pixel[2] as f64 / 255.0, pixel[1] as f64 / 255.0, pixel[0] as f64 / 255.0])
                        .collect::<Vec<_>>()
                });
                SelectObject(memory, previous);
                let _ = DeleteObject(HGDIOBJ(bitmap.0));
                pixels
            }
            _ => None,
        };
        let _ = DeleteDC(memory);
        ReleaseDC(None, screen);
        result
    }
}

fn luminance(color: [f64; 3]) -> f64 {
    0.2126 * color[0] + 0.7152 * color[1] + 0.0722 * color[2]
}

/// The most common colour.
fn background(pixels: &[[f64; 3]]) -> [f64; 3] {
    let mut buckets = std::collections::HashMap::<(u8, u8, u8), (usize, [f64; 3])>::new();
    for color in pixels {
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
    format!("#{:02x}{:02x}{:02x}{:02x}", channel(color[0]), channel(color[1]), channel(color[2]), channel(color[3]))
}
