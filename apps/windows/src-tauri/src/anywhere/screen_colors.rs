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

/// A caret a field draws itself, as seen on the screen (physical pixels).
#[derive(Clone, Copy, Debug, PartialEq)]
pub(super) struct CaretBar {
    pub(super) top: f64,
    pub(super) height: f64,
}

/// The caret a field draws itself standing on `rect`'s top edge (physical pixels): a thin vertical
/// bar, alone in its columns, its foot within a few pixels of that edge, near `rect`'s left. Where
/// an application reports a made-up caret under its real one (WeChat: a line of the default font
/// below it), this is the real one; None while it blinks off, or where there is none.
pub(super) fn caret_bar_above(rect: ScreenRect) -> Option<CaretBar> {
    let (scale, work) = super::panel::monitor_at(rect.x, rect.y);
    let side = (6.0 * scale).round() as i32;
    let reach = (super::MAX_TEXT_HEIGHT * 0.75 * scale).round() as i32;
    let below = (3.0 * scale).round().max(2.0) as i32;
    // the area above the rectangle's top edge, on the screen only (off it a capture is black)
    let (wanted_left, wanted_top) = (rect.x.round() as i32 - side, rect.y.round() as i32 - reach);
    let left = wanted_left.max(work.left);
    let top = wanted_top.max(work.top);
    let right = (wanted_left + side * 2 + 1).min(work.right);
    let bottom = (wanted_top + reach + below + 1).min(work.bottom);
    let (width, height) = (right - left, bottom - top);
    let (edge, center) = (reach - (top - wanted_top), side - (left - wanted_left));
    if width < 5 || height < 5 || !(0..height).contains(&edge) {
        return None;
    }
    let pixels = capture(left, top, width, height)?;
    let (bar_top, bar_bottom) = bar_in(&pixels, width, height, edge, center, scale)?;
    Some(CaretBar { top: (top + bar_top) as f64, height: (bar_bottom - bar_top + 1) as f64 })
}

/// The rows of the caret bar in `pixels` (`width` x `height`): ending within two rows or so of row
/// `edge` (the foot of a letter before the caret, on the baseline, is higher), nearest column
/// `center`, one to a few pixels wide, and a line of text tall.
fn bar_in(pixels: &[[f64; 3]], width: i32, height: i32, edge: i32, center: i32, scale: f64) -> Option<(i32, i32)> {
    if pixels.len() != (width * height) as usize || width < 5 || height < 5 {
        return None;
    }
    let background = background(pixels);
    let ink: Vec<bool> = pixels.iter().map(|&color| distance(color, background) > 0.28).collect();
    let at = |x: i32, y: i32| ink[(y * width + x) as usize];
    let tolerance = (1.5 * scale).round().max(2.0) as i32;
    let shortest = (super::MIN_TEXT_HEIGHT * 0.8 * scale).round() as i32;
    let tallest = (super::MAX_TEXT_HEIGHT * scale).round() as i32;
    // per column: the run of ink whose foot is the lowest ink near the edge
    let runs: Vec<Option<(i32, i32)>> = (0..width)
        .map(|x| {
            let foot = ((edge - tolerance).max(0)..=(edge + tolerance).min(height - 1)).rev().find(|&y| at(x, y))?;
            let mut head = foot;
            while head > 0 && at(x, head - 1) {
                head -= 1;
            }
            (foot - head + 1 >= shortest && foot - head < tallest).then_some((head, foot))
        })
        .collect();
    let widest = (3.0 * scale).ceil() as i32 + 1;
    let mut best: Option<(i32, i32, i32)> = None;
    let mut x = 0;
    while x < width {
        let Some((head, foot)) = runs[x as usize] else {
            x += 1;
            continue;
        };
        // adjacent columns with nearly the same run are one bar
        let (mut right, mut top, mut bottom) = (x, head, foot);
        while right + 1 < width {
            match runs[(right + 1) as usize] {
                Some((next_head, next_foot)) if (next_head - head).abs() <= 2 && (next_foot - foot).abs() <= 2 => {
                    right += 1;
                    top = top.min(next_head);
                    bottom = bottom.max(next_foot);
                }
                _ => break,
            }
        }
        // alone: the columns two pixels beside it are background along it (no glyph stroke)
        let clear = |column: i32| {
            column < 0
                || column >= width
                || (top..=bottom).filter(|&y| !at(column, y)).count() * 10 >= ((bottom - top + 1) * 8) as usize
        };
        if right - x + 1 <= widest && clear(x - 2) && clear(right + 2) {
            let off = ((x + right) / 2 - center).abs();
            if best.map(|(nearest, _, _)| off < nearest).unwrap_or(true) {
                best = Some((off, top, bottom));
            }
        }
        x = right + 1;
    }
    best.map(|(_, top, bottom)| (top, bottom))
}

fn distance(left: [f64; 3], right: [f64; 3]) -> f64 {
    ((left[0] - right[0]).powi(2) + (left[1] - right[1]).powi(2) + (left[2] - right[2]).powi(2)).sqrt()
}

/// The field's colours at `caret` (physical pixels), as #rrggbbaa.
pub(super) fn field_colors(caret: ScreenRect) -> Option<FieldColors> {
    let height = caret.height.max(10.0).round() as i32;
    let (scale, work) = super::panel::monitor_at(caret.x, caret.y);
    // The line right around the caret: what the cover must match is the field's colour where the
    // formula stands. (A wide stretch before the caret reached past the field's left edge, and for
    // a caret at the start of a line the cover took the colour of what is beside the field: a faint
    // grey box under the formula, 2026-11-15; on the Mac it was a chat's message box, 2026-11-11.)
    // Only what is on the screen: beyond its edge a capture is black, which would make a light
    // field's formula white (a maximized window's caret near the edge).
    let before = (14.0 * scale).round() as i32;
    let width = before + (44.0 * scale).round() as i32;
    let left = (caret.x.round() as i32 - before).max(work.left);
    let right = (caret.x.round() as i32 - before + width).min(work.right);
    let top = (caret.y.round() as i32).max(work.top);
    let bottom = (caret.y.round() as i32 + height).min(work.bottom);
    if right - left < 4 || bottom - top < 4 {
        return None;
    }
    let pixels = capture(left, top, right - left, bottom - top)?;
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

/// The most common colour: the one colour most pixels have exactly (a field's background is
/// flat), when a fifth of them have it; else the most common among colours taken coarsely (a
/// gradient, a translucent field), averaged. (Coarsely alone, the field's colour and the near
/// greys of its glyph edges or of what lies beside it fell together, and the cover was a shade off.)
fn background(pixels: &[[f64; 3]]) -> [f64; 3] {
    let mut exact = std::collections::HashMap::<(u8, u8, u8), usize>::new();
    for color in pixels {
        let key = ((color[0] * 255.0).round() as u8, (color[1] * 255.0).round() as u8, (color[2] * 255.0).round() as u8);
        *exact.entry(key).or_insert(0) += 1;
    }
    if let Some((key, count)) = exact.iter().max_by_key(|(_, count)| **count) {
        if *count * 5 >= pixels.len() {
            return [key.0 as f64 / 255.0, key.1 as f64 / 255.0, key.2 as f64 / 255.0];
        }
    }
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

#[cfg(test)]
mod tests {
    use super::*;

    const PAPER: [f64; 3] = [0.95, 0.95, 0.953];
    const GREEN: [f64; 3] = [0.03, 0.76, 0.38];
    const INK: [f64; 3] = [0.0, 0.0, 0.0];

    fn canvas(width: i32, height: i32) -> Vec<[f64; 3]> {
        vec![PAPER; (width * height) as usize]
    }

    fn fill(pixels: &mut [[f64; 3]], width: i32, x: std::ops::Range<i32>, y: std::ops::Range<i32>, color: [f64; 3]) {
        for row in y {
            for column in x.clone() {
                pixels[(row * width + column) as usize] = color;
            }
        }
    }

    #[test]
    fn wechats_caret_stands_on_its_made_up_one() {
        // WeChat at 150 %: a 2 px green caret, 27 px tall, its foot on the reported rectangle's top
        let (width, height, edge, center) = (19, 113, 108, 9);
        let mut pixels = canvas(width, height);
        fill(&mut pixels, width, 7..9, 81..108, GREEN);
        assert_eq!(bar_in(&pixels, width, height, edge, center, 1.5), Some((81, 107)));
        // after text: the letter before it stands on the baseline, five pixels higher
        fill(&mut pixels, width, 1..3, 80..103, INK);
        assert_eq!(bar_in(&pixels, width, height, edge, center, 1.5), Some((81, 107)));
    }

    #[test]
    fn glyphs_and_short_marks_are_no_caret() {
        let (width, height, edge, center) = (19, 113, 108, 9);
        // a letter's stem with its bowl beside it
        let mut glyph = canvas(width, height);
        fill(&mut glyph, width, 7..9, 85..108, INK);
        fill(&mut glyph, width, 9..15, 96..108, INK);
        assert_eq!(bar_in(&glyph, width, height, edge, center, 1.5), None);
        // a mark much shorter than a line of text
        let mut mark = canvas(width, height);
        fill(&mut mark, width, 7..9, 100..108, INK);
        assert_eq!(bar_in(&mark, width, height, edge, center, 1.5), None);
        // nothing (the caret blinked off)
        assert_eq!(bar_in(&canvas(width, height), width, height, edge, center, 1.5), None);
        // a bar that does not reach the edge: a letter's stem on the baseline
        let mut floating = canvas(width, height);
        fill(&mut floating, width, 7..9, 80..103, INK);
        assert_eq!(bar_in(&floating, width, height, edge, center, 1.5), None);
    }

    #[test]
    fn a_flat_field_gives_its_colour_exactly() {
        // a light grey field (#f3f3f3) and the paler near greys of glyph edges in the same coarse
        // bucket: the colour is the field's, not an average with them (the cover was a shade off the
        // field, a faint box)
        let field = [243.0 / 255.0; 3];
        let (width, height) = (60, 10);
        let mut pixels = vec![field; (width * height) as usize];
        for step in 0..12 {
            fill(&mut pixels, width, 6 + step..7 + step, 0..4, [0.96 + 0.003 * step as f64; 3]);
        }
        assert_eq!(background(&pixels), field);
    }

    #[test]
    fn a_gradient_gives_its_most_common_shade() {
        // no colour reaches a fifth: the coarse buckets decide, averaged
        let pixels: Vec<[f64; 3]> = (0..100).map(|index| [0.2 + 0.006 * index as f64; 3]).collect();
        let color = background(&pixels);
        assert!(color[0] > 0.2 && color[0] < 0.8, "{color:?}");
    }
}
