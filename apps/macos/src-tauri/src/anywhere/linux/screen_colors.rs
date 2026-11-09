//! The field's colours read off the screen (X11) around where the formula goes, so the formula is
//! drawn in a colour that reads on the field and its cover hides the field's own text (as on macOS,
//! screen_colors.rs there; read before the formula shows, so it is not in the picture).

use super::ScreenRect;
use gtk::prelude::*;

pub(super) fn can_read_screen() -> bool {
    std::env::var_os("WAYLAND_DISPLAY").is_none() || std::env::var("GDK_BACKEND").map(|backend| backend == "x11").unwrap_or(false)
}

pub(super) fn request_access() -> bool {
    can_read_screen()
}

pub(super) struct FieldColors {
    pub(super) text: String,
    pub(super) background: String,
}

pub(super) fn field_colors(caret: ScreenRect, _panel_window: u32) -> Option<FieldColors> {
    if !can_read_screen() {
        return None;
    }
    let app = super::panel::app()?;
    let (sender, receiver) = std::sync::mpsc::sync_channel(1);
    // the line around the caret, mostly before it (after it the field may still show a placeholder)
    let before = 160.0;
    let (x, y, width, height) = (caret.x - before, caret.y, before + 80.0, caret.height.max(10.0));
    app.run_on_main_thread(move || {
        let pixels = (|| {
            let root = gdk::Screen::default()?.root_window()?;
            let pixbuf = root.pixbuf(x.max(0.0) as i32, y.max(0.0) as i32, width as i32, height as i32)?;
            let channels = pixbuf.n_channels() as usize;
            let stride = pixbuf.rowstride() as usize;
            let bytes = pixbuf.read_pixel_bytes();
            let mut rgb = Vec::new();
            for row in 0..pixbuf.height() as usize {
                for column in 0..pixbuf.width() as usize {
                    let at = row * stride + column * channels;
                    if at + 2 < bytes.len() {
                        rgb.push([bytes[at] as f64 / 255.0, bytes[at + 1] as f64 / 255.0, bytes[at + 2] as f64 / 255.0]);
                    }
                }
            }
            Some(rgb)
        })();
        let _ = sender.send(pixels);
    })
    .ok()?;
    let pixels = receiver.recv_timeout(std::time::Duration::from_millis(500)).ok()??;
    if pixels.is_empty() {
        return None;
    }
    let background = background(&pixels);
    Some(FieldColors {
        text: hex(text_color_on(background)),
        background: hex([background[0], background[1], background[2], 1.0]),
    })
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
