//! The caret of a field that tells nothing about it, found on the screen:
//! WeChat draws its own text and shows Accessibility nothing; QQ's Chromium
//! reports the chat box but no rectangle for a place in it (the whole box, or
//! nothing). What every field does draw is its caret, blinking: a thin upright
//! bar that comes and goes. Only the application's own front window is
//! captured (no other window, not VisualTeX's panel), a few times within one
//! blink, and nothing is kept. Needs the Screen Recording permission (asked
//! for at startup, see mod.rs).
//!
//! Measured 2026-11-16 (macOS 27.0.1): WeChat's caret 2 x 17 pt and QQ's
//! 1 x 20 pt each came out as the only bar that changed in their window over
//! two seconds; both blink about twice a second.

use super::ScreenRect;
use std::ffi::{c_void, CString};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Arc;
use std::time::{Duration, Instant};

type CFTypeRef = *const c_void;

#[repr(C)]
#[derive(Clone, Copy, Debug, Default, PartialEq)]
struct CGRect {
    x: f64,
    y: f64,
    width: f64,
    height: f64,
}

/// CGWindowListCreateImage, looked up at run time (deprecated; a missing
/// symbol must not keep VisualTeX from starting).
type CreateImage = unsafe extern "C" fn(CGRect, u32, u32, u32) -> CFTypeRef;

/// kCGWindowListOptionOnScreenOnly | kCGWindowListExcludeDesktopElements
const LIST_ON_SCREEN: u32 = (1 << 0) | (1 << 4);
/// kCGWindowListOptionIncludingWindow: that window alone.
const LIST_INCLUDING_WINDOW: u32 = 1 << 3;
/// kCGWindowImageBoundsIgnoreFraming | kCGWindowImageBestResolution
const IMAGE_OPTIONS: u32 = (1 << 0) | (1 << 3);
const CF_NUMBER_SINT64: isize = 4;
const ALPHA_INFO_MASK: u32 = 0x1f;
const BYTE_ORDER_MASK: u32 = 0x7000;
const BYTE_ORDER_32_LITTLE: u32 = 2 << 12;
/// kCGImageAlphaPremultipliedLast, kCGImageAlphaLast, kCGImageAlphaNoneSkipLast
const ALPHA_LAST: [u32; 3] = [1, 3, 5];

/// How long the caret is looked for: a little more than one blink (the
/// caret shows or hides about every half second), so that it has changed at
/// least once.
pub(super) const SEARCH_TIME: Duration = Duration::from_millis(650);
/// Between two captures (a capture of a whole window takes some 40-50 ms).
const CAPTURE_PAUSE: Duration = Duration::from_millis(25);

#[link(name = "CoreGraphics", kind = "framework")]
unsafe extern "C" {
    static kCGWindowOwnerPID: CFTypeRef;
    static kCGWindowLayer: CFTypeRef;
    static kCGWindowNumber: CFTypeRef;
    static kCGWindowBounds: CFTypeRef;
    fn CGPreflightScreenCaptureAccess() -> bool;
    fn CGWindowListCopyWindowInfo(option: u32, relative_to_window: u32) -> CFTypeRef;
    fn CGRectMakeWithDictionaryRepresentation(dictionary: CFTypeRef, rect: *mut CGRect) -> bool;
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
    fn CFArrayGetCount(array: CFTypeRef) -> isize;
    fn CFArrayGetValueAtIndex(array: CFTypeRef, index: isize) -> *const c_void;
    fn CFDictionaryGetValue(dictionary: CFTypeRef, key: *const c_void) -> *const c_void;
    fn CFNumberGetValue(number: CFTypeRef, kind: isize, value: *mut c_void) -> u8;
    fn CFDataGetBytePtr(data: CFTypeRef) -> *const u8;
    fn CFDataGetLength(data: CFTypeRef) -> isize;
}

fn symbol(name: &str) -> Option<*mut c_void> {
    let name = CString::new(name).ok()?;
    let symbol = unsafe { libc::dlsym(libc::RTLD_DEFAULT, name.as_ptr()) };
    (!symbol.is_null()).then_some(symbol)
}

/// A window of an application: its number and frame (Quartz points).
#[derive(Clone, Copy, Debug)]
pub(super) struct Window {
    number: u32,
    frame: CGRect,
}

impl Window {
    pub(super) fn frame(&self) -> ScreenRect {
        ScreenRect {
            x: self.frame.x,
            y: self.frame.y,
            width: self.frame.width,
            height: self.frame.height,
        }
    }
}

/// The application's frontmost ordinary window on the screen (not a menu,
/// a tooltip or a tiny helper window).
pub(super) fn front_window(pid: i32) -> Option<Window> {
    unsafe {
        let list = CGWindowListCopyWindowInfo(LIST_ON_SCREEN, 0);
        if list.is_null() {
            return None;
        }
        let mut found = None;
        for index in 0..CFArrayGetCount(list) {
            let info = CFArrayGetValueAtIndex(list, index);
            let number = |key: CFTypeRef| -> Option<i64> {
                let value = CFDictionaryGetValue(info, key);
                if value.is_null() {
                    return None;
                }
                let mut out: i64 = 0;
                (CFNumberGetValue(value, CF_NUMBER_SINT64, (&mut out as *mut i64).cast()) != 0).then_some(out)
            };
            if number(kCGWindowOwnerPID) != Some(pid as i64) || number(kCGWindowLayer) != Some(0) {
                continue;
            }
            let Some(window) = number(kCGWindowNumber) else { continue };
            let bounds = CFDictionaryGetValue(info, kCGWindowBounds);
            let mut frame = CGRect::default();
            if bounds.is_null() || !CGRectMakeWithDictionaryRepresentation(bounds, &mut frame) {
                continue;
            }
            if frame.width < 160.0 || frame.height < 100.0 {
                continue;
            }
            found = Some(Window { number: window as u32, frame });
            break;
        }
        CFRelease(list);
        found
    }
}

/// A capture in grey (0-255), row by row from the top.
struct Frame {
    grey: Vec<u8>,
    width: usize,
    height: usize,
}

fn capture(create: CreateImage, window: &Window) -> Option<Frame> {
    let image = unsafe { create(window.frame, LIST_INCLUDING_WINDOW, window.number, IMAGE_OPTIONS) };
    if image.is_null() {
        return None;
    }
    let frame = grey_pixels(image);
    unsafe { CFRelease(image) };
    frame
}

fn grey_pixels(image: CFTypeRef) -> Option<Frame> {
    unsafe {
        if CGImageGetBitsPerPixel(image) != 32 {
            return None;
        }
        let width = CGImageGetWidth(image);
        let height = CGImageGetHeight(image);
        let row = CGImageGetBytesPerRow(image);
        let info = CGImageGetBitmapInfo(image);
        let provider = CGImageGetDataProvider(image);
        if provider.is_null() || width == 0 || height == 0 || row < width * 4 {
            return None;
        }
        let data = CGDataProviderCopyData(provider);
        if data.is_null() {
            return None;
        }
        let length = CFDataGetLength(data).max(0) as usize;
        if length < row * (height - 1) + width * 4 {
            CFRelease(data);
            return None;
        }
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
        let mut grey = Vec::with_capacity(width * height);
        for y in 0..height {
            let start = y * row;
            for x in 0..width {
                let at = start + x * 4;
                let value = 54 * bytes[at + r] as u32 + 183 * bytes[at + g] as u32 + 19 * bytes[at + b] as u32;
                grey.push((value >> 8) as u8);
            }
        }
        CFRelease(data);
        Some(Frame { grey, width, height })
    }
}

/// A thin upright bar where two captures differ, in pixels of the capture
/// (inclusive bounds).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct Bar {
    pub(super) x0: usize,
    pub(super) x1: usize,
    pub(super) y0: usize,
    pub(super) y1: usize,
}

impl Bar {
    fn near(&self, other: &Bar) -> bool {
        self.x0.abs_diff(other.x0) <= 2 && self.y0.abs_diff(other.y0) <= 2 && self.y1.abs_diff(other.y1) <= 2
    }
}

/// A pixel changed (a caret shown or hidden: its whole contrast).
const CHANGED: u8 = 48;
/// A pixel stayed what it was.
const STILL: u8 = 16;

/// The bars that tell a caret where `first` and `now` differ: changed runs
/// down a column (a text line's height: 8-64 pt), at most 4 pt wide, with
/// nothing changed just beside them (a glyph typed, a picture playing or a
/// scroll change wider areas). `scale`: pixels per point.
pub(super) fn changed_bars(first: &[u8], now: &[u8], width: usize, height: usize, scale: f64) -> Vec<Bar> {
    if first.len() != width * height || now.len() != width * height || width == 0 || height == 0 {
        return Vec::new();
    }
    let difference = |x: usize, y: usize| first[y * width + x].abs_diff(now[y * width + x]);
    let min_height = ((8.0 * scale).round() as usize).max(4);
    let max_height = (64.0 * scale).round() as usize;
    let max_width = ((4.0 * scale).round() as usize).max(1);
    // changed runs down each column (a pixel's gap allowed: anti-aliasing)
    let mut runs: Vec<(usize, usize, usize)> = Vec::new();
    for x in 0..width {
        let mut y = 0;
        while y < height {
            if difference(x, y) < CHANGED {
                y += 1;
                continue;
            }
            let start = y;
            let mut end = y;
            let mut gap = 0;
            let mut next = y + 1;
            while next < height {
                if difference(x, next) >= CHANGED {
                    end = next;
                    gap = 0;
                } else {
                    gap += 1;
                    if gap > 1 {
                        break;
                    }
                }
                next += 1;
            }
            if end + 1 - start >= min_height {
                runs.push((x, start, end));
            }
            y = end + 1;
        }
    }
    // runs side by side make one bar
    let mut bars: Vec<Bar> = Vec::new();
    for (x, y0, y1) in runs {
        let joined = bars
            .iter_mut()
            .rev()
            .find(|bar| bar.x1 + 1 >= x && y0.max(bar.y0) <= y1.min(bar.y1));
        match joined {
            Some(bar) => {
                bar.x1 = bar.x1.max(x);
                bar.y0 = bar.y0.min(y0);
                bar.y1 = bar.y1.max(y1);
            }
            None => bars.push(Bar { x0: x, x1: x, y0, y1 }),
        }
    }
    let isolated = |bar: &Bar| {
        let mut moved = 0;
        let mut total = 0;
        for y in bar.y0..=bar.y1 {
            for x in [bar.x0.wrapping_sub(3), bar.x0.wrapping_sub(2), bar.x1 + 2, bar.x1 + 3] {
                if x < width {
                    total += 1;
                    if difference(x, y) >= STILL {
                        moved += 1;
                    }
                }
            }
        }
        let above_below = [bar.y0.wrapping_sub(2), bar.y1 + 2]
            .into_iter()
            .filter(|&y| y < height)
            .all(|y| (bar.x0..=bar.x1).all(|x| difference(x, y) < CHANGED));
        moved * 10 <= total && above_below
    };
    bars.retain(|bar| {
        let bar_width = bar.x1 + 1 - bar.x0;
        let bar_height = bar.y1 + 1 - bar.y0;
        bar_width <= max_width && (min_height..=max_height).contains(&bar_height) && isolated(bar)
    });
    bars
}

/// Of the bars seen while looking (each with the number of captures it
/// changed in), the caret: the only one, else the one inside the field
/// Accessibility reported, else the one that changed most often.
fn choose(seen: &[(Bar, usize)], inside: impl Fn(&Bar) -> bool) -> Option<Bar> {
    match seen {
        [] => None,
        [(bar, _)] => Some(*bar),
        _ => {
            let within: Vec<&(Bar, usize)> = seen.iter().filter(|(bar, _)| inside(bar)).collect();
            let pool: Vec<&(Bar, usize)> = if within.is_empty() { seen.iter().collect() } else { within };
            let most = pool.iter().map(|(_, count)| *count).max()?;
            let mut best = pool.iter().filter(|(_, count)| *count == most);
            let first = best.next()?;
            // two bars that changed as often: no telling which is the caret
            best.next().is_none().then_some(first.0)
        }
    }
}

/// What the search found, for the strip and anywhere.log.
pub(super) struct Found {
    /// The caret as a zero-width rectangle on its line (Quartz points).
    pub(super) caret: ScreenRect,
    pub(super) report: String,
}

/// A search started when the shortcut is pressed, alongside Accessibility
/// (which answers within some 300 ms in an application that has nothing to
/// tell): the captures are taken meanwhile.
pub(super) struct Search {
    stop: Arc<AtomicBool>,
    worker: Option<std::thread::JoinHandle<Outcome>>,
    window: Option<Window>,
}

struct Outcome {
    seen: Vec<(Bar, usize)>,
    /// A bar that was the only change in a capture: taken at once.
    early: Option<Bar>,
    scale: f64,
    frames: usize,
    elapsed: Duration,
    failure: Option<&'static str>,
}

impl Search {
    /// Starts looking in application `pid`'s front window (no search without
    /// the Screen Recording permission or a window).
    pub(super) fn start(pid: i32) -> Self {
        let none = Self { stop: Arc::new(AtomicBool::new(true)), worker: None, window: None };
        if pid <= 0 || !unsafe { CGPreflightScreenCaptureAccess() } {
            return none;
        }
        let Some(window) = front_window(pid) else { return none };
        let Some(create) = symbol("CGWindowListCreateImage") else { return none };
        let create = unsafe { std::mem::transmute::<*mut c_void, CreateImage>(create) };
        let stop = Arc::new(AtomicBool::new(false));
        let flag = stop.clone();
        let worker = std::thread::Builder::new()
            .name("visualtex-screen-caret".to_string())
            .spawn(move || look(create, window, &flag))
            .ok();
        Self { stop, worker, window: Some(window) }
    }

    /// The window being searched (None: no search).
    pub(super) fn window(&self) -> Option<Window> {
        self.window
    }

    /// Accessibility found the caret: the captures stop.
    pub(super) fn cancel(mut self) {
        self.stop.store(true, Ordering::SeqCst);
        self.worker.take();
    }

    /// The caret found, waiting for the search to end; `field`: the field
    /// Accessibility reported (a bar outside it is no caret of that field).
    pub(super) fn finish(mut self, field: Option<ScreenRect>) -> Result<Found, String> {
        let window = self.window.ok_or_else(|| "no screen search (permission or window)".to_string())?;
        let worker = self.worker.take().ok_or_else(|| "no screen search thread".to_string())?;
        let outcome = worker.join().map_err(|_| "screen search failed".to_string())?;
        let summary = format!("frames={} ms={}", outcome.frames, outcome.elapsed.as_millis());
        if let Some(failure) = outcome.failure {
            return Err(format!("{failure} {summary}"));
        }
        let scale = outcome.scale;
        let to_screen = |bar: &Bar| ScreenRect {
            x: window.frame.x + bar.x0 as f64 / scale,
            y: window.frame.y + bar.y0 as f64 / scale,
            width: 0.0,
            height: (bar.y1 + 1 - bar.y0) as f64 / scale,
        };
        let inside = |bar: &Bar| {
            field.map_or(true, |field| {
                let caret = to_screen(bar);
                caret.x >= field.x - 2.0
                    && caret.x <= field.x + field.width + 2.0
                    && caret.y >= field.y - 2.0
                    && caret.y + caret.height <= field.y + field.height + 2.0
            })
        };
        let bar = match outcome.early {
            Some(bar) if inside(&bar) => Some(bar),
            _ => choose(&outcome.seen, &inside).filter(|bar| inside(bar)),
        };
        let Some(bar) = bar else {
            return Err(format!("no caret (bars seen: {}) {summary}", outcome.seen.len()));
        };
        let caret = to_screen(&bar);
        Ok(Found {
            caret,
            report: format!(
                "{:.0},{:.0} {:.1}x{:.1}pt {summary}",
                caret.x,
                caret.y,
                (bar.x1 + 1 - bar.x0) as f64 / scale,
                caret.height
            ),
        })
    }
}

fn look(create: CreateImage, window: Window, stop: &AtomicBool) -> Outcome {
    let started = Instant::now();
    let mut outcome = Outcome { seen: Vec::new(), early: None, scale: 1.0, frames: 0, elapsed: Duration::ZERO, failure: None };
    let Some(first) = capture(create, &window) else {
        outcome.failure = Some("window not captured");
        return outcome;
    };
    outcome.frames = 1;
    outcome.scale = first.width as f64 / window.frame.width.max(1.0);
    while !stop.load(Ordering::SeqCst) && started.elapsed() < SEARCH_TIME {
        std::thread::sleep(CAPTURE_PAUSE);
        let Some(now) = capture(create, &window) else { break };
        if now.width != first.width || now.height != first.height {
            // the window was resized or moved to another screen meanwhile
            break;
        }
        outcome.frames += 1;
        let bars = changed_bars(&first.grey, &now.grey, first.width, first.height, outcome.scale);
        for bar in &bars {
            match outcome.seen.iter_mut().find(|(seen, _)| seen.near(bar)) {
                Some((_, count)) => *count += 1,
                None => outcome.seen.push((*bar, 1)),
            }
        }
        if bars.len() == 1 {
            outcome.early = Some(bars[0]);
            break;
        }
    }
    outcome.elapsed = started.elapsed();
    outcome
}

#[cfg(test)]
mod tests {
    use super::*;

    const W: usize = 200;
    const H: usize = 120;

    fn field(value: u8) -> Vec<u8> {
        vec![value; W * H]
    }

    fn draw(frame: &mut [u8], x0: usize, x1: usize, y0: usize, y1: usize, value: u8) {
        for y in y0..=y1 {
            for x in x0..=x1 {
                frame[y * W + x] = value;
            }
        }
    }

    #[test]
    fn a_caret_shown_on_a_white_field_is_one_bar() {
        let first = field(250);
        let mut now = first.clone();
        draw(&mut now, 60, 61, 40, 73, 90);
        assert_eq!(changed_bars(&first, &now, W, H, 2.0), vec![Bar { x0: 60, x1: 61, y0: 40, y1: 73 }]);
    }

    #[test]
    fn a_light_caret_hidden_on_a_dark_field_is_one_bar() {
        let mut first = field(34);
        draw(&mut first, 100, 101, 20, 59, 233);
        let now = field(34);
        assert_eq!(changed_bars(&first, &now, W, H, 2.0), vec![Bar { x0: 100, x1: 101, y0: 20, y1: 59 }]);
    }

    #[test]
    fn a_wide_change_is_no_caret() {
        let first = field(250);
        let mut now = first.clone();
        // a glyph typed or a picture playing: wider than a caret
        draw(&mut now, 60, 75, 40, 73, 90);
        assert!(changed_bars(&first, &now, W, H, 2.0).is_empty());
    }

    #[test]
    fn a_bar_with_changes_beside_it_is_no_caret() {
        let first = field(250);
        let mut now = first.clone();
        draw(&mut now, 60, 61, 40, 73, 90);
        draw(&mut now, 64, 64, 40, 73, 200);
        assert!(changed_bars(&first, &now, W, H, 2.0).is_empty());
    }

    #[test]
    fn a_short_or_tall_bar_is_no_caret() {
        let first = field(250);
        let mut short = first.clone();
        draw(&mut short, 60, 61, 40, 50, 90);
        assert!(changed_bars(&first, &short, W, H, 2.0).is_empty());
        let mut tall = vec![250; W * 400];
        let first_tall = tall.clone();
        for y in 10..=200 {
            tall[y * W + 60] = 90;
        }
        assert!(changed_bars(&first_tall, &tall, W, 400, 2.0).is_empty());
    }

    #[test]
    fn an_anti_aliased_caret_with_a_faint_pixel_is_one_bar() {
        let first = field(250);
        let mut now = first.clone();
        draw(&mut now, 60, 60, 40, 73, 120);
        // one faint pixel down the bar (anti-aliasing)
        now[55 * W + 60] = 220;
        draw(&mut now, 61, 61, 40, 73, 190);
        assert_eq!(changed_bars(&first, &now, W, H, 2.0), vec![Bar { x0: 60, x1: 61, y0: 40, y1: 73 }]);
    }

    #[test]
    fn a_three_column_bar_stays_one_bar_beside_another_change() {
        let first = field(250);
        let mut now = first.clone();
        draw(&mut now, 60, 62, 40, 73, 90);
        // another bar further left that ends before the first one's columns
        draw(&mut now, 20, 20, 80, 115, 90);
        let bars = changed_bars(&first, &now, W, H, 2.0);
        assert!(bars.contains(&Bar { x0: 60, x1: 62, y0: 40, y1: 73 }), "{bars:?}");
        assert_eq!(bars.len(), 2, "{bars:?}");
    }

    #[test]
    fn the_bar_inside_the_field_or_the_most_frequent_is_chosen() {
        let a = Bar { x0: 10, x1: 11, y0: 10, y1: 40 };
        let b = Bar { x0: 100, x1: 101, y0: 60, y1: 90 };
        assert_eq!(choose(&[(a, 1)], |_| true), Some(a));
        assert_eq!(choose(&[(a, 1), (b, 3)], |_| true), Some(b));
        assert_eq!(choose(&[(a, 2), (b, 2)], |_| true), None);
        assert_eq!(choose(&[(a, 1), (b, 3)], |bar| bar.x0 < 50), Some(a));
    }
}
