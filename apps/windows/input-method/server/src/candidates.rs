//! The candidate window under the line: one row of numbered candidates, the highlighted one in
//! the accent colour, as the macOS input method draws it (CandidateWindow.swift there); and the
//! small "中" / "英" shown for a moment when Shift switches between Chinese and English. Both
//! live on this program's UI thread, never take the keyboard or the mouse, and are placed in
//! physical pixels (this process is per-monitor DPI aware).

use crate::rime::Candidate;
use std::cell::RefCell;
use std::sync::atomic::{AtomicIsize, Ordering};
use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{COLORREF, HWND, LPARAM, LRESULT, POINT, RECT, SIZE, WPARAM};
use windows::Win32::Graphics::Dwm::{
    DwmGetColorizationColor, DwmSetWindowAttribute, DWMWA_WINDOW_CORNER_PREFERENCE, DWMWCP_ROUND, DWMWCP_ROUNDSMALL,
};
use windows::Win32::Graphics::Gdi::{
    BeginPaint, BitBlt, CreateCompatibleBitmap, CreateCompatibleDC, CreateFontW, CreatePen, CreateSolidBrush, DeleteDC,
    DeleteObject, EndPaint, GetMonitorInfoW, GetStockObject, GetTextExtentPoint32W, InvalidateRect, MonitorFromPoint,
    RoundRect, SelectObject, SetBkMode, SetTextColor, TextOutW, CLEARTYPE_QUALITY, CLIP_DEFAULT_PRECIS,
    DEFAULT_CHARSET, FW_NORMAL, HDC, HFONT, HGDIOBJ, MONITORINFO, MONITOR_DEFAULTTONEAREST, NULL_BRUSH,
    OUT_DEFAULT_PRECIS, PAINTSTRUCT, PS_SOLID, SRCCOPY, TRANSPARENT,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Registry::{RegGetValueW, HKEY_CURRENT_USER, RRF_RT_REG_DWORD};
use windows::Win32::UI::HiDpi::{GetDpiForMonitor, MDT_EFFECTIVE_DPI};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, KillTimer, PostMessageW, RegisterClassExW, SetTimer, SetWindowPos, ShowWindow,
    CS_DROPSHADOW, HTTRANSPARENT, HWND_TOPMOST, MA_NOACTIVATE, SWP_NOACTIVATE, SW_HIDE, SW_SHOWNOACTIVATE, WM_APP,
    WM_MOUSEACTIVATE, WM_NCHITTEST, WM_PAINT, WM_TIMER, WNDCLASSEXW, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW,
    WS_EX_TOPMOST, WS_POPUP,
};

/// A rectangle in physical screen pixels: the caret's line (x at the caret).
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct Rect {
    pub x: i32,
    pub y: i32,
    pub width: i32,
    pub height: i32,
}

impl Rect {
    pub fn from_json(value: Option<&serde_json::Value>) -> Option<Rect> {
        let values: Vec<i64> = value?.as_array()?.iter().filter_map(|value| value.as_f64().map(|v| v.round() as i64)).collect();
        let &[x, y, width, height] = values.as_slice() else { return None };
        (height > 0 && height < 2000).then_some(Rect { x: x as i32, y: y as i32, width: width as i32, height: height as i32 })
    }
}

pub struct Show {
    pub owner: usize,
    pub candidates: Vec<Candidate>,
    pub labels: Vec<String>,
    pub highlighted: usize,
    pub caret: Rect,
}

pub struct Indicator {
    pub text: String,
    pub caret: Rect,
}

const WM_SHOW: u32 = WM_APP + 1;
const WM_HIDE: u32 = WM_APP + 2;
const WM_INDICATOR: u32 = WM_APP + 3;
const INDICATOR_TIMER: usize = 1;

static CANDIDATES: AtomicIsize = AtomicIsize::new(0);
static INDICATOR: AtomicIsize = AtomicIsize::new(0);

thread_local! {
    static MODEL: RefCell<Option<Model>> = const { RefCell::new(None) };
    static INDICATOR_TEXT: RefCell<(String, f32)> = const { RefCell::new((String::new(), 1.0)) };
}

/// What the candidate window shows, laid out.
struct Model {
    owner: usize,
    items: Vec<Item>,
    highlighted: usize,
    scale: f32,
    width: i32,
    height: i32,
}

struct Item {
    label: Vec<u16>,
    text: Vec<u16>,
    comment: Vec<u16>,
    left: i32,
    width: i32,
}

// MARK: Posting from the pipe threads

fn window(slot: &AtomicIsize) -> Option<HWND> {
    let raw = slot.load(Ordering::SeqCst);
    (raw != 0).then_some(HWND(raw as *mut _))
}

pub fn show(show: Show) {
    let Some(target) = window(&CANDIDATES) else { return };
    let message = Box::into_raw(Box::new(show));
    if unsafe { PostMessageW(Some(target), WM_SHOW, WPARAM(0), LPARAM(message as isize)) }.is_err() {
        drop(unsafe { Box::from_raw(message) });
    }
}

/// Hides the candidates (`owner` 0: whoever shows them).
pub fn hide(owner: usize) {
    if let Some(target) = window(&CANDIDATES) {
        let _ = unsafe { PostMessageW(Some(target), WM_HIDE, WPARAM(owner), LPARAM(0)) };
    }
}

pub fn indicate(indicator: Indicator) {
    let Some(target) = window(&INDICATOR) else { return };
    let message = Box::into_raw(Box::new(indicator));
    if unsafe { PostMessageW(Some(target), WM_INDICATOR, WPARAM(0), LPARAM(message as isize)) }.is_err() {
        drop(unsafe { Box::from_raw(message) });
    }
}

// MARK: Look

struct Palette {
    background: COLORREF,
    border: COLORREF,
    text: COLORREF,
    label: COLORREF,
    comment: COLORREF,
    accent: COLORREF,
    on_accent: COLORREF,
}

fn rgb(r: u8, g: u8, b: u8) -> COLORREF {
    COLORREF(r as u32 | (g as u32) << 8 | (b as u32) << 16)
}

fn dark_mode() -> bool {
    let mut value = 1u32;
    let mut size = std::mem::size_of::<u32>() as u32;
    let status = unsafe {
        RegGetValueW(
            HKEY_CURRENT_USER,
            w!("Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize"),
            w!("AppsUseLightTheme"),
            RRF_RT_REG_DWORD,
            None,
            Some((&mut value as *mut u32).cast()),
            Some(&mut size),
        )
    };
    status.is_ok() && value == 0
}

fn accent() -> COLORREF {
    let mut color = 0u32;
    let mut opaque = windows::core::BOOL(0);
    if unsafe { DwmGetColorizationColor(&mut color, &mut opaque) }.is_ok() && color != 0 {
        let (r, g, b) = ((color >> 16) as u8, (color >> 8) as u8, color as u8);
        // Too light an accent would not carry white text.
        if 0.2126 * r as f32 + 0.7152 * g as f32 + 0.0722 * b as f32 <= 170.0 {
            return rgb(r, g, b);
        }
    }
    rgb(0x00, 0x67, 0xc0)
}

fn palette() -> Palette {
    if dark_mode() {
        Palette {
            background: rgb(0x2c, 0x2c, 0x2c),
            border: rgb(0x48, 0x48, 0x48),
            text: rgb(0xf2, 0xf2, 0xf2),
            label: rgb(0xa8, 0xa8, 0xa8),
            comment: rgb(0x8c, 0x8c, 0x8c),
            accent: accent(),
            on_accent: rgb(0xff, 0xff, 0xff),
        }
    } else {
        Palette {
            background: rgb(0xff, 0xff, 0xff),
            border: rgb(0xd6, 0xd6, 0xd6),
            text: rgb(0x1f, 0x1f, 0x1f),
            label: rgb(0x70, 0x70, 0x70),
            comment: rgb(0x8a, 0x8a, 0x8a),
            accent: accent(),
            on_accent: rgb(0xff, 0xff, 0xff),
        }
    }
}

fn font(pixels: f32) -> HFONT {
    font_of_weight(pixels, FW_NORMAL.0 as i32)
}

fn font_of_weight(pixels: f32, weight: i32) -> HFONT {
    unsafe {
        CreateFontW(
            -(pixels.round() as i32),
            0,
            0,
            0,
            weight,
            0,
            0,
            0,
            DEFAULT_CHARSET,
            OUT_DEFAULT_PRECIS,
            CLIP_DEFAULT_PRECIS,
            CLEARTYPE_QUALITY,
            0,
            w!("Microsoft YaHei UI"),
        )
    }
}

struct Fonts {
    text: HFONT,
    label: HFONT,
    comment: HFONT,
}

impl Fonts {
    fn new(scale: f32) -> Fonts {
        Fonts { text: font(18.0 * scale), label: font(12.5 * scale), comment: font(12.5 * scale) }
    }
}

impl Drop for Fonts {
    fn drop(&mut self) {
        unsafe {
            let _ = DeleteObject(self.text.into());
            let _ = DeleteObject(self.label.into());
            let _ = DeleteObject(self.comment.into());
        }
    }
}

fn measure(dc: HDC, font: HFONT, text: &[u16]) -> SIZE {
    let mut size = SIZE::default();
    if text.is_empty() {
        return size;
    }
    unsafe {
        let previous = SelectObject(dc, font.into());
        let _ = GetTextExtentPoint32W(dc, text, &mut size);
        SelectObject(dc, previous);
    }
    size
}

/// The monitor's scale (1.0 at 96 DPI) and work area at a point.
fn monitor_at(x: i32, y: i32) -> (f32, RECT) {
    let monitor = unsafe { MonitorFromPoint(POINT { x, y }, MONITOR_DEFAULTTONEAREST) };
    let mut info = MONITORINFO { cbSize: std::mem::size_of::<MONITORINFO>() as u32, ..Default::default() };
    let _ = unsafe { GetMonitorInfoW(monitor, &mut info) };
    let (mut dpi_x, mut dpi_y) = (96u32, 96u32);
    let _ = unsafe { GetDpiForMonitor(monitor, MDT_EFFECTIVE_DPI, &mut dpi_x, &mut dpi_y) };
    (dpi_x.max(96) as f32 / 96.0, info.rcWork)
}

// Layout, in pixels at scale 1.
const INSET_X: f32 = 5.0;
const INSET_Y: f32 = 4.0;
const ITEM_HEIGHT: f32 = 28.0;
const PADDING: f32 = 7.0;
const GAP: f32 = 2.0;
const SPACING: f32 = 4.0;

fn layout(show: &Show, scale: f32) -> Model {
    let fonts = Fonts::new(scale);
    let dc = unsafe { CreateCompatibleDC(None) };
    let mut x = (INSET_X * scale).round() as i32;
    let mut items = Vec::with_capacity(show.candidates.len());
    for (index, candidate) in show.candidates.iter().enumerate() {
        let label = show.labels.get(index).cloned().unwrap_or_else(|| ((index + 1) % 10).to_string());
        let label: Vec<u16> = label.encode_utf16().collect();
        let text: Vec<u16> = candidate.text.encode_utf16().collect();
        let comment: Vec<u16> = candidate.comment.encode_utf16().collect();
        let mut width = PADDING * scale
            + measure(dc, fonts.label, &label).cx as f32
            + SPACING * scale
            + measure(dc, fonts.text, &text).cx as f32
            + PADDING * scale;
        if !comment.is_empty() {
            width += SPACING * scale + measure(dc, fonts.comment, &comment).cx as f32;
        }
        let width = width.ceil() as i32;
        items.push(Item { label, text, comment, left: x, width });
        x += width + (GAP * scale).round() as i32;
    }
    unsafe {
        let _ = DeleteDC(dc);
    }
    let width = (x - (GAP * scale).round() as i32 + (INSET_X * scale).round() as i32).max((40.0 * scale) as i32);
    let height = ((INSET_Y * 2.0 + ITEM_HEIGHT) * scale).round() as i32;
    Model { owner: show.owner, items, highlighted: show.highlighted, scale, width, height }
}

/// The window's place under the caret's line, above it when there is no room below.
fn place(caret: Rect, width: i32, height: i32, scale: f32, work: RECT) -> (i32, i32) {
    let gap = (3.0 * scale).round() as i32;
    let mut x = caret.x - (6.0 * scale).round() as i32;
    let mut y = caret.y + caret.height + gap;
    if y + height > work.bottom {
        y = caret.y - height - gap;
    }
    x = x.min(work.right - width).max(work.left);
    y = y.max(work.top);
    (x, y)
}

fn paint_candidates(dc: HDC, model: &Model) {
    let colors = palette();
    let fonts = Fonts::new(model.scale);
    unsafe {
        let background = CreateSolidBrush(colors.background);
        let border = CreatePen(PS_SOLID, 1, colors.border);
        let previous_brush = SelectObject(dc, background.into());
        let previous_pen = SelectObject(dc, border.into());
        let _ = windows::Win32::Graphics::Gdi::Rectangle(dc, 0, 0, model.width, model.height);
        SelectObject(dc, previous_pen);
        SelectObject(dc, previous_brush);
        let _ = DeleteObject(background.into());
        let _ = DeleteObject(border.into());
        SetBkMode(dc, TRANSPARENT);
        let top = (INSET_Y * model.scale).round() as i32;
        let item_height = (ITEM_HEIGHT * model.scale).round() as i32;
        let middle = top + item_height / 2;
        for (index, item) in model.items.iter().enumerate() {
            let highlighted = index == model.highlighted;
            if highlighted {
                let brush = CreateSolidBrush(colors.accent);
                let pen = CreatePen(PS_SOLID, 1, colors.accent);
                let previous_brush = SelectObject(dc, brush.into());
                let previous_pen = SelectObject(dc, pen.into());
                let radius = (6.0 * model.scale).round() as i32;
                let _ = RoundRect(dc, item.left, top, item.left + item.width, top + item_height, radius, radius);
                SelectObject(dc, previous_pen);
                SelectObject(dc, previous_brush);
                let _ = DeleteObject(brush.into());
                let _ = DeleteObject(pen.into());
            }
            let mut x = item.left + (PADDING * model.scale).round() as i32;
            for (font, text, color) in [
                (fonts.label, &item.label, if highlighted { colors.on_accent } else { colors.label }),
                (fonts.text, &item.text, if highlighted { colors.on_accent } else { colors.text }),
                (fonts.comment, &item.comment, if highlighted { colors.on_accent } else { colors.comment }),
            ] {
                if text.is_empty() {
                    continue;
                }
                let size = measure(dc, font, text);
                let previous = SelectObject(dc, font.into());
                SetTextColor(dc, color);
                let _ = TextOutW(dc, x, middle - size.cy / 2, text);
                SelectObject(dc, previous);
                x += size.cx + (SPACING * model.scale).round() as i32;
            }
        }
        let _ = GetStockObject(NULL_BRUSH);
    }
}

fn paint_indicator(dc: HDC, text: &[u16], scale: f32, width: i32, height: i32) {
    let colors = palette();
    unsafe {
        let background = CreateSolidBrush(colors.background);
        let border = CreatePen(PS_SOLID, 1, colors.border);
        let previous_brush = SelectObject(dc, background.into());
        let previous_pen = SelectObject(dc, border.into());
        let _ = windows::Win32::Graphics::Gdi::Rectangle(dc, 0, 0, width, height);
        SelectObject(dc, previous_pen);
        SelectObject(dc, previous_brush);
        let _ = DeleteObject(background.into());
        let _ = DeleteObject(border.into());
        // the macOS indicator's label: 14 pt, medium (input-method/Sources/ModeIndicator.swift)
        let font = font_of_weight(14.0 * scale, 500);
        let size = measure(dc, font, text);
        let previous = SelectObject(dc, font.into());
        SetBkMode(dc, TRANSPARENT);
        SetTextColor(dc, colors.text);
        let _ = TextOutW(dc, (width - size.cx) / 2, (height - size.cy) / 2, text);
        SelectObject(dc, previous);
        let _ = DeleteObject(font.into());
    }
}

/// Paints through an off-screen bitmap (no flicker).
fn buffered_paint(window: HWND, draw: impl FnOnce(HDC, i32, i32)) {
    unsafe {
        let mut paint = PAINTSTRUCT::default();
        let dc = BeginPaint(window, &mut paint);
        let mut rect = RECT::default();
        let _ = windows::Win32::UI::WindowsAndMessaging::GetClientRect(window, &mut rect);
        let (width, height) = (rect.right - rect.left, rect.bottom - rect.top);
        let memory = CreateCompatibleDC(Some(dc));
        let bitmap = CreateCompatibleBitmap(dc, width.max(1), height.max(1));
        let previous = SelectObject(memory, bitmap.into());
        draw(memory, width, height);
        let _ = BitBlt(dc, 0, 0, width, height, Some(memory), 0, 0, SRCCOPY);
        SelectObject(memory, previous);
        let _ = DeleteObject(HGDIOBJ(bitmap.0));
        let _ = DeleteDC(memory);
        let _ = EndPaint(window, &paint);
    }
}

fn move_to(window: HWND, x: i32, y: i32, width: i32, height: i32) {
    unsafe {
        let _ = SetWindowPos(window, Some(HWND_TOPMOST), x, y, width, height, SWP_NOACTIVATE);
        let _ = InvalidateRect(Some(window), None, false);
        let _ = ShowWindow(window, SW_SHOWNOACTIVATE);
    }
}

extern "system" fn candidates_proc(window: HWND, message: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match message {
        WM_SHOW => {
            let show = unsafe { Box::from_raw(lparam.0 as *mut Show) };
            if show.candidates.is_empty() {
                MODEL.with(|model| model.borrow_mut().take());
                unsafe {
                    let _ = ShowWindow(window, SW_HIDE);
                }
                return LRESULT(0);
            }
            let (scale, work) = monitor_at(show.caret.x, show.caret.y + show.caret.height);
            let model = layout(&show, scale);
            let (x, y) = place(show.caret, model.width, model.height, scale, work);
            let (width, height) = (model.width, model.height);
            MODEL.with(|slot| *slot.borrow_mut() = Some(model));
            move_to(window, x, y, width, height);
            LRESULT(0)
        }
        WM_HIDE => {
            let owner = wparam.0;
            let hidden = MODEL.with(|model| {
                let mut model = model.borrow_mut();
                let matches = owner == 0 || model.as_ref().map(|model| model.owner == owner).unwrap_or(true);
                if matches {
                    model.take();
                }
                matches
            });
            if hidden {
                unsafe {
                    let _ = ShowWindow(window, SW_HIDE);
                }
            }
            LRESULT(0)
        }
        WM_PAINT => {
            buffered_paint(window, |dc, _, _| {
                MODEL.with(|model| {
                    if let Some(model) = model.borrow().as_ref() {
                        paint_candidates(dc, model);
                    }
                })
            });
            LRESULT(0)
        }
        WM_NCHITTEST => LRESULT(HTTRANSPARENT as isize),
        WM_MOUSEACTIVATE => LRESULT(MA_NOACTIVATE as isize),
        _ => unsafe { DefWindowProcW(window, message, wparam, lparam) },
    }
}

extern "system" fn indicator_proc(window: HWND, message: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match message {
        WM_INDICATOR => {
            let indicator = unsafe { Box::from_raw(lparam.0 as *mut Indicator) };
            let (scale, work) = monitor_at(indicator.caret.x, indicator.caret.y + indicator.caret.height);
            // as on macOS: 30 x 26, 4 pixels under the caret (the formula's caret, on its line),
            // its left edge 4 pixels left of it; over the caret at the screen's bottom
            let (width, height) = ((30.0 * scale).round() as i32, (26.0 * scale).round() as i32);
            let gap = (4.0 * scale).round() as i32;
            let mut x = indicator.caret.x - gap;
            let mut y = indicator.caret.y + indicator.caret.height + gap;
            if y + height > work.bottom {
                y = indicator.caret.y - height - gap;
            }
            x = x.min(work.right - width).max(work.left);
            y = y.max(work.top);
            INDICATOR_TEXT.with(|slot| *slot.borrow_mut() = (indicator.text.clone(), scale));
            move_to(window, x, y, width, height);
            unsafe {
                SetTimer(Some(window), INDICATOR_TIMER, 800, None);
            }
            LRESULT(0)
        }
        WM_TIMER if wparam.0 == INDICATOR_TIMER => {
            unsafe {
                let _ = KillTimer(Some(window), INDICATOR_TIMER);
                let _ = ShowWindow(window, SW_HIDE);
            }
            LRESULT(0)
        }
        WM_PAINT => {
            buffered_paint(window, |dc, width, height| {
                INDICATOR_TEXT.with(|slot| {
                    let (text, scale) = &*slot.borrow();
                    let text: Vec<u16> = text.encode_utf16().collect();
                    paint_indicator(dc, &text, *scale, width, height);
                })
            });
            LRESULT(0)
        }
        WM_NCHITTEST => LRESULT(HTTRANSPARENT as isize),
        WM_MOUSEACTIVATE => LRESULT(MA_NOACTIVATE as isize),
        _ => unsafe { DefWindowProcW(window, message, wparam, lparam) },
    }
}

fn create_popup(class: PCWSTR, procedure: windows::Win32::UI::WindowsAndMessaging::WNDPROC, small_corners: bool) -> Option<HWND> {
    let instance = unsafe { GetModuleHandleW(None) }.ok()?;
    let class_info = WNDCLASSEXW {
        cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
        style: CS_DROPSHADOW,
        lpfnWndProc: procedure,
        hInstance: instance.into(),
        lpszClassName: class,
        ..Default::default()
    };
    unsafe { RegisterClassExW(&class_info) };
    let window = unsafe {
        CreateWindowExW(
            WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
            class,
            w!(""),
            WS_POPUP,
            0,
            0,
            10,
            10,
            None,
            None,
            Some(instance.into()),
            None,
        )
    }
    .ok()?;
    // Windows 11 rounds the corners (Windows 10 ignores this).
    let corner = if small_corners { DWMWCP_ROUNDSMALL } else { DWMWCP_ROUND };
    unsafe {
        let _ = DwmSetWindowAttribute(
            window,
            DWMWA_WINDOW_CORNER_PREFERENCE,
            (&corner as *const windows::Win32::Graphics::Dwm::DWM_WINDOW_CORNER_PREFERENCE).cast(),
            std::mem::size_of_val(&corner) as u32,
        );
    }
    Some(window)
}

/// Creates both windows (hidden) on the calling thread, which runs the message loop.
pub fn create() {
    if let Some(window) = create_popup(w!("VisualTeXInputCandidates"), Some(candidates_proc), false) {
        CANDIDATES.store(window.0 as isize, Ordering::SeqCst);
    }
    if let Some(window) = create_popup(w!("VisualTeXInputIndicator"), Some(indicator_proc), true) {
        INDICATOR.store(window.0 as isize, Ordering::SeqCst);
    }
}
