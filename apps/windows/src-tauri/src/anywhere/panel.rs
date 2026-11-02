//! The formula's window: a transparent, borderless WebView (anywhere.html) over the field's line.
//! With the input method it never takes the keyboard (the field keeps it and the input method
//! forwards the keys) and lets clicks through to the field; for a field without the input method
//! it is a strip on the caret's line that takes the keyboard until the formula is pasted. It is
//! left out of screen captures only while the field's colours are read off the screen. Placed in
//! physical pixels.

use super::{PanelLayout, ScreenRect, PANEL_TITLE, WINDOW_LABEL};
use tauri::webview::WebviewWindowBuilder;
use tauri::{AppHandle, Manager, WebviewUrl, WebviewWindow};
use windows::Win32::Foundation::{HWND, POINT, RECT};
use windows::Win32::Graphics::Gdi::{GetMonitorInfoW, MonitorFromPoint, MONITORINFO, MONITOR_DEFAULTTONEAREST};
use windows::Win32::UI::HiDpi::{GetDpiForMonitor, MDT_EFFECTIVE_DPI};
use windows::Win32::UI::WindowsAndMessaging::{
    GetWindowLongPtrW, GetWindowRect, SetWindowDisplayAffinity, SetWindowLongPtrW, SetWindowPos, ShowWindow, GWL_EXSTYLE, HWND_TOPMOST,
    SWP_NOACTIVATE, SWP_NOZORDER, SW_HIDE, SW_SHOWNOACTIVATE, WDA_EXCLUDEFROMCAPTURE, WDA_NONE, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW,
};

/// Where the window stands. InLine: the formula drawn in the field's line, its start on the
/// anchor (the field's composition) and its baseline on the line's. BelowLine: under the caret's
/// line, before the field reports the composition's place. OnLine: the pasting strip, its middle
/// on the caret's line and its formula starting at the caret.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum Placement {
    InLine,
    BelowLine,
    OnLine,
}

const BELOW_LINE_GAP: f64 = 4.0;

pub(super) fn ensure(app: &AppHandle) -> Result<WebviewWindow, String> {
    if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
        return Ok(window);
    }
    let window = WebviewWindowBuilder::new(app, WINDOW_LABEL, WebviewUrl::App("anywhere.html".into()))
        .title(PANEL_TITLE)
        .inner_size(96.0, 44.0)
        .resizable(false)
        .decorations(false)
        .transparent(true)
        .shadow(false)
        .always_on_top(true)
        .focused(false)
        .focusable(false)
        .skip_taskbar(true)
        .visible(false)
        .background_throttling(tauri::utils::config::BackgroundThrottlingPolicy::Disabled)
        .build()
        .map_err(|error| format!("Unable to create the formula window: {error}"))?;
    let _ = window.set_background_color(Some(tauri::window::Color(0, 0, 0, 0)));
    let _ = window.set_ignore_cursor_events(true);
    set_passive(&window, true);
    let events = app.clone();
    window.on_window_event(move |event| {
        if let tauri::WindowEvent::Focused(false) = event {
            super::handle_focus_lost(&events);
        }
    });
    super::log("formula window created");
    Ok(window)
}

/// Passive: never activated (the input method's formula); otherwise it can take the keyboard (the
/// pasting strip).
fn set_passive(window: &WebviewWindow, passive: bool) {
    if let Some(hwnd) = hwnd(window) {
        unsafe {
            let style = GetWindowLongPtrW(hwnd, GWL_EXSTYLE) | WS_EX_TOOLWINDOW.0 as isize;
            let style = if passive { style | WS_EX_NOACTIVATE.0 as isize } else { style & !(WS_EX_NOACTIVATE.0 as isize) };
            SetWindowLongPtrW(hwnd, GWL_EXSTYLE, style);
        }
    }
    let _ = window.set_focusable(!passive);
    let _ = window.set_ignore_cursor_events(passive);
}

/// Left out of screen captures while the field's colours are read off the screen (so the reading
/// sees the field, not the formula); otherwise in them, like any window (screenshots, recordings).
pub(super) fn exclude_from_capture(window: &WebviewWindow, exclude: bool) {
    if let Some(hwnd) = hwnd(window) {
        let affinity = if exclude { WDA_EXCLUDEFROMCAPTURE } else { WDA_NONE };
        let _ = unsafe { SetWindowDisplayAffinity(hwnd, affinity) };
    }
}

fn hwnd(window: &WebviewWindow) -> Option<HWND> {
    window.hwnd().ok().map(|handle| HWND(handle.0 as *mut _))
}

/// The scale (1.0 at 96 DPI) and work area of the monitor at a point (physical pixels).
pub(super) fn monitor_at(x: f64, y: f64) -> (f64, RECT) {
    let monitor = unsafe { MonitorFromPoint(POINT { x: x as i32, y: y as i32 }, MONITOR_DEFAULTTONEAREST) };
    let mut info = MONITORINFO { cbSize: std::mem::size_of::<MONITORINFO>() as u32, ..Default::default() };
    let _ = unsafe { GetMonitorInfoW(monitor, &mut info) };
    let (mut dpi_x, mut dpi_y) = (96u32, 96u32);
    let _ = unsafe { GetDpiForMonitor(monitor, MDT_EFFECTIVE_DPI, &mut dpi_x, &mut dpi_y) };
    (dpi_x.max(48) as f64 / 96.0, info.rcWork)
}

/// The window's frame (physical pixels) for the page's layout (CSS pixels) at `anchor`
/// (physical: x at the formula's start; y the baseline for InLine, the caret's top otherwise).
fn frame_for(anchor: ScreenRect, layout: PanelLayout, placement: Placement) -> (i32, i32, i32, i32) {
    let (scale, work) = monitor_at(anchor.x, anchor.y);
    let width = layout.width * scale;
    let height = layout.height * scale;
    let mut left = anchor.x - layout.anchor_x * scale;
    let mut top = match placement {
        Placement::InLine => anchor.y - layout.anchor_y * scale,
        Placement::BelowLine => anchor.y + anchor.height + BELOW_LINE_GAP * scale - layout.anchor_y * scale,
        Placement::OnLine => anchor.y + anchor.height / 2.0 - layout.anchor_y * scale,
    };
    if placement != Placement::InLine {
        left = left.min(work.right as f64 - width).max(work.left as f64);
        if placement == Placement::BelowLine && top + height > work.bottom as f64 {
            top = anchor.y - height - BELOW_LINE_GAP * scale;
        }
        // On the screen, the caret's line above the candidates when the strip rises.
        top = top.min(work.bottom as f64 - height).max(work.top as f64);
    }
    (left.round() as i32, top.round() as i32, width.round() as i32, height.round() as i32)
}

fn place(window: &WebviewWindow, anchor: ScreenRect, layout: PanelLayout, placement: Placement, show: bool) {
    let Some(hwnd) = hwnd(window) else { return };
    let (x, y, width, height) = frame_for(anchor, layout, placement);
    // Raised when it shows; following the formula keeps its place in the stack, so the input
    // method's pinyin list (also topmost, raised as it shows) stays over it.
    let order = if show { SWP_NOACTIVATE } else { SWP_NOACTIVATE | SWP_NOZORDER };
    unsafe {
        let _ = SetWindowPos(hwnd, Some(HWND_TOPMOST), x, y, width, height, order);
        // Moved onto a monitor of another scale, the window's WM_DPICHANGED handling resizes it
        // to keep its old logical size (cutting the formula or its candidates): the frame again,
        // now that it is on that monitor.
        let mut placed = RECT::default();
        if GetWindowRect(hwnd, &mut placed).is_ok()
            && (placed.left, placed.top, placed.right - placed.left, placed.bottom - placed.top) != (x, y, width, height)
        {
            super::log(&format!(
                "formula window reframed after a scale change: {}x{} -> {width}x{height}",
                placed.right - placed.left,
                placed.bottom - placed.top
            ));
            let _ = SetWindowPos(hwnd, None, x, y, width, height, SWP_NOACTIVATE | SWP_NOZORDER);
        }
        if show {
            let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
        }
    }
}

/// Shows the window at the formula's place, without taking the keyboard.
pub(super) fn present(window: &WebviewWindow, anchor: ScreenRect, layout: PanelLayout, placement: Placement) {
    set_passive(window, true);
    place(window, anchor, layout, placement, true);
}

/// Shows the pasting strip on the caret's line with the keyboard (the shortcut that opened it
/// lets VisualTeX take the foreground).
pub(super) fn present_focused(window: &WebviewWindow, anchor: ScreenRect, layout: PanelLayout) {
    set_passive(window, false);
    place(window, anchor, layout, Placement::OnLine, true);
    if let Some(hwnd) = hwnd(window) {
        unsafe {
            let _ = windows::Win32::UI::WindowsAndMessaging::SetForegroundWindow(hwnd);
        }
    }
    let _ = window.set_focus();
}

/// Follows the formula's place or its new size.
pub(super) fn reframe(window: &WebviewWindow, anchor: ScreenRect, layout: PanelLayout, placement: Placement) {
    place(window, anchor, layout, placement, false);
}

pub(super) fn park(window: &WebviewWindow) {
    set_passive(window, true);
    if let Some(hwnd) = hwnd(window) {
        unsafe {
            let _ = ShowWindow(hwnd, SW_HIDE);
        }
    }
}
