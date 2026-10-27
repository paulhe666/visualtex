//! What does not go through the input method: the Image output mode (the PNG on the clipboard,
//! then Ctrl+V into the field) and Ctrl+V in an open formula (the clipboard's image for OCR, or
//! its text).

use std::time::Duration;
use windows::core::w;
use windows::Win32::Foundation::{GlobalFree, HANDLE, HGLOBAL, HWND};
use windows::Win32::System::DataExchange::{
    CloseClipboard, EmptyClipboard, GetClipboardData, IsClipboardFormatAvailable, OpenClipboard, RegisterClipboardFormatW,
    SetClipboardData,
};
use windows::Win32::System::Memory::{GlobalAlloc, GlobalLock, GlobalSize, GlobalUnlock, GMEM_MOVEABLE};
use windows::Win32::System::Ole::{CF_DIB, CF_HDROP, CF_UNICODETEXT};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetAsyncKeyState, SendInput, INPUT, INPUT_0, INPUT_KEYBOARD, KEYBDINPUT, KEYEVENTF_KEYUP, VIRTUAL_KEY, VK_CONTROL,
    VK_LWIN, VK_MENU, VK_RWIN, VK_SHIFT, VK_V,
};
use windows::Win32::UI::Shell::{DragQueryFileW, HDROP};

const OPEN_RETRIES: usize = 25;

struct Clipboard;

impl Clipboard {
    fn open() -> Result<Clipboard, String> {
        for _ in 0..OPEN_RETRIES {
            if unsafe { OpenClipboard(None) }.is_ok() {
                return Ok(Clipboard);
            }
            std::thread::sleep(Duration::from_millis(20));
        }
        Err("The clipboard is in use by another program".to_string())
    }
}

impl Drop for Clipboard {
    fn drop(&mut self) {
        unsafe {
            let _ = CloseClipboard();
        }
    }
}

fn global_with(bytes: &[u8]) -> Result<HGLOBAL, String> {
    unsafe {
        let memory = GlobalAlloc(GMEM_MOVEABLE, bytes.len().max(1)).map_err(|error| error.to_string())?;
        let target = GlobalLock(memory);
        if target.is_null() {
            let _ = GlobalFree(Some(memory));
            return Err("Clipboard memory is unavailable".to_string());
        }
        std::ptr::copy_nonoverlapping(bytes.as_ptr(), target.cast::<u8>(), bytes.len());
        let _ = GlobalUnlock(memory);
        Ok(memory)
    }
}

fn set_data(format: u32, bytes: &[u8]) -> Result<(), String> {
    let memory = global_with(bytes)?;
    unsafe {
        if SetClipboardData(format, Some(HANDLE(memory.0))).is_err() {
            let _ = GlobalFree(Some(memory));
            return Err("The clipboard did not take the image".to_string());
        }
    }
    Ok(())
}

/// A PNG as a bottom-up 32-bit DIB (BITMAPINFOHEADER + pixels), for programs that read CF_DIB.
fn png_to_dib(png_bytes: &[u8]) -> Option<Vec<u8>> {
    let mut decoder = png::Decoder::new(std::io::Cursor::new(png_bytes));
    decoder.set_transformations(png::Transformations::EXPAND | png::Transformations::STRIP_16);
    let mut reader = decoder.read_info().ok()?;
    let mut buffer = vec![0u8; reader.output_buffer_size()];
    let frame = reader.next_frame(&mut buffer).ok()?;
    let (width, height) = (frame.width as usize, frame.height as usize);
    let channels = match frame.color_type {
        png::ColorType::Rgba => 4,
        png::ColorType::Rgb => 3,
        png::ColorType::GrayscaleAlpha => 2,
        png::ColorType::Grayscale => 1,
        _ => return None,
    };
    let mut dib = Vec::with_capacity(40 + width * height * 4);
    dib.extend_from_slice(&40u32.to_le_bytes());
    dib.extend_from_slice(&(width as i32).to_le_bytes());
    dib.extend_from_slice(&(height as i32).to_le_bytes());
    dib.extend_from_slice(&1u16.to_le_bytes());
    dib.extend_from_slice(&32u16.to_le_bytes());
    dib.extend_from_slice(&0u32.to_le_bytes()); // BI_RGB
    dib.extend_from_slice(&((width * height * 4) as u32).to_le_bytes());
    dib.extend_from_slice(&[0u8; 16]);
    for y in (0..height).rev() {
        let row = &buffer[y * frame.line_size..y * frame.line_size + width * channels];
        for x in 0..width {
            let pixel = &row[x * channels..x * channels + channels];
            let (r, g, b, a) = match channels {
                4 => (pixel[0], pixel[1], pixel[2], pixel[3]),
                3 => (pixel[0], pixel[1], pixel[2], 255),
                2 => (pixel[0], pixel[0], pixel[0], pixel[1]),
                _ => (pixel[0], pixel[0], pixel[0], 255),
            };
            // Programs that ignore alpha see the image over white.
            let blend = |channel: u8| ((channel as u32 * a as u32 + 255 * (255 - a as u32)) / 255) as u8;
            dib.extend_from_slice(&[blend(b), blend(g), blend(r), 255]);
        }
    }
    Some(dib)
}

/// The keys the user still holds (the commit's Enter chord) would join Ctrl+V: wait them out.
fn wait_for_released_modifiers() {
    for _ in 0..50 {
        let held = [VK_SHIFT, VK_MENU, VK_LWIN, VK_RWIN, VK_CONTROL]
            .iter()
            .any(|key| unsafe { GetAsyncKeyState(key.0 as i32) } < 0);
        if !held {
            return;
        }
        std::thread::sleep(Duration::from_millis(20));
    }
}

fn key(vk: VIRTUAL_KEY, up: bool) -> INPUT {
    INPUT {
        r#type: INPUT_KEYBOARD,
        Anonymous: INPUT_0 {
            ki: KEYBDINPUT { wVk: vk, dwFlags: if up { KEYEVENTF_KEYUP } else { Default::default() }, ..Default::default() },
        },
    }
}

/// Puts the PNG on the clipboard (PNG and DIB) and pastes it into the focused field.
pub(super) fn paste_png(png: Vec<u8>) -> Result<(), String> {
    {
        let _clipboard = Clipboard::open()?;
        unsafe { EmptyClipboard().map_err(|error| error.to_string())? };
        let png_format = unsafe { RegisterClipboardFormatW(w!("PNG")) };
        set_data(png_format, &png)?;
        if let Some(dib) = png_to_dib(&png) {
            set_data(CF_DIB.0 as u32, &dib)?;
        }
    }
    wait_for_released_modifiers();
    std::thread::sleep(Duration::from_millis(60));
    let inputs = [key(VK_CONTROL, false), key(VK_V, false), key(VK_V, true), key(VK_CONTROL, true)];
    let sent = unsafe { SendInput(&inputs, std::mem::size_of::<INPUT>() as i32) };
    if sent as usize != inputs.len() {
        return Err("Windows did not accept Ctrl+V".to_string());
    }
    Ok(())
}

pub(super) enum ClipboardContent {
    Image { bytes: Vec<u8>, extension: String },
    Text(String),
    Nothing,
}

fn data(format: u32) -> Option<Vec<u8>> {
    unsafe {
        let handle = GetClipboardData(format).ok()?;
        let memory = HGLOBAL(handle.0);
        let size = GlobalSize(memory);
        let pointer = GlobalLock(memory);
        if pointer.is_null() {
            return None;
        }
        let bytes = std::slice::from_raw_parts(pointer as *const u8, size).to_vec();
        let _ = GlobalUnlock(memory);
        Some(bytes)
    }
}

/// A DIB from the clipboard as a PNG.
fn dib_to_png(dib: &[u8]) -> Option<Vec<u8>> {
    if dib.len() < 40 {
        return None;
    }
    let read_i32 = |at: usize| i32::from_le_bytes(dib[at..at + 4].try_into().unwrap());
    let read_u16 = |at: usize| u16::from_le_bytes(dib[at..at + 2].try_into().unwrap());
    let header = read_i32(0) as usize;
    let width = read_i32(4);
    let height = read_i32(8);
    let bits = read_u16(14);
    let compression = read_i32(16);
    if width <= 0 || height == 0 || !(bits == 24 || bits == 32) || !(compression == 0 || compression == 3) {
        return None;
    }
    let (width, rows) = (width as usize, height.unsigned_abs() as usize);
    let stride = (width * bits as usize / 8 + 3) & !3;
    let offset = header + if compression == 3 && header == 40 { 12 } else { 0 };
    if dib.len() < offset + stride * rows {
        return None;
    }
    let mut rgb = Vec::with_capacity(width * rows * 3);
    for row in 0..rows {
        let source_row = if height > 0 { rows - 1 - row } else { row };
        let line = &dib[offset + source_row * stride..];
        for x in 0..width {
            let pixel = &line[x * bits as usize / 8..];
            rgb.extend_from_slice(&[pixel[2], pixel[1], pixel[0]]);
        }
    }
    let mut output = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut output, width as u32, rows as u32);
        encoder.set_color(png::ColorType::Rgb);
        encoder.set_depth(png::BitDepth::Eight);
        let mut writer = encoder.write_header().ok()?;
        writer.write_image_data(&rgb).ok()?;
    }
    Some(output)
}

/// The first copied file, when it is an image.
fn copied_image_file() -> Option<(Vec<u8>, String)> {
    let handle = unsafe { GetClipboardData(CF_HDROP.0 as u32) }.ok()?;
    let drop = HDROP(handle.0);
    let mut buffer = vec![0u16; 2048];
    let length = unsafe { DragQueryFileW(drop, 0, Some(&mut buffer)) } as usize;
    if length == 0 {
        return None;
    }
    let path = std::path::PathBuf::from(String::from_utf16_lossy(&buffer[..length]));
    let extension = path.extension()?.to_string_lossy().to_ascii_lowercase();
    if !["png", "jpg", "jpeg", "bmp", "gif", "webp", "tif", "tiff"].contains(&extension.as_str()) {
        return None;
    }
    let bytes = std::fs::read(&path).ok()?;
    Some((bytes, if extension == "jpeg" { "jpg".to_string() } else { extension }))
}

pub(super) fn read_clipboard() -> Result<ClipboardContent, String> {
    let _clipboard = Clipboard::open()?;
    let png_format = unsafe { RegisterClipboardFormatW(w!("PNG")) };
    if unsafe { IsClipboardFormatAvailable(png_format) }.is_ok() {
        if let Some(bytes) = data(png_format) {
            return Ok(ClipboardContent::Image { bytes, extension: "png".to_string() });
        }
    }
    if unsafe { IsClipboardFormatAvailable(CF_DIB.0 as u32) }.is_ok() {
        if let Some(bytes) = data(CF_DIB.0 as u32).and_then(|dib| dib_to_png(&dib)) {
            return Ok(ClipboardContent::Image { bytes, extension: "png".to_string() });
        }
    }
    if unsafe { IsClipboardFormatAvailable(CF_HDROP.0 as u32) }.is_ok() {
        if let Some((bytes, extension)) = copied_image_file() {
            return Ok(ClipboardContent::Image { bytes, extension });
        }
    }
    if unsafe { IsClipboardFormatAvailable(CF_UNICODETEXT.0 as u32) }.is_ok() {
        if let Some(bytes) = data(CF_UNICODETEXT.0 as u32) {
            let units: Vec<u16> = bytes.chunks_exact(2).map(|pair| u16::from_le_bytes([pair[0], pair[1]])).collect();
            let end = units.iter().position(|unit| *unit == 0).unwrap_or(units.len());
            return Ok(ClipboardContent::Text(String::from_utf16_lossy(&units[..end])));
        }
    }
    Ok(ClipboardContent::Nothing)
}

/// Brings `window` (the field's) back to the foreground after the screenshot selection, which
/// takes it for a moment.
pub(super) fn restore_foreground(window: isize) {
    use windows::Win32::System::Threading::{AttachThreadInput, GetCurrentThreadId};
    use windows::Win32::UI::WindowsAndMessaging::{GetForegroundWindow, GetWindowThreadProcessId, SetForegroundWindow};
    if window == 0 {
        return;
    }
    let target = HWND(window as *mut _);
    for _ in 0..10 {
        let foreground = unsafe { GetForegroundWindow() };
        if foreground == target {
            return;
        }
        unsafe {
            let current = GetCurrentThreadId();
            let foreground_thread = GetWindowThreadProcessId(foreground, None);
            let attached = foreground_thread != 0 && AttachThreadInput(current, foreground_thread, true).as_bool();
            let _ = SetForegroundWindow(target);
            if attached {
                let _ = AttachThreadInput(current, foreground_thread, false);
            }
        }
        std::thread::sleep(Duration::from_millis(60));
    }
}
