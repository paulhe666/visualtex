//! The pasting strip's delivery on Linux: the formula on the clipboard, then Ctrl+V into the field
//! (xdotool, X11); and the clipboard's image or text for a formula (Ctrl+V in it).

use gtk::prelude::*;
use std::process::Command;
use std::time::Duration;
use tauri::AppHandle;

fn on_main<T: Send + 'static>(app: &AppHandle, work: impl FnOnce() -> T + Send + 'static) -> Result<T, String> {
    let (sender, receiver) = std::sync::mpsc::sync_channel(1);
    app.run_on_main_thread(move || {
        let _ = sender.send(work());
    })
    .map_err(|error| format!("Unable to reach the main thread: {error}"))?;
    receiver
        .recv_timeout(Duration::from_secs(5))
        .map_err(|_| "The clipboard did not answer".to_string())
}

pub(super) fn copy(app: &AppHandle, text: String) -> Result<(), String> {
    on_main(app, move || {
        let clipboard = gtk::Clipboard::get(&gdk::SELECTION_CLIPBOARD);
        clipboard.set_text(&text);
        clipboard.store();
    })
}

/// Ctrl+V into the window with the keyboard (the field the formula was opened in).
fn press_paste() -> Result<(), String> {
    std::thread::sleep(Duration::from_millis(60));
    let status = Command::new("xdotool")
        .args(["key", "--clearmodifiers", "ctrl+v"])
        .status()
        .map_err(|_| "The formula is on the clipboard: paste it with Ctrl+V (xdotool is not installed)".to_string())?;
    if status.success() {
        Ok(())
    } else {
        Err("The formula is on the clipboard: paste it with Ctrl+V".to_string())
    }
}

pub(super) fn paste(app: &AppHandle, text: String) -> Result<(), String> {
    copy(app, text)?;
    press_paste()
}

pub(super) fn paste_png(app: &AppHandle, png: Vec<u8>) -> Result<(), String> {
    on_main(app, move || -> Result<(), String> {
        let loader = gdk_pixbuf::PixbufLoader::new();
        loader.write(&png).map_err(|error| error.to_string())?;
        loader.close().map_err(|error| error.to_string())?;
        let pixbuf = loader.pixbuf().ok_or("The formula image could not be read")?;
        let clipboard = gtk::Clipboard::get(&gdk::SELECTION_CLIPBOARD);
        clipboard.set_image(&pixbuf);
        clipboard.store();
        Ok(())
    })??;
    press_paste()
}

pub(super) enum ClipboardContent {
    Image { bytes: Vec<u8>, extension: String },
    Text(String),
    Nothing,
}

pub(super) fn read_clipboard(app: &AppHandle) -> Result<ClipboardContent, String> {
    on_main(app, || {
        let clipboard = gtk::Clipboard::get(&gdk::SELECTION_CLIPBOARD);
        if let Some(image) = clipboard.wait_for_image() {
            if let Ok(bytes) = image.save_to_bufferv("png", &[]) {
                return ClipboardContent::Image { bytes, extension: "png".to_string() };
            }
        }
        match clipboard.wait_for_text() {
            Some(text) if !text.is_empty() => ClipboardContent::Text(text.to_string()),
            _ => ClipboardContent::Nothing,
        }
    })
}
