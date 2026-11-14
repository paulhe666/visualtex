//! The user's clipboard around a WPS Presentation native equation: the add-in
//! copies WPS's own copy of the equation to paste it (module-wpp.js), so
//! VisualTeX puts back what the clipboard held before, as the PowerPoint path
//! does (commit_powerpoint_native) and VisualTeX for Windows'
//! PowerPointClipboard.Preserve.

use objc2::rc::Retained;
use objc2::runtime::ProtocolObject;
use objc2_app_kit::{NSPasteboard, NSPasteboardItem, NSPasteboardWriting};
use objc2_foundation::{NSArray, NSData, NSString};
use std::sync::mpsc;
use std::time::Duration;
use tauri::AppHandle;

struct Saved {
    items: Vec<Vec<(String, Vec<u8>)>>,
    change_count: isize,
}

fn on_main<T: Send + 'static>(app: &AppHandle, work: impl FnOnce() -> T + Send + 'static) -> Option<T> {
    let (sender, receiver) = mpsc::sync_channel(1);
    app.run_on_main_thread(move || {
        let _ = sender.send(work());
    })
    .ok()?;
    receiver.recv_timeout(Duration::from_secs(3)).ok()
}

fn snapshot() -> Saved {
    let pasteboard = NSPasteboard::generalPasteboard();
    let items = pasteboard
        .pasteboardItems()
        .map(|items| {
            items
                .iter()
                .map(|item| {
                    item.types()
                        .iter()
                        .filter_map(|kind| item.dataForType(&kind).map(|data| (kind.to_string(), data.to_vec())))
                        .collect()
                })
                .collect()
        })
        .unwrap_or_default();
    Saved { items, change_count: pasteboard.changeCount() }
}

/// Only when something wrote the clipboard since the snapshot.
fn restore(saved: Saved) {
    let pasteboard = NSPasteboard::generalPasteboard();
    if pasteboard.changeCount() == saved.change_count {
        return;
    }
    pasteboard.clearContents();
    let items: Vec<Retained<ProtocolObject<dyn NSPasteboardWriting>>> = saved
        .items
        .iter()
        .filter(|item| !item.is_empty())
        .map(|saved| {
            let item = NSPasteboardItem::new();
            for (kind, bytes) in saved {
                item.setData_forType(&NSData::with_bytes(bytes), &NSString::from_str(kind));
            }
            ProtocolObject::<dyn NSPasteboardWriting>::from_retained(item)
        })
        .collect();
    if !items.is_empty() {
        pasteboard.writeObjects(&NSArray::from_retained_slice(&items));
    }
}

/// Runs `work` and puts the clipboard back as it was before.
pub(crate) fn kept<T>(app: Option<&AppHandle>, work: impl FnOnce() -> T) -> T {
    let saved = app.and_then(|app| on_main(app, snapshot));
    let result = work();
    if let (Some(app), Some(saved)) = (app, saved) {
        let _ = on_main(app, move || restore(saved));
    }
    result
}
