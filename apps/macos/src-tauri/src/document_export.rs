// Documents to PDF (docs/document-editor/DESIGN.md §12) and the files a document is read from.
// - pdf_print_open / pdf_print_save / pdf_print_close: the print window (print.html) gets the
//   print copy of the page view from the main window; WKWebView's print operation writes it
//   to the file (runOperationModalForWindow; runOperation gives blank pages).
// - read_import_file: a document file chosen in the open dialog (its folder then finds its
//   pictures, read with read_image_file).
// - fetch_image_url: a picture a Markdown file names on the web.

use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};
use std::{fs, path::PathBuf};
use tauri::{AppHandle, Manager, WebviewUrl, WebviewWindowBuilder};

const PRINT_WINDOW: &str = "pdf-print";

/// A document file to open (the open dialog's choice), as base64.
#[tauri::command]
pub fn read_import_file(path: String) -> Result<String, String> {
    let target = PathBuf::from(path.trim());
    let extension = target
        .extension()
        .and_then(|value| value.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    if !target.is_absolute() || !matches!(extension.as_str(), "vtxdoc" | "json" | "md" | "markdown" | "tex" | "ltx" | "zip") {
        return Err("Not a document file".to_string());
    }
    let size = fs::metadata(&target).map_err(|error| format!("Unable to read the file: {error}"))?.len();
    if size > 300 * 1024 * 1024 {
        return Err("The file is too large".to_string());
    }
    let bytes = fs::read(&target).map_err(|error| format!("Unable to read the file: {error}"))?;
    Ok(BASE64_STANDARD.encode(bytes))
}

/// A picture on the web (http/https, at most 30 MB, an image by its type or its first bytes), as base64.
#[tauri::command]
pub async fn fetch_image_url(url: String) -> Result<String, String> {
    let parsed = reqwest::Url::parse(url.trim()).map_err(|error| format!("Invalid address: {error}"))?;
    if !matches!(parsed.scheme(), "http" | "https") {
        return Err("Only http and https pictures".to_string());
    }
    let client = reqwest::Client::builder()
        .timeout(std::time::Duration::from_secs(30))
        .user_agent("VisualTeX")
        .build()
        .map_err(|error| error.to_string())?;
    let response = client.get(parsed).send().await.map_err(|error| format!("Unable to fetch the picture: {error}"))?;
    if !response.status().is_success() {
        return Err(format!("Unable to fetch the picture: {}", response.status()));
    }
    if response.content_length().is_some_and(|length| length > 30 * 1024 * 1024) {
        return Err("The picture is too large".to_string());
    }
    let typed = response
        .headers()
        .get(reqwest::header::CONTENT_TYPE)
        .and_then(|value| value.to_str().ok())
        .is_some_and(|value| value.trim_start().to_ascii_lowercase().starts_with("image/"));
    let bytes = response.bytes().await.map_err(|error| format!("Unable to fetch the picture: {error}"))?;
    if bytes.len() > 30 * 1024 * 1024 {
        return Err("The picture is too large".to_string());
    }
    let looks_like_image = bytes.starts_with(&[0x89, b'P', b'N', b'G'])
        || bytes.starts_with(&[0xff, 0xd8])
        || bytes.starts_with(b"GIF8")
        || (bytes.len() > 12 && &bytes[8..12] == b"WEBP")
        || bytes.starts_with(b"BM");
    if !typed && !looks_like_image {
        return Err("Not a picture".to_string());
    }
    Ok(BASE64_STANDARD.encode(&bytes))
}

/// A new, hidden print window (an older one goes first).
#[tauri::command]
pub async fn pdf_print_open(app: AppHandle) -> Result<(), String> {
    if let Some(window) = app.get_webview_window(PRINT_WINDOW) {
        let _ = window.destroy();
    }
    WebviewWindowBuilder::new(&app, PRINT_WINDOW, WebviewUrl::App("print.html".into()))
        .title("VisualTeX PDF")
        .visible(false)
        .focused(false)
        .skip_taskbar(true)
        .inner_size(794.0, 1123.0)
        .build()
        .map_err(|error| format!("Unable to open the print window: {error}"))?;
    Ok(())
}

#[tauri::command]
pub fn pdf_print_close(app: AppHandle) {
    if let Some(window) = app.get_webview_window(PRINT_WINDOW) {
        let _ = window.destroy();
    }
}

fn pdf_target(path: &str) -> Result<PathBuf, String> {
    let target = PathBuf::from(path.trim());
    if !target.is_absolute() {
        return Err("Export path must be absolute".to_string());
    }
    let extension = target.extension().and_then(|value| value.to_str()).unwrap_or_default().to_ascii_lowercase();
    if extension != "pdf" {
        return Err("Unsupported export file extension".to_string());
    }
    let parent = target.parent().ok_or_else(|| "Export path has no parent directory".to_string())?;
    if !parent.is_dir() {
        return Err("Export directory does not exist".to_string());
    }
    Ok(target)
}

/// The print window's pages to `path` (paper `width_mm` x `height_mm`, no margins: the pages carry theirs).
#[tauri::command]
pub async fn pdf_print_save(app: AppHandle, path: String, width_mm: f64, height_mm: f64, title: String) -> Result<(), String> {
    let target = pdf_target(&path)?;
    if !(width_mm > 10.0 && width_mm < 2000.0 && height_mm > 10.0 && height_mm < 2000.0) {
        return Err("Unsupported paper size".to_string());
    }
    let window = app
        .get_webview_window(PRINT_WINDOW)
        .ok_or_else(|| "The print window is not open".to_string())?;
    // a file left from before must not pass for the new one
    if target.exists() {
        std::fs::remove_file(&target).map_err(|error| format!("Unable to replace the file: {error}"))?;
    }
    let paper = (width_mm * 72.0 / 25.4, height_mm * 72.0 / 25.4);
    #[cfg(target_os = "macos")]
    {
        macos::print_to_pdf(&window, target, paper, title).await
    }
    #[cfg(not(target_os = "macos"))]
    {
        let _ = (window, target, paper, title);
        Err("PDF export is not available here".to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn only_pdf_paths_in_existing_folders() {
        assert!(pdf_target("relative.pdf").is_err());
        assert!(pdf_target("/tmp/x.docx").is_err());
        assert!(pdf_target("/no-such-folder-vtx/x.pdf").is_err());
        assert!(pdf_target("/tmp/x.pdf").is_ok());
    }

    #[test]
    fn only_document_files_are_opened() {
        assert!(read_import_file("/etc/hosts".to_string()).is_err());
        assert!(read_import_file("relative.md".to_string()).is_err());
    }
}

#[cfg(target_os = "macos")]
mod macos {
    use std::{cell::Cell, cell::RefCell, ffi::c_void, path::PathBuf};

    use objc2::rc::Retained;
    use objc2::runtime::{AnyObject, Bool, NSObject};
    use objc2::{define_class, msg_send, sel, ClassType, DefinedClass, MainThreadMarker, MainThreadOnly};
    use objc2_app_kit::{NSPrintInfo, NSPrintJobSavingURL, NSPrintOperation, NSPrintSaveJob, NSWindow};
    use objc2_foundation::{NSPoint, NSRect, NSSize, NSString, NSURL};
    use objc2_web_kit::WKWebView;

    pub struct Ivars {
        done: Cell<Option<Box<dyn FnOnce(bool)>>>,
    }

    define_class!(
        #[unsafe(super(NSObject))]
        #[thread_kind = MainThreadOnly]
        #[name = "VisualTeXPdfPrintDelegate"]
        #[ivars = Ivars]
        pub struct PrintDelegate;

        impl PrintDelegate {
            #[unsafe(method(printOperationDidRun:success:contextInfo:))]
            fn did_run(&self, _operation: &NSPrintOperation, success: Bool, _context: *mut c_void) {
                if let Some(done) = self.ivars().done.take() {
                    done(success.as_bool());
                }
            }
        }
    );

    thread_local! {
        // the delegate lives until its operation has run
        static PENDING: RefCell<Option<Retained<PrintDelegate>>> = const { RefCell::new(None) };
    }

    pub async fn print_to_pdf(window: &tauri::WebviewWindow, out: PathBuf, paper: (f64, f64), title: String) -> Result<(), String> {
        let (tx, rx) = tokio::sync::oneshot::channel::<Result<(), String>>();
        window
            .with_webview(move |platform| unsafe {
                let Some(mtm) = MainThreadMarker::new() else {
                    let _ = tx.send(Err("not on the main thread".to_string()));
                    return;
                };
                let webview: &WKWebView = &*platform.inner().cast();
                let ns_window: &NSWindow = &*platform.ns_window().cast();
                // a fresh print info (the shared one is the app's print settings)
                let info: Retained<NSPrintInfo> = msg_send![NSPrintInfo::class(), new];
                info.setPaperSize(NSSize::new(paper.0, paper.1));
                info.setTopMargin(0.0);
                info.setBottomMargin(0.0);
                info.setLeftMargin(0.0);
                info.setRightMargin(0.0);
                info.setVerticallyCentered(false);
                info.setHorizontallyCentered(false);
                info.setJobDisposition(NSPrintSaveJob);
                let url = NSURL::fileURLWithPath(&NSString::from_str(&out.to_string_lossy()));
                let dictionary: Retained<AnyObject> = msg_send![&*info, dictionary];
                let _: () = msg_send![&*dictionary, setObject: &*url, forKey: NSPrintJobSavingURL];
                let operation = webview.printOperationWithPrintInfo(&info);
                operation.setShowsPrintPanel(false);
                operation.setShowsProgressPanel(false);
                operation.setJobTitle(Some(&NSString::from_str(&title)));
                if let Some(view) = operation.view() {
                    view.setFrame(NSRect::new(NSPoint::new(0.0, 0.0), NSSize::new(paper.0, paper.1)));
                }
                let written = out.clone();
                let tx = Cell::new(Some(tx));
                let delegate = PrintDelegate::alloc(mtm).set_ivars(Ivars {
                    done: Cell::new(Some(Box::new(move |success| {
                        // success is not proof: the file must be there (a missing folder "succeeds")
                        let ok = std::fs::metadata(&written).map(|m| m.len() > 0).unwrap_or(false);
                        if let Some(tx) = tx.take() {
                            let _ = tx.send(if success && ok { Ok(()) } else { Err("The print operation wrote no PDF".to_string()) });
                        }
                        PENDING.with(|pending| pending.borrow_mut().take());
                    }))),
                });
                let delegate: Retained<PrintDelegate> = msg_send![super(delegate), init];
                PENDING.with(|pending| *pending.borrow_mut() = Some(delegate.clone()));
                operation.runOperationModalForWindow_delegate_didRunSelector_contextInfo(
                    ns_window,
                    Some(&delegate),
                    Some(sel!(printOperationDidRun:success:contextInfo:)),
                    std::ptr::null_mut(),
                );
            })
            .map_err(|error| error.to_string())?;
        match tokio::time::timeout(std::time::Duration::from_secs(300), rx).await {
            Ok(Ok(result)) => result,
            Ok(Err(_)) => Err("The print operation ended without an answer".to_string()),
            Err(_) => Err("The print operation took too long".to_string()),
        }
    }
}
