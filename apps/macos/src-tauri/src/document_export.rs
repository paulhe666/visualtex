// Documents to Word and PDF (docs/document-editor/DESIGN.md §12).
// - document_export_omml: the document's formulas as Word equations, written by the shared
//   Word writer (omml_write), numbered ones with Word's own "#" number and the bookmark
//   references point to.
// - pdf_print_open / pdf_print_save / pdf_print_close: the print window (print.html) gets the
//   print copy of the page view from the main window; WKWebView's print operation writes it
//   to the file (runOperationModalForWindow; runOperation gives blank pages).

use serde::{Deserialize, Serialize};
use std::path::PathBuf;
use tauri::{AppHandle, Manager, WebviewUrl, WebviewWindowBuilder};

const PRINT_WINDOW: &str = "pdf-print";

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OmmlNumber {
    bookmark: String,
    sequence: String,
    chapter_level: u32,
    separator: String,
    literal_prefix: Option<String>,
}

#[derive(Deserialize)]
pub struct OmmlItem {
    mathml: String,
    number: Option<OmmlNumber>,
}

#[derive(Serialize, Default)]
pub struct OmmlAnswer {
    #[serde(skip_serializing_if = "Option::is_none")]
    omml: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    error: Option<String>,
}

fn omml_of(item: &OmmlItem, math_font: &str) -> Result<String, String> {
    let omml = crate::office::omml_write::transform_mathml_to_omml(&item.mathml)?;
    let omml = match &item.number {
        Some(number) => crate::office::omml_write::build_word_native_numbered_omml(
            &omml,
            "Equation",
            Some(&number.bookmark),
            number.chapter_level,
            if number.chapter_level > 0 { &number.separator } else { "" },
            if number.sequence.trim().is_empty() { "1" } else { &number.sequence },
            if number.chapter_level > 0 { number.literal_prefix.as_deref() } else { None },
            None,
        )?,
        None => omml,
    };
    crate::office::omml_write::insertion_omml(&omml, math_font)
}

#[tauri::command]
pub async fn document_export_omml(items: Vec<OmmlItem>, math_font: String) -> Vec<OmmlAnswer> {
    if items.len() > 20_000 {
        return Vec::new();
    }
    tauri::async_runtime::spawn_blocking(move || {
        items
            .iter()
            .map(|item| match omml_of(item, &math_font) {
                Ok(omml) => OmmlAnswer { omml: Some(omml), error: None },
                Err(error) => OmmlAnswer { omml: None, error: Some(error) },
            })
            .collect()
    })
    .await
    .unwrap_or_default()
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

    const MATHML: &str = r#"<math xmlns="http://www.w3.org/1998/Math/MathML" display="block"><mrow><mi>E</mi><mo>=</mo><mi>m</mi><msup><mi>c</mi><mn>2</mn></msup></mrow></math>"#;

    #[test]
    fn a_formula_becomes_one_word_equation() {
        let omml = omml_of(&OmmlItem { mathml: MATHML.to_string(), number: None }, "Cambria Math").unwrap();
        assert!(omml.starts_with("<m:oMath"), "{omml}");
        assert!(omml.contains("Cambria Math"));
    }

    #[test]
    fn a_numbered_formula_has_words_number_and_the_bookmark() {
        let number = OmmlNumber {
            bookmark: "VTEqNum_0123456789abcdef0123456789abcdef".to_string(),
            sequence: "3".to_string(),
            chapter_level: 0,
            separator: ".".to_string(),
            literal_prefix: None,
        };
        let omml = omml_of(&OmmlItem { mathml: MATHML.to_string(), number: Some(number) }, "Cambria Math").unwrap();
        assert!(omml.contains("eqArr"), "{omml}");
        assert!(omml.contains("VTEqNum_0123456789abcdef0123456789abcdef"));
        assert!(omml.contains("SEQ Equation"));
        if let Ok(path) = std::env::var("VTX_OMML_SAMPLE") {
            let plain = omml_of(&OmmlItem { mathml: MATHML.to_string(), number: None }, "Cambria Math").unwrap();
            std::fs::write(path, format!("{plain}\n{omml}\n")).unwrap();
        }
    }

    #[test]
    fn only_pdf_paths_in_existing_folders() {
        assert!(pdf_target("relative.pdf").is_err());
        assert!(pdf_target("/tmp/x.docx").is_err());
        assert!(pdf_target("/no-such-folder-vtx/x.pdf").is_err());
        assert!(pdf_target("/tmp/x.pdf").is_ok());
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
