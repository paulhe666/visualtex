// Documents to Word and PDF (docs/document-editor/DESIGN.md §12), and images dropped on a document.
// - document_export_omml: the document's formulas as Word equations, written by the shared Word writer
//   (office::omml_write), numbered ones with Word's own "#" number and the bookmark references point to.
// - pdf_print_open / pdf_print_save / pdf_print_close: the print window (print.html) gets the print copy of the
//   page view from the main window; WebView2's PrintToPdf writes it to the file.
// - read_image_file: an image file dropped from the system (Tauri's drag-drop event gives its path).

use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};
use serde::{Deserialize, Serialize};
use std::{fs, path::PathBuf};
use tauri::{AppHandle, Manager, WebviewUrl, WebviewWindowBuilder};

const PRINT_WINDOW: &str = "pdf-print";

#[tauri::command]
pub fn read_image_file(path: String) -> Result<String, String> {
    let target = PathBuf::from(path.trim());
    let extension = target
        .extension()
        .and_then(|value| value.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    if !target.is_absolute()
        || !matches!(extension.as_str(), "png" | "jpg" | "jpeg" | "gif" | "webp" | "bmp" | "svg" | "tif" | "tiff" | "avif")
    {
        return Err("Not an image file".to_string());
    }
    let size = fs::metadata(&target).map_err(|error| format!("Unable to read the image: {error}"))?.len();
    if size > 50 * 1024 * 1024 {
        return Err("The image is too large".to_string());
    }
    let bytes = fs::read(&target).map_err(|error| format!("Unable to read the image: {error}"))?;
    Ok(BASE64_STANDARD.encode(bytes))
}

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
    use crate::office::omml_write;
    let omml = omml_write::transform_mathml_to_omml(&item.mathml)?;
    let omml = match &item.number {
        Some(number) => omml_write::build_word_native_numbered_omml(
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
    omml_write::insertion_omml(&omml, math_font)
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
    if target.exists() {
        fs::remove_file(&target).map_err(|error| format!("Unable to replace the file: {error}"))?;
    }
    let _ = title;
    #[cfg(windows)]
    {
        windows_pdf::print_to_pdf(&window, target, (width_mm / 25.4, height_mm / 25.4)).await
    }
    #[cfg(not(windows))]
    {
        let _ = (window, target);
        Err("PDF export is not available here".to_string())
    }
}

#[cfg(windows)]
mod windows_pdf {
    use std::path::PathBuf;
    use webview2_com::Microsoft::Web::WebView2::Win32::{ICoreWebView2Environment6, ICoreWebView2_7};
    use webview2_com::PrintToPdfCompletedHandler;
    use windows::core::{Interface, HSTRING};

    /// WebView2's PrintToPdf (ICoreWebView2_7) with print settings of the paper's size in inches and no margins,
    /// header or footer; the file must be there afterwards.
    pub async fn print_to_pdf(window: &tauri::WebviewWindow, out: PathBuf, paper_in: (f64, f64)) -> Result<(), String> {
        let (tx, rx) = tokio::sync::oneshot::channel::<Result<(), String>>();
        // shared by the completion handler and the setup's error path (which says what failed)
        let tx = std::sync::Arc::new(std::sync::Mutex::new(Some(tx)));
        let tx_error = tx.clone();
        let written = out.clone();
        window
            .with_webview(move |platform| {
                let result = (|| -> Result<(), String> {
                    let controller = platform.controller();
                    let core = unsafe { controller.CoreWebView2() }.map_err(|e| e.to_string())?;
                    let core7: ICoreWebView2_7 = core.cast().map_err(|e| e.to_string())?;
                    let environment: ICoreWebView2Environment6 = platform.environment().cast().map_err(|e| e.to_string())?;
                    let settings = unsafe { environment.CreatePrintSettings() }.map_err(|e| e.to_string())?;
                    unsafe {
                        settings.SetPageWidth(paper_in.0).map_err(|e| e.to_string())?;
                        settings.SetPageHeight(paper_in.1).map_err(|e| e.to_string())?;
                        settings.SetMarginTop(0.0).map_err(|e| e.to_string())?;
                        settings.SetMarginBottom(0.0).map_err(|e| e.to_string())?;
                        settings.SetMarginLeft(0.0).map_err(|e| e.to_string())?;
                        settings.SetMarginRight(0.0).map_err(|e| e.to_string())?;
                        settings.SetScaleFactor(1.0).map_err(|e| e.to_string())?;
                        settings.SetShouldPrintBackgrounds(true).map_err(|e| e.to_string())?;
                        settings.SetShouldPrintHeaderAndFooter(false).map_err(|e| e.to_string())?;
                    }
                    let file = written.clone();
                    let handler = PrintToPdfCompletedHandler::create(Box::new(move |error_code, success| {
                        let ok = error_code.is_ok() && success;
                        let exists = std::fs::metadata(&file).map(|m| m.len() > 0).unwrap_or(false);
                        if let Some(tx) = tx.lock().ok().and_then(|mut t| t.take()) {
                            let _ = tx.send(if ok && exists { Ok(()) } else { Err("WebView2 wrote no PDF".to_string()) });
                        }
                        Ok(())
                    }));
                    let path = HSTRING::from(written.to_string_lossy().as_ref());
                    unsafe { core7.PrintToPdf(&path, &settings, &handler) }.map_err(|e| e.to_string())?;
                    Ok(())
                })();
                if let Err(error) = result {
                    if let Some(tx) = tx_error.lock().ok().and_then(|mut t| t.take()) {
                        let _ = tx.send(Err(format!("WebView2 could not print: {error}")));
                    }
                }
            })
            .map_err(|error| error.to_string())?;
        match tokio::time::timeout(std::time::Duration::from_secs(300), rx).await {
            Ok(Ok(result)) => result,
            Ok(Err(_)) => Err("The print operation ended without an answer".to_string()),
            Err(_) => Err("The print operation took too long".to_string()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const MATHML: &str = r#"<math xmlns="http://www.w3.org/1998/Math/MathML" display="block"><mrow><mi>E</mi><mo>=</mo><mi>m</mi><msup><mi>c</mi><mn>2</mn></msup></mrow></math>"#;

    #[test]
    fn a_formula_becomes_one_word_equation() {
        let omml = omml_of(&OmmlItem { mathml: MATHML.to_string(), number: None }, "Cambria Math").unwrap();
        assert!(omml.contains("oMath"), "{omml}");
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
    }

    #[test]
    fn only_pdf_paths_in_existing_folders() {
        assert!(pdf_target("relative.pdf").is_err());
        assert!(pdf_target(r"C:\Windows\x.docx").is_err());
        assert!(pdf_target(r"C:\no-such-folder-vtx\x.pdf").is_err());
        assert!(pdf_target(&std::env::temp_dir().join("x.pdf").to_string_lossy()).is_ok());
    }

    #[test]
    fn only_image_files_are_read() {
        assert!(read_image_file(r"C:\Windows\win.ini".to_string()).is_err());
        assert!(read_image_file("relative.png".to_string()).is_err());
    }
}
