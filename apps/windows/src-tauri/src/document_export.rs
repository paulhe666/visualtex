// Documents to PDF (docs/document-editor/DESIGN.md §12) and the files a document is read from.
// - pdf_print_open / pdf_print_save / pdf_print_close: the print window (print.html) gets the print copy of the
//   page view from the main window; WebView2's PrintToPdf writes it to the file.
// - read_image_file: an image file dropped from the system (Tauri's drag-drop event gives its path), or a picture a
//   Markdown / LaTeX file names next to it.
// - read_import_file: a document file chosen in the open dialog (its folder then finds its pictures).
// - fetch_image_url: a picture a Markdown file names on the web.

use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};
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
        let written = out.clone();
        // shared by the completion handler and the setup's error path (which says what failed)
        let tx = std::sync::Arc::new(std::sync::Mutex::new(Some(tx)));
        let tx_error = tx.clone();
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

    #[test]
    fn only_document_files_are_opened() {
        assert!(read_import_file(r"C:\Windows\win.ini".to_string()).is_err());
        assert!(read_import_file("relative.md".to_string()).is_err());
    }
}
