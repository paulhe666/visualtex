//! `render_latex` on macOS: VisualTeX draws formulas in its WebView. The companion asks the running
//! windows (the main window, the resident Office editors) through an event; the first answer wins
//! (`mcp_render_result`, src/office/shared/mcpRenderBridge.ts).
use crate::office::state::OfficeCompanionState;
use serde::Deserialize;
use serde_json::json;
use std::collections::HashMap;
use std::sync::{Mutex, OnceLock};
use std::time::Duration;
use tauri::Emitter;
use tokio::sync::oneshot;

pub const RENDER_EVENT: &str = "visualtex://mcp-render";

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RenderExport {
    #[serde(default, rename = "mathml")]
    pub math_ml: Option<String>,
    pub svg: String,
    #[serde(rename = "widthPx")]
    pub width: f64,
    #[serde(rename = "heightPx")]
    pub height: f64,
    #[serde(default, rename = "baselinePx")]
    pub baseline: Option<f64>,
    #[serde(default)]
    pub png_base64: Option<String>,
}

type Waiter = oneshot::Sender<Result<RenderExport, String>>;

static PENDING: OnceLock<Mutex<HashMap<String, Waiter>>> = OnceLock::new();

fn pending() -> std::sync::MutexGuard<'static, HashMap<String, Waiter>> {
    PENDING
        .get_or_init(|| Mutex::new(HashMap::new()))
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner())
}

pub async fn render(
    companion: &OfficeCompanionState,
    latex: &str,
    display: &str,
    font_size_pt: f64,
    png: bool,
    timeout: Duration,
) -> Result<RenderExport, (&'static str, String)> {
    let app = companion
        .app
        .clone()
        .ok_or(("render_unavailable", "This VisualTeX process has no window to draw formulas in".to_string()))?;
    let id = uuid::Uuid::new_v4().simple().to_string();
    let (sender, receiver) = oneshot::channel();
    pending().insert(id.clone(), sender);
    if let Err(error) = app.emit(
        RENDER_EVENT,
        json!({ "id": id, "latex": latex, "display": display, "fontSizePt": font_size_pt, "png": png }),
    ) {
        pending().remove(&id);
        return Err(("render_unavailable", error.to_string()));
    }
    let outcome = tokio::time::timeout(timeout, receiver).await;
    pending().remove(&id);
    match outcome {
        Ok(Ok(Ok(export))) => Ok(export),
        Ok(Ok(Err(message))) => Err(("latex_error", message)),
        Ok(Err(_)) => Err(("render_unavailable", "The drawing window closed".to_string())),
        Err(_) => Err((
            "render_unavailable",
            "No VisualTeX window answered (open VisualTeX once, or turn on Office background start)".to_string(),
        )),
    }
}

/// A window's answer to a render request (the first one is used).
#[tauri::command]
pub fn mcp_render_result(id: String, result: Option<RenderExport>, error: Option<String>) {
    let Some(waiter) = pending().remove(&id) else { return };
    let _ = waiter.send(match (result, error) {
        (Some(export), _) => Ok(export),
        (None, Some(message)) => Err(message),
        (None, None) => Err("The formula could not be drawn".to_string()),
    });
}
