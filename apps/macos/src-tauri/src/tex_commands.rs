//! The tex_* Tauri commands of VisualTeX's LaTeX documents (personal edition): thin wrappers around the
//! tex-bridge crate's `Service`, run off the async runtime (TeX runs take seconds). Add to src-tauri/src/ as
//! `tex_commands.rs` (Cargo.toml: `tex-bridge = { path = "tex-bridge" }`; register the state in setup and the commands in generate_handler!).

use std::sync::Arc;

use serde::Deserialize;
use tauri::{AppHandle, Manager, State};
use tex_bridge::service::{
    BuildRequest, CompileRequest, CompileResult, FileRequest, PageOut, PagesRequest, PdfOut, PreviewRequest, PreviewResult, ProjectInfo, Service, SynctexOut, TexStatus,
};

/// The app's compile service: build folders in the cache, projects made by the app (new documents, opened
/// archives) in the app's data folder.
pub struct TexState(pub Arc<Service>);

impl TexState {
    pub fn new(app: &AppHandle) -> Result<TexState, String> {
        let cache = app.path().app_cache_dir().map_err(|e| e.to_string())?.join("latex-builds");
        let projects = app.path().app_data_dir().map_err(|e| e.to_string())?.join("latex-projects");
        Ok(TexState(Arc::new(Service::new(cache, projects))))
    }
}

#[derive(Deserialize)]
pub struct NameRequest {
    #[serde(default)]
    pub name: String,
}

#[derive(Deserialize)]
pub struct DocRequest {
    pub doc: String,
}

async fn blocking<T: Send + 'static>(f: impl FnOnce() -> T + Send + 'static) -> Result<T, String> {
    tauri::async_runtime::spawn_blocking(f).await.map_err(|e| e.to_string())
}

#[tauri::command]
pub async fn tex_status(state: State<'_, TexState>) -> Result<TexStatus, String> {
    let s = state.0.clone();
    blocking(move || s.status()).await
}

#[tauri::command]
pub async fn tex_project_open(state: State<'_, TexState>, request: FileRequest) -> Result<ProjectInfo, String> {
    let s = state.0.clone();
    blocking(move || s.project_open(&request)).await?
}

#[tauri::command]
pub async fn tex_project_new(state: State<'_, TexState>, request: NameRequest) -> Result<ProjectInfo, String> {
    let s = state.0.clone();
    blocking(move || s.project_new(&request.name)).await?
}

#[tauri::command]
pub async fn tex_project_files(state: State<'_, TexState>, request: FileRequest) -> Result<ProjectInfo, String> {
    let s = state.0.clone();
    blocking(move || s.project_files(&request)).await?
}

#[tauri::command]
pub async fn tex_project_read(state: State<'_, TexState>, request: FileRequest) -> Result<String, String> {
    let s = state.0.clone();
    blocking(move || s.project_read(&request)).await?
}

#[tauri::command]
pub async fn tex_project_write(state: State<'_, TexState>, request: FileRequest) -> Result<(), String> {
    let s = state.0.clone();
    blocking(move || s.project_write(&request)).await?
}

#[tauri::command]
pub async fn tex_compile(state: State<'_, TexState>, request: CompileRequest) -> Result<CompileResult, String> {
    let s = state.0.clone();
    blocking(move || s.compile(&request)).await
}

#[tauri::command]
pub async fn tex_preview(state: State<'_, TexState>, request: PreviewRequest) -> Result<PreviewResult, String> {
    let s = state.0.clone();
    blocking(move || s.preview(&request)).await
}

#[tauri::command]
pub async fn tex_pages(state: State<'_, TexState>, request: PagesRequest) -> Result<Vec<PageOut>, String> {
    let s = state.0.clone();
    blocking(move || s.pages(&request)).await?
}

#[tauri::command]
pub async fn tex_pdf(state: State<'_, TexState>, request: BuildRequest) -> Result<PdfOut, String> {
    let s = state.0.clone();
    blocking(move || s.pdf(&request)).await?
}

#[tauri::command]
pub async fn tex_synctex(state: State<'_, TexState>, request: BuildRequest) -> Result<SynctexOut, String> {
    let s = state.0.clone();
    blocking(move || s.synctex(&request)).await?
}

#[tauri::command]
pub async fn tex_close(state: State<'_, TexState>, request: DocRequest) -> Result<(), String> {
    let s = state.0.clone();
    blocking(move || s.close(&request.doc)).await
}
