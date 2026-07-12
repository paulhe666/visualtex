use notify::{RecommendedWatcher, RecursiveMode, Watcher};
use parking_lot::Mutex;
use serde::Serialize;
use std::ffi::OsString;
use std::fs;
use std::path::{Component, Path, PathBuf};
use tauri::{AppHandle, Emitter, Manager, State};
use tauri_plugin_deep_link::DeepLinkExt;
use vt_core::{CoreService, EditOutcome};
use vt_models::{InstalledModelPackage, ModelKind, ModelPackageInspection};
use vt_protocol::{
    CompileArtifact, DocumentOcrResult, DocumentSnapshot, ExternalChangeReport,
    ExternalConflictOutcome, ExternalConflictResolution, ExternalFileChange, FileId,
    FormulaOcrResult, ForwardSearchResult, InverseSearchResult, LayoutMapArtifact,
    NodeAttributesPatch, NodeId, OcrWorkerHealth, PdfDocumentInfo, PdfRenderRequest,
    PdfRenderedImage, PdfTextHit, ProjectDependencyGraph, ProjectIndex, ProjectReplaceOutcome,
    ProjectReplacePlan, ProjectReplaceRequest, ProjectSearchMatch, ProjectSearchRequest,
    ProjectTemplateSummary, Revision, SymbolRenameRequest, TextEdit, ToolInfo,
};

#[derive(Default)]
struct AppState {
    core: Mutex<Option<CoreService>>,
    watcher: Mutex<Option<RecommendedWatcher>>,
    pending_deep_links: Mutex<Vec<String>>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ModelCatalog {
    models_root: PathBuf,
    installed: Vec<InstalledModelPackage>,
    active: Vec<InstalledModelPackage>,
}

fn models_root(app: &AppHandle) -> Result<PathBuf, String> {
    app.path()
        .app_data_dir()
        .map(|path| path.join("models"))
        .map_err(|error| error.to_string())
}

fn load_model_catalog(app: &AppHandle) -> Result<ModelCatalog, String> {
    let root = models_root(app)?;
    let installed = vt_models::list_installed(&root).map_err(|error| error.to_string())?;
    let mut active = Vec::new();
    for kind in [
        ModelKind::FormulaOcr,
        ModelKind::LayoutOcr,
        ModelKind::TextOcr,
        ModelKind::TableOcr,
    ] {
        if let Some(package) =
            vt_models::active_model(&root, kind).map_err(|error| error.to_string())?
        {
            active.push(package);
        }
    }
    Ok(ModelCatalog {
        models_root: root,
        installed,
        active,
    })
}

fn configured_ocr_worker(
    project_root: PathBuf,
    app: &AppHandle,
) -> Result<vt_ocr::OcrWorkerConfig, String> {
    let mut config = vt_ocr::OcrWorkerConfig::bundled(project_root);
    let root = models_root(app)?;
    if let Some(package) =
        vt_models::active_model(&root, ModelKind::FormulaOcr).map_err(|error| error.to_string())?
    {
        let model_dir = package.install_path.join(&package.manifest.entrypoint);
        config.formula_model_dir = Some(
            model_dir
                .canonicalize()
                .map_err(|error| error.to_string())?,
        );
        config.formula_model_name = Some(format!(
            "{}@{}",
            package.manifest.id, package.manifest.version
        ));
    }
    if let Some(package) =
        vt_models::active_model(&root, ModelKind::LayoutOcr).map_err(|error| error.to_string())?
    {
        let pipeline_config = package.install_path.join(&package.manifest.entrypoint);
        config.document_pipeline_config = Some(
            pipeline_config
                .canonicalize()
                .map_err(|error| error.to_string())?,
        );
        config.document_package_root = Some(
            package
                .install_path
                .canonicalize()
                .map_err(|error| error.to_string())?,
        );
        config.document_model_name = Some(format!(
            "{}@{}",
            package.manifest.id, package.manifest.version
        ));
    }
    Ok(config)
}

fn import_ocr_image(project_root: &Path, source: &Path) -> Result<PathBuf, String> {
    vt_ocr::import_image(project_root, source).map_err(|error| error.to_string())
}

fn with_core<T>(
    state: &State<'_, AppState>,
    f: impl FnOnce(&mut CoreService) -> Result<T, String>,
) -> Result<T, String> {
    let mut guard = state.core.lock();
    let core = guard
        .as_mut()
        .ok_or_else(|| "No project is open".to_owned())?;
    f(core)
}

fn validate_imported_ocr_image(project_root: &Path, source: &Path) -> Result<PathBuf, String> {
    let allowed_root = project_root
        .join(".visualtex/ocr-input")
        .canonicalize()
        .map_err(|error| error.to_string())?;
    let metadata = fs::symlink_metadata(source).map_err(|error| error.to_string())?;
    if metadata.file_type().is_symlink() || !metadata.is_file() {
        return Err("OCR source page must be a regular imported image file".to_owned());
    }
    let source = source.canonicalize().map_err(|error| error.to_string())?;
    if !source.starts_with(&allowed_root)
        || source.extension().and_then(|value| value.to_str()) != Some("png")
    {
        return Err(
            "OCR source page is outside the current project's validated OCR cache".to_owned(),
        );
    }
    Ok(source)
}

fn startup_project_argument(args: impl IntoIterator<Item = OsString>) -> Option<PathBuf> {
    let mut arguments = args.into_iter();
    let _ = arguments.next();
    while let Some(argument) = arguments.next() {
        if argument == "--project" {
            return arguments.next().map(PathBuf::from);
        }
        if let Some(value) = argument
            .to_str()
            .and_then(|value| value.strip_prefix("--project="))
        {
            return Some(PathBuf::from(value));
        }
        let argument_text = argument.to_string_lossy();
        if !argument_text.starts_with('-') && !argument_text.contains("://") {
            return Some(PathBuf::from(argument));
        }
    }
    None
}

fn is_source_event_path(path: &Path) -> bool {
    if path.components().any(|component| {
        matches!(component, Component::Normal(value) if value == ".visualtex" || value == ".git" || value == "target" || value == "node_modules")
    }) {
        return false;
    }
    matches!(
        path.extension().and_then(|value| value.to_str()),
        Some("tex" | "bib" | "sty" | "cls")
    )
}

fn install_core(
    core: CoreService,
    app: &AppHandle,
    state: &State<'_, AppState>,
) -> Result<DocumentSnapshot, String> {
    let snapshot = core.root_snapshot().map_err(|error| error.to_string())?;
    let root = core.project_root().to_path_buf();
    let app_handle = app.clone();
    let mut watcher =
        notify::recommended_watcher(move |result: Result<notify::Event, notify::Error>| {
            let Ok(event) = result else {
                return;
            };
            if event.paths.iter().any(|path| is_source_event_path(path)) {
                let _ = app_handle.emit("visualtex://project-source-changed", ());
            }
        })
        .map_err(|error| error.to_string())?;
    watcher
        .watch(&root, RecursiveMode::Recursive)
        .map_err(|error| error.to_string())?;
    *state.watcher.lock() = Some(watcher);
    *state.core.lock() = Some(core);
    Ok(snapshot)
}

fn ensure_deep_link_project(app: &AppHandle, project: &Path) -> Result<DocumentSnapshot, String> {
    let target = project.canonicalize().map_err(|error| error.to_string())?;
    let state = app.state::<AppState>();
    if let Some(core) = state.core.lock().as_ref() {
        let current = core
            .project_root()
            .canonicalize()
            .map_err(|error| error.to_string())?;
        if current != target {
            return Err(
                "A different project is already open. Use the visualstudio UI to switch projects so unsaved buffers are not discarded."
                    .to_owned(),
            );
        }
        return core.root_snapshot().map_err(|error| error.to_string());
    }
    let core = CoreService::open_project(&target).map_err(|error| error.to_string())?;
    install_core(core, app, &state)
}

fn emit_deep_link_error(app: &AppHandle, message: impl Into<String>) {
    let _ = app.emit("visualtex://deep-link-error", message.into());
}

fn queue_visualtex_uri(app: &AppHandle, value: String) {
    app.state::<AppState>()
        .pending_deep_links
        .lock()
        .push(value);
    let _ = app.emit("visualtex://deep-link-received", ());
}

fn handle_visualtex_uri(app: AppHandle, value: &str) {
    let action = match vt_uri::VisualTexUriAction::parse(value) {
        Ok(action) => action,
        Err(error) => {
            emit_deep_link_error(&app, error.to_string());
            return;
        }
    };
    let snapshot = match ensure_deep_link_project(&app, action.project()) {
        Ok(snapshot) => snapshot,
        Err(error) => {
            emit_deep_link_error(&app, error);
            return;
        }
    };
    let _ = app.emit("visualtex://project-opened", snapshot);

    match action {
        vt_uri::VisualTexUriAction::Open { .. } => {}
        vt_uri::VisualTexUriAction::ForwardSearch {
            source_file,
            line,
            column,
            pdf_path,
            ..
        } => {
            tauri::async_runtime::spawn(async move {
                let state = app.state::<AppState>();
                let project_root =
                    match with_core(&state, |core| Ok(core.project_root().to_path_buf())) {
                        Ok(project_root) => project_root,
                        Err(error) => {
                            emit_deep_link_error(&app, error);
                            return;
                        }
                    };
                let result = vt_synctex::forward_search(
                    &project_root,
                    &source_file,
                    line,
                    column,
                    &pdf_path,
                )
                .await;
                match result {
                    Ok(result) => {
                        let _ = app.emit("visualtex://forward-search-result", result);
                    }
                    Err(error) => emit_deep_link_error(&app, error.to_string()),
                }
            });
        }
        vt_uri::VisualTexUriAction::InverseSearch {
            pdf_path,
            page,
            x,
            y,
            ..
        } => {
            tauri::async_runtime::spawn(async move {
                let state = app.state::<AppState>();
                let project_root =
                    match with_core(&state, |core| Ok(core.project_root().to_path_buf())) {
                        Ok(project_root) => project_root,
                        Err(error) => {
                            emit_deep_link_error(&app, error);
                            return;
                        }
                    };
                let result = vt_synctex::inverse_search(&project_root, &pdf_path, page, x, y).await;
                match result {
                    Ok(result) => {
                        let _ = app.emit("visualtex://inverse-search-result", result);
                    }
                    Err(error) => emit_deep_link_error(&app, error.to_string()),
                }
            });
        }
    }
}

#[tauri::command]
fn drain_deep_links(state: State<'_, AppState>) -> Vec<String> {
    std::mem::take(&mut *state.pending_deep_links.lock())
}

#[tauri::command]
fn process_deep_link(value: String, app: AppHandle) {
    handle_visualtex_uri(app, &value);
}

#[tauri::command]
fn open_project(
    path: PathBuf,
    app: AppHandle,
    state: State<'_, AppState>,
) -> Result<DocumentSnapshot, String> {
    let core = CoreService::open_project(path).map_err(|error| error.to_string())?;
    install_core(core, &app, &state)
}

#[tauri::command]
fn init_project(
    path: PathBuf,
    app: AppHandle,
    state: State<'_, AppState>,
) -> Result<DocumentSnapshot, String> {
    let core = CoreService::init_project(path).map_err(|error| error.to_string())?;
    install_core(core, &app, &state)
}

#[tauri::command]
fn list_project_templates() -> Vec<ProjectTemplateSummary> {
    CoreService::project_templates()
}

#[tauri::command]
fn inspect_model_package(source: PathBuf) -> Result<ModelPackageInspection, String> {
    vt_models::inspect_package(source).map_err(|error| error.to_string())
}

#[tauri::command]
fn install_model_package(source: PathBuf, app: AppHandle) -> Result<InstalledModelPackage, String> {
    let root = models_root(&app)?;
    vt_models::install_package(source, root).map_err(|error| error.to_string())
}

#[tauri::command]
fn list_model_packages(app: AppHandle) -> Result<ModelCatalog, String> {
    load_model_catalog(&app)
}

#[tauri::command]
fn activate_model_package(
    kind: ModelKind,
    id: String,
    version: String,
    app: AppHandle,
) -> Result<ModelCatalog, String> {
    let root = models_root(&app)?;
    vt_models::set_active_model(root, kind, &id, &version).map_err(|error| error.to_string())?;
    load_model_catalog(&app)
}

#[tauri::command]
fn remove_model_package(
    id: String,
    version: String,
    app: AppHandle,
) -> Result<ModelCatalog, String> {
    let root = models_root(&app)?;
    vt_models::remove_installed(root, &id, &version).map_err(|error| error.to_string())?;
    load_model_catalog(&app)
}

#[tauri::command]
async fn ocr_health(app: AppHandle, state: State<'_, AppState>) -> Result<OcrWorkerHealth, String> {
    let project_root = with_core(&state, |core| Ok(core.project_root().to_path_buf()))?;
    let config = configured_ocr_worker(project_root, &app)?;
    let mut worker = vt_ocr::OcrWorker::spawn(config)
        .await
        .map_err(|error| error.to_string())?;
    let result = worker.health().await.map_err(|error| error.to_string());
    let _ = worker.shutdown().await;
    result
}

#[tauri::command]
async fn recognize_formula_image(
    source: PathBuf,
    app: AppHandle,
    state: State<'_, AppState>,
) -> Result<FormulaOcrResult, String> {
    let project_root = with_core(&state, |core| Ok(core.project_root().to_path_buf()))?;
    let import_root = project_root.clone();
    let imported =
        tauri::async_runtime::spawn_blocking(move || import_ocr_image(&import_root, &source))
            .await
            .map_err(|error| error.to_string())??;
    let config = configured_ocr_worker(project_root, &app)?;
    let mut worker = vt_ocr::OcrWorker::spawn(config)
        .await
        .map_err(|error| error.to_string())?;
    let result = worker
        .recognize_formula(imported)
        .await
        .map_err(|error| error.to_string());
    let _ = worker.shutdown().await;
    result
}

#[tauri::command]
async fn recognize_document_image(
    source: PathBuf,
    app: AppHandle,
    state: State<'_, AppState>,
) -> Result<DocumentOcrResult, String> {
    let project_root = with_core(&state, |core| Ok(core.project_root().to_path_buf()))?;
    let import_root = project_root.clone();
    let imported =
        tauri::async_runtime::spawn_blocking(move || import_ocr_image(&import_root, &source))
            .await
            .map_err(|error| error.to_string())??;
    let config = configured_ocr_worker(project_root, &app)?;
    let mut worker = vt_ocr::OcrWorker::spawn(config)
        .await
        .map_err(|error| error.to_string())?;
    let mut result = worker
        .recognize_document(&imported)
        .await
        .map_err(|error| error.to_string());
    let _ = worker.shutdown().await;
    if let Ok(result) = &mut result {
        result.image_path = Some(imported);
    }
    result
}

#[tauri::command]
fn create_ocr_project(
    target: PathBuf,
    source_image: PathBuf,
    latex_body: String,
    ocr_document: DocumentOcrResult,
    app: AppHandle,
    state: State<'_, AppState>,
) -> Result<DocumentSnapshot, String> {
    let project_root = with_core(&state, |core| Ok(core.project_root().to_path_buf()))?;
    let source_image = validate_imported_ocr_image(&project_root, &source_image)?;
    let core = CoreService::init_ocr_project(target, &latex_body, source_image, ocr_document)
        .map_err(|error| error.to_string())?;
    install_core(core, &app, &state)
}

#[tauri::command]
fn init_project_template(
    path: PathBuf,
    template_id: String,
    app: AppHandle,
    state: State<'_, AppState>,
) -> Result<DocumentSnapshot, String> {
    let core = CoreService::init_project_with_template(path, &template_id)
        .map_err(|error| error.to_string())?;
    install_core(core, &app, &state)
}

#[tauri::command]
fn root_snapshot(state: State<'_, AppState>) -> Result<DocumentSnapshot, String> {
    with_core(&state, |core| {
        core.root_snapshot().map_err(|error| error.to_string())
    })
}

#[tauri::command]
fn list_files(state: State<'_, AppState>) -> Result<Vec<PathBuf>, String> {
    with_core(&state, |core| Ok(core.list_files()))
}

#[tauri::command]
fn open_file(path: PathBuf, state: State<'_, AppState>) -> Result<DocumentSnapshot, String> {
    with_core(&state, |core| {
        core.open_file(path).map_err(|error| error.to_string())
    })
}

#[tauri::command]
fn apply_text_edit(edit: TextEdit, state: State<'_, AppState>) -> Result<EditOutcome, String> {
    with_core(&state, |core| {
        core.apply_text_edit(edit)
            .map_err(|error| error.to_string())
    })
}

#[tauri::command]
fn apply_visual_edit(
    file_id: FileId,
    base_revision: Revision,
    node_id: NodeId,
    content: String,
    state: State<'_, AppState>,
) -> Result<EditOutcome, String> {
    with_core(&state, |core| {
        core.apply_visual_edit(file_id, base_revision, node_id, content)
            .map_err(|error| error.to_string())
    })
}

#[tauri::command]
fn apply_node_attributes(
    file_id: FileId,
    base_revision: Revision,
    node_id: NodeId,
    patch: NodeAttributesPatch,
    state: State<'_, AppState>,
) -> Result<EditOutcome, String> {
    with_core(&state, |core| {
        core.apply_node_attributes(file_id, base_revision, node_id, patch)
            .map_err(|error| error.to_string())
    })
}

#[tauri::command]
fn undo(file_id: FileId, state: State<'_, AppState>) -> Result<EditOutcome, String> {
    with_core(&state, |core| {
        core.undo(file_id).map_err(|error| error.to_string())
    })
}

#[tauri::command]
fn redo(file_id: FileId, state: State<'_, AppState>) -> Result<EditOutcome, String> {
    with_core(&state, |core| {
        core.redo(file_id).map_err(|error| error.to_string())
    })
}

#[tauri::command]
fn save(file_id: FileId, state: State<'_, AppState>) -> Result<(), String> {
    with_core(&state, |core| {
        core.save(file_id).map_err(|error| error.to_string())
    })
}

#[tauri::command]
fn check_external_changes(state: State<'_, AppState>) -> Result<ExternalChangeReport, String> {
    with_core(&state, |core| {
        core.check_external_changes()
            .map_err(|error| error.to_string())
    })
}

#[tauri::command]
fn resolve_external_conflict(
    change: ExternalFileChange,
    resolution: ExternalConflictResolution,
    state: State<'_, AppState>,
) -> Result<ExternalConflictOutcome, String> {
    with_core(&state, |core| {
        core.resolve_external_conflict(&change, resolution)
            .map_err(|error| error.to_string())
    })
}

#[tauri::command]
fn project_index(state: State<'_, AppState>) -> Result<ProjectIndex, String> {
    with_core(&state, |core| {
        core.project_index().map_err(|error| error.to_string())
    })
}

#[tauri::command]
fn project_dependencies(state: State<'_, AppState>) -> Result<ProjectDependencyGraph, String> {
    with_core(&state, |core| Ok(core.project_dependencies()))
}

#[tauri::command]
fn search_project(
    request: ProjectSearchRequest,
    state: State<'_, AppState>,
) -> Result<Vec<ProjectSearchMatch>, String> {
    with_core(&state, |core| {
        core.search_project(&request)
            .map_err(|error| error.to_string())
    })
}

#[tauri::command]
fn preview_project_replace(
    request: ProjectReplaceRequest,
    state: State<'_, AppState>,
) -> Result<ProjectReplacePlan, String> {
    with_core(&state, |core| {
        core.preview_project_replace(&request)
            .map_err(|error| error.to_string())
    })
}

#[tauri::command]
fn preview_symbol_rename(
    request: SymbolRenameRequest,
    state: State<'_, AppState>,
) -> Result<ProjectReplacePlan, String> {
    with_core(&state, |core| {
        core.preview_symbol_rename(&request)
            .map_err(|error| error.to_string())
    })
}

#[tauri::command]
fn apply_project_replace(
    plan: ProjectReplacePlan,
    state: State<'_, AppState>,
) -> Result<ProjectReplaceOutcome, String> {
    with_core(&state, |core| {
        core.apply_project_replace(plan)
            .map_err(|error| error.to_string())
    })
}

#[tauri::command]
async fn compile_project(state: State<'_, AppState>) -> Result<CompileArtifact, String> {
    let request = with_core(&state, |core| {
        core.prepare_compile().map_err(|error| error.to_string())
    })?;
    vt_compiler::compile(request)
        .await
        .map_err(|error| error.to_string())
}

#[tauri::command]
async fn forward_search(
    source_file: PathBuf,
    line: u32,
    column: u32,
    pdf_path: PathBuf,
    state: State<'_, AppState>,
) -> Result<ForwardSearchResult, String> {
    let project_root = with_core(&state, |core| Ok(core.project_root().to_path_buf()))?;
    vt_synctex::forward_search(&project_root, &source_file, line, column, &pdf_path)
        .await
        .map_err(|error| error.to_string())
}

#[tauri::command]
async fn inverse_search(
    pdf_path: PathBuf,
    page: u32,
    x: f32,
    y: f32,
    state: State<'_, AppState>,
) -> Result<InverseSearchResult, String> {
    let project_root = with_core(&state, |core| Ok(core.project_root().to_path_buf()))?;
    vt_synctex::inverse_search(&project_root, &pdf_path, page, x, y)
        .await
        .map_err(|error| error.to_string())
}

#[tauri::command]
fn pdf_document_info(
    pdf_path: PathBuf,
    state: State<'_, AppState>,
) -> Result<PdfDocumentInfo, String> {
    with_core(&state, |core| {
        core.pdf_document_info(&pdf_path)
            .map_err(|error| error.to_string())
    })
}

#[tauri::command]
fn render_pdf(
    request: PdfRenderRequest,
    state: State<'_, AppState>,
) -> Result<PdfRenderedImage, String> {
    with_core(&state, |core| {
        core.render_pdf(request).map_err(|error| error.to_string())
    })
}

#[tauri::command]
fn pdf_text_hit(
    pdf_path: PathBuf,
    page_index: u32,
    x: f32,
    y: f32,
    state: State<'_, AppState>,
) -> Result<Option<PdfTextHit>, String> {
    with_core(&state, |core| {
        core.pdf_text_hit(&pdf_path, page_index, x, y)
            .map_err(|error| error.to_string())
    })
}

#[tauri::command]
async fn build_layout_map(
    pdf_path: PathBuf,
    state: State<'_, AppState>,
) -> Result<LayoutMapArtifact, String> {
    let request = with_core(&state, |core| {
        core.prepare_layout_map(&pdf_path)
            .map_err(|error| error.to_string())
    })?;
    vt_layout_map::build_layout_map(request)
        .await
        .map_err(|error| error.to_string())
}

#[tauri::command]
async fn detect_toolchain(state: State<'_, AppState>) -> Result<Vec<ToolInfo>, String> {
    with_core(&state, |_core| Ok(()))?;
    Ok(vt_compiler::detect_toolchain().await)
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .manage(AppState::default())
        .plugin(tauri_plugin_deep_link::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_opener::init())
        .setup(|app| {
            #[cfg(any(target_os = "linux", all(debug_assertions, windows)))]
            app.deep_link().register_all()?;

            if let Some(urls) = app.deep_link().get_current()? {
                for url in urls {
                    queue_visualtex_uri(app.handle(), url.to_string());
                }
            }
            let app_handle = app.handle().clone();
            app.deep_link().on_open_url(move |event| {
                for url in event.urls() {
                    queue_visualtex_uri(&app_handle, url.to_string());
                }
            });

            if let Some(path) = startup_project_argument(std::env::args_os()) {
                match CoreService::open_project(&path) {
                    Ok(core) => {
                        let state = app.state::<AppState>();
                        if let Err(error) = install_core(core, app.handle(), &state) {
                            eprintln!(
                                "visualstudio could not install startup project {}: {error}",
                                path.display()
                            );
                        }
                    }
                    Err(error) => {
                        eprintln!(
                            "visualstudio could not open startup project {}: {error}",
                            path.display()
                        );
                    }
                }
            }
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            drain_deep_links,
            process_deep_link,
            open_project,
            init_project,
            list_project_templates,
            inspect_model_package,
            install_model_package,
            list_model_packages,
            activate_model_package,
            remove_model_package,
            ocr_health,
            recognize_formula_image,
            recognize_document_image,
            create_ocr_project,
            init_project_template,
            root_snapshot,
            list_files,
            open_file,
            apply_text_edit,
            apply_visual_edit,
            apply_node_attributes,
            undo,
            redo,
            save,
            check_external_changes,
            resolve_external_conflict,
            project_index,
            project_dependencies,
            search_project,
            preview_project_replace,
            preview_symbol_rename,
            apply_project_replace,
            compile_project,
            forward_search,
            inverse_search,
            pdf_document_info,
            render_pdf,
            pdf_text_hit,
            build_layout_map,
            detect_toolchain
        ])
        .run(tauri::generate_context!())
        .expect("error while running visualstudio");
}

#[cfg(test)]
mod tests {
    use super::*;
    use image::{Rgba, RgbaImage};
    use tempfile::tempdir;

    #[test]
    fn startup_project_argument_supports_flag_equals_and_positional_paths() {
        assert_eq!(
            startup_project_argument([
                OsString::from("visualstudio"),
                OsString::from("--project"),
                OsString::from("论文 项目"),
            ]),
            Some(PathBuf::from("论文 项目"))
        );
        assert_eq!(
            startup_project_argument([
                OsString::from("visualstudio"),
                OsString::from("--project=/tmp/含 空格"),
            ]),
            Some(PathBuf::from("/tmp/含 空格"))
        );
        assert_eq!(
            startup_project_argument([
                OsString::from("visualstudio"),
                OsString::from("relative project"),
            ]),
            Some(PathBuf::from("relative project"))
        );
        assert_eq!(
            startup_project_argument([
                OsString::from("visualstudio"),
                OsString::from("visualtex://open?v=1&project=%2Ftmp%2Fpaper"),
            ]),
            None
        );
    }

    #[test]
    fn ocr_image_import_normalizes_and_deduplicates_png() {
        let project = tempdir().unwrap();
        let input = tempdir().unwrap();
        let source = input.path().join("formula.png");
        RgbaImage::from_pixel(24, 12, Rgba([255, 255, 255, 255]))
            .save(&source)
            .unwrap();

        let first = import_ocr_image(project.path(), &source).unwrap();
        let second = import_ocr_image(project.path(), &source).unwrap();
        assert_eq!(first, second);
        assert!(first.starts_with(project.path().canonicalize().unwrap()));
        assert_eq!(
            first.extension().and_then(|value| value.to_str()),
            Some("png")
        );
        let imported = image::open(&first).unwrap();
        assert_eq!((imported.width(), imported.height()), (24, 12));
    }

    #[test]
    fn ocr_image_import_rejects_oversized_file_before_decode() {
        let project = tempdir().unwrap();
        let input = tempdir().unwrap();
        let source = input.path().join("oversized.png");
        let file = fs::File::create(&source).unwrap();
        file.set_len(64 * 1024 * 1024 + 1).unwrap();
        let error = import_ocr_image(project.path(), &source).unwrap_err();
        assert!(error.contains("64 MiB"));
    }

    #[test]
    fn ocr_image_import_rejects_dimensions_over_limit() {
        let project = tempdir().unwrap();
        let input = tempdir().unwrap();
        let source = input.path().join("too-wide.png");
        RgbaImage::from_pixel(12_001, 1, Rgba([0, 0, 0, 255]))
            .save(&source)
            .unwrap();
        assert!(import_ocr_image(project.path(), &source).is_err());
    }

    #[cfg(unix)]
    #[test]
    fn ocr_image_import_rejects_symbolic_links() {
        use std::os::unix::fs::symlink;

        let project = tempdir().unwrap();
        let input = tempdir().unwrap();
        let target = input.path().join("target.png");
        let link = input.path().join("link.png");
        RgbaImage::from_pixel(4, 4, Rgba([0, 0, 0, 255]))
            .save(&target)
            .unwrap();
        symlink(&target, &link).unwrap();
        let error = import_ocr_image(project.path(), &link).unwrap_err();
        assert!(error.contains("symbolic link"));
    }
}
