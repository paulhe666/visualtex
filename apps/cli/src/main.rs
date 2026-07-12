use std::path::{Path, PathBuf};
use std::process::Command as ProcessCommand;

use clap::{Parser, Subcommand, ValueEnum};
use vt_core::CoreService;
use vt_rpc::RpcServer;

#[derive(Debug, Parser)]
#[command(
    name = "visualtex",
    version,
    about = "VisualTeX Next local CLI and bridge"
)]
struct Cli {
    /// Directory containing installed VisualTeX model packages.
    #[arg(long, global = true, value_name = "DIR")]
    models_root: Option<PathBuf>,
    #[command(subcommand)]
    command: Command,
}

#[derive(Clone, Copy, Debug, ValueEnum)]
enum CliModelKind {
    Formula,
    Layout,
    Text,
    Table,
}

impl From<CliModelKind> for vt_models::ModelKind {
    fn from(value: CliModelKind) -> Self {
        match value {
            CliModelKind::Formula => Self::FormulaOcr,
            CliModelKind::Layout => Self::LayoutOcr,
            CliModelKind::Text => Self::TextOcr,
            CliModelKind::Table => Self::TableOcr,
        }
    }
}

#[derive(Debug, Subcommand)]
enum Command {
    /// Create a standard offline LaTeX project.
    Init { path: PathBuf },
    /// Open a project in the VisualTeX desktop application.
    Open { path: PathBuf },
    /// Execute a versioned visualtex:// URI action.
    OpenUri { uri: String },
    /// Print a project snapshot as JSON.
    Inspect { path: PathBuf },
    /// Print the Tree-sitter-derived project include graph as JSON.
    Dependencies { path: PathBuf },
    /// Compile a project with its configured local TeX toolchain.
    Compile { path: PathBuf },
    /// Compare two PDFs using PDFium-rendered pixels.
    PdfDiff {
        left_pdf: PathBuf,
        right_pdf: PathBuf,
        #[arg(long, default_value_t = 1440)]
        width: u32,
        #[arg(long, default_value_t = 0)]
        tolerance: u8,
    },
    /// Build a shadow-instrumented node-to-PDF layout map.
    LayoutMap { path: PathBuf, pdf_path: PathBuf },
    /// Resolve a source line to one or more PDF rectangles using SyncTeX.
    ForwardSearch {
        path: PathBuf,
        source_file: PathBuf,
        line: u32,
        column: u32,
        pdf_path: PathBuf,
    },
    /// Resolve a PDF coordinate to a source location using SyncTeX.
    InverseSearch {
        path: PathBuf,
        pdf_path: PathBuf,
        page: u32,
        x: f32,
        y: f32,
    },
    /// Inspect and verify an offline model package without installing it.
    ModelInspect { source: PathBuf },
    /// Install a verified offline model package.
    ModelInstall { source: PathBuf },
    /// List installed offline model packages and the active selections.
    ModelList,
    /// Select an installed model version for a capability.
    ModelActivate {
        kind: CliModelKind,
        id: String,
        version: String,
    },
    /// Remove an installed model version.
    ModelRemove { id: String, version: String },
    /// Check the optional local OCR worker and configured backend.
    OcrHealth {
        path: PathBuf,
        #[arg(long)]
        mock: bool,
    },
    /// Recognize a formula image with the optional local OCR worker.
    OcrFormula {
        path: PathBuf,
        image: PathBuf,
        #[arg(long)]
        mock: bool,
    },
    /// Recognize a full document page with the optional local OCR worker.
    OcrDocument {
        path: PathBuf,
        image: PathBuf,
        #[arg(long)]
        mock: bool,
    },
    /// Detect local TeX tools.
    Doctor { path: Option<PathBuf> },
    /// Run a token-authenticated local JSON-RPC bridge for TeXstudio and adapters.
    BridgeServe { path: PathBuf },
    /// Print the active bridge discovery record for a project.
    BridgeStatus { path: PathBuf },
    /// Send one JSON-RPC method through the active local bridge.
    BridgeRequest {
        path: PathBuf,
        method: String,
        #[arg(long, default_value = "{}")]
        params: String,
        #[arg(long)]
        result_only: bool,
    },
    /// Compile through the active bridge after safely refreshing TeXstudio disk changes.
    BridgeCompile { path: PathBuf },
    /// Run forward SyncTeX through the active bridge.
    BridgeForwardSearch {
        path: PathBuf,
        source_file: PathBuf,
        line: u32,
        column: u32,
        pdf_path: PathBuf,
    },
    /// Run inverse SyncTeX through the active bridge.
    BridgeInverseSearch {
        path: PathBuf,
        pdf_path: PathBuf,
        page: u32,
        x: f32,
        y: f32,
    },
    /// Ask the active project bridge to shut down cleanly.
    BridgeShutdown { path: PathBuf },
    /// Run newline-delimited JSON-RPC 2.0 over stdin/stdout.
    Rpc {
        path: PathBuf,
        #[arg(long)]
        init: bool,
    },
}

fn resolve_models_root(explicit: Option<PathBuf>) -> anyhow::Result<PathBuf> {
    if let Some(path) = explicit {
        return Ok(path);
    }
    if let Some(path) = std::env::var_os("VISUALTEX_MODELS_ROOT") {
        return Ok(PathBuf::from(path));
    }
    if let Some(path) = std::env::var_os("LOCALAPPDATA") {
        return Ok(PathBuf::from(path).join("VisualTeX/models"));
    }
    if let Some(path) = std::env::var_os("HOME") {
        return Ok(PathBuf::from(path).join(".visualtex/models"));
    }
    Ok(std::env::current_dir()?.join(".visualtex-models"))
}

fn configured_ocr_worker(
    project_root: &PathBuf,
    models_root: &PathBuf,
    mock: bool,
) -> anyhow::Result<vt_ocr::OcrWorkerConfig> {
    let mut config = vt_ocr::OcrWorkerConfig::bundled(project_root);
    config.mock = mock;
    if !mock
        && let Some(package) =
            vt_models::active_model(models_root, vt_models::ModelKind::FormulaOcr)?
    {
        config.formula_model_dir = Some(
            package
                .install_path
                .join(&package.manifest.entrypoint)
                .canonicalize()?,
        );
        config.formula_model_name = Some(format!(
            "{}@{}",
            package.manifest.id, package.manifest.version
        ));
    }
    if !mock
        && let Some(package) =
            vt_models::active_model(models_root, vt_models::ModelKind::LayoutOcr)?
    {
        config.document_pipeline_config = Some(
            package
                .install_path
                .join(&package.manifest.entrypoint)
                .canonicalize()?,
        );
        config.document_package_root = Some(package.install_path.canonicalize()?);
        config.document_model_name = Some(format!(
            "{}@{}",
            package.manifest.id, package.manifest.version
        ));
    }
    Ok(config)
}

fn launch_desktop(project_root: &Path) -> anyhow::Result<()> {
    let project_root = project_root.canonicalize()?;
    let _ = CoreService::open_project(&project_root)?;

    for variable in ["VISUALSTUDIO_BIN", "VISUALTEX_DESKTOP_BIN"] {
        if let Some(executable) = std::env::var_os(variable) {
            ProcessCommand::new(executable)
                .arg("--project")
                .arg(&project_root)
                .spawn()?;
            return Ok(());
        }
    }

    if let Ok(current) = std::env::current_exe()
        && let Some(directory) = current.parent()
    {
        let file_names: &[&str] = if cfg!(windows) {
            &["visualstudio.exe", "visualtex-desktop.exe"]
        } else {
            &["visualstudio", "visualtex-desktop"]
        };
        for file_name in file_names {
            let sibling = directory.join(file_name);
            if sibling.is_file() {
                ProcessCommand::new(sibling)
                    .arg("--project")
                    .arg(&project_root)
                    .spawn()?;
                return Ok(());
            }
        }
    }

    #[cfg(target_os = "macos")]
    {
        let status = ProcessCommand::new("open")
            .arg("-a")
            .arg("visualstudio")
            .arg("--args")
            .arg("--project")
            .arg(&project_root)
            .status()?;
        if !status.success() {
            anyhow::bail!("macOS could not launch the visualstudio application bundle");
        }
        Ok(())
    }

    #[cfg(not(target_os = "macos"))]
    {
        ProcessCommand::new("visualstudio")
            .arg("--project")
            .arg(&project_root)
            .spawn()?;
        Ok(())
    }
}

async fn execute_visualtex_uri(uri: &str) -> anyhow::Result<Option<serde_json::Value>> {
    match vt_uri::VisualTexUriAction::parse(uri)? {
        vt_uri::VisualTexUriAction::Open { project } => {
            launch_desktop(&project)?;
            Ok(None)
        }
        vt_uri::VisualTexUriAction::ForwardSearch {
            project,
            source_file,
            line,
            column,
            pdf_path,
        } => {
            let core = CoreService::open_project(project)?;
            let result = core
                .forward_search(&source_file, line, column, &pdf_path)
                .await?;
            Ok(Some(serde_json::to_value(result)?))
        }
        vt_uri::VisualTexUriAction::InverseSearch {
            project,
            pdf_path,
            page,
            x,
            y,
        } => {
            let core = CoreService::open_project(project)?;
            let result = core.inverse_search(&pdf_path, page, x, y).await?;
            Ok(Some(serde_json::to_value(result)?))
        }
    }
}

async fn bridge_result(
    project_root: &Path,
    method: &str,
    params: serde_json::Value,
) -> anyhow::Result<serde_json::Value> {
    let response = vt_bridge::BridgeClient::request(
        project_root,
        vt_protocol::RpcRequest {
            jsonrpc: "2.0".into(),
            id: serde_json::json!(1),
            method: method.to_owned(),
            params,
        },
    )
    .await?;
    if let Some(error) = response.error {
        let data = error
            .data
            .map(|value| {
                format!(
                    "\n{}",
                    serde_json::to_string_pretty(&value).unwrap_or_default()
                )
            })
            .unwrap_or_default();
        anyhow::bail!(
            "VisualTeX bridge RPC error {}: {}{}",
            error.code,
            error.message,
            data
        );
    }
    Ok(response.result.unwrap_or(serde_json::Value::Null))
}

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    tracing_subscriber::fmt()
        .with_env_filter(
            tracing_subscriber::EnvFilter::try_from_default_env()
                .unwrap_or_else(|_| "visualtex=info".into()),
        )
        .with_writer(std::io::stderr)
        .init();

    let cli = Cli::parse();
    let models_root = resolve_models_root(cli.models_root)?;

    match cli.command {
        Command::Init { path } => {
            let core = CoreService::init_project(&path)?;
            println!("{}", serde_json::to_string_pretty(&core.root_snapshot()?)?);
        }
        Command::Open { path } => {
            launch_desktop(&path)?;
        }
        Command::OpenUri { uri } => {
            if let Some(result) = execute_visualtex_uri(&uri).await? {
                println!("{}", serde_json::to_string_pretty(&result)?);
            }
        }
        Command::Inspect { path } => {
            let core = CoreService::open_project(path)?;
            println!("{}", serde_json::to_string_pretty(&core.root_snapshot()?)?);
        }
        Command::Dependencies { path } => {
            let core = CoreService::open_project(path)?;
            println!(
                "{}",
                serde_json::to_string_pretty(&core.project_dependencies())?
            );
        }
        Command::Compile { path } => {
            let mut core = CoreService::open_project(path)?;
            let artifact = core.compile().await?;
            println!("{}", serde_json::to_string_pretty(&artifact)?);
            if !matches!(artifact.status, vt_protocol::CompileStatus::Succeeded) {
                std::process::exit(2);
            }
        }
        Command::PdfDiff {
            left_pdf,
            right_pdf,
            width,
            tolerance,
        } => {
            let cache = std::env::temp_dir().join("visualtex-pdf-diff-cache");
            let service = vt_pdf::PdfService::new(cache);
            let report = service.compare_documents(left_pdf, right_pdf, width, tolerance)?;
            println!("{}", serde_json::to_string_pretty(&report)?);
        }
        Command::LayoutMap { path, pdf_path } => {
            let mut core = CoreService::open_project(path)?;
            let artifact = core.build_layout_map(&pdf_path).await?;
            println!("{}", serde_json::to_string_pretty(&artifact)?);
            if artifact.compile_status != vt_protocol::CompileStatus::Succeeded {
                std::process::exit(2);
            }
        }
        Command::ForwardSearch {
            path,
            source_file,
            line,
            column,
            pdf_path,
        } => {
            let core = CoreService::open_project(path)?;
            let result = core
                .forward_search(&source_file, line, column, &pdf_path)
                .await?;
            println!("{}", serde_json::to_string_pretty(&result)?);
        }
        Command::InverseSearch {
            path,
            pdf_path,
            page,
            x,
            y,
        } => {
            let core = CoreService::open_project(path)?;
            let result = core.inverse_search(&pdf_path, page, x, y).await?;
            println!("{}", serde_json::to_string_pretty(&result)?);
        }
        Command::ModelInspect { source } => {
            println!(
                "{}",
                serde_json::to_string_pretty(&vt_models::inspect_package(source)?)?
            );
        }
        Command::ModelInstall { source } => {
            println!(
                "{}",
                serde_json::to_string_pretty(&vt_models::install_package(source, &models_root)?)?
            );
        }
        Command::ModelList => {
            println!(
                "{}",
                serde_json::to_string_pretty(&serde_json::json!({
                    "modelsRoot": models_root,
                    "installed": vt_models::list_installed(&models_root)?,
                    "active": {
                        "formulaOcr": vt_models::active_model(&models_root, vt_models::ModelKind::FormulaOcr)?,
                        "layoutOcr": vt_models::active_model(&models_root, vt_models::ModelKind::LayoutOcr)?,
                        "textOcr": vt_models::active_model(&models_root, vt_models::ModelKind::TextOcr)?,
                        "tableOcr": vt_models::active_model(&models_root, vt_models::ModelKind::TableOcr)?,
                    }
                }))?
            );
        }
        Command::ModelActivate { kind, id, version } => {
            println!(
                "{}",
                serde_json::to_string_pretty(&vt_models::set_active_model(
                    &models_root,
                    kind.into(),
                    &id,
                    &version
                )?)?
            );
        }
        Command::ModelRemove { id, version } => {
            println!(
                "{}",
                serde_json::to_string_pretty(&serde_json::json!({
                    "removed": vt_models::remove_installed(&models_root, &id, &version)?,
                    "id": id,
                    "version": version,
                }))?
            );
        }
        Command::OcrHealth { path, mock } => {
            let config = configured_ocr_worker(&path, &models_root, mock)?;
            let mut worker = vt_ocr::OcrWorker::spawn(config).await?;
            println!("{}", serde_json::to_string_pretty(&worker.health().await?)?);
            worker.shutdown().await?;
        }
        Command::OcrFormula { path, image, mock } => {
            let config = configured_ocr_worker(&path, &models_root, mock)?;
            let mut worker = vt_ocr::OcrWorker::spawn(config).await?;
            println!(
                "{}",
                serde_json::to_string_pretty(&worker.recognize_formula(image).await?)?
            );
            worker.shutdown().await?;
        }
        Command::OcrDocument { path, image, mock } => {
            let config = configured_ocr_worker(&path, &models_root, mock)?;
            let mut worker = vt_ocr::OcrWorker::spawn(config).await?;
            println!(
                "{}",
                serde_json::to_string_pretty(&worker.recognize_document(image).await?)?
            );
            worker.shutdown().await?;
        }
        Command::Doctor { path } => {
            let core = match path {
                Some(path) => CoreService::open_project(path)?,
                None => {
                    let current = std::env::current_dir()?;
                    CoreService::open_project(current)?
                }
            };
            println!(
                "{}",
                serde_json::to_string_pretty(&core.detect_toolchain().await)?
            );
        }
        Command::BridgeServe { path } => {
            let server = vt_bridge::BridgeServer::bind(path, Some(models_root.clone())).await?;
            println!("{}", serde_json::to_string_pretty(server.discovery())?);
            server.run().await?;
        }
        Command::BridgeStatus { path } => {
            println!(
                "{}",
                serde_json::to_string_pretty(&vt_bridge::BridgeClient::discovery(path)?)?
            );
        }
        Command::BridgeRequest {
            path,
            method,
            params,
            result_only,
        } => {
            let params = serde_json::from_str(&params)?;
            let response = vt_bridge::BridgeClient::request(
                path,
                vt_protocol::RpcRequest {
                    jsonrpc: "2.0".into(),
                    id: serde_json::json!(1),
                    method,
                    params,
                },
            )
            .await?;
            if let Some(error) = &response.error {
                eprintln!(
                    "VisualTeX bridge RPC error {}: {}",
                    error.code, error.message
                );
                if let Some(data) = &error.data {
                    eprintln!("{}", serde_json::to_string_pretty(data)?);
                }
                std::process::exit(2);
            }
            if result_only {
                println!(
                    "{}",
                    serde_json::to_string_pretty(
                        response.result.as_ref().unwrap_or(&serde_json::Value::Null)
                    )?
                );
            } else {
                println!("{}", serde_json::to_string_pretty(&response)?);
            }
        }
        Command::BridgeCompile { path } => {
            bridge_result(&path, "project.refreshFromDisk", serde_json::json!({})).await?;
            let result = bridge_result(&path, "project.compile", serde_json::json!({})).await?;
            println!("{}", serde_json::to_string_pretty(&result)?);
            if result.get("status").and_then(serde_json::Value::as_str) != Some("succeeded") {
                std::process::exit(2);
            }
        }
        Command::BridgeForwardSearch {
            path,
            source_file,
            line,
            column,
            pdf_path,
        } => {
            bridge_result(&path, "project.refreshFromDisk", serde_json::json!({})).await?;
            let result = bridge_result(
                &path,
                "synctex.forwardSearch",
                serde_json::json!({
                    "sourceFile": source_file,
                    "line": line,
                    "column": column,
                    "pdfPath": pdf_path,
                }),
            )
            .await?;
            println!("{}", serde_json::to_string_pretty(&result)?);
        }
        Command::BridgeInverseSearch {
            path,
            pdf_path,
            page,
            x,
            y,
        } => {
            let result = bridge_result(
                &path,
                "synctex.inverseSearch",
                serde_json::json!({
                    "pdfPath": pdf_path,
                    "page": page,
                    "x": x,
                    "y": y,
                }),
            )
            .await?;
            println!("{}", serde_json::to_string_pretty(&result)?);
        }
        Command::BridgeShutdown { path } => {
            let response = vt_bridge::BridgeClient::shutdown(path).await?;
            println!("{}", serde_json::to_string_pretty(&response)?);
            if response.error.is_some() {
                std::process::exit(2);
            }
        }
        Command::Rpc { path, init } => {
            let server = if init {
                RpcServer::init_with_models(path, Some(models_root.clone()))?
            } else {
                RpcServer::open_with_models(path, Some(models_root.clone()))?
            };
            server.run_stdio().await?;
        }
    }
    Ok(())
}
