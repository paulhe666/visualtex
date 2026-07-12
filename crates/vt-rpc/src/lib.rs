use std::path::{Path, PathBuf};

use anyhow::Context;
use serde::Deserialize;
use serde_json::{Value, json};
use tokio::io::{AsyncBufReadExt, AsyncWriteExt, BufReader};
use vt_core::CoreService;
use vt_models::ModelKind;
use vt_protocol::{
    ExternalConflictResolution, ExternalFileChange, FileId, NodeAttributesPatch, NodeId,
    PROTOCOL_VERSION, PdfRenderRequest, ProjectReplacePlan, ProjectReplaceRequest,
    ProjectSearchRequest, Revision, RpcError, RpcRequest, RpcResponse, SymbolRenameRequest,
    TextEdit,
};

pub struct RpcServer {
    core: CoreService,
    models_root: Option<PathBuf>,
}

impl RpcServer {
    pub fn open(root: impl AsRef<Path>) -> anyhow::Result<Self> {
        Self::open_with_models(root, None::<PathBuf>)
    }

    pub fn open_with_models(
        root: impl AsRef<Path>,
        models_root: impl Into<Option<PathBuf>>,
    ) -> anyhow::Result<Self> {
        Ok(Self {
            core: CoreService::open_project(root).context("failed to open VisualTeX project")?,
            models_root: models_root.into(),
        })
    }

    pub fn init(root: impl AsRef<Path>) -> anyhow::Result<Self> {
        Self::init_with_models(root, None::<PathBuf>)
    }

    pub fn init_with_models(
        root: impl AsRef<Path>,
        models_root: impl Into<Option<PathBuf>>,
    ) -> anyhow::Result<Self> {
        Ok(Self {
            core: CoreService::init_project(root)
                .context("failed to initialize VisualTeX project")?,
            models_root: models_root.into(),
        })
    }

    fn model_available(&self, kind: ModelKind) -> anyhow::Result<bool> {
        let Some(root) = &self.models_root else {
            return Ok(false);
        };
        Ok(vt_models::active_model(root, kind)?.is_some())
    }

    fn configured_ocr_worker(&self) -> anyhow::Result<vt_ocr::OcrWorkerConfig> {
        let mut config = vt_ocr::OcrWorkerConfig::bundled(self.core.project_root());
        let Some(root) = &self.models_root else {
            return Ok(config);
        };
        if let Some(package) = vt_models::active_model(root, ModelKind::FormulaOcr)? {
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
        if let Some(package) = vt_models::active_model(root, ModelKind::LayoutOcr)? {
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

    pub async fn run_stdio(mut self) -> anyhow::Result<()> {
        let stdin = tokio::io::stdin();
        let mut lines = BufReader::new(stdin).lines();
        let mut stdout = tokio::io::stdout();
        while let Some(line) = lines.next_line().await? {
            if line.trim().is_empty() {
                continue;
            }
            let response = match serde_json::from_str::<RpcRequest>(&line) {
                Ok(request) => self.handle(request).await,
                Err(error) => RpcResponse {
                    jsonrpc: "2.0".into(),
                    id: Value::Null,
                    result: None,
                    error: Some(RpcError {
                        code: -32700,
                        message: "Parse error".into(),
                        data: Some(json!({ "detail": error.to_string() })),
                    }),
                },
            };
            let mut encoded = serde_json::to_vec(&response)?;
            encoded.push(b'\n');
            stdout.write_all(&encoded).await?;
            stdout.flush().await?;
        }
        Ok(())
    }

    pub async fn handle(&mut self, request: RpcRequest) -> RpcResponse {
        let id = request.id.clone();
        let result = self.dispatch(&request.method, request.params).await;
        match result {
            Ok(result) => RpcResponse {
                jsonrpc: "2.0".into(),
                id,
                result: Some(result),
                error: None,
            },
            Err(error) => RpcResponse {
                jsonrpc: "2.0".into(),
                id,
                result: None,
                error: Some(RpcError {
                    code: -32000,
                    message: error.to_string(),
                    data: None,
                }),
            },
        }
    }

    async fn dispatch(&mut self, method: &str, params: Value) -> anyhow::Result<Value> {
        match method {
            "initialize" => {
                let formula_ocr = self.model_available(ModelKind::FormulaOcr)?;
                let document_ocr = self.model_available(ModelKind::LayoutOcr)?;
                Ok(json!({
                    "protocolVersion": PROTOCOL_VERSION,
                    "capabilities": {
                        "incrementalEdits": true,
                        "incrementalSyntaxTree": true,
                        "projectDependencyGraph": true,
                        "visualEdits": true,
                        "compile": true,
                        "toolchainDetection": true,
                        "undoRedo": true,
                        "pdfiumRendering": true,
                        "pdfTiles": true,
                        "shadowLayoutMap": true,
                        "formulaOcr": formula_ocr,
                        "documentOcr": document_ocr,
                        "offlineOnly": true
                    }
                }))
            }
            "project.rootSnapshot" => Ok(serde_json::to_value(self.core.root_snapshot()?)?),
            "project.listFiles" => Ok(serde_json::to_value(self.core.list_files())?),
            "project.openFile" => {
                #[derive(Deserialize)]
                #[serde(rename_all = "camelCase")]
                struct Params {
                    path: String,
                }
                let params: Params = serde_json::from_value(params)?;
                Ok(serde_json::to_value(self.core.open_file(params.path)?)?)
            }
            "document.applyEdit" => {
                let edit: TextEdit = serde_json::from_value(params)?;
                Ok(serde_json::to_value(self.core.apply_text_edit(edit)?)?)
            }
            "document.applyVisualEdit" => {
                #[derive(Deserialize)]
                #[serde(rename_all = "camelCase")]
                struct Params {
                    file_id: FileId,
                    base_revision: Revision,
                    node_id: NodeId,
                    content: String,
                }
                let params: Params = serde_json::from_value(params)?;
                Ok(serde_json::to_value(self.core.apply_visual_edit(
                    params.file_id,
                    params.base_revision,
                    params.node_id,
                    params.content,
                )?)?)
            }
            "document.applyNodeAttributes" => {
                #[derive(Deserialize)]
                #[serde(rename_all = "camelCase")]
                struct Params {
                    file_id: FileId,
                    base_revision: Revision,
                    node_id: NodeId,
                    patch: NodeAttributesPatch,
                }
                let params: Params = serde_json::from_value(params)?;
                Ok(serde_json::to_value(self.core.apply_node_attributes(
                    params.file_id,
                    params.base_revision,
                    params.node_id,
                    params.patch,
                )?)?)
            }
            "document.confirmSaved" => {
                #[derive(Deserialize)]
                #[serde(rename_all = "camelCase")]
                struct Params {
                    file_id: FileId,
                }
                let params: Params = serde_json::from_value(params)?;
                self.core.confirm_external_save(params.file_id)?;
                Ok(json!({ "ok": true }))
            }
            "document.undo" | "document.redo" | "document.save" => {
                #[derive(Deserialize)]
                #[serde(rename_all = "camelCase")]
                struct Params {
                    file_id: FileId,
                }
                let params: Params = serde_json::from_value(params)?;
                match method {
                    "document.undo" => Ok(serde_json::to_value(self.core.undo(params.file_id)?)?),
                    "document.redo" => Ok(serde_json::to_value(self.core.redo(params.file_id)?)?),
                    _ => {
                        self.core.save(params.file_id)?;
                        Ok(json!({ "ok": true }))
                    }
                }
            }
            "project.saveAll" => {
                self.core.save_all()?;
                Ok(json!({ "ok": true }))
            }
            "project.checkExternalChanges" => {
                Ok(serde_json::to_value(self.core.check_external_changes()?)?)
            }
            "project.refreshFromDisk" => {
                let report = self.core.check_external_changes()?;
                if !report.conflicts.is_empty() {
                    anyhow::bail!(
                        "cannot refresh from disk: {} file conflict(s) require explicit resolution",
                        report.conflicts.len()
                    );
                }
                Ok(serde_json::to_value(report)?)
            }
            "project.resolveExternalConflict" => {
                #[derive(Deserialize)]
                #[serde(rename_all = "camelCase")]
                struct Params {
                    change: ExternalFileChange,
                    resolution: ExternalConflictResolution,
                }
                let params: Params = serde_json::from_value(params)?;
                Ok(serde_json::to_value(self.core.resolve_external_conflict(
                    &params.change,
                    params.resolution,
                )?)?)
            }
            "project.index" => Ok(serde_json::to_value(self.core.project_index()?)?),
            "project.dependencies" => Ok(serde_json::to_value(self.core.project_dependencies())?),
            "project.search" => {
                let request: ProjectSearchRequest = serde_json::from_value(params)?;
                Ok(serde_json::to_value(self.core.search_project(&request)?)?)
            }
            "project.previewReplace" => {
                let request: ProjectReplaceRequest = serde_json::from_value(params)?;
                Ok(serde_json::to_value(
                    self.core.preview_project_replace(&request)?,
                )?)
            }
            "project.previewSymbolRename" => {
                let request: SymbolRenameRequest = serde_json::from_value(params)?;
                Ok(serde_json::to_value(
                    self.core.preview_symbol_rename(&request)?,
                )?)
            }
            "project.applyReplace" => {
                let plan: ProjectReplacePlan = serde_json::from_value(params)?;
                Ok(serde_json::to_value(
                    self.core.apply_project_replace(plan)?,
                )?)
            }
            "project.compile" => Ok(serde_json::to_value(self.core.compile().await?)?),
            "synctex.forwardSearch" => {
                #[derive(Deserialize)]
                #[serde(rename_all = "camelCase")]
                struct Params {
                    source_file: String,
                    line: u32,
                    column: u32,
                    pdf_path: String,
                }
                let params: Params = serde_json::from_value(params)?;
                Ok(serde_json::to_value(
                    self.core
                        .forward_search(
                            Path::new(&params.source_file),
                            params.line,
                            params.column,
                            Path::new(&params.pdf_path),
                        )
                        .await?,
                )?)
            }
            "synctex.inverseSearch" => {
                #[derive(Deserialize)]
                #[serde(rename_all = "camelCase")]
                struct Params {
                    pdf_path: String,
                    page: u32,
                    x: f32,
                    y: f32,
                }
                let params: Params = serde_json::from_value(params)?;
                Ok(serde_json::to_value(
                    self.core
                        .inverse_search(
                            Path::new(&params.pdf_path),
                            params.page,
                            params.x,
                            params.y,
                        )
                        .await?,
                )?)
            }
            "pdf.documentInfo" => {
                #[derive(Deserialize)]
                #[serde(rename_all = "camelCase")]
                struct Params {
                    pdf_path: String,
                }
                let params: Params = serde_json::from_value(params)?;
                Ok(serde_json::to_value(
                    self.core.pdf_document_info(Path::new(&params.pdf_path))?,
                )?)
            }
            "pdf.render" => {
                let request: PdfRenderRequest = serde_json::from_value(params)?;
                Ok(serde_json::to_value(self.core.render_pdf(request)?)?)
            }
            "layout.build" => {
                #[derive(Deserialize)]
                #[serde(rename_all = "camelCase")]
                struct Params {
                    pdf_path: String,
                }
                let params: Params = serde_json::from_value(params)?;
                Ok(serde_json::to_value(
                    self.core
                        .build_layout_map(Path::new(&params.pdf_path))
                        .await?,
                )?)
            }
            "ocr.health" => {
                let mut worker = vt_ocr::OcrWorker::spawn(self.configured_ocr_worker()?).await?;
                let formula = worker.health().await;
                let document = worker.document_health().await;
                let _ = worker.shutdown().await;
                Ok(json!({
                    "formula": formula?,
                    "document": document?,
                    "offlineOnly": true
                }))
            }
            "ocr.recognizeFormula" | "ocr.recognizeDocument" => {
                #[derive(Deserialize)]
                #[serde(rename_all = "camelCase")]
                struct Params {
                    source_path: String,
                }
                let params: Params = serde_json::from_value(params)?;
                let project_root = self.core.project_root().to_path_buf();
                let source = PathBuf::from(params.source_path);
                let imported = tokio::task::spawn_blocking(move || {
                    vt_ocr::import_project_image(project_root, source)
                })
                .await??;
                let mut worker = vt_ocr::OcrWorker::spawn(self.configured_ocr_worker()?).await?;
                if method == "ocr.recognizeFormula" {
                    let result = worker.recognize_formula(&imported).await;
                    let _ = worker.shutdown().await;
                    Ok(serde_json::to_value(result?)?)
                } else {
                    let result = worker.recognize_document(&imported).await;
                    let _ = worker.shutdown().await;
                    let mut result = result?;
                    result.image_path = Some(imported);
                    Ok(serde_json::to_value(result)?)
                }
            }
            "toolchain.detect" => Ok(serde_json::to_value(self.core.detect_toolchain().await)?),
            _ => anyhow::bail!("method not found: {method}"),
        }
    }
}

#[cfg(test)]
mod tests {
    use std::fs;

    use crate::RpcServer;
    use serde_json::json;
    use tempfile::tempdir;
    use vt_protocol::{PROTOCOL_VERSION, RpcRequest};

    #[tokio::test]
    async fn exposes_incremental_syntax_capabilities_and_dependency_graph() {
        let temp = tempdir().unwrap();
        fs::write(
            temp.path().join("main.tex"),
            "\\documentclass{article}\n\\begin{document}\n\\input{chapter}\n\\end{document}\n",
        )
        .unwrap();
        fs::write(temp.path().join("chapter.tex"), "Chapter text.\n").unwrap();
        let mut server = RpcServer::open(temp.path()).unwrap();

        let initialized = server
            .handle(RpcRequest {
                jsonrpc: "2.0".into(),
                id: json!(1),
                method: "initialize".into(),
                params: json!({ "protocolVersion": PROTOCOL_VERSION }),
            })
            .await;
        let capabilities = &initialized.result.unwrap()["capabilities"];
        assert_eq!(capabilities["incrementalSyntaxTree"], true);
        assert_eq!(capabilities["projectDependencyGraph"], true);

        let graph = server
            .handle(RpcRequest {
                jsonrpc: "2.0".into(),
                id: json!(2),
                method: "project.dependencies".into(),
                params: json!({}),
            })
            .await
            .result
            .unwrap();
        assert_eq!(graph["edges"].as_array().unwrap().len(), 1);
        assert_eq!(graph["edges"][0]["sourceFile"], "main.tex");
        assert_eq!(graph["edges"][0]["targetFile"], "chapter.tex");
        assert_eq!(graph["edges"][0]["resolved"], true);
        assert!(graph["cycles"].as_array().unwrap().is_empty());
    }
}
