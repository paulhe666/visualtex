use std::collections::{BTreeSet, HashMap};
use std::fs;
use std::path::{Path, PathBuf};
use std::time::Duration;

use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use vt_buffer::{BufferError, DocumentBuffer};
use vt_compiler::{CompileRequest, CompilerError};
use vt_index::IndexError;
use vt_latex_semantic::SemanticDocument;
use vt_latex_syntax::{SyntaxDocument, SyntaxError, build_dependency_graph};
use vt_layout_map::{LayoutMapError, LayoutMapRequest};
use vt_pdf::{PdfError, PdfService};
use vt_project::{Project, ProjectError, atomic_write};
use vt_protocol::{
    AppliedEdit, CompileArtifact, DocumentOcrResult, DocumentSnapshot, EditOrigin,
    ExternalChangeKind, ExternalChangeReport, ExternalConflictOutcome, ExternalConflictResolution,
    ExternalFileChange, FileId, ForwardSearchResult, InverseSearchResult, LayoutMapArtifact,
    NodeAttributesPatch, NodeId, OperationId, PdfDocumentInfo, PdfRect, PdfRenderRequest,
    PdfRenderedImage, PdfTextHit, PdfTextLine, ProjectDependencyGraph, ProjectIndex,
    ProjectReplaceFilePlan, ProjectReplaceOutcome, ProjectReplacePlan, ProjectReplaceRequest,
    ProjectSearchMatch, ProjectSearchRequest, ProjectTemplateSummary, ProjectTextReplacement,
    Revision, SymbolKind, SymbolRenameKind, SymbolRenameRequest, TextEdit, VisualNode, VisualPatch,
};

#[derive(Debug, thiserror::Error)]
pub enum CoreError {
    #[error(transparent)]
    Project(#[from] ProjectError),
    #[error(transparent)]
    Buffer(#[from] BufferError),
    #[error(transparent)]
    Compiler(#[from] CompilerError),
    #[error(transparent)]
    SyncTex(#[from] vt_synctex::SyncTexError),
    #[error(transparent)]
    Pdf(#[from] PdfError),
    #[error(transparent)]
    LayoutMap(#[from] LayoutMapError),
    #[error(transparent)]
    Syntax(#[from] SyntaxError),
    #[error(transparent)]
    Index(#[from] IndexError),
    #[error("semantic node not found: {0:?}")]
    NodeNotFound(NodeId),
    #[error("node {0:?} cannot be edited visually")]
    NodeNotEditable(NodeId),
    #[error("operation recovery failed: {0}")]
    Recovery(String),
    #[error("replacement plan is stale for {0}")]
    ReplacePlanStale(PathBuf),
    #[error("replacement plan is truncated; refine the query before applying it")]
    ReplacePlanTruncated,
    #[error("invalid symbol rename: {0}")]
    InvalidSymbolRename(String),
    #[error("the target symbol already exists: {0}")]
    SymbolRenameConflict(String),
    #[error(transparent)]
    Io(#[from] std::io::Error),
    #[error(transparent)]
    Json(#[from] serde_json::Error),
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct EditOutcome {
    pub revision: Revision,
    pub patch: VisualPatch,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct RecoveryRecord {
    path: PathBuf,
    base_disk_sha256: Option<String>,
    resulting_text: String,
    resulting_revision: Revision,
}

#[derive(Debug)]
pub struct CoreService {
    project: Project,
    syntax: HashMap<FileId, SyntaxDocument>,
    semantics: HashMap<FileId, SemanticDocument>,
    dependency_graph: ProjectDependencyGraph,
}

impl CoreService {
    pub fn open_project(root: impl AsRef<Path>) -> Result<Self, CoreError> {
        let project = Project::open(root)?;
        Self::from_project(project)
    }

    pub fn init_project(root: impl AsRef<Path>) -> Result<Self, CoreError> {
        let project = Project::init(root)?;
        Self::from_project(project)
    }

    pub fn init_project_with_template(
        root: impl AsRef<Path>,
        template_id: &str,
    ) -> Result<Self, CoreError> {
        let project = Project::init_with_template(root, template_id)?;
        Self::from_project(project)
    }

    pub fn init_ocr_project(
        root: impl AsRef<Path>,
        latex_body: &str,
        original_image: impl AsRef<Path>,
        ocr_document: DocumentOcrResult,
    ) -> Result<Self, CoreError> {
        let project = Project::init_ocr_project(root, latex_body, original_image, ocr_document)?;
        Self::from_project(project)
    }

    pub fn project_templates() -> Vec<ProjectTemplateSummary> {
        vt_project::built_in_templates()
    }

    fn from_project(project: Project) -> Result<Self, CoreError> {
        let mut service = Self {
            project,
            syntax: HashMap::new(),
            semantics: HashMap::new(),
            dependency_graph: ProjectDependencyGraph::default(),
        };
        service.recover_pending()?;
        service.open_and_parse_project_sources()?;
        Ok(service)
    }

    pub fn project_root(&self) -> &Path {
        &self.project.root
    }

    pub fn root_file_id(&self) -> Result<FileId, CoreError> {
        Ok(self.project.root_file_id()?)
    }

    pub fn list_files(&self) -> Vec<PathBuf> {
        self.project.list_source_files()
    }

    pub fn project_dependencies(&self) -> ProjectDependencyGraph {
        self.dependency_graph.clone()
    }

    pub fn open_file(&mut self, relative: impl AsRef<Path>) -> Result<DocumentSnapshot, CoreError> {
        let file_id = self.project.open_file(relative)?;
        self.refresh_document(file_id)?;
        self.snapshot(file_id)
    }

    pub fn snapshot(&self, file_id: FileId) -> Result<DocumentSnapshot, CoreError> {
        let buffer = self.project.buffer(file_id)?;
        let nodes = self
            .semantics
            .get(&file_id)
            .map(|document| document.nodes.clone())
            .unwrap_or_default();
        Ok(DocumentSnapshot {
            file_id,
            path: buffer.path.clone(),
            revision: buffer.revision,
            text: buffer.text(),
            dirty: buffer.dirty,
            nodes,
        })
    }

    pub fn root_snapshot(&self) -> Result<DocumentSnapshot, CoreError> {
        self.snapshot(self.root_file_id()?)
    }

    pub fn check_external_changes(&mut self) -> Result<ExternalChangeReport, CoreError> {
        let discovered_sources = self.discover_project_sources_without_graph()?;
        if discovered_sources {
            self.rebuild_dependency_graph();
        }
        let changes = self.project.external_changes()?;
        let mut report = ExternalChangeReport::default();
        for change in changes {
            if !change.buffer_dirty && change.kind == ExternalChangeKind::Modified {
                self.project.reload_external(change.file_id)?;
                self.refresh_document(change.file_id)?;
                report.reloaded.push(self.snapshot(change.file_id)?);
            } else {
                report.conflicts.push(change);
            }
        }
        if !report.reloaded.is_empty() {
            self.rewrite_recovery_snapshots()?;
        }
        Ok(report)
    }

    pub fn resolve_external_conflict(
        &mut self,
        change: &ExternalFileChange,
        resolution: ExternalConflictResolution,
    ) -> Result<ExternalConflictOutcome, CoreError> {
        let conflict_copy_path = match resolution {
            ExternalConflictResolution::ReloadDisk => {
                self.project.reload_external(change.file_id)?;
                self.refresh_document(change.file_id)?;
                None
            }
            ExternalConflictResolution::KeepBuffer => {
                self.project.accept_external_baseline(change.file_id)?;
                None
            }
            ExternalConflictResolution::SaveCopyAndReload => {
                let copy = self.project.save_conflict_copy(change.file_id)?;
                if change.kind == ExternalChangeKind::Deleted {
                    self.project.accept_external_baseline(change.file_id)?;
                } else {
                    self.project.reload_external(change.file_id)?;
                    self.refresh_document(change.file_id)?;
                }
                Some(copy)
            }
        };
        self.rewrite_recovery_snapshots()?;
        Ok(ExternalConflictOutcome {
            snapshot: self.snapshot(change.file_id)?,
            conflict_copy_path,
        })
    }

    pub fn apply_text_edit(&mut self, edit: TextEdit) -> Result<EditOutcome, CoreError> {
        let file_id = edit.file_id;
        let old_nodes = self
            .semantics
            .get(&file_id)
            .map(|document| document.nodes.clone())
            .unwrap_or_default();
        let mut candidate = self.project.buffer(file_id)?.clone();
        let old_source = candidate.text();
        let applied = candidate.apply(edit)?;
        self.commit_buffer_transition(file_id, old_nodes, old_source, candidate, applied)
    }

    pub fn apply_visual_edit(
        &mut self,
        file_id: FileId,
        base_revision: Revision,
        node_id: NodeId,
        content: String,
    ) -> Result<EditOutcome, CoreError> {
        let node = self
            .semantics
            .get(&file_id)
            .and_then(|document| document.nodes.iter().find(|node| node.id == node_id))
            .cloned()
            .ok_or(CoreError::NodeNotFound(node_id))?;
        let buffer = self.project.buffer(file_id)?;
        if buffer.revision != base_revision {
            return Err(BufferError::StaleRevision {
                expected: buffer.revision,
                received: base_revision,
            }
            .into());
        }
        let source = buffer.text();
        let (start_byte, end_byte, replacement) = serialize_visual_change(&source, &node, content)
            .ok_or(CoreError::NodeNotEditable(node_id))?;
        self.apply_text_edit(TextEdit {
            operation_id: OperationId::new(),
            origin: match node.kind {
                vt_protocol::NodeKind::InlineMath | vt_protocol::NodeKind::DisplayMath => {
                    EditOrigin::MathEditor
                }
                _ => EditOrigin::VisualEditor,
            },
            file_id,
            base_revision,
            start_byte,
            end_byte,
            replacement,
        })
    }

    pub fn apply_node_attributes(
        &mut self,
        file_id: FileId,
        base_revision: Revision,
        node_id: NodeId,
        patch: NodeAttributesPatch,
    ) -> Result<EditOutcome, CoreError> {
        let node = self
            .semantics
            .get(&file_id)
            .and_then(|document| document.nodes.iter().find(|node| node.id == node_id))
            .cloned()
            .ok_or(CoreError::NodeNotFound(node_id))?;
        if !matches!(
            node.kind,
            vt_protocol::NodeKind::Figure | vt_protocol::NodeKind::Table
        ) {
            return Err(CoreError::NodeNotEditable(node_id));
        }
        let buffer = self.project.buffer(file_id)?;
        if buffer.revision != base_revision {
            return Err(BufferError::StaleRevision {
                expected: buffer.revision,
                received: base_revision,
            }
            .into());
        }
        let source = buffer.text();
        let start = node.source.start_byte;
        let end = node.source.end_byte;
        let fragment = source
            .get(start..end)
            .ok_or(CoreError::NodeNotEditable(node_id))?;
        let replacement = apply_attribute_patch(fragment, &node, &patch)
            .ok_or(CoreError::NodeNotEditable(node_id))?;
        self.apply_text_edit(TextEdit {
            operation_id: OperationId::new(),
            origin: EditOrigin::VisualEditor,
            file_id,
            base_revision,
            start_byte: start,
            end_byte: end,
            replacement,
        })
    }

    pub fn undo(&mut self, file_id: FileId) -> Result<EditOutcome, CoreError> {
        let old_nodes = self.semantics[&file_id].nodes.clone();
        let mut candidate = self.project.buffer(file_id)?.clone();
        let old_source = candidate.text();
        let applied = candidate.undo()?;
        self.commit_buffer_transition(file_id, old_nodes, old_source, candidate, applied)
    }

    pub fn redo(&mut self, file_id: FileId) -> Result<EditOutcome, CoreError> {
        let old_nodes = self.semantics[&file_id].nodes.clone();
        let mut candidate = self.project.buffer(file_id)?.clone();
        let old_source = candidate.text();
        let applied = candidate.redo()?;
        self.commit_buffer_transition(file_id, old_nodes, old_source, candidate, applied)
    }

    pub fn save(&mut self, file_id: FileId) -> Result<(), CoreError> {
        self.project.save_file(file_id)?;
        self.rewrite_recovery_snapshots()?;
        Ok(())
    }

    pub fn save_all(&mut self) -> Result<(), CoreError> {
        self.project.save_all()?;
        self.rewrite_recovery_snapshots()?;
        Ok(())
    }

    pub fn confirm_external_save(&mut self, file_id: FileId) -> Result<(), CoreError> {
        self.project.confirm_external_save(file_id)?;
        self.rewrite_recovery_snapshots()?;
        Ok(())
    }

    pub fn prepare_compile(&mut self) -> Result<CompileRequest, CoreError> {
        self.save_all()?;
        let root_id = self.project.root_file_id()?;
        let revision = self.project.buffer(root_id)?.revision;
        Ok(CompileRequest {
            project_root: self.project.root.clone(),
            config: self.project.config.clone(),
            source_revision: revision,
            timeout: Duration::from_secs(120),
        })
    }

    pub async fn compile(&mut self) -> Result<CompileArtifact, CoreError> {
        Ok(vt_compiler::compile(self.prepare_compile()?).await?)
    }

    pub async fn detect_toolchain(&self) -> Vec<vt_protocol::ToolInfo> {
        vt_compiler::detect_toolchain().await
    }

    pub async fn forward_search(
        &self,
        source_file: &Path,
        line: u32,
        column: u32,
        pdf_path: &Path,
    ) -> Result<ForwardSearchResult, CoreError> {
        Ok(
            vt_synctex::forward_search(&self.project.root, source_file, line, column, pdf_path)
                .await?,
        )
    }

    pub async fn inverse_search(
        &self,
        pdf_path: &Path,
        page: u32,
        x: f32,
        y: f32,
    ) -> Result<InverseSearchResult, CoreError> {
        Ok(vt_synctex::inverse_search(&self.project.root, pdf_path, page, x, y).await?)
    }

    pub fn pdf_document_info(&self, pdf_path: &Path) -> Result<PdfDocumentInfo, CoreError> {
        let pdf_path = self.validated_pdf_path(pdf_path)?;
        let service = self.pdf_service();
        let info = service.document_info(pdf_path)?;
        service.purge_except(&info.fingerprint)?;
        Ok(info)
    }

    pub fn render_pdf(&self, mut request: PdfRenderRequest) -> Result<PdfRenderedImage, CoreError> {
        request.pdf_path = self.validated_pdf_path(&request.pdf_path)?;
        Ok(self.pdf_service().render(&request)?)
    }

    pub fn pdf_text_hit(
        &self,
        pdf_path: &Path,
        page_index: u32,
        x: f32,
        y: f32,
    ) -> Result<Option<PdfTextHit>, CoreError> {
        let pdf_path = self.validated_pdf_path(pdf_path)?;
        Ok(self.pdf_service().text_hit(pdf_path, page_index, x, y)?)
    }

    pub fn pdf_text_lines(
        &self,
        pdf_path: &Path,
        page_index: u32,
        regions: &[PdfRect],
    ) -> Result<Vec<PdfTextLine>, CoreError> {
        let pdf_path = self.validated_pdf_path(pdf_path)?;
        Ok(self
            .pdf_service()
            .text_lines(pdf_path, page_index, regions)?)
    }

    pub fn project_index(&self) -> Result<ProjectIndex, CoreError> {
        let mut symbols = Vec::new();
        for file in self.project.list_source_files() {
            let source = self.project.source_text(&file)?;
            symbols.extend(vt_index::index_file(&file, &source)?.symbols);
        }
        symbols.sort_by(|left, right| {
            left.file
                .cmp(&right.file)
                .then_with(|| left.start_byte.cmp(&right.start_byte))
        });
        Ok(ProjectIndex { symbols })
    }

    pub fn search_project(
        &self,
        request: &ProjectSearchRequest,
    ) -> Result<Vec<ProjectSearchMatch>, CoreError> {
        let global_limit = request.max_results.max(1);
        let mut matches = Vec::new();
        for file in self.project.list_source_files() {
            if matches.len() >= global_limit {
                break;
            }
            let source = self.project.source_text(&file)?;
            let mut local_request = request.clone();
            local_request.max_results = global_limit - matches.len();
            matches.extend(vt_index::search_file(&file, &source, &local_request)?);
        }
        matches.sort_by(|left, right| {
            left.file
                .cmp(&right.file)
                .then_with(|| left.start_byte.cmp(&right.start_byte))
        });
        Ok(matches)
    }

    pub fn preview_project_replace(
        &self,
        request: &ProjectReplaceRequest,
    ) -> Result<ProjectReplacePlan, CoreError> {
        if request.query.is_empty() {
            return Err(IndexError::EmptyQuery.into());
        }
        let limit = request.max_replacements.max(1);
        let mut remaining = limit.saturating_add(1);
        let mut collected = Vec::<(PathBuf, String, Vec<ProjectSearchMatch>)>::new();
        for file in self.project.list_source_files() {
            if remaining == 0 {
                break;
            }
            let source = self.project.source_text(&file)?;
            let matches = vt_index::search_file(
                &file,
                &source,
                &ProjectSearchRequest {
                    query: request.query.clone(),
                    case_sensitive: request.case_sensitive,
                    whole_word: request.whole_word,
                    max_results: remaining,
                },
            )?;
            remaining = remaining.saturating_sub(matches.len());
            if !matches.is_empty() {
                collected.push((file, source, matches));
            }
        }

        let discovered = collected
            .iter()
            .map(|(_, _, matches)| matches.len())
            .sum::<usize>();
        let truncated = discovered > limit || remaining == 0;
        let mut accepted = 0usize;
        let mut files = Vec::new();
        for (file, source, matches) in collected {
            if accepted >= limit {
                break;
            }
            let take = (limit - accepted).min(matches.len());
            let replacements = matches
                .into_iter()
                .take(take)
                .map(|matched| ProjectTextReplacement {
                    start_byte: matched.start_byte,
                    end_byte: matched.end_byte,
                    expected: source[matched.start_byte..matched.end_byte].to_owned(),
                    replacement: request.replacement.clone(),
                    line: matched.line,
                    column: matched.column,
                    preview: matched.preview,
                })
                .collect::<Vec<_>>();
            accepted += replacements.len();
            files.push(ProjectReplaceFilePlan {
                file,
                expected_sha256: sha256_text(&source),
                replacements,
            });
        }

        Ok(ProjectReplacePlan {
            plan_id: OperationId::new(),
            description: format!(
                "Replace {:?} with {:?} in project",
                request.query, request.replacement
            ),
            files,
            total_replacements: accepted,
            truncated,
        })
    }

    pub fn preview_symbol_rename(
        &self,
        request: &SymbolRenameRequest,
    ) -> Result<ProjectReplacePlan, CoreError> {
        validate_symbol_key(&request.old_key)?;
        validate_symbol_key(&request.new_key)?;
        if request.old_key == request.new_key {
            return Err(CoreError::InvalidSymbolRename(
                "old and new keys are identical".to_owned(),
            ));
        }
        let index = self.project_index()?;
        let (definition_kind, usage_kind) = match request.kind {
            SymbolRenameKind::Label => (SymbolKind::LabelDefinition, SymbolKind::Reference),
            SymbolRenameKind::Citation => (SymbolKind::BibliographyEntry, SymbolKind::Citation),
        };
        if index
            .symbols
            .iter()
            .any(|symbol| symbol.kind == definition_kind && symbol.key == request.new_key)
        {
            return Err(CoreError::SymbolRenameConflict(request.new_key.clone()));
        }
        if !index
            .symbols
            .iter()
            .any(|symbol| symbol.kind == definition_kind && symbol.key == request.old_key)
        {
            return Err(CoreError::InvalidSymbolRename(format!(
                "definition not found for {}",
                request.old_key
            )));
        }

        let mut grouped = HashMap::<PathBuf, Vec<_>>::new();
        for symbol in index.symbols.into_iter().filter(|symbol| {
            symbol.key == request.old_key
                && matches!(symbol.kind, kind if kind == definition_kind || kind == usage_kind)
        }) {
            grouped.entry(symbol.file.clone()).or_default().push(symbol);
        }

        let mut files = Vec::new();
        let mut total_replacements = 0usize;
        for (file, mut symbols) in grouped {
            let source = self.project.source_text(&file)?;
            symbols.sort_by_key(|symbol| symbol.start_byte);
            let replacements = symbols
                .into_iter()
                .map(|symbol| ProjectTextReplacement {
                    start_byte: symbol.start_byte,
                    end_byte: symbol.end_byte,
                    expected: request.old_key.clone(),
                    replacement: request.new_key.clone(),
                    line: symbol.line,
                    column: symbol.column,
                    preview: source_line_preview(&source, symbol.start_byte, symbol.end_byte, 180),
                })
                .collect::<Vec<_>>();
            total_replacements += replacements.len();
            files.push(ProjectReplaceFilePlan {
                file,
                expected_sha256: sha256_text(&source),
                replacements,
            });
        }
        files.sort_by(|left, right| left.file.cmp(&right.file));
        Ok(ProjectReplacePlan {
            plan_id: OperationId::new(),
            description: format!(
                "Rename {:?} symbol {:?} to {:?}",
                request.kind, request.old_key, request.new_key
            ),
            files,
            total_replacements,
            truncated: false,
        })
    }

    pub fn apply_project_replace(
        &mut self,
        plan: ProjectReplacePlan,
    ) -> Result<ProjectReplaceOutcome, CoreError> {
        if plan.truncated {
            return Err(CoreError::ReplacePlanTruncated);
        }

        struct PreparedFile {
            file: PathBuf,
            file_id: FileId,
            base_revision: Revision,
            old_len: usize,
            next_text: String,
        }

        let mut prepared = Vec::new();
        let mut backups = Vec::new();
        for file_plan in &plan.files {
            let source = self.project.source_text(&file_plan.file)?;
            if sha256_text(&source) != file_plan.expected_sha256 {
                return Err(CoreError::ReplacePlanStale(file_plan.file.clone()));
            }
            let next_text =
                apply_planned_replacements(&source, &file_plan.file, &file_plan.replacements)?;
            let file_id = self.project.open_file(&file_plan.file)?;
            let buffer = self.project.buffer(file_id)?;
            backups.push(buffer.clone());
            prepared.push(PreparedFile {
                file: file_plan.file.clone(),
                file_id,
                base_revision: buffer.revision,
                old_len: buffer.len_bytes(),
                next_text,
            });
        }

        let mut changed_files = Vec::new();
        for prepared_file in &prepared {
            let result = self.apply_text_edit(TextEdit {
                operation_id: OperationId::new(),
                origin: EditOrigin::Refactor,
                file_id: prepared_file.file_id,
                base_revision: prepared_file.base_revision,
                start_byte: 0,
                end_byte: prepared_file.old_len,
                replacement: prepared_file.next_text.clone(),
            });
            if let Err(error) = result {
                for backup in backups {
                    let file_id = backup.file_id;
                    self.project.restore_buffer(backup);
                    self.refresh_document(file_id)?;
                }
                self.rewrite_recovery_snapshots()?;
                return Err(error);
            }
            changed_files.push(prepared_file.file.clone());
        }

        Ok(ProjectReplaceOutcome {
            plan_id: plan.plan_id,
            changed_files,
            total_replacements: plan.total_replacements,
        })
    }

    pub fn prepare_layout_map(
        &mut self,
        authoritative_pdf: &Path,
    ) -> Result<LayoutMapRequest, CoreError> {
        self.save_all()?;
        let authoritative_pdf = self.validated_pdf_path(authoritative_pdf)?;
        let root_id = self.project.root_file_id()?;
        let buffer = self.project.buffer(root_id)?;
        let nodes = self
            .semantics
            .get(&root_id)
            .map(|document| document.nodes.clone())
            .unwrap_or_default();
        Ok(LayoutMapRequest {
            project_root: self.project.root.clone(),
            config: self.project.config.clone(),
            source_revision: buffer.revision,
            source_file: buffer.path.clone(),
            source_text: buffer.text(),
            nodes,
            authoritative_pdf,
        })
    }

    pub async fn build_layout_map(
        &mut self,
        authoritative_pdf: &Path,
    ) -> Result<LayoutMapArtifact, CoreError> {
        Ok(vt_layout_map::build_layout_map(self.prepare_layout_map(authoritative_pdf)?).await?)
    }

    fn pdf_service(&self) -> PdfService {
        PdfService::new(self.project.root.join(".visualtex/cache/pdf"))
    }

    fn validated_pdf_path(&self, pdf_path: &Path) -> Result<PathBuf, CoreError> {
        let candidate = if pdf_path.is_absolute() {
            pdf_path.to_path_buf()
        } else {
            self.project.root.join(pdf_path)
        };
        let canonical = candidate.canonicalize()?;
        if !canonical.starts_with(&self.project.root) {
            return Err(ProjectError::PathEscape(pdf_path.to_path_buf()).into());
        }
        let build_root = self
            .project
            .resolve(&self.project.config.output_directory)?;
        let canonical_build_root = build_root.canonicalize()?;
        if !canonical.starts_with(&canonical_build_root) {
            return Err(CoreError::Recovery(format!(
                "PDF is outside the configured build directory: {}",
                canonical.display()
            )));
        }
        Ok(canonical)
    }

    fn open_and_parse_project_sources(&mut self) -> Result<(), CoreError> {
        self.discover_project_sources_without_graph()?;
        for file_id in self.project.open_file_ids() {
            if !self.syntax.contains_key(&file_id) {
                self.refresh_document_without_graph(file_id)?;
            }
        }
        self.rebuild_dependency_graph();
        Ok(())
    }

    fn discover_project_sources_without_graph(&mut self) -> Result<bool, CoreError> {
        let mut discovered = false;
        let source_paths = self.project.list_source_files();
        for path in source_paths
            .into_iter()
            .filter(|path| is_latex_source_path(path))
        {
            let file_id = self.project.open_file(path)?;
            if !self.syntax.contains_key(&file_id) {
                self.refresh_document_without_graph(file_id)?;
                discovered = true;
            }
        }
        Ok(discovered)
    }

    fn refresh_document(&mut self, file_id: FileId) -> Result<(), CoreError> {
        self.refresh_document_without_graph(file_id)?;
        self.rebuild_dependency_graph();
        Ok(())
    }

    fn refresh_document_without_graph(&mut self, file_id: FileId) -> Result<(), CoreError> {
        let source = self.project.buffer(file_id)?.text();
        let syntax = SyntaxDocument::parse(&source)?;
        let semantics = SemanticDocument::parse(file_id, &source);
        self.syntax.insert(file_id, syntax);
        self.semantics.insert(file_id, semantics);
        Ok(())
    }

    fn commit_buffer_transition(
        &mut self,
        file_id: FileId,
        old_nodes: Vec<VisualNode>,
        old_source: String,
        candidate: DocumentBuffer,
        applied: AppliedEdit,
    ) -> Result<EditOutcome, CoreError> {
        let new_source = candidate.text();
        if let Some(document) = self.syntax.get_mut(&file_id) {
            document.apply_edit(
                &old_source,
                &new_source,
                applied.start_byte,
                applied.old_end_byte,
                applied.new_end_byte,
            )?;
        } else {
            self.syntax
                .insert(file_id, SyntaxDocument::parse(&new_source)?);
        }
        let semantics = SemanticDocument::parse(file_id, &new_source);
        self.project.restore_buffer(candidate);
        self.semantics.insert(file_id, semantics);
        self.rebuild_dependency_graph();
        self.rewrite_recovery_snapshots()?;
        let new_nodes = self.semantics[&file_id].nodes.clone();
        Ok(EditOutcome {
            revision: applied.new_revision,
            patch: diff_nodes(applied.new_revision, &old_nodes, &new_nodes),
        })
    }

    fn rebuild_dependency_graph(&mut self) {
        let known_files = self
            .syntax
            .keys()
            .filter_map(|file_id| {
                self.project
                    .buffer(*file_id)
                    .ok()
                    .map(|buffer| buffer.path.clone())
            })
            .collect::<BTreeSet<_>>();
        let documents = self.syntax.iter().filter_map(|(file_id, document)| {
            self.project
                .buffer(*file_id)
                .ok()
                .map(|buffer| (buffer.path.clone(), document))
        });
        self.dependency_graph = build_dependency_graph(documents, &known_files);
    }

    fn recovery_path(&self) -> PathBuf {
        self.project
            .root
            .join(".visualtex/recovery/operations.jsonl")
    }

    fn recover_pending(&mut self) -> Result<(), CoreError> {
        let path = self.recovery_path();
        if !path.is_file() {
            return Ok(());
        }
        let contents = fs::read_to_string(&path)?;
        if contents.trim().is_empty() {
            return Ok(());
        }

        let mut latest = HashMap::<PathBuf, RecoveryRecord>::new();
        for line in contents.lines().filter(|line| !line.trim().is_empty()) {
            match serde_json::from_str::<RecoveryRecord>(line) {
                Ok(record) => {
                    latest.insert(record.path.clone(), record);
                }
                Err(error) => {
                    let corrupt = self
                        .project
                        .root
                        .join(".visualtex/recovery/conflicts/corrupt-recovery.jsonl");
                    atomic_write(&corrupt, contents.as_bytes())?;
                    return Err(CoreError::Recovery(format!(
                        "invalid recovery record: {error}; preserved at {}",
                        corrupt.display()
                    )));
                }
            }
        }

        for record in latest.into_values() {
            let file_id = self.project.open_file(&record.path)?;
            let full_path = self.project.resolve(&record.path)?;
            let disk_hash = sha256_file_optional(&full_path)?;
            if disk_hash != record.base_disk_sha256 {
                self.write_recovery_conflict(&record)?;
                continue;
            }
            if self.project.buffer(file_id)?.text() != record.resulting_text {
                self.project
                    .buffer_mut(file_id)?
                    .apply_external_text(record.resulting_text)?;
            }
        }
        self.rewrite_recovery_snapshots()
    }

    fn rewrite_recovery_snapshots(&self) -> Result<(), CoreError> {
        let mut bytes = Vec::new();
        for (_file_id, path, revision, text) in self.project.dirty_documents() {
            let full_path = self.project.resolve(&path)?;
            let record = RecoveryRecord {
                path,
                base_disk_sha256: sha256_file_optional(&full_path)?,
                resulting_text: text,
                resulting_revision: revision,
            };
            serde_json::to_writer(&mut bytes, &record)?;
            bytes.push(b'\n');
        }
        atomic_write(&self.recovery_path(), &bytes)?;
        Ok(())
    }

    fn write_recovery_conflict(&self, record: &RecoveryRecord) -> Result<(), CoreError> {
        let safe_name = record
            .path
            .to_string_lossy()
            .chars()
            .map(|character| {
                if character.is_ascii_alphanumeric() || matches!(character, '.' | '-' | '_') {
                    character
                } else {
                    '_'
                }
            })
            .collect::<String>();
        let conflict = self
            .project
            .root
            .join(".visualtex/recovery/conflicts")
            .join(format!("{safe_name}.recovered.tex"));
        atomic_write(&conflict, record.resulting_text.as_bytes())?;
        Ok(())
    }
}

fn sha256_file(path: &Path) -> Result<String, std::io::Error> {
    let bytes = fs::read(path)?;
    Ok(format!("{:x}", Sha256::digest(bytes)))
}

fn sha256_file_optional(path: &Path) -> Result<Option<String>, std::io::Error> {
    if !path.exists() {
        return Ok(None);
    }
    sha256_file(path).map(Some)
}

fn sha256_text(text: &str) -> String {
    format!("{:x}", Sha256::digest(text.as_bytes()))
}

fn validate_symbol_key(key: &str) -> Result<(), CoreError> {
    if key.is_empty() || key.trim() != key {
        return Err(CoreError::InvalidSymbolRename(
            "symbol keys must be non-empty and have no surrounding whitespace".to_owned(),
        ));
    }
    if key.chars().any(|character| {
        character.is_whitespace() || matches!(character, '{' | '}' | ',' | '\\' | '%')
    }) {
        return Err(CoreError::InvalidSymbolRename(format!(
            "symbol key contains an unsafe character: {key:?}"
        )));
    }
    Ok(())
}

fn apply_planned_replacements(
    source: &str,
    file: &Path,
    replacements: &[ProjectTextReplacement],
) -> Result<String, CoreError> {
    let mut ordered = replacements.to_vec();
    ordered.sort_by_key(|replacement| (replacement.start_byte, replacement.end_byte));
    let mut previous_end = 0usize;
    for replacement in &ordered {
        if replacement.start_byte < previous_end
            || replacement.start_byte > replacement.end_byte
            || replacement.end_byte > source.len()
            || !source.is_char_boundary(replacement.start_byte)
            || !source.is_char_boundary(replacement.end_byte)
            || source[replacement.start_byte..replacement.end_byte] != replacement.expected
        {
            return Err(CoreError::ReplacePlanStale(file.to_path_buf()));
        }
        previous_end = replacement.end_byte;
    }

    let mut result = source.to_owned();
    for replacement in ordered.into_iter().rev() {
        result.replace_range(
            replacement.start_byte..replacement.end_byte,
            &replacement.replacement,
        );
    }
    Ok(result)
}

fn source_line_preview(source: &str, start: usize, end: usize, max_chars: usize) -> String {
    let line_start = source[..start].rfind('\n').map_or(0, |index| index + 1);
    let line_end = source[end..]
        .find('\n')
        .map_or(source.len(), |relative| end + relative);
    let line = source[line_start..line_end].trim();
    if line.chars().count() <= max_chars {
        return line.to_owned();
    }
    let match_offset = source[line_start..start].chars().count();
    let half = max_chars / 2;
    let skip = match_offset.saturating_sub(half);
    let preview = line.chars().skip(skip).take(max_chars).collect::<String>();
    if skip > 0 {
        format!("…{preview}")
    } else {
        format!("{preview}…")
    }
}

fn is_latex_source_path(path: &Path) -> bool {
    matches!(
        path.extension().and_then(|extension| extension.to_str()),
        Some("tex" | "sty" | "cls" | "ltx")
    )
}

fn diff_nodes(revision: Revision, old: &[VisualNode], new: &[VisualNode]) -> VisualPatch {
    let old_by_id = old
        .iter()
        .map(|node| (node.id, node))
        .collect::<HashMap<_, _>>();
    let new_by_id = new
        .iter()
        .map(|node| (node.id, node))
        .collect::<HashMap<_, _>>();
    let removed = old_by_id
        .keys()
        .filter(|id| !new_by_id.contains_key(id))
        .copied()
        .collect::<Vec<_>>();
    let upserted = new
        .iter()
        .filter(|node| old_by_id.get(&node.id).is_none_or(|old| *old != *node))
        .cloned()
        .collect::<Vec<_>>();

    if removed.len() + upserted.len() > new.len().saturating_mul(3) / 4 {
        VisualPatch::Reset {
            revision,
            nodes: new.to_vec(),
        }
    } else {
        VisualPatch::Replace {
            revision,
            removed,
            upserted,
        }
    }
}

fn serialize_visual_change(
    source: &str,
    node: &VisualNode,
    content: String,
) -> Option<(usize, usize, String)> {
    use vt_protocol::NodeKind;
    let start = node.source.start_byte;
    let end = node.source.end_byte;
    if start > end || end > source.len() {
        return None;
    }
    match node.kind {
        NodeKind::Paragraph => Some((start, end, content)),
        NodeKind::Title
        | NodeKind::Author
        | NodeKind::Section
        | NodeKind::Subsection
        | NodeKind::Citation
        | NodeKind::Reference
        | NodeKind::Footnote
        | NodeKind::Bibliography => {
            let fragment = &source[start..end];
            let open = fragment.find('{')?;
            let close = fragment.rfind('}')?;
            (open < close).then(|| (start + open + 1, start + close, content))
        }
        NodeKind::InlineMath => {
            let fragment = &source[start..end];
            if fragment.starts_with('$') && fragment.ends_with('$') {
                let delimiter = if fragment.starts_with("$$") { 2 } else { 1 };
                Some((start + delimiter, end - delimiter, content))
            } else {
                None
            }
        }
        NodeKind::DisplayMath => {
            let fragment = &source[start..end];
            if (fragment.starts_with("\\[") && fragment.ends_with("\\]"))
                || (fragment.starts_with("$$") && fragment.ends_with("$$"))
            {
                Some((start + 2, end - 2, content))
            } else if fragment.starts_with("\\begin{") {
                let begin_end = fragment.find('}')? + 1;
                let close_start = fragment.rfind("\\end{")?;
                Some((start + begin_end, start + close_start, content))
            } else {
                None
            }
        }
        NodeKind::Abstract
        | NodeKind::Figure
        | NodeKind::Table
        | NodeKind::List
        | NodeKind::Theorem => {
            let fragment = &source[start..end];
            let begin_end = fragment.find('}')? + 1;
            let close_start = fragment.rfind("\\end{")?;
            Some((start + begin_end, start + close_start, content))
        }
        _ => None,
    }
}

fn apply_attribute_patch(
    fragment: &str,
    node: &VisualNode,
    patch: &NodeAttributesPatch,
) -> Option<String> {
    use vt_protocol::NodeKind;

    let environment = node.command.as_deref()?;
    let mut result = fragment.to_owned();
    if let Some(placement) = &patch.placement {
        result = replace_environment_option(&result, environment, placement)?;
    }
    match node.kind {
        NodeKind::Figure => {
            if patch.image_path.is_some() || patch.image_width.is_some() {
                result = replace_includegraphics(
                    &result,
                    patch.image_path.as_deref(),
                    patch.image_width.as_deref(),
                )?;
            }
        }
        NodeKind::Table => {
            if patch.column_spec.is_some() || patch.table_rows.is_some() {
                result = replace_tabular(
                    &result,
                    patch.column_spec.as_deref(),
                    patch.table_rows.as_deref(),
                )?;
            }
        }
        _ => return None,
    }
    if let Some(caption) = &patch.caption {
        result = replace_or_insert_command(&result, environment, "caption", caption)?;
    }
    if let Some(label) = &patch.label {
        result = replace_or_insert_command(&result, environment, "label", label)?;
    }
    Some(result)
}

fn replace_environment_option(fragment: &str, environment: &str, value: &str) -> Option<String> {
    let marker = format!("\\begin{{{environment}}}");
    let start = fragment.find(&marker)?;
    let cursor = skip_ascii_whitespace(fragment, start + marker.len());
    let mut result = fragment.to_owned();
    if let Some(end) = balanced_delimiter_end(fragment, cursor, b'[', b']') {
        if value.trim().is_empty() {
            result.replace_range(cursor..end, "");
        } else {
            result.replace_range(cursor + 1..end - 1, value.trim());
        }
    } else if !value.trim().is_empty() {
        result.insert_str(cursor, &format!("[{}]", value.trim()));
    }
    Some(result)
}

fn replace_or_insert_command(
    fragment: &str,
    environment: &str,
    command: &str,
    value: &str,
) -> Option<String> {
    if let Some((command_start, content_start, content_end, command_end)) =
        command_group_range(fragment, command)
    {
        let mut result = fragment.to_owned();
        if value.trim().is_empty() {
            result.replace_range(command_start..command_end, "");
        } else {
            result.replace_range(content_start..content_end, value);
        }
        return Some(result);
    }
    if value.trim().is_empty() {
        return Some(fragment.to_owned());
    }
    let close = format!("\\end{{{environment}}}");
    let insertion = fragment.rfind(&close)?;
    let mut result = fragment.to_owned();
    result.insert_str(insertion, &format!("\n\\{command}{{{value}}}\n"));
    Some(result)
}

fn replace_includegraphics(
    fragment: &str,
    image_path: Option<&str>,
    image_width: Option<&str>,
) -> Option<String> {
    let marker = "\\includegraphics";
    let start = fragment.find(marker)?;
    let cursor = skip_ascii_whitespace(fragment, start + marker.len());
    let option_end = balanced_delimiter_end(fragment, cursor, b'[', b']');
    let path_open = skip_ascii_whitespace(fragment, option_end.unwrap_or(cursor));
    let path_end = balanced_delimiter_end(fragment, path_open, b'{', b'}')?;
    let mut result = fragment.to_owned();

    if let Some(path) = image_path {
        result.replace_range(path_open + 1..path_end - 1, path.trim());
    }
    if let Some(width) = image_width {
        result = replace_includegraphics_width(&result, width)?;
    }
    Some(result)
}

fn replace_includegraphics_width(fragment: &str, width: &str) -> Option<String> {
    let marker = "\\includegraphics";
    let start = fragment.find(marker)?;
    let cursor = skip_ascii_whitespace(fragment, start + marker.len());
    let option_end = balanced_delimiter_end(fragment, cursor, b'[', b']');
    let mut options = option_end
        .map(|end| {
            fragment[cursor + 1..end - 1]
                .split(',')
                .map(str::trim)
                .filter(|value| !value.is_empty())
                .map(str::to_owned)
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    options.retain(|option| {
        option
            .split_once('=')
            .is_none_or(|(key, _)| key.trim() != "width")
    });
    if !width.trim().is_empty() {
        options.push(format!("width={}", width.trim()));
    }
    let mut result = fragment.to_owned();
    match (option_end, options.is_empty()) {
        (Some(end), true) => result.replace_range(cursor..end, ""),
        (Some(end), false) => result.replace_range(cursor + 1..end - 1, &options.join(",")),
        (None, false) => result.insert_str(cursor, &format!("[{}]", options.join(","))),
        (None, true) => {}
    }
    Some(result)
}

fn replace_tabular(
    fragment: &str,
    column_spec: Option<&str>,
    rows: Option<&[Vec<String>]>,
) -> Option<String> {
    let marker = "\\begin{tabular}";
    let start = fragment.find(marker)?;
    let spec_open = skip_ascii_whitespace(fragment, start + marker.len());
    let spec_end = balanced_delimiter_end(fragment, spec_open, b'{', b'}')?;
    let mut result = fragment.to_owned();
    if let Some(spec) = column_spec {
        result.replace_range(spec_open + 1..spec_end - 1, spec.trim());
    }
    if let Some(rows) = rows {
        let start = result.find(marker)?;
        let spec_open = skip_ascii_whitespace(&result, start + marker.len());
        let spec_end = balanced_delimiter_end(&result, spec_open, b'{', b'}')?;
        let body_end = result[spec_end..].find("\\end{tabular}")? + spec_end;
        result.replace_range(spec_end..body_end, &serialize_table_rows(rows));
    }
    Some(result)
}

fn serialize_table_rows(rows: &[Vec<String>]) -> String {
    if rows.is_empty() {
        return "\n".to_owned();
    }
    let body = rows
        .iter()
        .map(|row| format!("{} \\\\", row.join(" & ")))
        .collect::<Vec<_>>()
        .join("\n");
    format!("\n{body}\n")
}

fn command_group_range(fragment: &str, command: &str) -> Option<(usize, usize, usize, usize)> {
    let marker = format!("\\{command}");
    let command_start = fragment.find(&marker)?;
    let mut cursor = skip_ascii_whitespace(fragment, command_start + marker.len());
    if let Some(end) = balanced_delimiter_end(fragment, cursor, b'[', b']') {
        cursor = skip_ascii_whitespace(fragment, end);
    }
    let group_end = balanced_delimiter_end(fragment, cursor, b'{', b'}')?;
    Some((command_start, cursor + 1, group_end - 1, group_end))
}

fn skip_ascii_whitespace(source: &str, mut cursor: usize) -> usize {
    while source
        .as_bytes()
        .get(cursor)
        .is_some_and(u8::is_ascii_whitespace)
    {
        cursor += 1;
    }
    cursor
}

fn balanced_delimiter_end(source: &str, open: usize, opening: u8, closing: u8) -> Option<usize> {
    if source.as_bytes().get(open) != Some(&opening) {
        return None;
    }
    let bytes = source.as_bytes();
    let mut depth = 0usize;
    for index in open..bytes.len() {
        if index > 0 && bytes[index - 1] == b'\\' {
            continue;
        }
        if bytes[index] == opening {
            depth += 1;
        } else if bytes[index] == closing {
            depth = depth.checked_sub(1)?;
            if depth == 0 {
                return Some(index + 1);
            }
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;
    use vt_protocol::NodeKind;

    #[test]
    fn source_and_visual_edits_share_one_buffer() {
        let temp = tempdir().unwrap();
        let mut core = CoreService::init_project(temp.path()).unwrap();
        let snapshot = core.root_snapshot().unwrap();
        let section = snapshot
            .nodes
            .iter()
            .find(|node| node.kind == NodeKind::Section)
            .unwrap();
        core.apply_visual_edit(
            snapshot.file_id,
            snapshot.revision,
            section.id,
            "方法".into(),
        )
        .unwrap();
        let updated = core.root_snapshot().unwrap();
        assert!(updated.text.contains("\\section{方法}"));
    }

    #[test]
    fn compile_request_is_detached_while_core_remains_available() {
        let temp = tempdir().unwrap();
        let mut core = CoreService::init_project(temp.path()).unwrap();
        let expected_root = core.project_root().to_path_buf();
        let request = core.prepare_compile().unwrap();

        assert_eq!(request.project_root, expected_root);
        assert_eq!(
            request.source_revision,
            core.root_snapshot().unwrap().revision
        );
        assert!(!core.root_snapshot().unwrap().dirty);
    }

    #[test]
    fn incremental_syntax_and_dependency_graph_follow_edit_undo_and_redo() {
        let temp = tempdir().unwrap();
        fs::create_dir(temp.path().join("chapters")).unwrap();
        fs::write(
            temp.path().join("main.tex"),
            "\\documentclass{article}\n\\begin{document}\n\\input{chapters/intro}\n\\end{document}\n",
        )
        .unwrap();
        fs::write(
            temp.path().join("chapters/intro.tex"),
            "\\section{Intro}\n\\input{../main}\n",
        )
        .unwrap();

        let mut core = CoreService::open_project(temp.path()).unwrap();
        let graph = core.project_dependencies();
        assert_eq!(graph.edges.len(), 2);
        assert!(graph.edges.iter().all(|edge| edge.resolved));
        assert_eq!(graph.cycles.len(), 1);
        assert_eq!(core.syntax.len(), 2);

        let snapshot = core.root_snapshot().unwrap();
        let start = snapshot.text.find("chapters/intro").unwrap();
        core.apply_text_edit(TextEdit {
            operation_id: OperationId::new(),
            origin: EditOrigin::SourceEditor,
            file_id: snapshot.file_id,
            base_revision: snapshot.revision,
            start_byte: start,
            end_byte: start + "chapters/intro".len(),
            replacement: "missing".into(),
        })
        .unwrap();

        let updated = core.root_snapshot().unwrap();
        let syntax = &core.syntax[&updated.file_id];
        assert_eq!(syntax.source_len(), updated.text.len());
        assert_eq!(
            syntax.root_sexp(),
            SyntaxDocument::parse(&updated.text).unwrap().root_sexp()
        );
        let graph = core.project_dependencies();
        assert_eq!(graph.edges.len(), 2);
        assert_eq!(graph.edges.iter().filter(|edge| edge.resolved).count(), 1);
        assert!(graph.cycles.is_empty());

        core.undo(updated.file_id).unwrap();
        assert_eq!(core.project_dependencies().cycles.len(), 1);
        core.redo(updated.file_id).unwrap();
        assert!(core.project_dependencies().cycles.is_empty());

        fs::write(
            temp.path().join("missing.tex"),
            "Newly discovered source.\n",
        )
        .unwrap();
        core.check_external_changes().unwrap();
        let graph = core.project_dependencies();
        assert_eq!(graph.edges.iter().filter(|edge| edge.resolved).count(), 2);
        assert_eq!(core.syntax.len(), 3);
    }

    #[test]
    fn project_replace_preview_and_apply_are_transactional_and_unsaved() {
        let temp = tempdir().unwrap();
        fs::write(
            temp.path().join("main.tex"),
            "\\documentclass{article}\n\\begin{document}\nalpha 中文 alpha\n\\input{chapter}\n\\end{document}\n",
        )
        .unwrap();
        fs::write(temp.path().join("chapter.tex"), "alpha beta\n").unwrap();
        let original_main = fs::read_to_string(temp.path().join("main.tex")).unwrap();
        let mut core = CoreService::open_project(temp.path()).unwrap();
        let plan = core
            .preview_project_replace(&ProjectReplaceRequest {
                query: "alpha".into(),
                replacement: "omega".into(),
                case_sensitive: true,
                whole_word: true,
                max_replacements: 20,
            })
            .unwrap();
        assert_eq!(plan.total_replacements, 3);
        assert!(!plan.truncated);
        assert_eq!(plan.files.len(), 2);

        let outcome = core.apply_project_replace(plan).unwrap();
        assert_eq!(outcome.total_replacements, 3);
        assert_eq!(
            core.search_project(&ProjectSearchRequest {
                query: "omega".into(),
                case_sensitive: true,
                whole_word: true,
                max_results: 20,
            })
            .unwrap()
            .len(),
            3
        );
        assert!(
            core.search_project(&ProjectSearchRequest {
                query: "alpha".into(),
                case_sensitive: true,
                whole_word: true,
                max_results: 20,
            })
            .unwrap()
            .is_empty()
        );
        assert_eq!(
            fs::read_to_string(temp.path().join("main.tex")).unwrap(),
            original_main,
            "project replacements remain in authoritative dirty buffers until save"
        );
        core.save_all().unwrap();
        assert!(
            fs::read_to_string(temp.path().join("chapter.tex"))
                .unwrap()
                .contains("omega beta")
        );
    }

    #[test]
    fn stale_or_truncated_replace_plans_are_rejected() {
        let temp = tempdir().unwrap();
        fs::write(
            temp.path().join("main.tex"),
            "\\documentclass{article}\n\\begin{document}\nkey key key\n\\end{document}\n",
        )
        .unwrap();
        let mut core = CoreService::open_project(temp.path()).unwrap();
        let truncated = core
            .preview_project_replace(&ProjectReplaceRequest {
                query: "key".into(),
                replacement: "value".into(),
                case_sensitive: true,
                whole_word: true,
                max_replacements: 2,
            })
            .unwrap();
        assert!(truncated.truncated);
        assert!(matches!(
            core.apply_project_replace(truncated),
            Err(CoreError::ReplacePlanTruncated)
        ));

        let plan = core
            .preview_project_replace(&ProjectReplaceRequest {
                query: "key".into(),
                replacement: "value".into(),
                case_sensitive: true,
                whole_word: true,
                max_replacements: 10,
            })
            .unwrap();
        let snapshot = core.root_snapshot().unwrap();
        core.apply_text_edit(TextEdit {
            operation_id: OperationId::new(),
            origin: EditOrigin::SourceEditor,
            file_id: snapshot.file_id,
            base_revision: snapshot.revision,
            start_byte: snapshot.text.find("key").unwrap(),
            end_byte: snapshot.text.find("key").unwrap() + 3,
            replacement: "changed".into(),
        })
        .unwrap();
        assert!(matches!(
            core.apply_project_replace(plan),
            Err(CoreError::ReplacePlanStale(path)) if path == Path::new("main.tex")
        ));
    }

    #[test]
    fn symbol_rename_updates_only_typed_definitions_and_usages() {
        let temp = tempdir().unwrap();
        fs::write(
            temp.path().join("main.tex"),
            "\\documentclass{article}\n\\begin{document}\n\\label{sec:old}\nSee \\ref{sec:old}; literal sec:old. Cite \\cite{paper-old}.\n\\end{document}\n",
        )
        .unwrap();
        fs::write(
            temp.path().join("refs.bib"),
            "@article{paper-old, title={Old key}}\n",
        )
        .unwrap();
        let mut core = CoreService::open_project(temp.path()).unwrap();
        let label_plan = core
            .preview_symbol_rename(&SymbolRenameRequest {
                kind: SymbolRenameKind::Label,
                old_key: "sec:old".into(),
                new_key: "sec:new".into(),
            })
            .unwrap();
        assert_eq!(label_plan.total_replacements, 2);
        core.apply_project_replace(label_plan).unwrap();
        let main = core.project.source_text("main.tex").unwrap();
        assert!(main.contains("\\label{sec:new}"));
        assert!(main.contains("\\ref{sec:new}"));
        assert!(main.contains("literal sec:old"));

        let citation_plan = core
            .preview_symbol_rename(&SymbolRenameRequest {
                kind: SymbolRenameKind::Citation,
                old_key: "paper-old".into(),
                new_key: "paper-new".into(),
            })
            .unwrap();
        assert_eq!(citation_plan.total_replacements, 2);
        core.apply_project_replace(citation_plan).unwrap();
        assert!(
            core.project
                .source_text("main.tex")
                .unwrap()
                .contains("\\cite{paper-new}")
        );
        assert!(
            core.project
                .source_text("refs.bib")
                .unwrap()
                .contains("@article{paper-new,")
        );

        assert!(matches!(
            core.preview_symbol_rename(&SymbolRenameRequest {
                kind: SymbolRenameKind::Citation,
                old_key: "paper-new".into(),
                new_key: "bad key".into(),
            }),
            Err(CoreError::InvalidSymbolRename(_))
        ));
    }

    #[test]
    fn project_index_and_search_use_unsaved_buffers() {
        let temp = tempdir().unwrap();
        fs::write(
            temp.path().join("main.tex"),
            "\\documentclass{article}\n\\begin{document}\n\\label{sec:old}\nSee \\ref{sec:old}.\n\\end{document}\n",
        )
        .unwrap();
        fs::write(
            temp.path().join("refs.bib"),
            "@article{paper2026, title={Current paper}}\n",
        )
        .unwrap();
        let mut core = CoreService::open_project(temp.path()).unwrap();
        let snapshot = core.root_snapshot().unwrap();
        let start = snapshot.text.find("sec:old").unwrap();
        core.apply_text_edit(TextEdit {
            operation_id: OperationId::new(),
            origin: EditOrigin::SourceEditor,
            file_id: snapshot.file_id,
            base_revision: snapshot.revision,
            start_byte: start,
            end_byte: start + "sec:old".len(),
            replacement: "sec:unsaved".into(),
        })
        .unwrap();

        let index = core.project_index().unwrap();
        assert!(index.symbols.iter().any(|symbol| {
            symbol.kind == vt_protocol::SymbolKind::LabelDefinition && symbol.key == "sec:unsaved"
        }));
        assert!(!index.symbols.iter().any(|symbol| symbol.key == "sec:old"
            && symbol.kind == vt_protocol::SymbolKind::LabelDefinition));
        assert!(index.symbols.iter().any(|symbol| {
            symbol.kind == vt_protocol::SymbolKind::BibliographyEntry && symbol.key == "paper2026"
        }));

        let matches = core
            .search_project(&ProjectSearchRequest {
                query: "sec:unsaved".into(),
                case_sensitive: true,
                whole_word: false,
                max_results: 20,
            })
            .unwrap();
        assert_eq!(matches.len(), 1);
        assert_eq!(matches[0].file, PathBuf::from("main.tex"));
    }

    #[test]
    fn figure_and_table_attribute_edits_preserve_unknown_source() {
        let temp = tempdir().unwrap();
        fs::write(
            temp.path().join("main.tex"),
            r#"\documentclass{article}
\begin{document}
\begin{figure}[ht]
\custombefore{unchanged}
\includegraphics[keepaspectratio,width=0.5\linewidth]{old.png}
\caption[Short]{Old caption}
\label{fig:old}
\end{figure}
\begin{table}[t]
\caption{Old table}
\begin{tabular}{lc}
A & 1 \\
B & 2 \\
\end{tabular}
\end{table}
\end{document}
"#,
        )
        .unwrap();
        let mut core = CoreService::open_project(temp.path()).unwrap();
        let snapshot = core.root_snapshot().unwrap();
        let figure = snapshot
            .nodes
            .iter()
            .find(|node| node.kind == NodeKind::Figure)
            .unwrap();
        core.apply_node_attributes(
            snapshot.file_id,
            snapshot.revision,
            figure.id,
            NodeAttributesPatch {
                placement: Some("H".into()),
                caption: Some("Updated caption".into()),
                label: Some(String::new()),
                image_path: Some("figures/new image.pdf".into()),
                image_width: Some("0.8\\linewidth".into()),
                ..NodeAttributesPatch::default()
            },
        )
        .unwrap();
        let after_figure = core.root_snapshot().unwrap();
        assert!(after_figure.text.contains("\\custombefore{unchanged}"));
        assert!(after_figure.text.contains("\\begin{figure}[H]"));
        assert!(after_figure.text.contains(
            "\\includegraphics[keepaspectratio,width=0.8\\linewidth]{figures/new image.pdf}"
        ));
        assert!(
            after_figure
                .text
                .contains("\\caption[Short]{Updated caption}")
        );
        assert!(!after_figure.text.contains("\\label{fig:old}"));

        let table = after_figure
            .nodes
            .iter()
            .find(|node| node.kind == NodeKind::Table && node.command.as_deref() == Some("table"))
            .unwrap();
        core.apply_node_attributes(
            after_figure.file_id,
            after_figure.revision,
            table.id,
            NodeAttributesPatch {
                caption: Some("Updated table".into()),
                column_spec: Some("cc".into()),
                table_rows: Some(vec![
                    vec!["X".into(), "10".into()],
                    vec!["Y".into(), "20".into()],
                ]),
                ..NodeAttributesPatch::default()
            },
        )
        .unwrap();
        let updated = core.root_snapshot().unwrap();
        assert!(updated.text.contains("\\caption{Updated table}"));
        assert!(updated.text.contains("\\begin{tabular}{cc}"));
        assert!(updated.text.contains("X & 10 \\\\"));
        assert!(updated.text.contains("Y & 20 \\\\"));
    }

    #[test]
    fn display_math_overlay_rewrites_only_environment_body() {
        let temp = tempdir().unwrap();
        fs::write(
            temp.path().join("main.tex"),
            "\\documentclass{article}\n\\begin{document}\n\\begin{equation}\na+b=c\n\\end{equation}\n\\end{document}\n",
        )
        .unwrap();
        let mut core = CoreService::open_project(temp.path()).unwrap();
        let snapshot = core.root_snapshot().unwrap();
        let display_math = snapshot
            .nodes
            .iter()
            .find(|node| node.kind == NodeKind::DisplayMath)
            .unwrap();
        core.apply_visual_edit(
            snapshot.file_id,
            snapshot.revision,
            display_math.id,
            "E=mc^2".into(),
        )
        .unwrap();
        let updated = core.root_snapshot().unwrap();
        assert!(
            updated
                .text
                .contains("\\begin{equation}E=mc^2\\end{equation}")
        );
        assert_eq!(updated.text.matches("\\begin{equation}").count(), 1);
        assert_eq!(updated.text.matches("\\end{equation}").count(), 1);
    }

    #[test]
    fn clean_external_change_auto_reloads_and_refreshes_semantics() {
        let temp = tempdir().unwrap();
        let mut core = CoreService::init_project(temp.path()).unwrap();
        fs::write(
            temp.path().join("main.tex"),
            "\\documentclass{article}\n\\begin{document}\n\\section{External}\nDisk version\n\\end{document}\n",
        )
        .unwrap();
        let report = core.check_external_changes().unwrap();
        assert_eq!(report.reloaded.len(), 1);
        assert!(report.conflicts.is_empty());
        let snapshot = core.root_snapshot().unwrap();
        assert!(snapshot.text.contains("Disk version"));
        assert!(snapshot.nodes.iter().any(|node| {
            node.kind == NodeKind::Section && node.text.as_deref() == Some("External")
        }));
        assert!(!snapshot.dirty);
    }

    #[test]
    fn dirty_external_change_requires_resolution_and_can_save_copy() {
        let temp = tempdir().unwrap();
        let mut core = CoreService::init_project(temp.path()).unwrap();
        let snapshot = core.root_snapshot().unwrap();
        core.apply_text_edit(TextEdit {
            operation_id: OperationId::new(),
            origin: EditOrigin::SourceEditor,
            file_id: snapshot.file_id,
            base_revision: snapshot.revision,
            start_byte: snapshot.text.len(),
            end_byte: snapshot.text.len(),
            replacement: "% LOCAL BUFFER\n".into(),
        })
        .unwrap();
        fs::write(
            temp.path().join("main.tex"),
            "\\documentclass{article}\n\\begin{document}\nDISK VERSION\n\\end{document}\n",
        )
        .unwrap();
        let report = core.check_external_changes().unwrap();
        assert!(report.reloaded.is_empty());
        assert_eq!(report.conflicts.len(), 1);
        assert!(report.conflicts[0].buffer_dirty);
        assert!(core.root_snapshot().unwrap().text.contains("LOCAL BUFFER"));

        let outcome = core
            .resolve_external_conflict(
                &report.conflicts[0],
                ExternalConflictResolution::SaveCopyAndReload,
            )
            .unwrap();
        assert!(outcome.snapshot.text.contains("DISK VERSION"));
        assert!(!outcome.snapshot.dirty);
        let copy = outcome.conflict_copy_path.unwrap();
        assert!(
            fs::read_to_string(temp.path().join(copy))
                .unwrap()
                .contains("LOCAL BUFFER")
        );
    }

    #[test]
    fn deleted_external_file_can_be_retained_and_recreated() {
        let temp = tempdir().unwrap();
        let mut core = CoreService::init_project(temp.path()).unwrap();
        fs::remove_file(temp.path().join("main.tex")).unwrap();
        let report = core.check_external_changes().unwrap();
        assert_eq!(report.conflicts.len(), 1);
        assert_eq!(report.conflicts[0].kind, ExternalChangeKind::Deleted);
        let outcome = core
            .resolve_external_conflict(&report.conflicts[0], ExternalConflictResolution::KeepBuffer)
            .unwrap();
        assert!(outcome.snapshot.dirty);
        core.save(outcome.snapshot.file_id).unwrap();
        assert!(temp.path().join("main.tex").is_file());
        assert!(!core.root_snapshot().unwrap().dirty);
    }

    #[test]
    fn restores_unsaved_text_after_reopen() {
        let temp = tempdir().unwrap();
        {
            let mut core = CoreService::init_project(temp.path()).unwrap();
            let snapshot = core.root_snapshot().unwrap();
            core.apply_text_edit(TextEdit {
                operation_id: OperationId::new(),
                origin: EditOrigin::SourceEditor,
                file_id: snapshot.file_id,
                base_revision: snapshot.revision,
                start_byte: 0,
                end_byte: 0,
                replacement: "% unsaved recovery marker\n".into(),
            })
            .unwrap();
        }

        let reopened = CoreService::open_project(temp.path()).unwrap();
        let snapshot = reopened.root_snapshot().unwrap();
        assert!(snapshot.text.starts_with("% unsaved recovery marker"));
        assert!(snapshot.dirty);
    }

    #[test]
    fn external_disk_change_creates_conflict_copy_instead_of_overwriting() {
        let temp = tempdir().unwrap();
        {
            let mut core = CoreService::init_project(temp.path()).unwrap();
            let snapshot = core.root_snapshot().unwrap();
            core.apply_text_edit(TextEdit {
                operation_id: OperationId::new(),
                origin: EditOrigin::SourceEditor,
                file_id: snapshot.file_id,
                base_revision: snapshot.revision,
                start_byte: 0,
                end_byte: 0,
                replacement: "% recovered version\n".into(),
            })
            .unwrap();
        }
        fs::write(
            temp.path().join("main.tex"),
            "\\documentclass{article}\n\\begin{document}\nexternal version\n\\end{document}\n",
        )
        .unwrap();

        let reopened = CoreService::open_project(temp.path()).unwrap();
        let snapshot = reopened.root_snapshot().unwrap();
        assert!(snapshot.text.contains("external version"));
        assert!(!snapshot.text.contains("recovered version"));
        let conflict = temp
            .path()
            .join(".visualtex/recovery/conflicts/main.tex.recovered.tex");
        assert!(
            fs::read_to_string(conflict)
                .unwrap()
                .contains("recovered version")
        );
    }

    #[test]
    fn diff_is_incremental_for_small_change() {
        let file_id = FileId::new();
        let first = SemanticDocument::parse(file_id, "\\section{A}");
        let second = SemanticDocument::parse(file_id, "\\section{B}");
        let patch = diff_nodes(Revision(1), &first.nodes, &second.nodes);
        assert!(matches!(
            patch,
            VisualPatch::Reset { .. } | VisualPatch::Replace { .. }
        ));
    }
}
