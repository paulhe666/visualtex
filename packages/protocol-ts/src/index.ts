import { convertFileSrc, invoke } from "@tauri-apps/api/core";

export type UUID = string;
export type Revision = number;

export type EditOrigin =
  | "sourceEditor"
  | "visualEditor"
  | "mathEditor"
  | "ocr"
  | "externalFileChange"
  | "formatter"
  | "refactor"
  | "undoRedo"
  | "plugin";

export type NodeKind =
  | "document"
  | "preamble"
  | "title"
  | "author"
  | "abstract"
  | "section"
  | "subsection"
  | "paragraph"
  | "text"
  | "inline_math"
  | "display_math"
  | "figure"
  | "table"
  | "list"
  | "theorem"
  | "citation"
  | "reference"
  | "footnote"
  | "bibliography"
  | "raw_latex";

export type SupportLevel = "native" | "partial" | "opaque" | "unstable";

export interface SourceSpan {
  fileId: UUID;
  startByte: number;
  endByte: number;
}

export interface NodeAttributes {
  placement: string | null;
  caption: string | null;
  label: string | null;
  imagePath: string | null;
  imageWidth: string | null;
  columnSpec: string | null;
  tableRows: string[][];
}

export interface NodeAttributesPatch {
  placement?: string;
  caption?: string;
  label?: string;
  imagePath?: string;
  imageWidth?: string;
  columnSpec?: string;
  tableRows?: string[][];
}

export interface VisualNode {
  id: UUID;
  kind: NodeKind;
  support: SupportLevel;
  source: SourceSpan;
  children: UUID[];
  text: string | null;
  command: string | null;
  attributes: NodeAttributes;
}

export type ExternalChangeKind = "modified" | "deleted";

export interface ExternalFileChange {
  fileId: UUID;
  path: string;
  kind: ExternalChangeKind;
  bufferDirty: boolean;
}

export interface ExternalChangeReport {
  reloaded: DocumentSnapshot[];
  conflicts: ExternalFileChange[];
}

export type ExternalConflictResolution =
  | "reload_disk"
  | "keep_buffer"
  | "save_copy_and_reload";

export interface DocumentSnapshot {
  fileId: UUID;
  path: string;
  revision: Revision;
  text: string;
  dirty: boolean;
  nodes: VisualNode[];
}

export interface ExternalConflictOutcome {
  snapshot: DocumentSnapshot;
  conflictCopyPath: string | null;
}

export interface TextEdit {
  operationId: UUID;
  origin: EditOrigin;
  fileId: UUID;
  baseRevision: Revision;
  startByte: number;
  endByte: number;
  replacement: string;
}

export type VisualPatch =
  | { reset: { revision: Revision; nodes: VisualNode[] } }
  | { replace: { revision: Revision; removed: UUID[]; upserted: VisualNode[] } };

export interface EditOutcome {
  revision: Revision;
  patch: VisualPatch;
}

export type DiagnosticSeverity = "error" | "warning" | "information" | "hint";

export interface Diagnostic {
  severity: DiagnosticSeverity;
  message: string;
  file: string | null;
  line: number | null;
  column: number | null;
  code: string | null;
}

export type SymbolKind =
  | "label_definition"
  | "reference"
  | "citation"
  | "bibliography_entry"
  | "macro_definition"
  | "package";

export interface ProjectSymbol {
  kind: SymbolKind;
  key: string;
  file: string;
  startByte: number;
  endByte: number;
  line: number;
  column: number;
  detail: string | null;
}

export interface ProjectIndex {
  symbols: ProjectSymbol[];
}

export type DependencyKind = "input" | "include" | "subfile" | "subfile_include";

export interface ProjectDependencyEdge {
  sourceFile: string;
  targetFile: string;
  rawPath: string;
  kind: DependencyKind;
  startByte: number;
  endByte: number;
  resolved: boolean;
}

export interface ProjectDependencyGraph {
  edges: ProjectDependencyEdge[];
  cycles: string[][];
}

export interface ProjectTemplateSummary {
  id: string;
  name: string;
  description: string;
  engine: string;
  rootFile: string;
}

export interface ProjectSearchRequest {
  query: string;
  caseSensitive: boolean;
  wholeWord: boolean;
  maxResults: number;
}

export interface ProjectSearchMatch {
  file: string;
  startByte: number;
  endByte: number;
  line: number;
  column: number;
  preview: string;
}

export interface ProjectReplaceRequest {
  query: string;
  replacement: string;
  caseSensitive: boolean;
  wholeWord: boolean;
  maxReplacements: number;
}

export type SymbolRenameKind = "label" | "citation";

export interface SymbolRenameRequest {
  kind: SymbolRenameKind;
  oldKey: string;
  newKey: string;
}

export interface ProjectTextReplacement {
  startByte: number;
  endByte: number;
  expected: string;
  replacement: string;
  line: number;
  column: number;
  preview: string;
}

export interface ProjectReplaceFilePlan {
  file: string;
  expectedSha256: string;
  replacements: ProjectTextReplacement[];
}

export interface ProjectReplacePlan {
  planId: UUID;
  description: string;
  files: ProjectReplaceFilePlan[];
  totalReplacements: number;
  truncated: boolean;
}

export interface ProjectReplaceOutcome {
  planId: UUID;
  changedFiles: string[];
  totalReplacements: number;
}

export type CompileStatus = "queued" | "running" | "succeeded" | "failed" | "cancelled" | "timedOut";

export interface CompileArtifact {
  buildId: UUID;
  sourceRevision: Revision;
  pdfPath: string | null;
  synctexPath: string | null;
  diagnostics: Diagnostic[];
  status: CompileStatus;
  startedAt: string;
  finishedAt: string | null;
  stdout: string;
  stderr: string;
}

export interface PdfPageInfo {
  index: number;
  widthPoints: number;
  heightPoints: number;
  rotationDegrees: number;
}

export interface PdfDocumentInfo {
  pdfPath: string;
  fingerprint: string;
  byteLen: number;
  pages: PdfPageInfo[];
}

export interface PdfPixelRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface PdfRenderRequest {
  pdfPath: string;
  pageIndex: number;
  targetWidthPixels: number;
  tile: PdfPixelRect | null;
  grayscale: boolean;
}

export interface PdfRenderedImage {
  pdfFingerprint: string;
  pageIndex: number;
  pageWidthPixels: number;
  pageHeightPixels: number;
  imageWidthPixels: number;
  imageHeightPixels: number;
  tile: PdfPixelRect | null;
  cachePath: string;
  cacheHit: boolean;
}

export interface PdfPixelDiffPage {
  pageIndex: number;
  widthPixels: number;
  heightPixels: number;
  changedPixels: number;
  totalPixels: number;
  changedRatio: number;
  maximumChannelDelta: number;
  meanAbsoluteChannelDelta: number;
}

export interface PdfPixelDiffReport {
  pageCountMatches: boolean;
  leftPageCount: number;
  rightPageCount: number;
  tolerance: number;
  pages: PdfPixelDiffPage[];
  maximumChangedRatio: number;
}

export interface PdfRect {
  page: number;
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface PdfTextGlyph {
  index: number;
  text: string;
  rect: PdfRect;
  fontName: string;
  fontSizePoints: number;
}

export interface PdfTextHit {
  pageIndex: number;
  glyphIndex: number;
  glyph: PdfTextGlyph;
  lineGlyphs: PdfTextGlyph[];
}

export type MappingConfidence = "exact" | "high" | "medium" | "low" | "unmapped";
export type MappingMethod =
  | "shadow_marker_and_sync_tex"
  | "shadow_marker"
  | "sync_tex"
  | "none";

export interface PdfPoint {
  page: number;
  x: number;
  y: number;
}

export interface LayoutBox {
  nodeId: UUID;
  source: SourceSpan;
  rects: PdfRect[];
  startMarker: PdfPoint | null;
  endMarker: PdfPoint | null;
  confidence: MappingConfidence;
  method: MappingMethod;
}

export interface LayoutMapArtifact {
  buildId: UUID;
  sourceRevision: Revision;
  shadowRoot: string;
  shadowPdfPath: string | null;
  compileStatus: CompileStatus;
  diagnostics: Diagnostic[];
  boxes: LayoutBox[];
  pixelDiff: PdfPixelDiffReport | null;
}

export interface ForwardSearchResult {
  pdfPath: string;
  boxes: PdfRect[];
}

export interface InverseSearchResult {
  sourcePath: string;
  line: number;
  column: number | null;
  offset: number | null;
}

export type ModelKind = "formula_ocr" | "layout_ocr" | "text_ocr" | "table_ocr";

export interface ModelFileDigest {
  path: string;
  sha256: string;
  byteLen: number;
}

export interface ModelPackageManifest {
  schemaVersion: number;
  id: string;
  kind: ModelKind;
  version: string;
  backend: string;
  entrypoint: string;
  files: ModelFileDigest[];
  metadata: Record<string, string>;
}

export interface InstalledModelPackage {
  manifest: ModelPackageManifest;
  installPath: string;
  installedSha256: string;
  totalBytes: number;
}

export interface ModelPackageInspection {
  manifest: ModelPackageManifest;
  sourcePath: string;
  computedFiles: ModelFileDigest[];
  packageSha256: string;
  totalBytes: number;
}

export interface ModelCatalog {
  modelsRoot: string;
  installed: InstalledModelPackage[];
  active: InstalledModelPackage[];
}

export interface OcrWorkerHealth {
  available: boolean;
  backend: string;
  modelVersion: string | null;
  detail: string | null;
}

export interface OcrCandidate {
  latex: string;
  confidence: number;
  backend: string;
}

export interface FormulaOcrResult {
  candidates: OcrCandidate[];
  modelVersion: string | null;
  warnings: string[];
}

export interface OcrRegion {
  kind: string;
  x: number;
  y: number;
  width: number;
  height: number;
  text: string | null;
  latex: string | null;
  confidence: number;
}

export interface DocumentOcrResult {
  imagePath: string | null;
  pageWidth: number;
  pageHeight: number;
  regions: OcrRegion[];
  readingOrder: number[];
  modelVersion: string | null;
  warnings: string[];
}

export interface ToolInfo {
  name: string;
  path: string | null;
  version: string | null;
  available: boolean;
}

export const createOperationId = (): UUID => crypto.randomUUID();

export const desktopApi = {
  drainDeepLinks(): Promise<string[]> {
    return invoke("drain_deep_links");
  },
  processDeepLink(value: string): Promise<void> {
    return invoke("process_deep_link", { value });
  },
  openProject(path: string): Promise<DocumentSnapshot> {
    return invoke("open_project", { path });
  },
  initProject(path: string): Promise<DocumentSnapshot> {
    return invoke("init_project", { path });
  },
  listProjectTemplates(): Promise<ProjectTemplateSummary[]> {
    return invoke("list_project_templates");
  },
  initProjectTemplate(path: string, templateId: string): Promise<DocumentSnapshot> {
    return invoke("init_project_template", { path, templateId });
  },
  inspectModelPackage(source: string): Promise<ModelPackageInspection> {
    return invoke("inspect_model_package", { source });
  },
  installModelPackage(source: string): Promise<InstalledModelPackage> {
    return invoke("install_model_package", { source });
  },
  listModelPackages(): Promise<ModelCatalog> {
    return invoke("list_model_packages");
  },
  activateModelPackage(
    kind: ModelKind,
    id: string,
    version: string,
  ): Promise<ModelCatalog> {
    return invoke("activate_model_package", { kind, id, version });
  },
  removeModelPackage(id: string, version: string): Promise<ModelCatalog> {
    return invoke("remove_model_package", { id, version });
  },
  ocrHealth(): Promise<OcrWorkerHealth> {
    return invoke("ocr_health");
  },
  recognizeFormulaImage(source: string): Promise<FormulaOcrResult> {
    return invoke("recognize_formula_image", { source });
  },
  recognizeDocumentImage(source: string): Promise<DocumentOcrResult> {
    return invoke("recognize_document_image", { source });
  },
  createOcrProject(
    target: string,
    sourceImage: string,
    latexBody: string,
    ocrDocument: DocumentOcrResult,
  ): Promise<DocumentSnapshot> {
    return invoke("create_ocr_project", { target, sourceImage, latexBody, ocrDocument });
  },
  rootSnapshot(): Promise<DocumentSnapshot> {
    return invoke("root_snapshot");
  },
  listFiles(): Promise<string[]> {
    return invoke("list_files");
  },
  openFile(path: string): Promise<DocumentSnapshot> {
    return invoke("open_file", { path });
  },
  applyTextEdit(edit: TextEdit): Promise<EditOutcome> {
    return invoke("apply_text_edit", { edit });
  },
  applyVisualEdit(
    fileId: UUID,
    baseRevision: Revision,
    nodeId: UUID,
    content: string,
  ): Promise<EditOutcome> {
    return invoke("apply_visual_edit", { fileId, baseRevision, nodeId, content });
  },
  applyNodeAttributes(
    fileId: UUID,
    baseRevision: Revision,
    nodeId: UUID,
    patch: NodeAttributesPatch,
  ): Promise<EditOutcome> {
    return invoke("apply_node_attributes", { fileId, baseRevision, nodeId, patch });
  },
  undo(fileId: UUID): Promise<EditOutcome> {
    return invoke("undo", { fileId });
  },
  redo(fileId: UUID): Promise<EditOutcome> {
    return invoke("redo", { fileId });
  },
  save(fileId: UUID): Promise<void> {
    return invoke("save", { fileId });
  },
  checkExternalChanges(): Promise<ExternalChangeReport> {
    return invoke("check_external_changes");
  },
  resolveExternalConflict(
    change: ExternalFileChange,
    resolution: ExternalConflictResolution,
  ): Promise<ExternalConflictOutcome> {
    return invoke("resolve_external_conflict", { change, resolution });
  },
  projectIndex(): Promise<ProjectIndex> {
    return invoke("project_index");
  },
  projectDependencies(): Promise<ProjectDependencyGraph> {
    return invoke("project_dependencies");
  },
  searchProject(request: ProjectSearchRequest): Promise<ProjectSearchMatch[]> {
    return invoke("search_project", { request });
  },
  previewProjectReplace(request: ProjectReplaceRequest): Promise<ProjectReplacePlan> {
    return invoke("preview_project_replace", { request });
  },
  previewSymbolRename(request: SymbolRenameRequest): Promise<ProjectReplacePlan> {
    return invoke("preview_symbol_rename", { request });
  },
  applyProjectReplace(plan: ProjectReplacePlan): Promise<ProjectReplaceOutcome> {
    return invoke("apply_project_replace", { plan });
  },
  compile(): Promise<CompileArtifact> {
    return invoke("compile_project");
  },
  forwardSearch(
    sourceFile: string,
    line: number,
    column: number,
    pdfPath: string,
  ): Promise<ForwardSearchResult> {
    return invoke("forward_search", { sourceFile, line, column, pdfPath });
  },
  inverseSearch(
    pdfPath: string,
    page: number,
    x: number,
    y: number,
  ): Promise<InverseSearchResult> {
    return invoke("inverse_search", { pdfPath, page, x, y });
  },
  pdfDocumentInfo(pdfPath: string): Promise<PdfDocumentInfo> {
    return invoke("pdf_document_info", { pdfPath });
  },
  renderPdf(request: PdfRenderRequest): Promise<PdfRenderedImage> {
    return invoke("render_pdf", { request });
  },
  pdfTextHit(pdfPath: string, pageIndex: number, x: number, y: number): Promise<PdfTextHit | null> {
    return invoke("pdf_text_hit", { pdfPath, pageIndex, x, y });
  },
  buildLayoutMap(pdfPath: string): Promise<LayoutMapArtifact> {
    return invoke("build_layout_map", { pdfPath });
  },
  detectToolchain(): Promise<ToolInfo[]> {
    return invoke("detect_toolchain");
  },
  fileUrl(path: string): string {
    return convertFileSrc(path);
  },
};
