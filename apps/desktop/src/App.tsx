import { useEffect, useMemo, useRef, useState } from "react";
import { listen } from "@tauri-apps/api/event";
import { open } from "@tauri-apps/plugin-dialog";
import {
  BookOpen,
  ChevronRight,
  CircleAlert,
  FileCode2,
  FilePlus2,
  FolderOpen,
  ScanText,
  Braces,
  Hammer,
  Library,
  Package,
  PackagePlus,
  Redo2,
  Save,
  ScanLine,
  Search,
  Settings2,
  Trash2,
  Undo2,
} from "lucide-react";
import { MathNodeEditor } from "@visualtex/math-editor";
import { PdfViewer } from "@visualtex/pdf-viewer";
import {
  createOperationId,
  desktopApi,
  type CompileArtifact,
  type DocumentSnapshot,
  type DocumentOcrResult,
  type EditOutcome,
  type ExternalConflictResolution,
  type ExternalFileChange,
  type FormulaOcrResult,
  type ForwardSearchResult,
  type InverseSearchResult,
  type InstalledModelPackage,
  type LayoutMapArtifact,
  type ModelCatalog,
  type ModelKind,
  type ModelPackageInspection,
  type NodeAttributesPatch,
  type OcrWorkerHealth,
  type PdfRect,
  type ProjectIndex,
  type ProjectReplacePlan,
  type ProjectSearchMatch,
  type ProjectSymbol,
  type ProjectTemplateSummary,
  type ToolInfo,
  type VisualNode,
  type VisualPatch,
} from "@visualtex/protocol";
import {
  SourceEditor,
  sourcePositionAtUtf8Byte,
  type SourceChange,
  type SourceCompletion,
  type SourceReveal,
} from "@visualtex/source-editor";
import {
  OCR_REGION_KIND_OPTIONS,
  changeOcrRegionKind,
  documentOcrToLatex,
  moveReadingOrder,
  ocrRegionContent,
  updateOcrRegionContent,
} from "./documentOcr";

function patchNodes(nodes: VisualNode[], patch: VisualPatch): VisualNode[] {
  if ("reset" in patch) return patch.reset.nodes;
  const removed = new Set(patch.replace.removed);
  const byId = new Map(nodes.filter((node) => !removed.has(node.id)).map((node) => [node.id, node]));
  for (const node of patch.replace.upserted) byId.set(node.id, node);
  return [...byId.values()].sort(
    (left, right) => left.source.startByte - right.source.startByte || left.source.endByte - right.source.endByte,
  );
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function modelKindLabel(kind: ModelKind): string {
  switch (kind) {
    case "formula_ocr": return "公式 OCR";
    case "layout_ocr": return "版面分析";
    case "text_ocr": return "文字 OCR";
    case "table_ocr": return "表格 OCR";
  }
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 ** 2) return `${(bytes / 1024).toFixed(1)} KiB`;
  if (bytes < 1024 ** 3) return `${(bytes / 1024 ** 2).toFixed(1)} MiB`;
  return `${(bytes / 1024 ** 3).toFixed(2)} GiB`;
}

interface PendingVisualEdit {
  nodeId: string;
  content: string;
}

export default function App() {
  const [snapshot, setSnapshot] = useState<DocumentSnapshot | null>(null);
  const [files, setFiles] = useState<string[]>([]);
  const [artifact, setArtifact] = useState<CompileArtifact | null>(null);
  const [tools, setTools] = useState<ToolInfo[]>([]);
  const [selectedNode, setSelectedNode] = useState<VisualNode | null>(null);
  const [rightTab, setRightTab] = useState<"pdf" | "structure">("pdf");
  const [bottomTab, setBottomTab] = useState<"diagnostics" | "log" | "toolchain" | "models">("diagnostics");
  const [busy, setBusy] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);
  const [autoCompile, setAutoCompile] = useState(true);
  const [pdfHighlights, setPdfHighlights] = useState<PdfRect[]>([]);
  const [sourceReveal, setSourceReveal] = useState<SourceReveal | null>(null);
  const [layoutMap, setLayoutMap] = useState<LayoutMapArtifact | null>(null);
  const [leftTab, setLeftTab] = useState<"files" | "search" | "symbols">("files");
  const [projectIndex, setProjectIndex] = useState<ProjectIndex>({ symbols: [] });
  const [searchQuery, setSearchQuery] = useState("");
  const [searchResults, setSearchResults] = useState<ProjectSearchMatch[]>([]);
  const [searchCaseSensitive, setSearchCaseSensitive] = useState(false);
  const [searchWholeWord, setSearchWholeWord] = useState(false);
  const [searching, setSearching] = useState(false);
  const [replacementText, setReplacementText] = useState("");
  const [replacePlan, setReplacePlan] = useState<ProjectReplacePlan | null>(null);
  const [renameTarget, setRenameTarget] = useState<ProjectSymbol | null>(null);
  const [renameKey, setRenameKey] = useState("");
  const [projectTemplates, setProjectTemplates] = useState<ProjectTemplateSummary[]>([]);
  const [templateDialogOpen, setTemplateDialogOpen] = useState(false);
  const [externalConflicts, setExternalConflicts] = useState<ExternalFileChange[]>([]);
  const [modelCatalog, setModelCatalog] = useState<ModelCatalog | null>(null);
  const [ocrWorkerHealth, setOcrWorkerHealth] = useState<OcrWorkerHealth | null>(null);
  const [pendingModelInstall, setPendingModelInstall] = useState<ModelPackageInspection | null>(null);
  const [pendingModelRemoval, setPendingModelRemoval] = useState<InstalledModelPackage | null>(null);
  const [formulaOcrResult, setFormulaOcrResult] = useState<FormulaOcrResult | null>(null);
  const [formulaOcrLatex, setFormulaOcrLatex] = useState("");
  const [formulaOcrSourceName, setFormulaOcrSourceName] = useState("");
  const [formulaInsertMode, setFormulaInsertMode] = useState<"inline" | "display">("inline");
  const [documentOcrResult, setDocumentOcrResult] = useState<DocumentOcrResult | null>(null);
  const [documentOcrSourceName, setDocumentOcrSourceName] = useState("");
  const [documentOcrDraft, setDocumentOcrDraft] = useState("");
  const [selectedOcrRegion, setSelectedOcrRegion] = useState(0);
  const [sourceCursorOffset, setSourceCursorOffset] = useState(0);

  const snapshotRef = useRef(snapshot);
  const revisionRef = useRef(0);
  const editQueueRef = useRef<Promise<void>>(Promise.resolve());
  const pendingVisualRef = useRef<PendingVisualEdit | null>(null);
  const visualTimerRef = useRef<number | null>(null);
  const compileTimerRef = useRef<number | null>(null);
  const compilingRef = useRef(false);
  const pendingCompileRef = useRef(false);
  const cursorTimerRef = useRef<number | null>(null);
  const indexTimerRef = useRef<number | null>(null);
  const searchTimerRef = useRef<number | null>(null);
  const externalTimerRef = useRef<number | null>(null);

  snapshotRef.current = snapshot;

  const adoptSnapshot = (next: DocumentSnapshot) => {
    snapshotRef.current = next;
    revisionRef.current = next.revision;
    setSnapshot(next);
  };

  const hydrateOpenedProject = async (next: DocumentSnapshot) => {
    adoptSnapshot(next);
    const [sourceFiles, index] = await Promise.all([
      desktopApi.listFiles(),
      desktopApi.projectIndex(),
    ]);
    setFiles(sourceFiles);
    setProjectIndex(index);
    setSearchQuery("");
    setSearchResults([]);
    setLeftTab("files");
    setExternalConflicts([]);
  };

  const refreshCurrentFile = async () => {
    const current = snapshotRef.current;
    if (!current) return;
    adoptSnapshot(await desktopApi.openFile(current.path));
  };

  const refreshProjectIndex = async () => {
    try {
      setProjectIndex(await desktopApi.projectIndex());
    } catch (cause) {
      setError(`项目索引失败：${errorMessage(cause)}`);
    }
  };

  const openProject = async (create: boolean, templateId = "article") => {
    const path = await open({ directory: true, multiple: false, title: create ? "选择新项目文件夹" : "打开 LaTeX 项目" });
    if (!path || Array.isArray(path)) return;
    setBusy(create ? "正在创建项目" : "正在打开项目");
    setError(null);
    try {
      const next = create
        ? await desktopApi.initProjectTemplate(path, templateId)
        : await desktopApi.openProject(path);
      await hydrateOpenedProject(next);
      setArtifact(null);
      setLayoutMap(null);
      setPdfHighlights([]);
      setSourceReveal(null);
      setSelectedNode(null);
      setExternalConflicts([]);
    } catch (cause) {
      setError(errorMessage(cause));
    } finally {
      setBusy(null);
    }
  };

  const openTemplateChooser = async () => {
    setBusy("正在加载项目模板");
    setError(null);
    try {
      setProjectTemplates(await desktopApi.listProjectTemplates());
      setTemplateDialogOpen(true);
    } catch (cause) {
      setError(`无法加载项目模板：${errorMessage(cause)}`);
    } finally {
      setBusy(null);
    }
  };

  const createFromTemplate = async (templateId: string) => {
    setTemplateDialogOpen(false);
    await openProject(true, templateId);
  };

  useEffect(() => {
    let disposed = false;
    void desktopApi.rootSnapshot()
      .then(async (next) => {
        if (disposed) return;
        await hydrateOpenedProject(next);
        if (disposed) return;
      })
      .catch((cause) => {
        const message = errorMessage(cause);
        if (!disposed && !message.includes("No project is open")) {
          setError(`启动项目载入失败：${message}`);
        }
      });
    return () => {
      disposed = true;
    };
  }, []);

  const handleEditOutcome = (fileId: string, outcome: EditOutcome) => {
    revisionRef.current = Math.max(revisionRef.current, outcome.revision);
    setSnapshot((current) => {
      if (!current || current.fileId !== fileId) return current;
      const next = {
        ...current,
        revision: Math.max(current.revision, outcome.revision),
        nodes: patchNodes(current.nodes, outcome.patch),
        dirty: true,
      };
      snapshotRef.current = next;
      return next;
    });
  };

  const enqueueSourceChange = (change: SourceChange, nextText: string) => {
    const current = snapshotRef.current;
    if (!current) return;
    const fileId = current.fileId;
    const baseRevision = revisionRef.current;
    revisionRef.current += 1;
    const optimistic: DocumentSnapshot = {
      ...current,
      text: nextText,
      revision: revisionRef.current,
      dirty: true,
    };
    snapshotRef.current = optimistic;
    setSnapshot(optimistic);

    editQueueRef.current = editQueueRef.current
      .then(async () => {
        const outcome = await desktopApi.applyTextEdit({
          operationId: createOperationId(),
          origin: "sourceEditor",
          fileId,
          baseRevision,
          startByte: change.startByte,
          endByte: change.endByte,
          replacement: change.replacement,
        });
        handleEditOutcome(fileId, outcome);
      })
      .catch(async (cause) => {
        setError(`源码同步失败：${errorMessage(cause)}`);
        try {
          await refreshCurrentFile();
        } catch (refreshError) {
          setError(`源码同步失败且无法重新载入：${errorMessage(refreshError)}`);
        }
      });
  };

  const enqueuePendingVisual = () => {
    const pending = pendingVisualRef.current;
    pendingVisualRef.current = null;
    if (!pending) return;
    editQueueRef.current = editQueueRef.current
      .then(async () => {
        const current = snapshotRef.current;
        if (!current) return;
        const node = current.nodes.find((candidate) => candidate.id === pending.nodeId);
        if (!node) return;
        const outcome = await desktopApi.applyVisualEdit(
          current.fileId,
          current.revision,
          node.id,
          pending.content,
        );
        handleEditOutcome(current.fileId, outcome);
        await refreshCurrentFile();
      })
      .catch(async (cause) => {
        setError(`可视化同步失败：${errorMessage(cause)}`);
        await refreshCurrentFile().catch(() => undefined);
      });
  };

  const commitVisualEdit = (nodeId: string, content: string) => {
    pendingVisualRef.current = { nodeId, content };
    if (visualTimerRef.current !== null) {
      window.clearTimeout(visualTimerRef.current);
      visualTimerRef.current = null;
    }
    enqueuePendingVisual();
  };

  const deleteVisualNode = async (node: VisualNode) => {
    try {
      await flushEdits();
      const current = snapshotRef.current;
      if (!current) return;
      const currentNode = current.nodes.find((candidate) => candidate.id === node.id)
        ?? current.nodes.find(
          (candidate) => candidate.kind === node.kind
            && candidate.source.startByte === node.source.startByte
            && candidate.source.endByte === node.source.endByte,
        );
      if (!currentNode) throw new Error("页面节点已因重排失效，请重新点击最新 PDF 区域");
      const outcome = await desktopApi.applyTextEdit({
        operationId: createOperationId(),
        origin: currentNode.kind === "inline_math" || currentNode.kind === "display_math"
          ? "mathEditor"
          : "visualEditor",
        fileId: current.fileId,
        baseRevision: current.revision,
        startByte: currentNode.source.startByte,
        endByte: currentNode.source.endByte,
        replacement: "",
      });
      handleEditOutcome(current.fileId, outcome);
      setSelectedNode(null);
      setLayoutMap(null);
      setPdfHighlights([]);
      await refreshCurrentFile();
      setNotice("节点已从 LaTeX 源码中删除；可以使用撤销恢复。重新编译后页面会更新。");
    } catch (cause) {
      setError(`删除节点失败：${errorMessage(cause)}`);
      await refreshCurrentFile().catch(() => undefined);
    }
  };

  const commitNodeAttributes = async (node: VisualNode, patch: NodeAttributesPatch) => {
    try {
      await flushEdits();
      const current = snapshotRef.current;
      if (!current) return;
      const currentNode = current.nodes.find((candidate) => candidate.id === node.id)
        ?? current.nodes.find(
          (candidate) => candidate.kind === node.kind && candidate.source.startByte === node.source.startByte,
        );
      if (!currentNode) throw new Error("页面节点已因重排失效，请重新点击最新 PDF 区域");
      const outcome = await desktopApi.applyNodeAttributes(
        current.fileId,
        current.revision,
        currentNode.id,
        patch,
      );
      handleEditOutcome(current.fileId, outcome);
      await refreshCurrentFile();
    } catch (cause) {
      setError(`属性同步失败：${errorMessage(cause)}`);
      await refreshCurrentFile().catch(() => undefined);
    }
  };

  const flushEdits = async () => {
    if (visualTimerRef.current !== null) {
      window.clearTimeout(visualTimerRef.current);
      visualTimerRef.current = null;
    }
    enqueuePendingVisual();
    await editQueueRef.current;
  };

  const checkExternalChanges = async () => {
    if (!snapshotRef.current) return;
    try {
      await flushEdits();
      const report = await desktopApi.checkExternalChanges();
      if (report.reloaded.length > 0) {
        const current = snapshotRef.current;
        const reloadedCurrent = current
          ? report.reloaded.find((item) => item.fileId === current.fileId)
          : undefined;
        if (reloadedCurrent) adoptSnapshot(reloadedCurrent);
        setArtifact(null);
        setLayoutMap(null);
        setPdfHighlights([]);
        setNotice(
          `已自动载入 ${report.reloaded.length} 个由外部编辑器修改的干净文件。`,
        );
        setFiles(await desktopApi.listFiles());
        await refreshProjectIndex();
      }
      if (report.conflicts.length > 0) {
        setExternalConflicts((current) => {
          const merged = new Map(current.map((change) => [change.fileId, change]));
          for (const change of report.conflicts) merged.set(change.fileId, change);
          return [...merged.values()].sort((left, right) => left.path.localeCompare(right.path));
        });
      }
    } catch (cause) {
      const message = errorMessage(cause);
      if (message.includes("No project is open") && snapshotRef.current) {
        if (externalTimerRef.current !== null) window.clearTimeout(externalTimerRef.current);
        externalTimerRef.current = window.setTimeout(() => {
          externalTimerRef.current = null;
          void checkExternalChanges();
        }, 320);
      } else {
        setError(`检查外部文件变化失败：${message}`);
      }
    }
  };

  const resolveExternalConflict = async (
    change: ExternalFileChange,
    resolution: ExternalConflictResolution,
  ) => {
    setBusy("正在处理外部文件冲突");
    setError(null);
    try {
      await flushEdits();
      const outcome = await desktopApi.resolveExternalConflict(change, resolution);
      if (snapshotRef.current?.fileId === outcome.snapshot.fileId) {
        adoptSnapshot(outcome.snapshot);
      }
      setExternalConflicts((current) =>
        current.filter((candidate) => candidate.fileId !== change.fileId),
      );
      setArtifact(null);
      setLayoutMap(null);
      setPdfHighlights([]);
      setFiles(await desktopApi.listFiles());
      await refreshProjectIndex();
      const action =
        resolution === "reload_disk"
          ? "已载入磁盘版本"
          : resolution === "keep_buffer"
            ? "已保留当前未保存缓冲；下一次保存将明确覆盖磁盘版本"
            : outcome.conflictCopyPath
              ? `已保存冲突副本 ${outcome.conflictCopyPath}`
              : "已保存冲突副本";
      setNotice(`${change.path}：${action}。`);
    } catch (cause) {
      setError(`处理外部文件冲突失败：${errorMessage(cause)}`);
    } finally {
      setBusy(null);
    }
  };

  const previewReplacement = async () => {
    const query = searchQuery.trim();
    if (!query) return;
    setBusy("正在生成替换预览");
    setError(null);
    try {
      await flushEdits();
      setReplacePlan(await desktopApi.previewProjectReplace({
        query,
        replacement: replacementText,
        caseSensitive: searchCaseSensitive,
        wholeWord: searchWholeWord,
        maxReplacements: 5_000,
      }));
    } catch (cause) {
      setError(`无法生成替换预览：${errorMessage(cause)}`);
    } finally {
      setBusy(null);
    }
  };

  const previewRename = async () => {
    if (!renameTarget || !renameKey.trim()) return;
    setBusy("正在检查符号重命名");
    setError(null);
    try {
      await flushEdits();
      setReplacePlan(await desktopApi.previewSymbolRename({
        kind: renameTarget.kind === "label_definition" ? "label" : "citation",
        oldKey: renameTarget.key,
        newKey: renameKey.trim(),
      }));
      setRenameTarget(null);
      setRenameKey("");
    } catch (cause) {
      setError(`无法重命名符号：${errorMessage(cause)}`);
    } finally {
      setBusy(null);
    }
  };

  const applyReplacementPlan = async () => {
    if (!replacePlan || replacePlan.truncated) return;
    setBusy("正在应用项目事务");
    setError(null);
    setNotice(null);
    try {
      await flushEdits();
      const outcome = await desktopApi.applyProjectReplace(replacePlan);
      setReplacePlan(null);
      setArtifact(null);
      setLayoutMap(null);
      setPdfHighlights([]);
      await refreshCurrentFile();
      await refreshProjectIndex();
      setFiles(await desktopApi.listFiles());
      setNotice(`已在 ${outcome.changedFiles.length} 个文件中完成 ${outcome.totalReplacements} 处修改；修改尚未保存到磁盘。`);
      if (autoCompile) {
        window.setTimeout(() => void compile(true), 120);
      }
    } catch (cause) {
      setError(`项目事务失败，未应用部分修改：${errorMessage(cause)}`);
    } finally {
      setBusy(null);
    }
  };

  const save = async () => {
    const current = snapshotRef.current;
    if (!current) return;
    setBusy("正在保存");
    setError(null);
    try {
      await flushEdits();
      await desktopApi.save(current.fileId);
      await refreshCurrentFile();
    } catch (cause) {
      setError(errorMessage(cause));
    } finally {
      setBusy(null);
    }
  };

  const compile = async (silent = false) => {
    if (!snapshotRef.current) return;
    if (compilingRef.current) {
      pendingCompileRef.current = true;
      if (!silent) setNotice("当前编译完成后将自动编译最新内容。");
      return;
    }

    compilingRef.current = true;
    if (!silent) setBusy("正在编译真实 PDF");
    setError(null);
    try {
      await flushEdits();
      const nextArtifact = await desktopApi.compile();
      await refreshCurrentFile();

      if (nextArtifact.sourceRevision !== snapshotRef.current?.revision) {
        pendingCompileRef.current = true;
        return;
      }

      setArtifact(nextArtifact);
      setLayoutMap(null);
      setPdfHighlights([]);
      if (nextArtifact.status === "succeeded" && nextArtifact.pdfPath) {
        try {
          const nextLayoutMap = await desktopApi.buildLayoutMap(nextArtifact.pdfPath);
          if (nextLayoutMap.sourceRevision !== snapshotRef.current?.revision) {
            pendingCompileRef.current = true;
          } else if (nextLayoutMap.compileStatus !== "succeeded") {
            setLayoutMap(null);
            setError("结构化页面映射编译失败；请在编译日志中检查影子工程诊断。");
          } else {
            setLayoutMap(nextLayoutMap);
            const editableCount = nextLayoutMap.boxes.filter(
              (box) => box.confidence === "exact" || box.confidence === "high",
            ).length;
            if (editableCount === 0) {
              setNotice("PDF 已生成，但当前文档没有通过排版一致性校验的可直接编辑节点；仍可双击页面跳转源码。");
            }
          }
        } catch (layoutError) {
          setError(`PDF 节点映射失败，预览仍可使用：${errorMessage(layoutError)}`);
        }
      } else {
        setBottomTab("diagnostics");
      }
    } catch (cause) {
      setError(errorMessage(cause));
    } finally {
      compilingRef.current = false;
      if (!silent) setBusy(null);
      if (pendingCompileRef.current && snapshotRef.current) {
        pendingCompileRef.current = false;
        window.setTimeout(() => void compile(true), 0);
      }
    }
  };

  const undoRedo = async (mode: "undo" | "redo") => {
    const current = snapshotRef.current;
    if (!current) return;
    try {
      await flushEdits();
      const outcome = await desktopApi[mode](current.fileId);
      handleEditOutcome(current.fileId, outcome);
      await refreshCurrentFile();
    } catch (cause) {
      setError(errorMessage(cause));
    }
  };

  const openFile = async (path: string): Promise<boolean> => {
    try {
      await flushEdits();
      adoptSnapshot(await desktopApi.openFile(path));
      setSelectedNode(null);
      return true;
    } catch (cause) {
      setError(errorMessage(cause));
      return false;
    }
  };

  const revealSourceLocation = async (path: string, line: number, column?: number | null) => {
    if (!(await openFile(path))) return;
    setSourceReveal({
      line,
      column,
      requestId: Date.now(),
    });
  };

  const handleVisualNodeSelect = (node: VisualNode) => {
    setSelectedNode(node);
    const current = snapshotRef.current;
    if (!current || node.source.fileId !== current.fileId) return;
    const position = sourcePositionAtUtf8Byte(current.text, node.source.startByte);
    setSourceReveal({
      line: position.line,
      column: position.column,
      requestId: Date.now(),
      focus: false,
    });
  };

  const handleInverseSearch = async (result: InverseSearchResult) => {
    const normalizedSource = result.sourcePath.replaceAll("\\", "/").replace("/./", "/");
    const match = files.find((file) => {
      const normalizedFile = file.replaceAll("\\", "/");
      return normalizedSource === normalizedFile || normalizedSource.endsWith(`/${normalizedFile}`);
    });
    if (!match) {
      setError(`SyncTeX 返回的文件不在当前项目列表中：${result.sourcePath}`);
      return;
    }
    await revealSourceLocation(match, result.line, result.column);
  };

  useEffect(() => {
    let disposed = false;
    const unlisteners: Array<() => void> = [];

    const consumeDeepLinks = async () => {
      const values = await desktopApi.drainDeepLinks();
      for (const value of values) {
        if (disposed) return;
        await desktopApi.processDeepLink(value);
      }
    };

    const register = async () => {
      const listeners = await Promise.all([
        listen<DocumentSnapshot>("visualtex://project-opened", ({ payload }) => {
          if (disposed) return;
          setArtifact(null);
          setLayoutMap(null);
          setPdfHighlights([]);
          setSourceReveal(null);
          setSelectedNode(null);
          void hydrateOpenedProject(payload).catch((cause) => {
            setError(`深链接项目载入失败：${errorMessage(cause)}`);
          });
        }),
        listen("visualtex://deep-link-received", () => {
          if (!disposed) void consumeDeepLinks().catch((cause) => {
            setError(`深链接消费失败：${errorMessage(cause)}`);
          });
        }),
        listen<string>("visualtex://deep-link-error", ({ payload }) => {
          if (!disposed) setError(`深链接失败：${payload}`);
        }),
        listen<ForwardSearchResult>("visualtex://forward-search-result", ({ payload }) => {
          if (disposed) return;
          const now = new Date().toISOString();
          const sourceRevision = snapshotRef.current?.revision ?? 0;
          setArtifact((current) => current?.pdfPath === payload.pdfPath
            ? { ...current, sourceRevision }
            : {
                buildId: `deep-link-${Date.now()}`,
                sourceRevision,
                pdfPath: payload.pdfPath,
                synctexPath: null,
                diagnostics: [],
                status: "succeeded",
                startedAt: now,
                finishedAt: now,
                stdout: "",
                stderr: "",
              });
          setPdfHighlights(payload.boxes);
          setRightTab("pdf");
        }),
        listen<InverseSearchResult>("visualtex://inverse-search-result", ({ payload }) => {
          if (disposed) return;
          void desktopApi.listFiles().then(async (sourceFiles) => {
            if (disposed) return;
            setFiles(sourceFiles);
            const normalizedSource = payload.sourcePath.replaceAll("\\", "/").replace("/./", "/");
            const match = sourceFiles.find((file) => {
              const normalizedFile = file.replaceAll("\\", "/");
              return normalizedSource === normalizedFile || normalizedSource.endsWith(`/${normalizedFile}`);
            });
            if (!match) {
              setError(`SyncTeX 返回的文件不在当前项目列表中：${payload.sourcePath}`);
              return;
            }
            await revealSourceLocation(match, payload.line, payload.column);
          }).catch((cause) => setError(`反向 SyncTeX 失败：${errorMessage(cause)}`));
        }),
      ]);
      if (disposed) {
        listeners.forEach((stop) => stop());
        return;
      }
      unlisteners.push(...listeners);
      await consumeDeepLinks();
    };

    void register().catch((cause) => setError(`深链接监听启动失败：${errorMessage(cause)}`));
    return () => {
      disposed = true;
      unlisteners.forEach((stop) => stop());
    };
  }, []);

  const handleSourceCursor = (utf16Offset: number) => {
    setSourceCursorOffset(utf16Offset);
    if (cursorTimerRef.current !== null) window.clearTimeout(cursorTimerRef.current);
    cursorTimerRef.current = window.setTimeout(() => {
      cursorTimerRef.current = null;
      const current = snapshotRef.current;
      const currentArtifact = artifact;
      if (!current || !currentArtifact?.pdfPath || currentArtifact.status !== "succeeded") return;
      if (currentArtifact.sourceRevision !== current.revision) {
        setPdfHighlights([]);
        return;
      }
      const prefix = current.text.slice(0, Math.max(0, Math.min(current.text.length, utf16Offset)));
      const lastLineBreak = prefix.lastIndexOf("\n");
      const line = prefix.split("\n").length;
      const column = prefix.length - lastLineBreak;
      void desktopApi
        .forwardSearch(current.path, line, column, currentArtifact.pdfPath)
        .then((result) => setPdfHighlights(result.boxes))
        .catch(() => setPdfHighlights([]));
    }, 180);
  };

  const inspectToolchain = async () => {
    setBusy("正在检测工具链");
    try {
      await flushEdits();
      setTools(await desktopApi.detectToolchain());
      setBottomTab("toolchain");
    } catch (cause) {
      setError(errorMessage(cause));
    } finally {
      setBusy(null);
    }
  };

  const refreshModelSettings = async (openTab = false) => {
    setBusy("正在读取本地模型");
    setError(null);
    try {
      const catalog = await desktopApi.listModelPackages();
      setModelCatalog(catalog);
      if (snapshotRef.current) {
        setOcrWorkerHealth(await desktopApi.ocrHealth());
      } else {
        setOcrWorkerHealth(null);
      }
      if (openTab) setBottomTab("models");
    } catch (cause) {
      setError(`模型设置读取失败：${errorMessage(cause)}`);
    } finally {
      setBusy(null);
    }
  };

  const chooseModelPackage = async () => {
    const selected = await open({
      directory: true,
      multiple: false,
      title: "选择包含 visualtex-model.json 的离线模型包",
    });
    if (typeof selected !== "string") return;
    setBusy("正在校验模型包");
    setError(null);
    try {
      setPendingModelInstall(await desktopApi.inspectModelPackage(selected));
    } catch (cause) {
      setError(`模型包校验失败：${errorMessage(cause)}`);
    } finally {
      setBusy(null);
    }
  };

  const installPendingModel = async () => {
    if (!pendingModelInstall) return;
    setBusy("正在安装离线模型");
    setError(null);
    try {
      const installed = await desktopApi.installModelPackage(pendingModelInstall.sourcePath);
      setPendingModelInstall(null);
      setNotice(
        `已安装 ${installed.manifest.id}@${installed.manifest.version}，文件校验和 ${installed.installedSha256.slice(0, 12)}…`,
      );
      await refreshModelSettings(true);
    } catch (cause) {
      setError(`模型安装失败：${errorMessage(cause)}`);
    } finally {
      setBusy(null);
    }
  };

  const activateModel = async (model: InstalledModelPackage) => {
    setBusy("正在切换活动模型");
    setError(null);
    try {
      const catalog = await desktopApi.activateModelPackage(
        model.manifest.kind,
        model.manifest.id,
        model.manifest.version,
      );
      setModelCatalog(catalog);
      if (snapshotRef.current) setOcrWorkerHealth(await desktopApi.ocrHealth());
      setNotice(`已启用 ${model.manifest.id}@${model.manifest.version}。`);
    } catch (cause) {
      setError(`模型切换失败：${errorMessage(cause)}`);
    } finally {
      setBusy(null);
    }
  };

  const removePendingModel = async () => {
    if (!pendingModelRemoval) return;
    setBusy("正在卸载模型");
    setError(null);
    try {
      const catalog = await desktopApi.removeModelPackage(
        pendingModelRemoval.manifest.id,
        pendingModelRemoval.manifest.version,
      );
      setModelCatalog(catalog);
      setPendingModelRemoval(null);
      if (snapshotRef.current) setOcrWorkerHealth(await desktopApi.ocrHealth());
      setNotice("模型已从本机应用数据目录中移除。");
    } catch (cause) {
      setError(`模型卸载失败：${errorMessage(cause)}`);
    } finally {
      setBusy(null);
    }
  };

  const chooseFormulaImage = async () => {
    const selected = await open({
      directory: false,
      multiple: false,
      title: "选择需要识别的公式图片",
      filters: [{ name: "Images", extensions: ["png", "jpg", "jpeg", "webp", "bmp", "tif", "tiff"] }],
    });
    if (typeof selected !== "string") return;
    setBusy("正在离线识别公式");
    setError(null);
    try {
      const result = await desktopApi.recognizeFormulaImage(selected);
      setFormulaOcrResult(result);
      setFormulaOcrLatex(result.candidates[0]?.latex ?? "");
      setFormulaOcrSourceName(selected.replaceAll("\\", "/").split("/").pop() ?? selected);
      setFormulaInsertMode("inline");
    } catch (cause) {
      setError(`公式 OCR 失败：${errorMessage(cause)}`);
      await refreshModelSettings(true).catch(() => undefined);
    } finally {
      setBusy(null);
    }
  };

  const insertRecognizedFormula = async () => {
    const current = snapshotRef.current;
    const latex = formulaOcrLatex.trim();
    if (!current || !latex) return;
    setBusy("正在写回识别公式");
    setError(null);
    try {
      await flushEdits();
      const latest = snapshotRef.current;
      if (!latest) return;
      const selectedMath = selectedNode
        ? latest.nodes.find(
          (node) => node.id === selectedNode.id && ["inline_math", "display_math"].includes(node.kind),
        )
        : undefined;
      let outcome: EditOutcome;
      if (selectedMath?.support === "native") {
        outcome = await desktopApi.applyVisualEdit(
          latest.fileId,
          latest.revision,
          selectedMath.id,
          latex,
        );
      } else {
        const safeOffset = Math.max(0, Math.min(latest.text.length, sourceCursorOffset));
        const byteOffset = new TextEncoder().encode(latest.text.slice(0, safeOffset)).length;
        const replacement = formulaInsertMode === "display"
          ? `\n\\begin{equation}\n  ${latex}\n\\end{equation}\n`
          : `$${latex}$`;
        outcome = await desktopApi.applyTextEdit({
          operationId: createOperationId(),
          origin: "ocr",
          fileId: latest.fileId,
          baseRevision: latest.revision,
          startByte: byteOffset,
          endByte: byteOffset,
          replacement,
        });
      }
      handleEditOutcome(latest.fileId, outcome);
      setFormulaOcrResult(null);
      setFormulaOcrLatex("");
      await refreshCurrentFile();
      setNotice("识别公式已写入当前未保存缓冲，可使用统一撤销恢复。");
    } catch (cause) {
      setError(`公式写回失败：${errorMessage(cause)}`);
      await refreshCurrentFile().catch(() => undefined);
    } finally {
      setBusy(null);
    }
  };

  const chooseDocumentImage = async () => {
    const selected = await open({
      directory: false,
      multiple: false,
      title: "选择需要进行整页 OCR 的文档图片",
      filters: [{ name: "Images", extensions: ["png", "jpg", "jpeg", "webp", "bmp", "tif", "tiff"] }],
    });
    if (typeof selected !== "string") return;
    setBusy("正在分析页面版面与阅读顺序");
    setError(null);
    try {
      const result = await desktopApi.recognizeDocumentImage(selected);
      setDocumentOcrResult(result);
      setDocumentOcrDraft(documentOcrToLatex(result));
      setSelectedOcrRegion(result.readingOrder[0] ?? 0);
      setDocumentOcrSourceName(selected.replaceAll("\\", "/").split("/").pop() ?? selected);
    } catch (cause) {
      setError(`整页 OCR 失败：${errorMessage(cause)}`);
      await refreshModelSettings(true).catch(() => undefined);
    } finally {
      setBusy(null);
    }
  };

  const updateDocumentOcrRegion = (index: number, value: string) => {
    setDocumentOcrResult((current) => {
      if (!current || !current.regions[index]) return current;
      const regions = current.regions.map((region, regionIndex) =>
        regionIndex === index ? updateOcrRegionContent(region, value) : region,
      );
      return { ...current, regions };
    });
  };

  const updateDocumentOcrRegionKind = (index: number, kind: string) => {
    setDocumentOcrResult((current) => {
      if (!current || !current.regions[index]) return current;
      const regions = current.regions.map((region, regionIndex) =>
        regionIndex === index ? changeOcrRegionKind(region, kind) : region,
      );
      return { ...current, regions };
    });
  };

  const moveDocumentOcrRegion = (regionIndex: number, delta: -1 | 1) => {
    setDocumentOcrResult((current) => current
      ? { ...current, readingOrder: moveReadingOrder(current.readingOrder, regionIndex, delta) }
      : current);
  };

  const rebuildDocumentOcrDraft = () => {
    if (documentOcrResult) setDocumentOcrDraft(documentOcrToLatex(documentOcrResult));
  };

  const createDocumentOcrProject = async () => {
    const draft = documentOcrDraft.trim();
    const result = documentOcrResult;
    const sourceImage = result?.imagePath;
    if (!draft || !result || !sourceImage) return;
    const target = await open({
      directory: true,
      multiple: false,
      title: "选择一个空文件夹生成 OCR LaTeX 项目",
    });
    if (typeof target !== "string") return;
    setBusy("正在原子生成 OCR LaTeX 项目");
    setError(null);
    try {
      await flushEdits();
      const next = await desktopApi.createOcrProject(target, sourceImage, draft, result);
      adoptSnapshot(next);
      setFiles(await desktopApi.listFiles());
      setProjectIndex(await desktopApi.projectIndex());
      setDocumentOcrResult(null);
      setDocumentOcrDraft("");
      setArtifact(null);
      setLayoutMap(null);
      setPdfHighlights([]);
      setSelectedNode(null);
      setSourceReveal(null);
      setSearchQuery("");
      setSearchResults([]);
      setLeftTab("files");
      setNotice("OCR 项目已生成：原图、结构化识别 JSON、可编辑正文和独立原页对照文件均已保留。");
      if (autoCompile) {
        window.setTimeout(() => void compile(true), 120);
      }
    } catch (cause) {
      setError(`生成 OCR 项目失败：${errorMessage(cause)}`);
    } finally {
      setBusy(null);
    }
  };

  const insertDocumentOcrDraft = async () => {
    const draft = documentOcrDraft.trim();
    if (!draft) return;
    setBusy("正在写入整页 OCR 草稿");
    setError(null);
    try {
      await flushEdits();
      const latest = snapshotRef.current;
      if (!latest) return;
      const safeOffset = Math.max(0, Math.min(latest.text.length, sourceCursorOffset));
      const byteOffset = new TextEncoder().encode(latest.text.slice(0, safeOffset)).length;
      const outcome = await desktopApi.applyTextEdit({
        operationId: createOperationId(),
        origin: "ocr",
        fileId: latest.fileId,
        baseRevision: latest.revision,
        startByte: byteOffset,
        endByte: byteOffset,
        replacement: `\n${draft}\n`,
      });
      handleEditOutcome(latest.fileId, outcome);
      setDocumentOcrResult(null);
      setDocumentOcrDraft("");
      await refreshCurrentFile();
      setNotice("整页 OCR 草稿已写入当前未保存缓冲；请编译检查后再保存。");
    } catch (cause) {
      setError(`整页 OCR 写回失败：${errorMessage(cause)}`);
      await refreshCurrentFile().catch(() => undefined);
    } finally {
      setBusy(null);
    }
  };

  useEffect(() => {
    if (!snapshot) return;
    let disposed = false;
    let unlisten: (() => void) | undefined;
    void listen("visualtex://project-source-changed", () => {
      if (externalTimerRef.current !== null) window.clearTimeout(externalTimerRef.current);
      externalTimerRef.current = window.setTimeout(() => {
        externalTimerRef.current = null;
        if (!disposed) void checkExternalChanges();
      }, 180);
    }).then((stop) => {
      if (disposed) stop();
      else unlisten = stop;
    }).catch((cause) => setError(`无法启动文件监听：${errorMessage(cause)}`));
    return () => {
      disposed = true;
      unlisten?.();
      if (externalTimerRef.current !== null) {
        window.clearTimeout(externalTimerRef.current);
        externalTimerRef.current = null;
      }
    };
  }, [Boolean(snapshot)]);

  useEffect(() => {
    if (!autoCompile || !snapshot?.dirty) return;
    if (compileTimerRef.current !== null) window.clearTimeout(compileTimerRef.current);
    compileTimerRef.current = window.setTimeout(() => {
      compileTimerRef.current = null;
      void compile(true);
    }, 1400);
    return () => {
      if (compileTimerRef.current !== null) window.clearTimeout(compileTimerRef.current);
    };
  }, [snapshot?.text, autoCompile]);

  useEffect(() => {
    if (!snapshot) return;
    if (indexTimerRef.current !== null) window.clearTimeout(indexTimerRef.current);
    indexTimerRef.current = window.setTimeout(() => {
      indexTimerRef.current = null;
      void editQueueRef.current.then(refreshProjectIndex);
    }, 360);
    return () => {
      if (indexTimerRef.current !== null) window.clearTimeout(indexTimerRef.current);
    };
  }, [snapshot?.revision, files.length]);

  useEffect(() => {
    const query = searchQuery.trim();
    if (!snapshot || query.length === 0) {
      setSearchResults([]);
      setSearching(false);
      return;
    }
    if (searchTimerRef.current !== null) window.clearTimeout(searchTimerRef.current);
    setSearching(true);
    let cancelled = false;
    searchTimerRef.current = window.setTimeout(() => {
      searchTimerRef.current = null;
      void editQueueRef.current
        .then(() => desktopApi.searchProject({
          query,
          caseSensitive: searchCaseSensitive,
          wholeWord: searchWholeWord,
          maxResults: 250,
        }))
        .then((results) => {
          if (!cancelled) setSearchResults(results);
        })
        .catch((cause) => {
          if (!cancelled) setError(`项目搜索失败：${errorMessage(cause)}`);
        })
        .finally(() => {
          if (!cancelled) setSearching(false);
        });
    }, 220);
    return () => {
      cancelled = true;
      if (searchTimerRef.current !== null) window.clearTimeout(searchTimerRef.current);
    };
  }, [searchQuery, searchCaseSensitive, searchWholeWord, snapshot?.revision]);

  useEffect(() => {
    if (!selectedNode || !snapshot) return;
    const exact = snapshot.nodes.find((node) => node.id === selectedNode.id);
    if (exact) {
      if (exact !== selectedNode) setSelectedNode(exact);
      return;
    }
    const replacement = snapshot.nodes.find(
      (node) => node.kind === selectedNode.kind && node.source.startByte === selectedNode.source.startByte,
    );
    setSelectedNode(replacement ?? null);
  }, [snapshot?.nodes]);

  const sourceCompletions = useMemo<SourceCompletion[]>(() => {
    const builtInCommands = [
      "\\section", "\\subsection", "\\label", "\\ref", "\\eqref", "\\cite",
      "\\begin", "\\end", "\\includegraphics", "\\caption", "\\item",
    ].map((label) => ({ label, kind: "command" as const, detail: "LaTeX" }));
    const indexed = projectIndex.symbols.flatMap<SourceCompletion>((symbol) => {
      const detail = `${symbol.file}:${symbol.line}`;
      switch (symbol.kind) {
        case "label_definition":
          return [{ label: symbol.key, kind: "label", detail }];
        case "bibliography_entry":
          return [{ label: symbol.key, kind: "citation", detail: symbol.detail ?? detail }];
        case "macro_definition":
          return [{ label: symbol.key, kind: "command", detail }];
        default:
          return [];
      }
    });
    const unique = new Map<string, SourceCompletion>();
    for (const completion of [...builtInCommands, ...indexed]) {
      unique.set(`${completion.kind}:${completion.label}`, completion);
    }
    return [...unique.values()].sort((left, right) =>
      left.kind.localeCompare(right.kind) || left.label.localeCompare(right.label),
    );
  }, [projectIndex]);

  const navigableSymbols = useMemo(
    () => projectIndex.symbols.filter((symbol) =>
      ["label_definition", "bibliography_entry", "macro_definition"].includes(symbol.kind),
    ),
    [projectIndex],
  );

  const mathNode = selectedNode && ["inline_math", "display_math"].includes(selectedNode.kind) ? selectedNode : null;
  const diagnostics = artifact?.diagnostics ?? [];
  const activeModelKeys = new Set(
    (modelCatalog?.active ?? []).map(
      (model) => `${model.manifest.kind}:${model.manifest.id}:${model.manifest.version}`,
    ),
  );
  const documentOcrLowConfidenceCount = documentOcrResult?.regions.filter(
    (region) => region.kind.toLowerCase() !== "ignore" && region.confidence < 0.65,
  ).length ?? 0;
  const title = snapshot ? `${snapshot.path}${snapshot.dirty ? " •" : ""}` : "未打开项目";
  const templateChooser = templateDialogOpen ? (
    <div className="transaction-backdrop" role="presentation">
      <section className="template-dialog" role="dialog" aria-modal="true" aria-label="选择项目模板">
        <header>
          <div>
            <p className="eyebrow">新建项目</p>
            <h2>选择标准 LaTeX 项目模板</h2>
            <p>创建时不会覆盖目标文件夹中已有的同名文件。</p>
          </div>
          <button type="button" className="dialog-close" onClick={() => setTemplateDialogOpen(false)}>×</button>
        </header>
        <div className="template-grid">
          {projectTemplates.map((template) => (
            <button type="button" key={template.id} onClick={() => void createFromTemplate(template.id)}>
              <span className="template-engine">{template.engine}</span>
              <strong>{template.name}</strong>
              <span>{template.description}</span>
              <code>{template.rootFile}</code>
            </button>
          ))}
        </div>
      </section>
    </div>
  ) : null;

  if (!snapshot) {
    return (
      <>
      {templateChooser}
      <main className="welcome-shell">
        <section className="welcome-card">
          <div className="brand-mark">VS</div>
          <p className="eyebrow">visualstudio</p>
          <h1>真实 PDF 排版驱动的离线论文编辑器</h1>
          <p className="welcome-copy">LaTeX 源文件始终是唯一权威内容。源码与真实编译页面双向同步，标题、正文和公式可在页面中原位编辑。</p>
          <div className="welcome-actions">
            <button className="primary" onClick={() => void openProject(false)}><FolderOpen size={18} />打开项目</button>
            <button onClick={() => void openTemplateChooser()}><FilePlus2 size={18} />创建项目</button>
          </div>
          {busy && <p className="status-line">{busy}…</p>}
          {error && <p className="error-banner"><CircleAlert size={17} />{error}</p>}
        </section>
      </main>
      </>
    );
  }

  return (
    <main className="app-shell">
      {templateChooser}
      <header className="topbar">
        <div className="brand-compact"><span>VS</span><strong>visualstudio</strong></div>
        <div className="document-title" title={snapshot.path}>{title}</div>
        <div className="toolbar">
          <button title="打开项目" onClick={() => void openProject(false)}><FolderOpen size={17} /></button>
          <button title="新建项目" onClick={() => void openTemplateChooser()}><FilePlus2 size={17} /></button>
          <button title="保存" onClick={() => void save()}><Save size={17} /></button>
          <button title="撤销" onClick={() => void undoRedo("undo")}><Undo2 size={17} /></button>
          <button title="重做" onClick={() => void undoRedo("redo")}><Redo2 size={17} /></button>
          <button className="compile-button" onClick={() => void compile(false)}><Hammer size={17} />编译</button>
          <label className="toggle"><input type="checkbox" checked={autoCompile} onChange={(event) => setAutoCompile(event.target.checked)} />自动编译</label>
          <button title="公式图片 OCR" onClick={() => void chooseFormulaImage()}><ScanLine size={17} /></button>
          <button title="整页文档 OCR" onClick={() => void chooseDocumentImage()}><ScanText size={17} /></button>
          <button title="本地模型" onClick={() => void refreshModelSettings(true)}><Package size={17} /></button>
          <button title="工具链" onClick={() => void inspectToolchain()}><Settings2 size={17} /></button>
        </div>
      </header>

      {error && <div className="error-banner app-error"><CircleAlert size={16} />{error}<button onClick={() => setError(null)}>×</button></div>}
      {notice && <div className="notice-banner"><span>{notice}</span><button onClick={() => setNotice(null)}>×</button></div>}
      {busy && <div className="busy-line">{busy}…</div>}

      {externalConflicts[0] && (
        <div className="transaction-backdrop" role="presentation">
          <section className="external-conflict-dialog" role="dialog" aria-modal="true" aria-label="外部文件修改冲突">
            <header>
              <div>
                <p className="eyebrow">外部文件冲突</p>
                <h2>{externalConflicts[0].path}</h2>
                <p>
                  {externalConflicts[0].kind === "deleted"
                    ? "这个文件已被外部程序删除，但 visualstudio 中仍保留当前缓冲。"
                    : "磁盘文件已被外部程序修改，同时 visualstudio 中存在未保存内容。"}
                </p>
              </div>
              <span className="conflict-count">1 / {externalConflicts.length}</span>
            </header>
            <div className="external-conflict-actions">
              {externalConflicts[0].kind === "modified" && (
                <button
                  type="button"
                  className="danger-choice"
                  onClick={() => void resolveExternalConflict(externalConflicts[0]!, "reload_disk")}
                >
                  <strong>载入磁盘版本</strong>
                  <span>放弃 visualstudio 中尚未保存的内容。</span>
                </button>
              )}
              <button
                type="button"
                onClick={() => void resolveExternalConflict(externalConflicts[0]!, "keep_buffer")}
              >
                <strong>{externalConflicts[0].kind === "deleted" ? "保留并重新创建文件" : "保留当前缓冲"}</strong>
                <span>
                  {externalConflicts[0].kind === "deleted"
                    ? "继续使用内存中的内容，下一次保存时重新创建该文件。"
                    : "接受当前磁盘状态为基线；下一次保存时明确覆盖磁盘版本。"}
                </span>
              </button>
              <button
                type="button"
                className="recommended-choice"
                onClick={() => void resolveExternalConflict(externalConflicts[0]!, "save_copy_and_reload")}
              >
                <strong>
                  {externalConflicts[0].kind === "deleted"
                    ? "保存本地冲突副本并保留缓冲"
                    : "保存本地冲突副本后载入磁盘"}
                </strong>
                <span>先将 visualstudio 缓冲写入 .visualtex/conflicts，再采用安全处理。</span>
              </button>
            </div>
            <footer>
              <p>在选择处理方式前，visualstudio 不会覆盖磁盘文件。</p>
            </footer>
          </section>
        </div>
      )}

      {documentOcrResult && (
        <div className="transaction-backdrop" role="presentation">
          <section className="document-ocr-dialog" role="dialog" aria-modal="true" aria-label="整页 OCR 校对">
            <header>
              <div>
                <p className="eyebrow">整页 OCR 与版面校对</p>
                <h2>{documentOcrSourceName || "文档图片"}</h2>
                <p>
                  {documentOcrResult.modelVersion ?? "本地布局模型"} · {documentOcrResult.regions.length} 个区域
                  {documentOcrLowConfidenceCount > 0 && ` · ${documentOcrLowConfidenceCount} 个低置信度区域`}
                </p>
              </div>
              <button type="button" className="dialog-close" onClick={() => setDocumentOcrResult(null)}>×</button>
            </header>
            <div className="document-ocr-content">
              {documentOcrResult.warnings.length > 0 && (
                <div className="formula-ocr-warnings document-warnings">
                  {documentOcrResult.warnings.map((warning) => <span key={warning}>{warning}</span>)}
                </div>
              )}
              <div className="document-ocr-preview">
                <div
                  className="document-ocr-canvas"
                  style={{
                    aspectRatio: `${Math.max(1, documentOcrResult.pageWidth)} / ${Math.max(1, documentOcrResult.pageHeight)}`,
                  }}
                >
                  {documentOcrResult.imagePath && (
                    <img src={desktopApi.fileUrl(documentOcrResult.imagePath)} alt="OCR document input" draggable={false} />
                  )}
                  {documentOcrResult.regions.map((region, index) => {
                    const order = documentOcrResult.readingOrder.indexOf(index);
                    const ignored = region.kind.toLowerCase() === "ignore";
                    return (
                      <button
                        type="button"
                        key={index}
                        className={`document-region-box${selectedOcrRegion === index ? " active" : ""}${ignored ? " ignored" : ""}${!ignored && region.confidence < 0.65 ? " low-confidence" : ""}`}
                        title={`${order >= 0 ? order + 1 : "-"}. ${region.kind}`}
                        onClick={() => setSelectedOcrRegion(index)}
                        style={{
                          left: `${region.x / Math.max(1, documentOcrResult.pageWidth) * 100}%`,
                          top: `${region.y / Math.max(1, documentOcrResult.pageHeight) * 100}%`,
                          width: `${region.width / Math.max(1, documentOcrResult.pageWidth) * 100}%`,
                          height: `${region.height / Math.max(1, documentOcrResult.pageHeight) * 100}%`,
                        }}
                      ><span>{order >= 0 ? order + 1 : "-"}</span></button>
                    );
                  })}
                </div>
              </div>
              <div className="document-region-editor">
                <div className="document-region-list">
                  {documentOcrResult.readingOrder.map((regionIndex, order) => {
                    const region = documentOcrResult.regions[regionIndex];
                    if (!region) return null;
                    const normalizedKind = region.kind.toLowerCase();
                    const knownKind = OCR_REGION_KIND_OPTIONS.some((option) => option.value === normalizedKind);
                    const formulaRegion = normalizedKind.includes("formula");
                    const ignoredRegion = normalizedKind === "ignore";
                    return (
                      <article
                        className={`${selectedOcrRegion === regionIndex ? "active" : ""}${region.confidence < 0.65 && !ignoredRegion ? " low-confidence" : ""}`}
                        key={regionIndex}
                      >
                        <div className="document-region-toolbar">
                          <button className="document-region-select" type="button" onClick={() => setSelectedOcrRegion(regionIndex)}>
                            <strong>{order + 1}. 区域 {regionIndex + 1}</strong>
                            <span>{Math.round(region.confidence * 100)}%</span>
                          </button>
                          <select
                            value={normalizedKind}
                            onFocus={() => setSelectedOcrRegion(regionIndex)}
                            onChange={(event) => updateDocumentOcrRegionKind(regionIndex, event.target.value)}
                            aria-label={`区域 ${regionIndex + 1} 类型`}
                          >
                            {!knownKind && <option value={normalizedKind}>{region.kind}</option>}
                            {OCR_REGION_KIND_OPTIONS.map((option) => (
                              <option value={option.value} key={option.value}>{option.label}</option>
                            ))}
                          </select>
                          <div className="document-region-order-buttons">
                            <button
                              type="button"
                              title="在阅读顺序中上移"
                              disabled={order === 0}
                              onClick={() => moveDocumentOcrRegion(regionIndex, -1)}
                            >↑</button>
                            <button
                              type="button"
                              title="在阅读顺序中下移"
                              disabled={order === documentOcrResult.readingOrder.length - 1}
                              onClick={() => moveDocumentOcrRegion(regionIndex, 1)}
                            >↓</button>
                          </div>
                        </div>
                        {ignoredRegion ? (
                          <p className="document-region-ignored">该区域不会进入 LaTeX 草稿。</p>
                        ) : formulaRegion ? (
                          <div className="document-region-formula" onFocus={() => setSelectedOcrRegion(regionIndex)}>
                            <MathNodeEditor
                              value={ocrRegionContent(region)}
                              onChange={(value) => updateDocumentOcrRegion(regionIndex, value)}
                            />
                            <textarea
                              value={ocrRegionContent(region)}
                              rows={3}
                              onFocus={() => setSelectedOcrRegion(regionIndex)}
                              onChange={(event) => updateDocumentOcrRegion(regionIndex, event.target.value)}
                              aria-label={`区域 ${regionIndex + 1} 公式 LaTeX`}
                              spellCheck={false}
                            />
                          </div>
                        ) : (
                          <textarea
                            value={ocrRegionContent(region)}
                            rows={normalizedKind.includes("table") ? 8 : 3}
                            onFocus={() => setSelectedOcrRegion(regionIndex)}
                            onChange={(event) => updateDocumentOcrRegion(regionIndex, event.target.value)}
                            spellCheck={false}
                          />
                        )}
                      </article>
                    );
                  })}
                </div>
              </div>
              <div className="document-latex-draft">
                <header>
                  <strong>LaTeX 草稿</strong>
                  <button type="button" onClick={rebuildDocumentOcrDraft}>根据校对区域重新生成</button>
                </header>
                <textarea
                  value={documentOcrDraft}
                  onChange={(event) => setDocumentOcrDraft(event.target.value)}
                  spellCheck={false}
                />
              </div>
            </div>
            <footer>
              <span>草稿将插入源码光标位置；页面框、低置信度和表格内容仍需人工检查。</span>
              <button type="button" onClick={() => setDocumentOcrResult(null)}>取消</button>
              <button
                type="button"
                disabled={!documentOcrDraft.trim() || !documentOcrResult.imagePath}
                onClick={() => void createDocumentOcrProject()}
              >
                生成新 OCR 项目
              </button>
              <button type="button" className="primary" disabled={!documentOcrDraft.trim()} onClick={() => void insertDocumentOcrDraft()}>
                写入当前缓冲
              </button>
            </footer>
          </section>
        </div>
      )}

      {formulaOcrResult && (
        <div className="transaction-backdrop" role="presentation">
          <section className="formula-ocr-dialog" role="dialog" aria-modal="true" aria-label="公式 OCR 校对">
            <header>
              <div>
                <p className="eyebrow">离线公式 OCR 校对</p>
                <h2>{formulaOcrSourceName || "公式图片"}</h2>
                <p>{formulaOcrResult.modelVersion ?? "本地模型"} · 识别结果必须校对后才会写入源码</p>
              </div>
              <button type="button" className="dialog-close" onClick={() => setFormulaOcrResult(null)}>×</button>
            </header>
            <div className="formula-ocr-content">
              {formulaOcrResult.warnings.length > 0 && (
                <div className="formula-ocr-warnings">
                  {formulaOcrResult.warnings.map((warning) => <span key={warning}>{warning}</span>)}
                </div>
              )}
              <div className="formula-candidate-list">
                {formulaOcrResult.candidates.length === 0 ? (
                  <p className="muted">模型没有返回可用候选。可以在下方手动输入 LaTeX 后写入。</p>
                ) : formulaOcrResult.candidates.map((candidate, index) => (
                  <button
                    type="button"
                    className={formulaOcrLatex === candidate.latex ? "active" : ""}
                    key={`${candidate.latex}-${index}`}
                    onClick={() => setFormulaOcrLatex(candidate.latex)}
                  >
                    <span>候选 {index + 1}</span>
                    <code>{candidate.latex}</code>
                    <em>{Math.round(candidate.confidence * 100)}%</em>
                  </button>
                ))}
              </div>
              <div className="formula-correction-editor">
                <label>校对后的 LaTeX</label>
                <MathNodeEditor value={formulaOcrLatex} onChange={setFormulaOcrLatex} />
                <textarea
                  value={formulaOcrLatex}
                  onChange={(event) => setFormulaOcrLatex(event.target.value)}
                  aria-label="公式 LaTeX 源码校对"
                  spellCheck={false}
                />
              </div>
              <div className="formula-insert-target">
                {mathNode?.support === "native" ? (
                  <p>将替换当前选择的 <strong>{mathNode.kind}</strong> 公式节点；原内容仍可通过撤销恢复。</p>
                ) : (
                  <>
                    <p>当前未选择可安全编辑的公式节点，将插入到源码光标位置。</p>
                    <label><input type="radio" checked={formulaInsertMode === "inline"} onChange={() => setFormulaInsertMode("inline")} />行内公式 <code>$...$</code></label>
                    <label><input type="radio" checked={formulaInsertMode === "display"} onChange={() => setFormulaInsertMode("display")} />独立 equation 环境</label>
                  </>
                )}
              </div>
            </div>
            <footer>
              <button type="button" onClick={() => setFormulaOcrResult(null)}>取消</button>
              <button type="button" className="primary" disabled={!formulaOcrLatex.trim()} onClick={() => void insertRecognizedFormula()}>
                写入未保存缓冲
              </button>
            </footer>
          </section>
        </div>
      )}

      {pendingModelInstall && (
        <div className="transaction-backdrop" role="presentation">
          <section className="model-confirm-dialog" role="dialog" aria-modal="true" aria-label="确认安装离线模型">
            <header>
              <div>
                <p className="eyebrow">离线模型校验通过</p>
                <h2>{pendingModelInstall.manifest.id}@{pendingModelInstall.manifest.version}</h2>
                <p>{modelKindLabel(pendingModelInstall.manifest.kind)} · {pendingModelInstall.manifest.backend}</p>
              </div>
              <button type="button" className="dialog-close" onClick={() => setPendingModelInstall(null)}>×</button>
            </header>
            <div className="model-confirm-details">
              <dl>
                <dt>来源目录</dt><dd><code>{pendingModelInstall.sourcePath}</code></dd>
                <dt>模型入口</dt><dd><code>{pendingModelInstall.manifest.entrypoint}</code></dd>
                <dt>文件数量</dt><dd>{pendingModelInstall.computedFiles.length}</dd>
                <dt>总大小</dt><dd>{formatBytes(pendingModelInstall.totalBytes)}</dd>
                <dt>包摘要</dt><dd><code>{pendingModelInstall.packageSha256}</code></dd>
              </dl>
              <p className="model-security-note">
                安装过程会再次复制并校验每个文件；符号链接、路径穿越和校验和不一致都会被拒绝。运行 OCR 时不会自动联网下载模型。
              </p>
            </div>
            <footer>
              <button type="button" onClick={() => setPendingModelInstall(null)}>取消</button>
              <button type="button" className="primary" onClick={() => void installPendingModel()}>
                安装到应用数据目录
              </button>
            </footer>
          </section>
        </div>
      )}

      {pendingModelRemoval && (
        <div className="transaction-backdrop" role="presentation">
          <section className="model-confirm-dialog compact" role="dialog" aria-modal="true" aria-label="确认卸载模型">
            <header>
              <div>
                <p className="eyebrow">卸载本地模型</p>
                <h2>{pendingModelRemoval.manifest.id}@{pendingModelRemoval.manifest.version}</h2>
                <p>将删除 {formatBytes(pendingModelRemoval.totalBytes)} 的本地模型文件，不会修改任何 LaTeX 项目。</p>
              </div>
              <button type="button" className="dialog-close" onClick={() => setPendingModelRemoval(null)}>×</button>
            </header>
            <footer>
              <button type="button" onClick={() => setPendingModelRemoval(null)}>取消</button>
              <button type="button" className="danger" onClick={() => void removePendingModel()}>确认卸载</button>
            </footer>
          </section>
        </div>
      )}

      {replacePlan && (
        <div className="transaction-backdrop" role="presentation">
          <section className="transaction-dialog" role="dialog" aria-modal="true" aria-label="项目修改预览">
            <header>
              <div>
                <p className="eyebrow">项目事务预览</p>
                <h2>{replacePlan.description}</h2>
              </div>
              <button type="button" className="dialog-close" onClick={() => setReplacePlan(null)}>×</button>
            </header>
            <div className="transaction-summary">
              <strong>{replacePlan.totalReplacements}</strong> 处修改，涉及 <strong>{replacePlan.files.length}</strong> 个文件
              {replacePlan.truncated && <span className="transaction-warning">结果超过上限，必须缩小搜索范围后才能应用。</span>}
            </div>
            <div className="transaction-files">
              {replacePlan.files.map((filePlan) => (
                <section key={filePlan.file}>
                  <h3>{filePlan.file}<span>{filePlan.replacements.length}</span></h3>
                  {filePlan.replacements.slice(0, 60).map((replacement, index) => (
                    <button
                      type="button"
                      key={`${replacement.startByte}-${index}`}
                      onClick={() => void revealSourceLocation(filePlan.file, replacement.line, replacement.column)}
                    >
                      <code>{replacement.line}:{replacement.column}</code>
                      <span>{replacement.preview}</span>
                      <small><del>{replacement.expected}</del> → <ins>{replacement.replacement || "∅"}</ins></small>
                    </button>
                  ))}
                  {filePlan.replacements.length > 60 && <p className="muted">该文件另有 {filePlan.replacements.length - 60} 处修改未在列表中展开。</p>}
                </section>
              ))}
            </div>
            <footer>
              <button type="button" onClick={() => setReplacePlan(null)}>取消</button>
              <button
                type="button"
                className="primary"
                disabled={replacePlan.truncated || replacePlan.totalReplacements === 0}
                onClick={() => void applyReplacementPlan()}
              >应用到未保存缓冲</button>
            </footer>
          </section>
        </div>
      )}

      <section className="workspace-grid">
        <aside className="file-panel panel">
          <div className="panel-tabs left-panel-tabs">
            <button className={leftTab === "files" ? "active" : ""} onClick={() => setLeftTab("files")} title="项目文件"><FileCode2 size={14} />文件</button>
            <button className={leftTab === "search" ? "active" : ""} onClick={() => setLeftTab("search")} title="跨文件搜索"><Search size={14} />搜索</button>
            <button className={leftTab === "symbols" ? "active" : ""} onClick={() => setLeftTab("symbols")} title="标签、文献与宏"><Braces size={14} />符号</button>
          </div>
          {leftTab === "files" && (
            <nav className="file-list">
              {files.map((file) => (
                <button className={file === snapshot.path ? "active" : ""} key={file} onClick={() => void openFile(file)}>
                  <ChevronRight size={13} /><span>{file}</span>
                </button>
              ))}
            </nav>
          )}
          {leftTab === "search" && (
            <div className="project-search">
              <label className="project-search-box">
                <Search size={14} />
                <input
                  autoFocus
                  value={searchQuery}
                  placeholder="搜索全部项目文件"
                  onChange={(event) => setSearchQuery(event.target.value)}
                />
              </label>
              <div className="project-search-options">
                <label><input type="checkbox" checked={searchCaseSensitive} onChange={(event) => setSearchCaseSensitive(event.target.checked)} />区分大小写</label>
                <label><input type="checkbox" checked={searchWholeWord} onChange={(event) => setSearchWholeWord(event.target.checked)} />全词匹配</label>
              </div>
              <div className="project-replace-box">
                <input
                  value={replacementText}
                  placeholder="替换为（留空表示删除）"
                  onChange={(event) => setReplacementText(event.target.value)}
                />
                <button type="button" disabled={!searchQuery.trim()} onClick={() => void previewReplacement()}>预览替换</button>
              </div>
              <div className="project-search-summary">
                {searching ? "正在搜索…" : searchQuery.trim() ? `${searchResults.length} 个结果` : "输入文字开始搜索"}
              </div>
              <div className="search-result-list">
                {searchResults.map((result, index) => (
                  <button
                    key={`${result.file}-${result.startByte}-${index}`}
                    onClick={() => void revealSourceLocation(result.file, result.line, result.column)}
                  >
                    <span className="search-result-location">{result.file}:{result.line}:{result.column}</span>
                    <span className="search-result-preview">{result.preview}</span>
                  </button>
                ))}
              </div>
            </div>
          )}
          {leftTab === "symbols" && (
            <div className="symbol-browser">
              <div className="symbol-summary"><Library size={14} />{navigableSymbols.length} 个可导航符号</div>
              {renameTarget && (
                <div className="symbol-rename-form">
                  <span>重命名 <code>{renameTarget.key}</code></span>
                  <input autoFocus value={renameKey} placeholder="新的键名" onChange={(event) => setRenameKey(event.target.value)} />
                  <div>
                    <button type="button" onClick={() => { setRenameTarget(null); setRenameKey(""); }}>取消</button>
                    <button type="button" className="primary" disabled={!renameKey.trim()} onClick={() => void previewRename()}>预览</button>
                  </div>
                </div>
              )}
              <div className="symbol-list">
                {navigableSymbols.map((symbol, index) => (
                  <div className="symbol-row" key={`${symbol.kind}-${symbol.file}-${symbol.startByte}-${index}`}>
                    <button
                      className="symbol-nav"
                      onClick={() => void revealSourceLocation(symbol.file, symbol.line, symbol.column)}
                    >
                      <span className={`symbol-kind ${symbol.kind}`}>
                        {symbol.kind === "label_definition" ? "标签" : symbol.kind === "bibliography_entry" ? "文献" : "宏"}
                      </span>
                      <span className="symbol-key">{symbol.key}</span>
                      <span className="symbol-location">{symbol.file}:{symbol.line}</span>
                    </button>
                    {symbol.kind !== "macro_definition" && (
                      <button
                        type="button"
                        className="symbol-rename-button"
                        title="安全重命名定义与全部引用"
                        onClick={() => { setRenameTarget(symbol); setRenameKey(symbol.key); }}
                      >重命名</button>
                    )}
                  </div>
                ))}
              </div>
            </div>
          )}
        </aside>

        <section className="editor-panel panel">
          <div className="panel-tabs"><span className="active">源码</span><span>rev {snapshot.revision}</span></div>
          <SourceEditor
            value={snapshot.text}
            reveal={sourceReveal}
            completions={sourceCompletions}
            onChange={enqueueSourceChange}
            onUndo={() => void undoRedo("undo")}
            onRedo={() => void undoRedo("redo")}
            onCursor={handleSourceCursor}
          />
        </section>

        <section className="preview-panel panel">
          <div className="panel-tabs">
            <button className={rightTab === "pdf" ? "active" : ""} onClick={() => setRightTab("pdf")}><BookOpen size={15} />PDF 预览</button>
            <button className={rightTab === "structure" ? "active" : ""} onClick={() => setRightTab("structure")}>结构化编辑</button>
          </div>
          {artifact?.pdfPath && artifact.status === "succeeded" ? (
            <PdfViewer
              pdfPath={artifact.pdfPath}
              buildKey={artifact.buildId}
              highlights={pdfHighlights}
              layoutBoxes={layoutMap?.boxes ?? []}
              nodes={snapshot.nodes}
              sourceText={snapshot.text}
              sourcePath={snapshot.path}
              editable={rightTab === "structure"}
              onNodeSelect={handleVisualNodeSelect}
              onNodeCommit={(node, content) => commitVisualEdit(node.id, content)}
              onNodeDelete={(node) => void deleteVisualNode(node)}
              onNodeAttributesCommit={(node, patch) => void commitNodeAttributes(node, patch)}
              onInverseSearch={(result) => void handleInverseSearch(result)}
              onError={(message) => setError(`PDF 页面失败：${message}`)}
            />
          ) : (
            <div className="empty-preview">
              <Hammer size={28} />
              <p>{rightTab === "structure" ? "先编译生成真实页面，再直接点击页面中的标题、正文和公式进行编辑" : "编译后在此显示由 PDFium 渲染的真实 PDF 页面"}</p>
              <button onClick={() => void compile(false)}>立即编译</button>
            </div>
          )}
        </section>

        <section className="bottom-panel panel">
          <div className="panel-tabs">
            <button className={bottomTab === "diagnostics" ? "active" : ""} onClick={() => setBottomTab("diagnostics")}>诊断 {diagnostics.length > 0 && `(${diagnostics.length})`}</button>
            <button className={bottomTab === "log" ? "active" : ""} onClick={() => setBottomTab("log")}>编译日志</button>
            <button className={bottomTab === "toolchain" ? "active" : ""} onClick={() => setBottomTab("toolchain")}>工具链</button>
            <button className={bottomTab === "models" ? "active" : ""} onClick={() => void refreshModelSettings(true)}>本地模型</button>
          </div>
          <div className="bottom-content">
            {bottomTab === "diagnostics" && (
              diagnostics.length ? diagnostics.map((diagnostic, index) => (
                <div className={`diagnostic ${diagnostic.severity}`} key={`${diagnostic.message}-${index}`}>
                  <strong>{diagnostic.severity}</strong>
                  <span>{diagnostic.message}</span>
                  {diagnostic.file && <code>{diagnostic.file}{diagnostic.line ? `:${diagnostic.line}` : ""}</code>}
                </div>
              )) : <p className="muted">当前没有编译诊断。</p>
            )}
            {bottomTab === "log" && <pre>{artifact ? `${artifact.stdout}\n${artifact.stderr}` : "尚未编译。"}</pre>}
            {bottomTab === "toolchain" && (
              tools.length ? <div className="tool-grid">{tools.map((tool) => (
                <div key={tool.name} className={tool.available ? "tool available" : "tool missing"}>
                  <strong>{tool.name}</strong><span>{tool.available ? tool.version ?? tool.path : "未安装"}</span>
                </div>
              ))}</div> : <p className="muted">点击右上角设置按钮检测本机 TeX 工具链。</p>
            )}
            {bottomTab === "models" && (
              <div className="model-settings-panel">
                <header className="model-settings-header">
                  <div>
                    <strong>离线模型管理</strong>
                    <span>模型只从本地校验包安装，识别时不会自动联网下载。</span>
                    {modelCatalog && <code>{modelCatalog.modelsRoot}</code>}
                  </div>
                  <div>
                    <button type="button" onClick={() => void refreshModelSettings(false)}>刷新</button>
                    <button type="button" className="primary" onClick={() => void chooseModelPackage()}><PackagePlus size={14} />安装模型包</button>
                  </div>
                </header>
                <div className={`ocr-health-card ${ocrWorkerHealth?.available ? "available" : "unavailable"}`}>
                  <span className="health-dot" />
                  <div>
                    <strong>公式 OCR 后端：{ocrWorkerHealth?.available ? "可用" : "未配置"}</strong>
                    <span>{ocrWorkerHealth?.backend ?? "paddleocr-formula"}{ocrWorkerHealth?.modelVersion ? ` · ${ocrWorkerHealth.modelVersion}` : ""}</span>
                    {ocrWorkerHealth?.detail && <small>{ocrWorkerHealth.detail}</small>}
                  </div>
                </div>
                <div className="installed-model-list">
                  {(modelCatalog?.installed ?? []).length === 0 ? (
                    <div className="empty-models">
                      <Package size={22} />
                      <p>尚未安装离线模型。请选择包含 <code>visualtex-model.json</code> 的模型包目录。</p>
                    </div>
                  ) : (modelCatalog?.installed ?? []).map((model) => {
                    const active = activeModelKeys.has(
                      `${model.manifest.kind}:${model.manifest.id}:${model.manifest.version}`,
                    );
                    return (
                      <article className={`installed-model-card${active ? " active" : ""}`} key={`${model.manifest.id}-${model.manifest.version}`}>
                        <div className="installed-model-title">
                          <span>{modelKindLabel(model.manifest.kind)}</span>
                          {active && <em>活动</em>}
                        </div>
                        <strong>{model.manifest.id}@{model.manifest.version}</strong>
                        <small>{model.manifest.backend} · {formatBytes(model.totalBytes)} · SHA-256 {model.installedSha256.slice(0, 12)}…</small>
                        {Object.keys(model.manifest.metadata).length > 0 && (
                          <div className="model-metadata">
                            {Object.entries(model.manifest.metadata).map(([key, value]) => <span key={key}>{key}: {value}</span>)}
                          </div>
                        )}
                        <div className="installed-model-actions">
                          <button type="button" disabled={active} onClick={() => void activateModel(model)}>{active ? "当前使用" : "设为活动"}</button>
                          <button type="button" className="danger-ghost" onClick={() => setPendingModelRemoval(model)}><Trash2 size={13} />卸载</button>
                        </div>
                      </article>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        </section>

        <aside className="inspector-panel panel">
          <div className="panel-heading">节点检查器</div>
          {selectedNode ? (
            <div className="inspector-content">
              <dl><dt>类型</dt><dd>{selectedNode.kind}</dd><dt>支持级别</dt><dd>{selectedNode.support}</dd><dt>源码范围</dt><dd>{selectedNode.source.startByte}–{selectedNode.source.endByte}</dd></dl>
              {mathNode ? (
                <div className="inspector-formula-summary">
                  <span>LaTeX 内容</span>
                  <code>{mathNode.text ?? ""}</code>
                  <p className="muted">公式只在“结构化编辑”的浮动工作台中修改，避免两个编辑器同时写入同一节点。</p>
                </div>
              ) : <p className="muted">在“结构化编辑”中点击正文、公式、图片或表格，可打开对应的浮动编辑工作台。</p>}
            </div>
          ) : <p className="muted inspector-empty">选择一个结构节点以查看属性。</p>}
        </aside>
      </section>
    </main>
  );
}
