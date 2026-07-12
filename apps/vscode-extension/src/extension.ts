import * as vscode from "vscode";
import * as path from "node:path";
import { homedir } from "node:os";
import { randomUUID } from "node:crypto";
import { CoreClient, type CoreConnectionStatus } from "./coreClient";
import { visualTextReplacement } from "./editorMapping";
import {
  documentOcrToLatex,
  type DocumentOcrResult,
  type FormulaOcrResult,
} from "./ocrMapping";
import { createWebviewHtml } from "./webview";

interface VisualNode {
  id: string;
  kind: string;
  support: "native" | "partial" | "opaque" | "unstable";
  source: { fileId: string; startByte: number; endByte: number };
  text: string | null;
}

interface Snapshot {
  fileId: string;
  path: string;
  revision: number;
  text: string;
  dirty: boolean;
  nodes: VisualNode[];
}

interface Diagnostic {
  severity: string;
  message: string;
  file: string | null;
  line: number | null;
  column: number | null;
}

interface CompileArtifact {
  buildId: string;
  status: string;
  pdfPath: string | null;
  diagnostics: Diagnostic[];
}

interface PdfPageInfo {
  index: number;
  widthPoints: number;
  heightPoints: number;
  rotationDegrees: number;
}

interface PdfDocumentInfo {
  pdfPath: string;
  fingerprint: string;
  byteLen: number;
  pages: PdfPageInfo[];
}

interface PdfRenderedImage {
  pageIndex: number;
  pageWidthPixels: number;
  pageHeightPixels: number;
  imageWidthPixels: number;
  imageHeightPixels: number;
  cachePath: string;
}

interface PdfRect {
  page: number;
  x: number;
  y: number;
  width: number;
  height: number;
}

interface ForwardSearchResult {
  pdfPath: string;
  boxes: PdfRect[];
}

interface InverseSearchResult {
  sourcePath: string;
  line: number;
  column: number | null;
}

interface WebviewSession {
  reconnect(): Promise<void>;
}

class VisualTexEditorProvider implements vscode.CustomTextEditorProvider {
  public static readonly viewType = "visualtex.visualEditor";
  private readonly sessions = new Set<WebviewSession>();

  constructor(private readonly context: vscode.ExtensionContext) {}

  async reconnectAll(): Promise<void> {
    await Promise.allSettled([...this.sessions].map((session) => session.reconnect()));
  }

  async resolveCustomTextEditor(
    document: vscode.TextDocument,
    webviewPanel: vscode.WebviewPanel,
  ): Promise<void> {
    const workspace = vscode.workspace.getWorkspaceFolder(document.uri);
    if (!workspace) throw new Error("VisualTeX requires the LaTeX file to be inside a workspace folder.");
    if (!vscode.workspace.isTrusted) {
      throw new Error("VisualTeX Core is disabled until the workspace is trusted.");
    }
    if (vscode.env.remoteName) {
      throw new Error(
        `VisualTeX desktop Core does not currently support VS Code Remote (${vscode.env.remoteName}). Open the project in a local VS Code window.`,
      );
    }

    const configuration = vscode.workspace.getConfiguration("visualtex");
    const corePath = configuration.get<string>("corePath", "visualtex");
    const modelsRootSetting = configuration.get<string>("modelsRoot", "").trim();
    const modelsArguments = modelsRootSetting
      ? ["--models-root", resolveConfiguredPath(modelsRootSetting, workspace.uri.fsPath)]
      : [];
    const client = new CoreClient(corePath, workspace.uri.fsPath, modelsArguments);
    const relativePath = path.relative(workspace.uri.fsPath, document.uri.fsPath).split(path.sep).join("/");
    let snapshot: Snapshot;
    let artifact: CompileArtifact | null = null;
    let pdfInfo: PdfDocumentInfo | null = null;
    let currentPage = 0;
    let syncQueue: Promise<void> = Promise.resolve();
    let disposed = false;
    let initialized = false;
    let selectionTimer: NodeJS.Timeout | null = null;
    let sourceSelection = initialSourceSelection(document);

    webviewPanel.webview.options = {
      enableScripts: true,
      localResourceRoots: [workspace.uri],
    };
    webviewPanel.webview.html = createWebviewHtml(webviewPanel.webview);

    const post = async (message: unknown): Promise<boolean> => {
      if (disposed) return false;
      return webviewPanel.webview.postMessage(message);
    };

    const postSnapshot = () => post({ type: "snapshot", snapshot });

    const openCoreSnapshot = () => client.request<Snapshot>(
      "project.openFile",
      { path: relativePath },
      { retryAfterReconnect: true },
    );

    const syncTextIntoCore = async (nextText: string): Promise<void> => {
      if (snapshot.text === nextText) return;
      await client.request("document.applyEdit", {
        operationId: randomUUID(),
        origin: "sourceEditor",
        fileId: snapshot.fileId,
        baseRevision: snapshot.revision,
        startByte: 0,
        endByte: Buffer.byteLength(snapshot.text, "utf8"),
        replacement: nextText,
      });
      snapshot = await openCoreSnapshot();
    };

    const resynchronizeFromDocument = async (): Promise<void> => {
      snapshot = await openCoreSnapshot();
      await syncTextIntoCore(document.getText());
      snapshot = await openCoreSnapshot();
      await postSnapshot();
    };

    const enqueueDocumentSync = (nextText: string): Promise<void> => {
      syncQueue = syncQueue
        .then(async () => {
          if (disposed) return;
          await syncTextIntoCore(nextText);
          await postSnapshot();
        })
        .catch(async (error) => {
          await post({ type: "coreStatus", status: { state: "failed", detail: String(error) } });
          void vscode.window.showErrorMessage(`VisualTeX synchronization failed: ${String(error)}`);
        });
      return syncQueue;
    };

    const confirmVsCodeSave = async (): Promise<void> => {
      await enqueueDocumentSync(document.getText());
      await client.request("document.confirmSaved", { fileId: snapshot.fileId });
      snapshot = await openCoreSnapshot();
      await postSnapshot();
    };

    const renderPdfPage = async (pageIndex: number, requestedWidth: number): Promise<void> => {
      if (!artifact?.pdfPath || !pdfInfo) return;
      const page = Math.max(0, Math.min(pdfInfo.pages.length - 1, pageIndex));
      const targetWidthPixels = Math.max(320, Math.min(3_000, Math.round(requestedWidth)));
      const rendered = await client.request<PdfRenderedImage>("pdf.render", {
        pdfPath: artifact.pdfPath,
        pageIndex: page,
        targetWidthPixels,
        tile: null,
        grayscale: false,
      }, { retryAfterReconnect: true });
      currentPage = page;
      await post({
        type: "pdfPage",
        pageIndex: page,
        imageUri: webviewPanel.webview.asWebviewUri(vscode.Uri.file(rendered.cachePath)).toString(),
        rendered,
      });
    };

    const loadPdf = async (nextArtifact: CompileArtifact): Promise<void> => {
      artifact = nextArtifact;
      pdfInfo = null;
      await post({ type: "compileResult", artifact: nextArtifact, pdfInfo: null });
      if (nextArtifact.status !== "succeeded" || !nextArtifact.pdfPath) return;
      pdfInfo = await client.request<PdfDocumentInfo>(
        "pdf.documentInfo",
        { pdfPath: nextArtifact.pdfPath },
        { retryAfterReconnect: true },
      );
      await post({ type: "compileResult", artifact: nextArtifact, pdfInfo });
      await renderPdfPage(0, 1_300);
    };

    const compileProject = async (): Promise<void> => {
      await document.save();
      await confirmVsCodeSave();
      const nextArtifact = await client.request<CompileArtifact>("project.compile", {}, { timeoutMs: 180_000 });
      await loadPdf(nextArtifact);
    };

    const forwardSearch = async (selection: vscode.Selection): Promise<void> => {
      if (!artifact?.pdfPath || artifact.status !== "succeeded") return;
      const result = await client.request<ForwardSearchResult>("synctex.forwardSearch", {
        sourceFile: relativePath,
        line: selection.active.line + 1,
        column: selection.active.character + 1,
        pdfPath: artifact.pdfPath,
      }, { retryAfterReconnect: true });
      await post({ type: "pdfHighlights", boxes: result.boxes });
      const highlightedPage = result.boxes[0]?.page;
      if (highlightedPage && highlightedPage - 1 !== currentPage) {
        await renderPdfPage(highlightedPage - 1, 1_300);
      }
    };

    const revealInverseSearch = async (page: number, x: number, y: number): Promise<void> => {
      if (!artifact?.pdfPath) return;
      const result = await client.request<InverseSearchResult>("synctex.inverseSearch", {
        pdfPath: artifact.pdfPath,
        page,
        x,
        y,
      }, { retryAfterReconnect: true });
      const sourceUri = path.isAbsolute(result.sourcePath)
        ? vscode.Uri.file(result.sourcePath)
        : vscode.Uri.joinPath(workspace.uri, ...result.sourcePath.split(/[\\/]/));
      const sourceDocument = await vscode.workspace.openTextDocument(sourceUri);
      const editor = await vscode.window.showTextDocument(sourceDocument, {
        preview: false,
        viewColumn: vscode.ViewColumn.One,
      });
      const line = Math.max(0, Math.min(sourceDocument.lineCount - 1, result.line - 1));
      const column = Math.max(0, (result.column ?? 1) - 1);
      const position = new vscode.Position(line, Math.min(column, sourceDocument.lineAt(line).text.length));
      editor.selection = new vscode.Selection(position, position);
      editor.revealRange(new vscode.Range(position, position), vscode.TextEditorRevealType.InCenter);
    };

    const selectWorkspaceImage = async (title: string): Promise<{
      relativePath: string;
      imageUri: string;
    } | null> => {
      const selected = await vscode.window.showOpenDialog({
        title,
        canSelectMany: false,
        canSelectFiles: true,
        canSelectFolders: false,
        defaultUri: workspace.uri,
        filters: {
          Images: ["png", "jpg", "jpeg", "webp", "bmp", "tif", "tiff"],
        },
      });
      const uri = selected?.[0];
      if (!uri || uri.scheme !== "file") return null;
      const relative = path.relative(workspace.uri.fsPath, uri.fsPath);
      if (relative.startsWith("..") || path.isAbsolute(relative)) {
        throw new Error("OCR images must be inside the current trusted workspace.");
      }
      return {
        relativePath: relative.split(path.sep).join("/"),
        imageUri: webviewPanel.webview.asWebviewUri(uri).toString(),
      };
    };

    const insertAtSourceSelection = async (text: string): Promise<void> => {
      const startOffset = document.offsetAt(sourceSelection.start);
      const edit = new vscode.WorkspaceEdit();
      edit.replace(document.uri, sourceSelection, text);
      if (!(await vscode.workspace.applyEdit(edit))) {
        throw new Error("VS Code rejected the OCR WorkspaceEdit");
      }
      const nextPosition = document.positionAt(startOffset + text.length);
      sourceSelection = new vscode.Selection(nextPosition, nextPosition);
    };

    const runFormulaOcr = async (): Promise<void> => {
      const selected = await selectWorkspaceImage("选择工作区内的公式图片");
      if (!selected) return;
      await post({ type: "ocrStatus", message: "正在离线识别公式…" });
      const result = await client.request<FormulaOcrResult>(
        "ocr.recognizeFormula",
        { sourcePath: selected.relativePath },
        { timeoutMs: 180_000 },
      );
      await post({
        type: "formulaOcrResult",
        result,
        imageUri: selected.imageUri,
      });
    };

    const runDocumentOcr = async (): Promise<void> => {
      const selected = await selectWorkspaceImage("选择工作区内的整页扫描图片");
      if (!selected) return;
      await post({ type: "ocrStatus", message: "正在离线分析版面与阅读顺序…" });
      const result = await client.request<DocumentOcrResult>(
        "ocr.recognizeDocument",
        { sourcePath: selected.relativePath },
        { timeoutMs: 300_000 },
      );
      const normalizedUri = result.imagePath
        ? webviewPanel.webview.asWebviewUri(vscode.Uri.file(result.imagePath)).toString()
        : selected.imageUri;
      await post({
        type: "documentOcrResult",
        result,
        imageUri: normalizedUri,
      });
    };

    const refreshOcrHealth = async (): Promise<void> => {
      try {
        const health = await client.request("ocr.health", {}, {
          retryAfterReconnect: true,
          timeoutMs: 30_000,
        });
        await post({ type: "ocrHealth", health });
      } catch (error) {
        await post({ type: "ocrHealth", health: null, error: String(error) });
      }
    };

    const reconnectSession = async (): Promise<void> => {
      await client.restart();
      await resynchronizeFromDocument();
      if (artifact) await loadPdf(artifact);
      await refreshOcrHealth();
    };
    const session: WebviewSession = { reconnect: reconnectSession };
    this.sessions.add(session);

    const statusSubscription = client.onStatus((status: CoreConnectionStatus) => {
      void post({ type: "coreStatus", status });
      if (status.state === "connected" && initialized && !disposed) {
        syncQueue = syncQueue
          .then(resynchronizeFromDocument)
          .catch((error) => void vscode.window.showErrorMessage(`VisualTeX reconnect failed: ${String(error)}`));
      }
    });

    await client.connect();
    snapshot = await openCoreSnapshot();
    await syncTextIntoCore(document.getText());
    snapshot = await openCoreSnapshot();
    initialized = true;
    await postSnapshot();

    const documentListener = vscode.workspace.onDidChangeTextDocument((event) => {
      if (event.document.uri.toString() === document.uri.toString()) {
        void enqueueDocumentSync(event.document.getText());
      }
    });

    const saveListener = vscode.workspace.onDidSaveTextDocument((saved) => {
      if (saved.uri.toString() !== document.uri.toString()) return;
      syncQueue = syncQueue
        .then(confirmVsCodeSave)
        .catch((error) => void vscode.window.showErrorMessage(`VisualTeX save confirmation failed: ${String(error)}`));
    });

    const selectionListener = vscode.window.onDidChangeTextEditorSelection((event) => {
      if (event.textEditor.document.uri.toString() !== document.uri.toString()) return;
      sourceSelection = event.selections[0] ?? event.textEditor.selection;
      if (selectionTimer) clearTimeout(selectionTimer);
      selectionTimer = setTimeout(() => {
        selectionTimer = null;
        void forwardSearch(sourceSelection).catch(() => undefined);
      }, 180);
    });

    const messageListener = webviewPanel.webview.onDidReceiveMessage(async (message: unknown) => {
      if (!message || typeof message !== "object") return;
      const payload = message as Record<string, unknown>;
      try {
        if (payload.type === "ready") {
          await postSnapshot();
          await post({ type: "coreStatus", status: { state: "connected" } });
          if (artifact) await loadPdf(artifact);
          await refreshOcrHealth();
        } else if (payload.type === "visualEdit") {
          const node = snapshot.nodes.find((candidate) => candidate.id === payload.nodeId);
          if (!node || typeof payload.content !== "string") return;
          const replacement = visualReplacement(document, node, payload.content);
          if (!replacement) {
            void vscode.window.showWarningMessage("This LaTeX node is only editable in source mode.");
            return;
          }
          const edit = new vscode.WorkspaceEdit();
          edit.replace(document.uri, replacement.range, replacement.text);
          if (!(await vscode.workspace.applyEdit(edit))) throw new Error("VS Code rejected the WorkspaceEdit");
        } else if (payload.type === "compile") {
          await compileProject();
        } else if (payload.type === "renderPdfPage") {
          const pageIndex = typeof payload.pageIndex === "number" ? payload.pageIndex : 0;
          const width = typeof payload.width === "number" ? payload.width : 1_300;
          await renderPdfPage(pageIndex, width);
        } else if (payload.type === "inverseSearch") {
          if (
            typeof payload.page === "number"
            && typeof payload.x === "number"
            && typeof payload.y === "number"
          ) {
            await revealInverseSearch(payload.page, payload.x, payload.y);
          }
        } else if (payload.type === "chooseFormulaOcr") {
          await runFormulaOcr();
        } else if (payload.type === "chooseDocumentOcr") {
          await runDocumentOcr();
        } else if (payload.type === "refreshOcrHealth") {
          await refreshOcrHealth();
        } else if (payload.type === "insertFormulaOcr") {
          if (typeof payload.latex !== "string" || payload.latex.length > 100_000) {
            throw new Error("Invalid formula OCR LaTeX payload");
          }
          const mode = payload.mode === "display" ? "display" : payload.mode === "raw" ? "raw" : "inline";
          const latex = payload.latex.trim();
          if (!latex) return;
          await insertAtSourceSelection(
            mode === "display" ? `\\[\n${latex}\n\\]` : mode === "raw" ? latex : `$${latex}$`,
          );
        } else if (payload.type === "insertDocumentOcr") {
          const reviewed = parseDocumentOcrResult(payload.result);
          const latex = documentOcrToLatex(reviewed).trim();
          if (!latex) throw new Error("Reviewed document OCR result contains no insertable content");
          await insertAtSourceSelection(latex);
        } else if (payload.type === "reconnect") {
          await reconnectSession();
        }
      } catch (error) {
        await post({ type: "operationError", message: String(error) });
        void vscode.window.showErrorMessage(`VisualTeX: ${String(error)}`);
      }
    });

    webviewPanel.onDidDispose(() => {
      disposed = true;
      if (selectionTimer) clearTimeout(selectionTimer);
      this.sessions.delete(session);
      documentListener.dispose();
      saveListener.dispose();
      selectionListener.dispose();
      messageListener.dispose();
      statusSubscription.dispose();
      client.dispose();
    });
  }
}

function resolveConfiguredPath(value: string, workspaceRoot: string): string {
  if (value === "~") return homedir();
  if (value.startsWith(`~${path.sep}`) || value.startsWith("~/")) {
    return path.join(homedir(), value.slice(2));
  }
  return path.isAbsolute(value) ? value : path.resolve(workspaceRoot, value);
}

function initialSourceSelection(document: vscode.TextDocument): vscode.Selection {
  const active = vscode.window.activeTextEditor;
  if (active?.document.uri.toString() === document.uri.toString()) {
    return active.selection;
  }
  const end = document.positionAt(document.getText().length);
  return new vscode.Selection(end, end);
}

function parseDocumentOcrResult(value: unknown): DocumentOcrResult {
  if (!value || typeof value !== "object") throw new Error("Invalid document OCR payload");
  const record = value as Record<string, unknown>;
  const pageWidth = finiteNumber(record.pageWidth, "pageWidth");
  const pageHeight = finiteNumber(record.pageHeight, "pageHeight");
  if (pageWidth <= 0 || pageHeight <= 0 || pageWidth > 100_000 || pageHeight > 100_000) {
    throw new Error("Invalid document OCR page dimensions");
  }
  if (!Array.isArray(record.regions) || record.regions.length > 5_000) {
    throw new Error("Invalid or excessive document OCR regions");
  }
  const regions = record.regions.map((item, index) => {
    if (!item || typeof item !== "object") throw new Error(`Invalid OCR region ${index}`);
    const region = item as Record<string, unknown>;
    const kind = boundedString(region.kind, `regions[${index}].kind`, 128);
    const text = nullableBoundedString(region.text, `regions[${index}].text`, 2_000_000);
    const latex = nullableBoundedString(region.latex, `regions[${index}].latex`, 2_000_000);
    const confidence = finiteNumber(region.confidence, `regions[${index}].confidence`);
    return {
      kind,
      x: finiteNumber(region.x, `regions[${index}].x`),
      y: finiteNumber(region.y, `regions[${index}].y`),
      width: finiteNumber(region.width, `regions[${index}].width`),
      height: finiteNumber(region.height, `regions[${index}].height`),
      text,
      latex,
      confidence: Math.max(0, Math.min(1, confidence)),
    };
  });
  if (!Array.isArray(record.readingOrder) || record.readingOrder.length > regions.length * 2 + 32) {
    throw new Error("Invalid document OCR reading order");
  }
  const readingOrder = record.readingOrder.map((item, index) => {
    if (!Number.isInteger(item) || Number(item) < 0 || Number(item) >= regions.length) {
      throw new Error(`Invalid reading-order index at ${index}`);
    }
    return Number(item);
  });
  const warnings = Array.isArray(record.warnings)
    ? record.warnings.slice(0, 200).map((warning, index) =>
      boundedString(warning, `warnings[${index}]`, 10_000))
    : [];
  return {
    imagePath: nullableBoundedString(record.imagePath, "imagePath", 32_768),
    pageWidth,
    pageHeight,
    regions,
    readingOrder,
    modelVersion: nullableBoundedString(record.modelVersion, "modelVersion", 1_024),
    warnings,
  };
}

function finiteNumber(value: unknown, name: string): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`Invalid ${name}`);
  }
  return value;
}

function boundedString(value: unknown, name: string, maximumLength: number): string {
  if (typeof value !== "string" || value.length > maximumLength) {
    throw new Error(`Invalid ${name}`);
  }
  return value;
}

function nullableBoundedString(
  value: unknown,
  name: string,
  maximumLength: number,
): string | null {
  if (value === null || value === undefined) return null;
  return boundedString(value, name, maximumLength);
}

export function visualReplacement(
  document: vscode.TextDocument,
  node: VisualNode,
  content: string,
): { range: vscode.Range; text: string } | null {
  const replacement = visualTextReplacement(document.getText(), node, content);
  if (!replacement) return null;
  return {
    range: new vscode.Range(
      document.positionAt(replacement.startUtf16),
      document.positionAt(replacement.endUtf16),
    ),
    text: replacement.text,
  };
}

function webviewHtml(webview: vscode.Webview): string {
  const nonce = randomUUID().replaceAll("-", "");
  return `<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${webview.cspSource} data:; style-src 'unsafe-inline'; script-src 'nonce-${nonce}';">
<style>
:root{color-scheme:light dark}*{box-sizing:border-box}body{margin:0;font-family:var(--vscode-font-family);color:var(--vscode-foreground);background:var(--vscode-editor-background)}header.app{position:sticky;top:0;z-index:5;display:flex;align-items:center;gap:8px;padding:8px 10px;border-bottom:1px solid var(--vscode-panel-border);background:var(--vscode-editor-background)}button,select{font:inherit}button{color:var(--vscode-button-foreground);background:var(--vscode-button-background);border:0;border-radius:3px;padding:5px 9px;cursor:pointer}button.secondary{color:var(--vscode-foreground);background:var(--vscode-button-secondaryBackground)}button:disabled{opacity:.45}.status{margin-left:auto;color:var(--vscode-descriptionForeground);font-size:11px}.tabs{display:flex;gap:3px}.tabs button.active{outline:2px solid var(--vscode-focusBorder)}main{min-height:calc(100vh - 45px)}.panel{display:none}.panel.active{display:block}.editor-panel{padding:0 14px 32px}.node{border:1px solid var(--vscode-panel-border);border-radius:6px;margin:10px 0;overflow:hidden}.label{font-size:10px;text-transform:uppercase;color:var(--vscode-descriptionForeground);background:var(--vscode-sideBar-background);padding:5px 8px}.node textarea{display:block;width:100%;min-height:56px;resize:vertical;padding:9px;border:0;outline:0;color:var(--vscode-input-foreground);background:var(--vscode-input-background);font:inherit;line-height:1.55}.node[data-kind=title] textarea{font-size:1.5em;font-weight:700;text-align:center}.node[data-kind=section] textarea{font-size:1.2em;font-weight:700}.readonly{opacity:.7}.diagnostics{padding:4px 14px 12px}.diagnostic{margin:5px 0;color:var(--vscode-errorForeground);font-size:11px}.pdf-panel{height:calc(100vh - 46px);grid-template-rows:auto minmax(0,1fr)}.pdf-panel.active{display:grid}.pdf-toolbar{display:flex;align-items:center;justify-content:center;gap:6px;padding:6px;border-bottom:1px solid var(--vscode-panel-border)}.pdf-scroll{min-height:0;overflow:auto;padding:20px;background:var(--vscode-sideBar-background)}.pdf-page{position:relative;width:min(100%,1000px);margin:0 auto;background:white;box-shadow:0 4px 18px rgba(0,0,0,.25)}.pdf-page img{display:block;width:100%;height:auto;user-select:none}.highlight-layer{position:absolute;inset:0;pointer-events:none}.highlight{position:absolute;border:1px solid #f0a000;background:rgba(255,203,57,.32)}.empty{padding:40px;text-align:center;color:var(--vscode-descriptionForeground)}.error-line{padding:6px 10px;color:var(--vscode-errorForeground);border-bottom:1px solid var(--vscode-panel-border);font-size:11px}
</style>
</head>
<body>
<header class="app"><strong>VisualTeX Next</strong><div class="tabs"><button id="editorTab" class="secondary active">结构</button><button id="pdfTab" class="secondary">PDF</button></div><button id="compile">编译</button><button id="reconnect" class="secondary">重连 Core</button><span class="status" id="status"></span></header>
<div id="error"></div>
<main><section id="editorPanel" class="panel editor-panel active"><div id="nodes"></div><div id="diagnostics" class="diagnostics"></div></section><section id="pdfPanel" class="panel pdf-panel"><div class="pdf-toolbar"><button id="previous" class="secondary">上一页</button><span id="pageLabel">0 / 0</span><button id="next" class="secondary">下一页</button><select id="zoom"><option value="0.8">80%</option><option value="1" selected>100%</option><option value="1.3">130%</option><option value="1.7">170%</option></select></div><div class="pdf-scroll" id="pdfScroll"><div id="pdfEmpty" class="empty">编译成功后显示 PDFium 渲染页面。</div><div id="pdfPage" class="pdf-page" hidden><img id="pdfImage" alt="Rendered PDF page"><div id="highlightLayer" class="highlight-layer"></div></div></div></section></main>
<script nonce="${nonce}">
const vscode=acquireVsCodeApi();const nodes=document.getElementById('nodes');const status=document.getElementById('status');const diagnostics=document.getElementById('diagnostics');const errorBox=document.getElementById('error');const editorPanel=document.getElementById('editorPanel');const pdfPanel=document.getElementById('pdfPanel');const editorTab=document.getElementById('editorTab');const pdfTab=document.getElementById('pdfTab');const pdfPage=document.getElementById('pdfPage');const pdfImage=document.getElementById('pdfImage');const pdfEmpty=document.getElementById('pdfEmpty');const pageLabel=document.getElementById('pageLabel');const previous=document.getElementById('previous');const next=document.getElementById('next');const zoom=document.getElementById('zoom');const highlightLayer=document.getElementById('highlightLayer');const pdfScroll=document.getElementById('pdfScroll');let pdfInfo=null;let currentPage=0;let highlights=[];
function setTab(tab){const pdf=tab==='pdf';editorPanel.classList.toggle('active',!pdf);pdfPanel.classList.toggle('active',pdf);editorTab.classList.toggle('active',!pdf);pdfTab.classList.toggle('active',pdf);if(pdf&&pdfInfo)requestPage(currentPage)}
function render(snapshot){status.textContent=snapshot.path+' · rev '+snapshot.revision+(snapshot.dirty?' · dirty':'');nodes.replaceChildren();for(const node of snapshot.nodes){if(node.kind==='document'||node.text===null)continue;const box=document.createElement('section');box.className='node'+(node.support==='opaque'||node.support==='unstable'?' readonly':'');box.dataset.kind=node.kind;const label=document.createElement('div');label.className='label';label.textContent=node.kind+' · '+node.support;const editor=document.createElement('textarea');editor.value=node.text;editor.disabled=node.support==='opaque'||node.support==='unstable';editor.addEventListener('change',()=>vscode.postMessage({type:'visualEdit',nodeId:node.id,content:editor.value}));box.append(label,editor);nodes.append(box)}}
function renderDiagnostics(items){diagnostics.replaceChildren();for(const item of items||[]){const row=document.createElement('div');row.className='diagnostic';row.textContent=(item.file?item.file+':':'')+(item.line||'')+' '+item.message;diagnostics.append(row)}}
function renderPdfState(){const pages=pdfInfo?.pages?.length||0;pageLabel.textContent=pages?String(currentPage+1)+' / '+String(pages):'0 / 0';previous.disabled=currentPage<=0;next.disabled=!pages||currentPage>=pages-1;pdfEmpty.hidden=pages>0;pdfPage.hidden=pages===0}
function requestPage(page){if(!pdfInfo)return;currentPage=Math.max(0,Math.min(pdfInfo.pages.length-1,page));renderPdfState();const width=Math.max(640,Math.min(3000,Math.round(pdfScroll.clientWidth*devicePixelRatio*Number(zoom.value))));vscode.postMessage({type:'renderPdfPage',pageIndex:currentPage,width})}
function renderHighlights(){highlightLayer.replaceChildren();if(!pdfInfo)return;const page=pdfInfo.pages[currentPage];if(!page)return;for(const box of highlights){if(box.page!==currentPage+1)continue;const item=document.createElement('span');item.className='highlight';item.style.left=(box.x/page.widthPoints*100)+'%';item.style.top=(box.y/page.heightPoints*100)+'%';item.style.width=(box.width/page.widthPoints*100)+'%';item.style.height=(box.height/page.heightPoints*100)+'%';highlightLayer.append(item)}}
window.addEventListener('message',event=>{const message=event.data;if(message.type==='snapshot')render(message.snapshot);else if(message.type==='coreStatus'){const state=message.status.state;status.textContent=state==='connected'?'Core 已连接':state==='connecting'?'Core 连接中…':'Core '+state+(message.status.detail?' · '+message.status.detail:'')}else if(message.type==='compileResult'){renderDiagnostics(message.artifact.diagnostics);pdfInfo=message.pdfInfo;currentPage=0;renderPdfState();if(pdfInfo)setTab('pdf')}else if(message.type==='pdfPage'){currentPage=message.pageIndex;pdfImage.src=message.imageUri;renderPdfState();renderHighlights()}else if(message.type==='pdfHighlights'){highlights=message.boxes||[];renderHighlights()}else if(message.type==='operationError'){errorBox.className='error-line';errorBox.textContent=message.message}});
pdfImage.addEventListener('dblclick',event=>{if(!pdfInfo)return;const page=pdfInfo.pages[currentPage];const bounds=pdfImage.getBoundingClientRect();vscode.postMessage({type:'inverseSearch',page:currentPage+1,x:(event.clientX-bounds.left)/bounds.width*page.widthPoints,y:(event.clientY-bounds.top)/bounds.height*page.heightPoints})});
editorTab.addEventListener('click',()=>setTab('editor'));pdfTab.addEventListener('click',()=>setTab('pdf'));document.getElementById('compile').addEventListener('click',()=>vscode.postMessage({type:'compile'}));document.getElementById('reconnect').addEventListener('click',()=>vscode.postMessage({type:'reconnect'}));previous.addEventListener('click',()=>requestPage(currentPage-1));next.addEventListener('click',()=>requestPage(currentPage+1));zoom.addEventListener('change',()=>requestPage(currentPage));let resizeTimer;window.addEventListener('resize',()=>{clearTimeout(resizeTimer);resizeTimer=setTimeout(()=>{if(pdfPanel.classList.contains('active')&&pdfInfo)requestPage(currentPage)},180)});vscode.postMessage({type:'ready'});
</script></body></html>`;
}

export function activate(context: vscode.ExtensionContext): void {
  const provider = new VisualTexEditorProvider(context);
  context.subscriptions.push(
    vscode.window.registerCustomEditorProvider(VisualTexEditorProvider.viewType, provider, {
      webviewOptions: { retainContextWhenHidden: true },
      supportsMultipleEditorsPerDocument: true,
    }),
    vscode.commands.registerCommand("visualtex.openVisualEditor", async () => {
      const editor = vscode.window.activeTextEditor;
      if (!editor || editor.document.languageId !== "latex") {
        void vscode.window.showInformationMessage("Open a .tex document first.");
        return;
      }
      await vscode.commands.executeCommand("vscode.openWith", editor.document.uri, VisualTexEditorProvider.viewType);
    }),
    vscode.commands.registerCommand("visualtex.compile", async () => {
      await vscode.commands.executeCommand("visualtex.openVisualEditor");
    }),
    vscode.commands.registerCommand("visualtex.reconnectCore", async () => {
      await provider.reconnectAll();
    }),
  );
}

export function deactivate(): void {}
