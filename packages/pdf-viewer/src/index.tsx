import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type KeyboardEvent,
  type MouseEvent,
  type PointerEvent,
  type RefObject,
} from "react";
import { createPortal } from "react-dom";
import {
  Check,
  ChevronLeft,
  ChevronRight,
  LoaderCircle,
  Minus,
  Plus,
  RotateCcw,
  X,
} from "lucide-react";
import { MathNodeEditor } from "@visualtex/math-editor";
import {
  desktopApi,
  type InverseSearchResult,
  type LayoutBox,
  type NodeAttributes,
  type NodeAttributesPatch,
  type PdfDocumentInfo,
  type PdfPageInfo,
  type PdfRect,
  type PdfRenderedImage,
  type PdfTextGlyph,
  type PdfTextHit,
  type VisualNode,
} from "@visualtex/protocol";
import "./styles.css";

export interface PdfViewerProps {
  pdfPath: string;
  buildKey?: string;
  highlights?: PdfRect[];
  layoutBoxes?: LayoutBox[];
  nodes?: VisualNode[];
  sourceText?: string;
  sourcePath?: string;
  editable?: boolean;
  onInverseSearch?: (result: InverseSearchResult) => void;
  onNodeSelect?: (node: VisualNode) => void;
  onNodeCommit?: (node: VisualNode, content: string) => void;
  onNodeDelete?: (node: VisualNode) => void;
  onNodeAttributesCommit?: (node: VisualNode, patch: NodeAttributesPatch) => void;
  onError?: (message: string) => void;
}

export interface ViewportAnchor {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export interface DirectEditState {
  pageIndex: number;
  node: VisualNode;
  layout: LayoutBox;
  rect: PdfRect;
  hitRect: PdfRect;
  viewportAnchor?: ViewportAnchor;
  draft: string;
  attributes: NodeAttributes;
}

interface RenderedPageProps {
  info: PdfDocumentInfo;
  page: PdfPageInfo;
  cssWidth: number;
  grayscale: boolean;
  rootRef: RefObject<HTMLDivElement | null>;
  highlights: PdfRect[];
  layoutBoxes: LayoutBox[];
  nodes: VisualNode[];
  sourceText?: string;
  sourcePath?: string;
  editable: boolean;
  directEdit: DirectEditState | null;
  onDirectEdit: (edit: DirectEditState | null) => void;
  onNodeSelect?: (node: VisualNode) => void;
  onNodeCommit?: (node: VisualNode, content: string) => void;
  onNodeDelete?: (node: VisualNode) => void;
  onNodeAttributesCommit?: (node: VisualNode, patch: NodeAttributesPatch) => void;
  onVisible: (pageIndex: number) => void;
  onInverseSearch?: (result: InverseSearchResult) => void;
  onError?: (message: string) => void;
}

interface ThumbnailProps {
  info: PdfDocumentInfo;
  page: PdfPageInfo;
  rootRef: RefObject<HTMLDivElement | null>;
  active: boolean;
  onSelect: () => void;
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function useObservedVisibility(
  targetRef: RefObject<Element | null>,
  rootRef: RefObject<Element | null>,
  rootMargin: string,
  onVisible?: () => void,
): boolean {
  const [visible, setVisible] = useState(false);
  const onVisibleRef = useRef(onVisible);
  onVisibleRef.current = onVisible;

  useEffect(() => {
    const target = targetRef.current;
    if (!target) return;
    if (!("IntersectionObserver" in window)) {
      setVisible(true);
      onVisibleRef.current?.();
      return;
    }
    const observer = new IntersectionObserver(
      (entries) => {
        const entry = entries[0];
        if (!entry) return;
        setVisible(entry.isIntersecting);
        if (entry.isIntersecting) onVisibleRef.current?.();
      },
      {
        root: rootRef.current,
        rootMargin,
        threshold: [0, 0.15, 0.5],
      },
    );
    observer.observe(target);
    return () => observer.disconnect();
  }, [rootRef, rootMargin, targetRef]);

  return visible;
}

function useRenderedImage(
  info: PdfDocumentInfo,
  pageIndex: number,
  targetWidthPixels: number,
  grayscale: boolean,
  enabled: boolean,
  onError?: (message: string) => void,
): PdfRenderedImage | null {
  const [rendered, setRendered] = useState<PdfRenderedImage | null>(null);
  const requestKey = `${info.fingerprint}:${pageIndex}:${targetWidthPixels}:${grayscale}`;

  useEffect(() => {
    if (!enabled) return;
    let cancelled = false;
    void desktopApi
      .renderPdf({
        pdfPath: info.pdfPath,
        pageIndex,
        targetWidthPixels,
        tile: null,
        grayscale,
      })
      .then((next) => {
        if (!cancelled) setRendered(next);
      })
      .catch((error: unknown) => {
        if (!cancelled) onError?.(messageOf(error));
      });
    return () => {
      cancelled = true;
    };
  }, [requestKey, enabled, info.pdfPath, onError]);

  useEffect(() => {
    if (enabled || !rendered) return;
    const timer = window.setTimeout(() => setRendered(null), 8_000);
    return () => window.clearTimeout(timer);
  }, [enabled, rendered]);

  return rendered;
}

export function isDirectlyEditable(node: VisualNode, layout: LayoutBox): boolean {
  if (layout.confidence !== "exact" && layout.confidence !== "high") return false;
  if (node.support === "native") {
    return [
      "section",
      "subsection",
      "paragraph",
      "inline_math",
      "display_math",
    ].includes(node.kind);
  }
  return node.support === "partial"
    && ["paragraph", "figure", "table"].includes(node.kind);
}

function containsPoint(rect: PdfRect, x: number, y: number, padding = 1.25): boolean {
  if (rect.width <= 0 || rect.height <= 0) return false;
  return (
    x >= rect.x - padding &&
    x <= rect.x + rect.width + padding &&
    y >= rect.y - padding &&
    y <= rect.y + rect.height + padding
  );
}

function rectArea(rect: PdfRect): number {
  return Math.max(0, rect.width) * Math.max(0, rect.height);
}

function unionRects(rects: PdfRect[]): PdfRect | null {
  const usable = rects.filter((rect) => rect.width > 0 && rect.height > 0);
  if (usable.length === 0) return null;
  const page = usable[0]!.page;
  const left = Math.min(...usable.map((rect) => rect.x));
  const top = Math.min(...usable.map((rect) => rect.y));
  const right = Math.max(...usable.map((rect) => rect.x + rect.width));
  const bottom = Math.max(...usable.map((rect) => rect.y + rect.height));
  return { page, x: left, y: top, width: right - left, height: bottom - top };
}

function normalizePath(value: string): string {
  return value.replaceAll("\\", "/").replace(/\/\.\//g, "/").toLocaleLowerCase();
}

function utf8Length(value: string): number {
  return new TextEncoder().encode(value).length;
}

export function sourceLineByteRange(source: string, line: number): { start: number; end: number } | null {
  if (!Number.isFinite(line) || line < 1) return null;
  let currentLine = 1;
  let characterStart = 0;
  for (let index = 0; index < source.length && currentLine < line; index += 1) {
    if (source[index] === "\n") {
      currentLine += 1;
      characterStart = index + 1;
    }
  }
  if (currentLine !== line) return null;
  const newline = source.indexOf("\n", characterStart);
  const characterEnd = newline < 0 ? source.length : newline;
  return {
    start: utf8Length(source.slice(0, characterStart)),
    end: utf8Length(source.slice(0, characterEnd)),
  };
}

export function sourceByteOffsetAtLineColumn(
  source: string,
  line: number,
  column?: number | null,
): number | null {
  const range = sourceLineByteRange(source, line);
  if (!range) return null;
  if (column === null || column === undefined || column <= 1) return range.start;
  const lines = source.replace(/\r\n?/g, "\n").split("\n");
  const lineText = lines[line - 1];
  if (lineText === undefined) return null;
  const prefix = Array.from(lineText).slice(0, Math.max(0, column - 1)).join("");
  return range.start + utf8Length(prefix);
}

function stateForNode(
  pageIndex: number,
  node: VisualNode,
  layout: LayoutBox,
  x: number,
  y: number,
): DirectEditState | null {
  const pageRects = layout.rects.filter((rect) => rect.page === pageIndex + 1);
  const rect = unionRects(pageRects);
  if (!rect) return null;
  const hitRect = pageRects
    .filter((candidate) => containsPoint(candidate, x, y))
    .sort((left, right) => rectArea(left) - rectArea(right))[0]
    ?? pageRects
      .slice()
      .sort((left, right) => {
        const leftDistance = Math.hypot(left.x + left.width / 2 - x, left.y + left.height / 2 - y);
        const rightDistance = Math.hypot(right.x + right.width / 2 - x, right.y + right.height / 2 - y);
        return leftDistance - rightDistance;
      })[0];
  if (!hitRect) return null;
  return {
    pageIndex,
    node,
    layout,
    rect,
    hitRect,
    draft: node.text ?? "",
    attributes: structuredClone(node.attributes),
  };
}

export function findDirectEdit(
  pageIndex: number,
  x: number,
  y: number,
  layoutBoxes: LayoutBox[],
  nodes: VisualNode[],
): DirectEditState | null {
  const nodesById = new Map(nodes.map((node) => [node.id, node]));
  const candidates: DirectEditState[] = [];
  for (const layout of layoutBoxes) {
    const node = nodesById.get(layout.nodeId);
    if (!node || !isDirectlyEditable(node, layout)) continue;
    const pageRects = layout.rects.filter((rect) => rect.page === pageIndex + 1);
    const hitRect = pageRects
      .filter((rect) => containsPoint(rect, x, y))
      .sort((left, right) => rectArea(left) - rectArea(right))[0];
    if (!hitRect) continue;
    const state = stateForNode(pageIndex, node, layout, x, y);
    if (state) candidates.push({ ...state, hitRect });
  }
  candidates.sort((left, right) => {
    const hitArea = rectArea(left.hitRect) - rectArea(right.hitRect);
    if (Math.abs(hitArea) > 0.01) return hitArea;
    const leftSpan = left.node.source.endByte - left.node.source.startByte;
    const rightSpan = right.node.source.endByte - right.node.source.startByte;
    return leftSpan - rightSpan;
  });
  return candidates[0] ?? null;
}

const latexSymbolMap: Record<string, string> = {
  alpha: "α",
  beta: "β",
  gamma: "γ",
  delta: "δ",
  theta: "θ",
  lambda: "λ",
  mu: "μ",
  pi: "π",
  sigma: "σ",
  phi: "φ",
  psi: "ψ",
  omega: "ω",
  Delta: "Δ",
  Sigma: "Σ",
  Omega: "Ω",
  cdot: "*",
  times: "*",
  leq: "≤",
  geq: "≥",
  neq: "≠",
  approx: "≈",
  infty: "∞",
  partial: "∂",
  nabla: "∇",
  to: "→",
  rightarrow: "→",
  leftarrow: "←",
  pm: "±",
};

function normalizeVisibleCharacter(value: string): string {
  return value
    .normalize("NFKC")
    .replace(/[−–—]/g, "-")
    .replace(/[⋅·×]/g, "*")
    .replace(/[‐‑]/g, "-")
    .replace(/\s+/g, "")
    .replace(/[{}_$^]/g, "")
    .toLocaleLowerCase();
}

export function normalizeLatexForPdfMatch(value: string): string {
  let normalized = value
    .replace(/\\(?:left|right|displaystyle|textstyle|scriptstyle|scriptscriptstyle)\b/g, "")
    .replace(/\\(?:,|;|:|!|quad|qquad| )/g, "")
    .replace(/\\begin\{[^}]+\}|\\end\{[^}]+\}/g, "")
    .replace(/\\(mathrm|mathbf|mathit|mathsf|mathtt|operatorname|text)\s*\{/g, "{")
    .replace(/\\([A-Za-z]+)/g, (_match, command: string) => latexSymbolMap[command] ?? "")
    .replace(/\\./g, "");
  normalized = normalizeVisibleCharacter(normalized);
  return normalized.replace(/[()[\]|]/g, "");
}

function normalizedGlyphSequence(glyphs: PdfTextGlyph[]): {
  text: string;
  glyphIndices: number[];
} {
  let text = "";
  const glyphIndices: number[] = [];
  for (const glyph of glyphs) {
    const normalized = normalizeVisibleCharacter(glyph.text).replace(/[()[\]|]/g, "");
    for (const character of normalized) {
      text += character;
      glyphIndices.push(glyph.index);
    }
  }
  return { text, glyphIndices };
}

export function inlineFormulaNodeAtTextHit(
  hit: PdfTextHit,
  candidates: VisualNode[],
): VisualNode | null {
  const sequence = normalizedGlyphSequence(hit.lineGlyphs);
  if (!sequence.text) return null;
  let best: { node: VisualNode; length: number } | null = null;
  for (const node of candidates) {
    if (node.kind !== "inline_math" || !node.text) continue;
    const formula = normalizeLatexForPdfMatch(node.text);
    if (!formula) continue;
    let offset = sequence.text.indexOf(formula);
    while (offset >= 0) {
      const indexes = sequence.glyphIndices.slice(offset, offset + formula.length);
      if (indexes.includes(hit.glyphIndex)) {
        if (!best || formula.length > best.length) best = { node, length: formula.length };
        break;
      }
      offset = sequence.text.indexOf(formula, offset + 1);
    }
  }
  return best?.node ?? null;
}

export function findDirectEditFromSource(
  pageIndex: number,
  x: number,
  y: number,
  sourceText: string,
  sourcePath: string,
  result: InverseSearchResult,
  textHit: PdfTextHit | null,
  layoutBoxes: LayoutBox[],
  nodes: VisualNode[],
): DirectEditState | null {
  const normalizedSource = normalizePath(result.sourcePath);
  const normalizedExpected = normalizePath(sourcePath);
  if (
    normalizedSource !== normalizedExpected
    && !normalizedSource.endsWith(`/${normalizedExpected}`)
  ) {
    return null;
  }
  const lineRange = sourceLineByteRange(sourceText, result.line);
  const offset = sourceByteOffsetAtLineColumn(sourceText, result.line, result.column);
  if (!lineRange || offset === null) return null;
  const nodesById = new Map(nodes.map((node) => [node.id, node]));
  const hasColumn = result.column !== null && result.column !== undefined && result.column >= 0;
  const unsorted = layoutBoxes
    .map((layout) => ({ layout, node: nodesById.get(layout.nodeId) }))
    .filter((entry): entry is { layout: LayoutBox; node: VisualNode } =>
      Boolean(entry.node && isDirectlyEditable(entry.node, entry.layout)),
    )
    .filter(({ node }) =>
      (node.source.startByte <= offset && offset <= node.source.endByte)
      || (node.source.startByte < lineRange.end && node.source.endByte > lineRange.start),
    );
  const formulaAtHit = !hasColumn && textHit
    ? inlineFormulaNodeAtTextHit(textHit, unsorted.map((entry) => entry.node))
    : null;
  const matching = unsorted.sort((left, right) => {
      const sourceContains = (entry: { node: VisualNode }) =>
        entry.node.source.startByte <= offset && offset <= entry.node.source.endByte;
      const geometryHit = (entry: { layout: LayoutBox }) =>
        entry.layout.rects.some((rect) =>
          rect.page === pageIndex + 1 && containsPoint(rect, x, y),
        );
      const leftContains = sourceContains(left);
      const rightContains = sourceContains(right);
      const leftGeometry = geometryHit(left);
      const rightGeometry = geometryHit(right);
      if (hasColumn && leftContains !== rightContains) return leftContains ? -1 : 1;
      if (formulaAtHit) {
        const leftFormula = left.node.id === formulaAtHit.id;
        const rightFormula = right.node.id === formulaAtHit.id;
        if (leftFormula !== rightFormula) return leftFormula ? -1 : 1;
      } else if (!hasColumn) {
        const leftParagraph = left.node.kind === "paragraph";
        const rightParagraph = right.node.kind === "paragraph";
        const leftInlineMath = left.node.kind === "inline_math";
        const rightInlineMath = right.node.kind === "inline_math";
        if (leftParagraph && rightInlineMath) return -1;
        if (rightParagraph && leftInlineMath) return 1;
      }
      if (leftGeometry !== rightGeometry) return leftGeometry ? -1 : 1;
      const leftHitArea = Math.min(
        ...left.layout.rects
          .filter((rect) => rect.page === pageIndex + 1 && containsPoint(rect, x, y))
          .map(rectArea),
        Number.POSITIVE_INFINITY,
      );
      const rightHitArea = Math.min(
        ...right.layout.rects
          .filter((rect) => rect.page === pageIndex + 1 && containsPoint(rect, x, y))
          .map(rectArea),
        Number.POSITIVE_INFINITY,
      );
      if (leftHitArea !== rightHitArea) return leftHitArea - rightHitArea;
      const leftSpan = left.node.source.endByte - left.node.source.startByte;
      const rightSpan = right.node.source.endByte - right.node.source.startByte;
      return leftSpan - rightSpan;
    });
  for (const { node, layout } of matching) {
    const state = stateForNode(pageIndex, node, layout, x, y);
    if (state) return state;
  }
  return null;
}

function attributesToPatch(attributes: NodeAttributes): NodeAttributesPatch {
  return {
    placement: attributes.placement ?? "",
    caption: attributes.caption ?? "",
    label: attributes.label ?? "",
    imagePath: attributes.imagePath ?? "",
    imageWidth: attributes.imageWidth ?? "",
    columnSpec: attributes.columnSpec ?? "",
    tableRows: attributes.tableRows,
  };
}

function AttributeEditor({
  edit,
  onChange,
}: {
  edit: DirectEditState;
  onChange: (attributes: NodeAttributes) => void;
}) {
  const update = (patch: Partial<NodeAttributes>) => onChange({ ...edit.attributes, ...patch });
  const field = (
    label: string,
    value: string | null,
    onValue: (value: string) => void,
    placeholder?: string,
  ) => (
    <label className="vt-pdf-attribute-field">
      <span>{label}</span>
      <input value={value ?? ""} placeholder={placeholder} onChange={(event) => onValue(event.target.value)} />
    </label>
  );

  if (edit.node.kind === "figure") {
    const widthMatch = edit.attributes.imageWidth?.match(/([0-9]*\.?[0-9]+)\\(?:line|text)width/);
    const widthRatio = Math.max(0.1, Math.min(1, Number(widthMatch?.[1] ?? 0.8)));
    return (
      <div className="vt-pdf-attribute-grid vt-image-visual-editor">
        <label className="vt-pdf-attribute-field">
          <span>浮动位置</span>
          <select
            value={edit.attributes.placement ?? "htbp"}
            onChange={(event) => update({ placement: event.target.value })}
          >
            <option value="htbp">自动 htbp</option>
            <option value="H">固定当前位置 H</option>
            <option value="t">页顶 t</option>
            <option value="b">页底 b</option>
            <option value="p">浮动页 p</option>
          </select>
        </label>
        <label className="vt-pdf-attribute-field">
          <span>图片宽度 · {Math.round(widthRatio * 100)}%</span>
          <input
            type="range"
            min="0.1"
            max="1"
            step="0.05"
            value={widthRatio}
            onChange={(event) => update({ imageWidth: `${Number(event.target.value).toFixed(2)}\\linewidth` })}
          />
        </label>
        <div className="wide vt-image-width-presets">
          {[0.25, 0.5, 0.75, 1].map((ratio) => (
            <button
              type="button"
              key={ratio}
              className={Math.abs(widthRatio - ratio) < 0.01 ? "active" : ""}
              onClick={() => update({ imageWidth: `${ratio.toFixed(2)}\\linewidth` })}
            >
              {Math.round(ratio * 100)}%
            </button>
          ))}
        </div>
        <div className="wide">{field("图片路径", edit.attributes.imagePath, (imagePath) => update({ imagePath }))}</div>
        <div className="wide">{field("图注", edit.attributes.caption, (caption) => update({ caption }))}</div>
        <div className="wide">{field("标签", edit.attributes.label, (label) => update({ label }), "fig:example")}</div>
        <p className="wide vt-image-editor-note">
          页面缩放手柄和这里的宽度滑块都会修改 <code>\\includegraphics[width=…]</code>；LaTeX 浮动体不支持任意像素坐标拖动。
        </p>
      </div>
    );
  }

  const rows = edit.attributes.tableRows.map((row) => row.join(" & ")).join("\n");
  return (
    <div className="vt-pdf-attribute-grid">
      {field("位置", edit.attributes.placement, (placement) => update({ placement }), "htbp")}
      {field("列格式", edit.attributes.columnSpec, (columnSpec) => update({ columnSpec }), "lcr")}
      <div className="wide">{field("表注", edit.attributes.caption, (caption) => update({ caption }))}</div>
      <div className="wide">{field("标签", edit.attributes.label, (label) => update({ label }), "tab:example")}</div>
      <label className="vt-pdf-attribute-field wide">
        <span>单元格（每行一行，使用 & 分列）</span>
        <textarea
          value={rows}
          rows={Math.max(4, edit.attributes.tableRows.length)}
          onChange={(event) =>
            update({
              tableRows: event.target.value
                .split("\n")
                .filter((row) => row.trim().length > 0)
                .map((row) => row.split("&").map((cell) => cell.trim())),
            })
          }
        />
      </label>
    </div>
  );
}

function DirectEditOverlay({
  edit,
  onChange,
  onAttributesChange,
  onCancel,
  onCommit,
  onDelete,
}: {
  edit: DirectEditState;
  onChange: (draft: string) => void;
  onAttributesChange: (attributes: NodeAttributes) => void;
  onCancel: () => void;
  onCommit: () => void;
  onDelete?: () => void;
}) {
  const panelRef = useRef<HTMLDivElement | null>(null);
  const [position, setPosition] = useState({ left: 16, top: 16, width: 640, ready: false });
  const isMath = edit.node.kind === "inline_math" || edit.node.kind === "display_math";
  const label = isMath
    ? "公式"
    : edit.node.kind === "paragraph"
      ? "正文"
      : edit.node.kind === "figure"
        ? "图片"
        : edit.node.kind === "table"
          ? "表格"
          : "标题";
  const contentDirty = edit.draft !== (edit.node.text ?? "");
  const attributesDirty = JSON.stringify(edit.attributes) !== JSON.stringify(edit.node.attributes);
  const dirty = edit.node.kind === "figure" || edit.node.kind === "table"
    ? attributesDirty
    : contentDirty;

  useLayoutEffect(() => {
    const update = () => {
      const panel = panelRef.current;
      const anchor = edit.viewportAnchor ?? {
        left: window.innerWidth / 2,
        right: window.innerWidth / 2,
        top: window.innerHeight / 2,
        bottom: window.innerHeight / 2,
      };
      const preferredWidth = isMath ? 760 : edit.node.kind === "paragraph" ? 620 : 680;
      const width = Math.max(320, Math.min(preferredWidth, window.innerWidth - 24));
      const measuredHeight = panel?.getBoundingClientRect().height ?? 420;
      const left = Math.max(12, Math.min(anchor.left, window.innerWidth - width - 12));
      const below = anchor.bottom + 10;
      const above = anchor.top - measuredHeight - 10;
      const top = below + measuredHeight <= window.innerHeight - 12
        ? below
        : above >= 12
          ? above
          : Math.max(12, Math.min(below, window.innerHeight - measuredHeight - 12));
      setPosition({ left, top, width, ready: true });
    };
    update();
    const frame = window.requestAnimationFrame(update);
    window.addEventListener("resize", update);
    return () => {
      window.cancelAnimationFrame(frame);
      window.removeEventListener("resize", update);
    };
  }, [edit.node.kind, edit.viewportAnchor, isMath]);

  const handleKeyboard = (event: KeyboardEvent<HTMLDivElement>) => {
    if (event.key === "Escape") {
      event.preventDefault();
      onCancel();
    } else if (event.key === "Enter" && (event.metaKey || event.ctrlKey)) {
      event.preventDefault();
      onCommit();
    }
  };

  return createPortal(
    <div className="vt-pdf-editor-layer">
      <div
        ref={panelRef}
        className={`vt-pdf-direct-overlay ${edit.node.kind}`}
        style={{
          position: "fixed",
          left: position.left,
          top: position.top,
          width: position.width,
          visibility: position.ready ? "visible" : "hidden",
        }}
        role="dialog"
        aria-modal="true"
        aria-label={`编辑${label}`}
        onKeyDown={handleKeyboard}
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="vt-pdf-direct-label">
          <span>编辑{label}</span>
          <span>{dirty ? "未应用 · " : ""}{edit.layout.confidence.toUpperCase()}</span>
        </header>
        <div className="vt-pdf-direct-body">
          {edit.node.kind === "figure" || edit.node.kind === "table" ? (
            <AttributeEditor edit={edit} onChange={onAttributesChange} />
          ) : isMath ? (
            <MathNodeEditor
              value={edit.draft}
              autoFocus
              showToolbar
              showCandidates
              onChange={onChange}
            />
          ) : edit.node.kind === "paragraph" ? (
            <div className="vt-paragraph-editor">
              {edit.node.support === "partial" && (
                <p>
                  该段包含行内公式或其他 LaTeX 命令；普通文字可直接增删，保留 <code>$…$</code> 等内嵌语法即可。
                </p>
              )}
              <textarea
                autoFocus
                value={edit.draft}
                onChange={(event) => onChange(event.target.value)}
                aria-label="编辑页面正文"
                rows={Math.max(7, edit.draft.split("\n").length + 2)}
              />
            </div>
          ) : (
            <input
              autoFocus
              value={edit.draft}
              onChange={(event) => onChange(event.target.value)}
              aria-label="编辑页面节点"
            />
          )}
        </div>
        <footer className="vt-pdf-direct-actions">
          <div>
            {(isMath || edit.node.kind === "paragraph") && (
              <button type="button" className="danger" onClick={() => onChange("")}>
                清空内容
              </button>
            )}
            {onDelete && (
              <button
                type="button"
                className="danger solid"
                onClick={() => {
                  const description = isMath
                    ? "整个公式（包括公式分隔符）"
                    : `整个${label}节点`;
                  if (window.confirm(`确定删除${description}吗？该操作可以撤销。`)) onDelete();
                }}
              >
                删除整个节点
              </button>
            )}
          </div>
          <div>
            <button type="button" title="取消 (Esc)" onClick={onCancel}><X size={14} />取消</button>
            <button type="button" className="confirm" title="写回源码 (⌘/Ctrl+Enter)" onClick={onCommit}>
              <Check size={14} />应用并重新排版
            </button>
          </div>
        </footer>
      </div>
    </div>,
    document.body,
  );
}

function RenderedPage({
  info,
  page,
  cssWidth,
  grayscale,
  rootRef,
  highlights,
  layoutBoxes,
  nodes,
  sourceText,
  sourcePath,
  editable,
  directEdit,
  onDirectEdit,
  onNodeSelect,
  onNodeCommit,
  onNodeAttributesCommit,
  onVisible,
  onInverseSearch,
  onError,
}: RenderedPageProps) {
  const pageRef = useRef<HTMLElement | null>(null);
  const clickSequenceRef = useRef(0);
  const [resolvingHit, setResolvingHit] = useState(false);
  const visible = useObservedVisibility(pageRef, rootRef, "900px 0px", () => onVisible(page.index));
  const rotated = page.rotationDegrees === 90 || page.rotationDegrees === 270;
  const widthPoints = rotated ? page.heightPoints : page.widthPoints;
  const heightPoints = rotated ? page.widthPoints : page.heightPoints;
  const cssHeight = Math.max(1, cssWidth * (heightPoints / widthPoints));
  const pixelRatio = Math.min(window.devicePixelRatio || 1, 2.5);
  const targetWidthPixels = Math.max(96, Math.min(8_192, Math.round(cssWidth * pixelRatio)));
  const rendered = useRenderedImage(info, page.index, targetWidthPixels, grayscale, visible, onError);
  const pageHighlights = highlights.filter((box) => box.page === page.index + 1);

  const pdfCoordinates = (event: MouseEvent<HTMLElement>) => {
    if (!pageRef.current) return null;
    const bounds = pageRef.current.getBoundingClientRect();
    return {
      x: ((event.clientX - bounds.left) / bounds.width) * widthPoints,
      y: ((event.clientY - bounds.top) / bounds.height) * heightPoints,
    };
  };

  const viewportAnchorForRect = (rect: PdfRect): ViewportAnchor | undefined => {
    const element = pageRef.current;
    if (!element) return undefined;
    const bounds = element.getBoundingClientRect();
    const left = bounds.left + (rect.x / widthPoints) * bounds.width;
    const top = bounds.top + (rect.y / heightPoints) * bounds.height;
    const right = left + (rect.width / widthPoints) * bounds.width;
    const bottom = top + (rect.height / heightPoints) * bounds.height;
    return { left, top, right, bottom };
  };

  const handleSingleClick = async (event: MouseEvent<HTMLElement>) => {
    if (!editable || directEdit) return;
    const point = pdfCoordinates(event);
    if (!point) return;
    const sequence = ++clickSequenceRef.current;
    setResolvingHit(true);
    let inverseResult: InverseSearchResult | null = null;
    let textHit: PdfTextHit | null = null;
    let hit: DirectEditState | null = null;
    const [inverseOutcome, textOutcome] = await Promise.allSettled([
      desktopApi.inverseSearch(
        info.pdfPath,
        page.index + 1,
        point.x,
        point.y,
      ),
      desktopApi.pdfTextHit(info.pdfPath, page.index, point.x, point.y),
    ]);
    if (inverseOutcome.status === "fulfilled") inverseResult = inverseOutcome.value;
    if (textOutcome.status === "fulfilled") textHit = textOutcome.value;
    if (inverseResult && sourceText !== undefined && sourcePath) {
      hit = findDirectEditFromSource(
        page.index,
        point.x,
        point.y,
        sourceText,
        sourcePath,
        inverseResult,
        textHit,
        layoutBoxes,
        nodes,
      );
    }
    if (sequence !== clickSequenceRef.current) {
      setResolvingHit(false);
      return;
    }
    hit ??= findDirectEdit(page.index, point.x, point.y, layoutBoxes, nodes);
    setResolvingHit(false);
    if (!hit) {
      if (inverseResult) onInverseSearch?.(inverseResult);
      return;
    }
    const next = { ...hit, viewportAnchor: viewportAnchorForRect(hit.hitRect) };
    onDirectEdit(next);
    onNodeSelect?.(next.node);
  };

  const handleImageResize = (event: PointerEvent<HTMLButtonElement>) => {
    if (!directEdit || directEdit.node.kind !== "figure" || !pageRef.current) return;
    event.preventDefault();
    event.stopPropagation();
    const bounds = pageRef.current.getBoundingClientRect();
    const startX = event.clientX;
    const startCssWidth = Math.max(24, (directEdit.hitRect.width / widthPoints) * bounds.width);
    const widthMatch = directEdit.attributes.imageWidth?.match(/([0-9]*\.?[0-9]+)\\(?:line|text)width/);
    const startRatio = Math.max(0.1, Math.min(1, Number(widthMatch?.[1] ?? 0.8)));
    const move = (moveEvent: globalThis.PointerEvent) => {
      const nextCssWidth = Math.max(24, startCssWidth + moveEvent.clientX - startX);
      const ratio = Math.max(0.1, Math.min(1, startRatio * (nextCssWidth / startCssWidth)));
      onDirectEdit({
        ...directEdit,
        attributes: {
          ...directEdit.attributes,
          imageWidth: `${ratio.toFixed(2)}\\linewidth`,
        },
      });
    };
    const finish = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", finish);
      window.removeEventListener("pointercancel", finish);
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", finish, { once: true });
    window.addEventListener("pointercancel", finish, { once: true });
  };

  const handleDoubleClick = async (event: MouseEvent<HTMLElement>) => {
    if (editable || !onInverseSearch) return;
    const point = pdfCoordinates(event);
    if (!point) return;
    try {
      onInverseSearch(await desktopApi.inverseSearch(info.pdfPath, page.index + 1, point.x, point.y));
    } catch (error) {
      onError?.(messageOf(error));
    }
  };

  return (
    <article
      ref={pageRef}
      className={`vt-pdf-page${editable ? " editable" : ""}${resolvingHit ? " resolving-hit" : ""}`}
      data-page-index={page.index}
      style={{ width: cssWidth, height: cssHeight }}
      onClick={handleSingleClick}
      onDoubleClick={(event) => void handleDoubleClick(event)}
      aria-label={`PDF page ${page.index + 1}`}
    >
      {rendered ? (
        <img
          draggable={false}
          src={desktopApi.fileUrl(rendered.cachePath)}
          alt={`Rendered PDF page ${page.index + 1}`}
          width={rendered.imageWidthPixels}
          height={rendered.imageHeightPixels}
        />
      ) : (
        <div className="vt-pdf-page-loading">
          {visible && <LoaderCircle size={22} className="spin" />}
          <span>第 {page.index + 1} 页</span>
        </div>
      )}
      <div className="vt-pdf-highlight-layer" aria-hidden="true">
        {pageHighlights.map((box, index) => (
          <span
            key={`${box.x}-${box.y}-${index}`}
            className="vt-pdf-highlight"
            style={{
              left: `${(box.x / widthPoints) * 100}%`,
              top: `${(box.y / heightPoints) * 100}%`,
              width: `${(box.width / widthPoints) * 100}%`,
              height: `${(box.height / heightPoints) * 100}%`,
            }}
          />
        ))}
      </div>
      {editable && directEdit?.pageIndex === page.index && (
        <span
          className={`vt-pdf-selected-node ${directEdit.node.kind}`}
          style={{
            left: `${(directEdit.hitRect.x / widthPoints) * 100}%`,
            top: `${(directEdit.hitRect.y / heightPoints) * 100}%`,
            width: `${(directEdit.hitRect.width / widthPoints) * 100}%`,
            height: `${(directEdit.hitRect.height / heightPoints) * 100}%`,
          }}
        >
          {directEdit.node.kind === "figure" && (
            <button
              type="button"
              className="vt-pdf-image-resize-handle"
              title="拖动修改图片宽度"
              onPointerDown={handleImageResize}
            />
          )}
        </span>
      )}
      <span className="vt-pdf-page-number">{page.index + 1}</span>
    </article>
  );
}

function Thumbnail({ info, page, rootRef, active, onSelect }: ThumbnailProps) {
  const ref = useRef<HTMLButtonElement | null>(null);
  const visible = useObservedVisibility(ref, rootRef, "300px", undefined);
  const rendered = useRenderedImage(info, page.index, 180, false, visible, undefined);
  const rotated = page.rotationDegrees === 90 || page.rotationDegrees === 270;
  const ratio = rotated
    ? page.widthPoints / page.heightPoints
    : page.heightPoints / page.widthPoints;
  const height = 132 * ratio;

  return (
    <button
      ref={ref}
      type="button"
      className={`vt-pdf-thumbnail${active ? " active" : ""}`}
      onClick={onSelect}
      aria-label={`跳转到第 ${page.index + 1} 页`}
    >
      <span className="vt-pdf-thumbnail-canvas" style={{ height }}>
        {rendered && <img draggable={false} src={desktopApi.fileUrl(rendered.cachePath)} alt="" />}
      </span>
      <span>{page.index + 1}</span>
    </button>
  );
}

export function PdfViewer({
  pdfPath,
  buildKey,
  highlights = [],
  layoutBoxes = [],
  nodes = [],
  sourceText,
  sourcePath,
  editable = false,
  onInverseSearch,
  onNodeSelect,
  onNodeCommit,
  onNodeDelete,
  onNodeAttributesCommit,
  onError,
}: PdfViewerProps) {
  const [info, setInfo] = useState<PdfDocumentInfo | null>(null);
  const [loading, setLoading] = useState(true);
  const [zoom, setZoom] = useState(1);
  const [grayscale, setGrayscale] = useState(false);
  const [activePage, setActivePage] = useState(0);
  const [viewportWidth, setViewportWidth] = useState(900);
  const [directEdit, setDirectEdit] = useState<DirectEditState | null>(null);
  const scrollRef = useRef<HTMLDivElement | null>(null);
  const thumbnailRef = useRef<HTMLDivElement | null>(null);
  const pageRefs = useRef(new Map<number, HTMLElement>());
  const errorRef = useRef(onError);
  errorRef.current = onError;

  useEffect(() => {
    if (!editable) setDirectEdit(null);
  }, [editable]);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setInfo(null);
    setActivePage(0);
    setDirectEdit(null);
    void desktopApi
      .pdfDocumentInfo(pdfPath)
      .then((next) => {
        if (!cancelled) setInfo(next);
      })
      .catch((error: unknown) => {
        if (!cancelled) errorRef.current?.(messageOf(error));
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [pdfPath, buildKey]);

  useEffect(() => {
    const element = scrollRef.current;
    if (!element) return;
    const observer = new ResizeObserver((entries) => {
      const width = entries[0]?.contentRect.width;
      if (width) setViewportWidth(width);
    });
    observer.observe(element);
    return () => observer.disconnect();
  }, []);

  const fitWidth = Math.max(320, Math.min(1_200, viewportWidth - 72));
  const cssWidth = fitWidth * zoom;

  const scrollToPage = useCallback((pageIndex: number) => {
    const element = pageRefs.current.get(pageIndex);
    element?.scrollIntoView({ behavior: "smooth", block: "start" });
    setActivePage(pageIndex);
  }, []);

  const visibleHighlights = useMemo(() => highlights.filter((box) => box.page > 0), [highlights]);
  const highlightedPage = visibleHighlights[0]?.page ? visibleHighlights[0].page - 1 : null;

  useEffect(() => {
    if (highlightedPage === null || highlightedPage === activePage) return;
    scrollToPage(highlightedPage);
  }, [highlightedPage, activePage, scrollToPage]);

  if (loading) {
    return <div className="vt-pdf-viewer-state"><LoaderCircle className="spin" />正在读取 PDF 页面…</div>;
  }
  if (!info || info.pages.length === 0) {
    return <div className="vt-pdf-viewer-state">PDF 没有可显示的页面。</div>;
  }

  const directlyEditableCount = layoutBoxes.filter((layout) => {
    const node = nodes.find((candidate) => candidate.id === layout.nodeId);
    return Boolean(node && isDirectlyEditable(node, layout));
  }).length;

  return (
    <section className="vt-pdf-viewer">
      <header className="vt-pdf-toolbar">
        <div className="vt-pdf-toolbar-group">
          <button type="button" onClick={() => scrollToPage(Math.max(0, activePage - 1))} disabled={activePage === 0}>
            <ChevronLeft size={16} />
          </button>
          <span>{activePage + 1} / {info.pages.length}</span>
          <button
            type="button"
            onClick={() => scrollToPage(Math.min(info.pages.length - 1, activePage + 1))}
            disabled={activePage >= info.pages.length - 1}
          >
            <ChevronRight size={16} />
          </button>
        </div>
        <div className="vt-pdf-toolbar-group">
          <button type="button" onClick={() => setZoom((value) => Math.max(0.5, value - 0.1))}><Minus size={16} /></button>
          <span>{Math.round(zoom * 100)}%</span>
          <button type="button" onClick={() => setZoom((value) => Math.min(3, value + 0.1))}><Plus size={16} /></button>
          <button type="button" title="适合宽度" onClick={() => setZoom(1)}><RotateCcw size={15} /></button>
        </div>
        <label className="vt-pdf-grayscale">
          <input type="checkbox" checked={grayscale} onChange={(event) => setGrayscale(event.target.checked)} />
          灰度
        </label>
        {editable && (
          <span className="vt-pdf-editable-count" title="当前可安全直接编辑的页面节点">
            可编辑 {directlyEditableCount}
          </span>
        )}
        <span className="vt-pdf-fingerprint" title={info.fingerprint}>{info.fingerprint.slice(0, 10)}</span>
      </header>

      <div className="vt-pdf-body">
        <aside ref={thumbnailRef} className="vt-pdf-thumbnails">
          {info.pages.map((page) => (
            <Thumbnail
              key={page.index}
              info={info}
              page={page}
              rootRef={thumbnailRef}
              active={activePage === page.index}
              onSelect={() => scrollToPage(page.index)}
            />
          ))}
        </aside>
        <div ref={scrollRef} className="vt-pdf-scroll">
          <div className="vt-pdf-pages">
            {info.pages.map((page) => (
              <div
                key={page.index}
                ref={(element) => {
                  if (element) pageRefs.current.set(page.index, element);
                  else pageRefs.current.delete(page.index);
                }}
              >
                <RenderedPage
                  info={info}
                  page={page}
                  cssWidth={cssWidth}
                  grayscale={grayscale}
                  rootRef={scrollRef}
                  highlights={visibleHighlights}
                  layoutBoxes={layoutBoxes}
                  nodes={nodes}
                  sourceText={sourceText}
                  sourcePath={sourcePath}
                  editable={editable}
                  directEdit={directEdit}
                  onDirectEdit={setDirectEdit}
                  onNodeSelect={onNodeSelect}
                  onNodeCommit={onNodeCommit}
                  onNodeAttributesCommit={onNodeAttributesCommit}
                  onVisible={setActivePage}
                  onInverseSearch={onInverseSearch}
                  onError={onError}
                />
              </div>
            ))}
          </div>
        </div>
      </div>
      <footer className="vt-pdf-hint">
        {editable
          ? "页面保持真实编译排版；单击正文、标题、公式、图片或表格后，在浮动工作台中编辑并重新排版。"
          : "双击页面内容可通过 SyncTeX 跳转到对应源码。"}
      </footer>
      {editable && directEdit && (
        <DirectEditOverlay
          edit={directEdit}
          onChange={(draft) => setDirectEdit({ ...directEdit, draft })}
          onAttributesChange={(attributes) => setDirectEdit({ ...directEdit, attributes })}
          onCancel={() => setDirectEdit(null)}
          onDelete={onNodeDelete ? () => {
            onNodeDelete(directEdit.node);
            setDirectEdit(null);
          } : undefined}
          onCommit={() => {
            if (directEdit.node.kind === "figure" || directEdit.node.kind === "table") {
              onNodeAttributesCommit?.(
                directEdit.node,
                attributesToPatch(directEdit.attributes),
              );
            } else {
              onNodeCommit?.(directEdit.node, directEdit.draft);
            }
            setDirectEdit(null);
          }}
        />
      )}
    </section>
  );
}
