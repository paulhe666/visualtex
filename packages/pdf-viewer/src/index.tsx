import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type KeyboardEvent,
  type MouseEvent,
  type RefObject,
} from "react";
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
  type VisualNode,
} from "@visualtex/protocol";
import "./styles.css";

export interface PdfViewerProps {
  pdfPath: string;
  buildKey?: string;
  highlights?: PdfRect[];
  layoutBoxes?: LayoutBox[];
  nodes?: VisualNode[];
  editable?: boolean;
  onInverseSearch?: (result: InverseSearchResult) => void;
  onNodeSelect?: (node: VisualNode) => void;
  onNodeCommit?: (node: VisualNode, content: string) => void;
  onNodeAttributesCommit?: (node: VisualNode, patch: NodeAttributesPatch) => void;
  onError?: (message: string) => void;
}

export interface DirectEditState {
  pageIndex: number;
  node: VisualNode;
  layout: LayoutBox;
  rect: PdfRect;
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
  editable: boolean;
  directEdit: DirectEditState | null;
  onDirectEdit: (edit: DirectEditState | null) => void;
  onNodeSelect?: (node: VisualNode) => void;
  onNodeCommit?: (node: VisualNode, content: string) => void;
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
  return node.support === "partial" && ["figure", "table"].includes(node.kind);
}

function containsPoint(rect: PdfRect, x: number, y: number): boolean {
  if (rect.width <= 0 || rect.height <= 0) return false;
  const padding = Math.max(2, Math.min(rect.height * 0.3, 6));
  return (
    x >= rect.x - padding &&
    x <= rect.x + rect.width + padding &&
    y >= rect.y - padding &&
    y <= rect.y + rect.height + padding
  );
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
    if (!pageRects.some((rect) => containsPoint(rect, x, y))) continue;
    const rect = unionRects(pageRects);
    if (!rect) continue;
    candidates.push({
      pageIndex,
      node,
      layout,
      rect,
      draft: node.text ?? "",
      attributes: structuredClone(node.attributes),
    });
  }
  candidates.sort((left, right) => {
    const leftArea = left.rect.width * left.rect.height;
    const rightArea = right.rect.width * right.rect.height;
    const kindPriority = (node: VisualNode) =>
      node.kind === "inline_math" || node.kind === "display_math" ? 0 : 1;
    return kindPriority(left.node) - kindPriority(right.node) || leftArea - rightArea;
  });
  return candidates[0] ?? null;
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
    return (
      <div className="vt-pdf-attribute-grid">
        {field("位置", edit.attributes.placement, (placement) => update({ placement }), "htbp")}
        {field("宽度", edit.attributes.imageWidth, (imageWidth) => update({ imageWidth }), "0.8\\linewidth")}
        <div className="wide">{field("图片路径", edit.attributes.imagePath, (imagePath) => update({ imagePath }))}</div>
        <div className="wide">{field("图注", edit.attributes.caption, (caption) => update({ caption }))}</div>
        <div className="wide">{field("标签", edit.attributes.label, (label) => update({ label }), "fig:example")}</div>
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
  page,
  onChange,
  onAttributesChange,
  onCancel,
  onCommit,
}: {
  edit: DirectEditState;
  page: PdfPageInfo;
  onChange: (draft: string) => void;
  onAttributesChange: (attributes: NodeAttributes) => void;
  onCancel: () => void;
  onCommit: () => void;
}) {
  const handleKeyboard = (event: KeyboardEvent<HTMLDivElement>) => {
    if (event.key === "Escape") {
      event.preventDefault();
      onCancel();
    } else if (event.key === "Enter" && (event.metaKey || event.ctrlKey)) {
      event.preventDefault();
      onCommit();
    }
  };
  const top = Math.max(0, (edit.rect.y / page.heightPoints) * 100);
  const left = Math.max(0, (edit.rect.x / page.widthPoints) * 100);
  const width = Math.max(16, Math.min(100 - left, (edit.rect.width / page.widthPoints) * 100));
  const height = Math.max(4, Math.min(100 - top, (edit.rect.height / page.heightPoints) * 100));

  return (
    <div
      className={`vt-pdf-direct-overlay ${edit.node.kind}`}
      style={{ left: `${left}%`, top: `${top}%`, width: `${width}%`, minHeight: `${height}%` }}
      onKeyDown={handleKeyboard}
      onClick={(event) => event.stopPropagation()}
      onDoubleClick={(event) => event.stopPropagation()}
    >
      <div className="vt-pdf-direct-label">
        <span>{edit.node.kind === "inline_math" || edit.node.kind === "display_math" ? "公式" : edit.node.kind === "paragraph" ? "正文" : edit.node.kind === "figure" ? "图片属性" : edit.node.kind === "table" ? "表格属性" : "标题"}</span>
        <span>{edit.layout.confidence}</span>
      </div>
      <div className="vt-pdf-direct-body">
        {edit.node.kind === "figure" || edit.node.kind === "table" ? (
          <AttributeEditor edit={edit} onChange={onAttributesChange} />
        ) : edit.node.kind === "inline_math" || edit.node.kind === "display_math" ? (
          <MathNodeEditor value={edit.draft} autoFocus onChange={onChange} />
        ) : edit.node.kind === "paragraph" ? (
          <textarea
            autoFocus
            value={edit.draft}
            onChange={(event) => onChange(event.target.value)}
            aria-label="编辑页面正文"
            rows={Math.max(3, edit.draft.split("\n").length)}
          />
        ) : (
          <input
            autoFocus
            value={edit.draft}
            onChange={(event) => onChange(event.target.value)}
            aria-label="编辑页面节点"
          />
        )}
      </div>
      <div className="vt-pdf-direct-actions">
        <button type="button" title="取消 (Esc)" onClick={onCancel}><X size={14} /></button>
        <button type="button" className="confirm" title="写回源码 (⌘/Ctrl+Enter)" onClick={onCommit}><Check size={14} /></button>
      </div>
    </div>
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
  const visible = useObservedVisibility(pageRef, rootRef, "900px 0px", () => onVisible(page.index));
  const rotated = page.rotationDegrees === 90 || page.rotationDegrees === 270;
  const widthPoints = rotated ? page.heightPoints : page.widthPoints;
  const heightPoints = rotated ? page.widthPoints : page.heightPoints;
  const displayPage = { ...page, widthPoints, heightPoints };
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

  const handleSingleClick = (event: MouseEvent<HTMLElement>) => {
    if (!editable) return;
    const point = pdfCoordinates(event);
    if (!point) return;
    const hit = findDirectEdit(page.index, point.x, point.y, layoutBoxes, nodes);
    if (!hit) return;
    onDirectEdit(hit);
    onNodeSelect?.(hit.node);
  };

  const handleDoubleClick = async (event: MouseEvent<HTMLElement>) => {
    if (!onInverseSearch) return;
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
      className={`vt-pdf-page${editable ? " editable" : ""}`}
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
        <DirectEditOverlay
          edit={directEdit}
          page={displayPage}
          onChange={(draft) => onDirectEdit({ ...directEdit, draft })}
          onAttributesChange={(attributes) => onDirectEdit({ ...directEdit, attributes })}
          onCancel={() => onDirectEdit(null)}
          onCommit={() => {
            if (directEdit.node.kind === "figure" || directEdit.node.kind === "table") {
              onNodeAttributesCommit?.(
                directEdit.node,
                attributesToPatch(directEdit.attributes),
              );
            } else {
              onNodeCommit?.(directEdit.node, directEdit.draft);
            }
            onDirectEdit(null);
          }}
        />
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
  editable = false,
  onInverseSearch,
  onNodeSelect,
  onNodeCommit,
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
          ? "页面保持真实编译排版；单击正文、标题、公式、图片或表格原位编辑，写回源码后自动重新编译。"
          : "双击页面内容可通过 SyncTeX 跳转到对应源码。"}
      </footer>
    </section>
  );
}
