import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import {
  ChevronDown,
  Copy,
  Eraser,
  Eye,
  EyeOff,
  FilePlus2,
  Hand,
  MousePointer2,
  Redo2,
  Trash2,
  Undo2,
  X,
} from "lucide-react";
import { restoreCustomSymbolDesignerDocument } from "../archive.ts";
import {
  compileCustomSymbolDesignerArtwork,
  customSymbolDefinitionFromDesignerDocument,
  glyphLayerFromAsset,
  measureCustomSymbolBox,
} from "../compiler.ts";
import {
  createEmptyCustomSymbolDesignerDocument,
  type CustomSymbolDesignerDocument,
  type CustomSymbolDesignerLayer,
  type CustomSymbolGeometryPreset,
  type CustomSymbolGlyphAsset,
} from "../designerTypes.ts";
import { createCustomSymbolEraserLayer, createCustomSymbolGeometryLayer, createGlyphLayerSlices } from "../geometry.ts";
import { glyphAssetFromEngineDrawing, measureMaterialItalic, textStyleVariant } from "../glyphAsset.ts";
import { isLatexMaterial } from "../symbolBox.ts";
import type { TextStyleMaterial } from "../tex.ts";
import { createDesignerId } from "../ids.ts";
import type { CustomSymbolDesignerPlatform } from "../platform.ts";
import { customSymbolSvgMarkup } from "../svg.ts";
import type { CustomSymbolDefinition } from "../symbolTypes.ts";
import { createSystemFontGlyphAssetAsync, systemFontPresetById, type SystemMathFontId } from "../systemGlyphs.ts";
import { customSymbolTexExport } from "../tex.ts";
import { CodePanel } from "./CodePanel.tsx";
import { DesignerCanvas, type DesignerTool } from "./DesignerCanvas.tsx";
import { InspectorPanel, type LayerAction } from "./InspectorPanel.tsx";
import {
  MATH_AXIS,
  baselineOf,
  contentBounds,
  layerWorldBounds,
  translated,
  withAnchorPosition,
  withCenteredOrigin,
} from "./layerMath.ts";
import { LayersPanel } from "./LayersPanel.tsx";
import { MaterialsPanel, normalizeMaterialLatex, shapeLabel } from "./MaterialsPanel.tsx";
import { RegistrationPanel, commandProblem } from "./RegistrationPanel.tsx";
import { useDesignerHistory } from "./useDesignerHistory.ts";
import "./designer.css";

export interface CustomSymbolDesignerProps {
  open: boolean;
  language: "cn" | "en";
  platform: CustomSymbolDesignerPlatform;
  onClose: () => void;
}

const REFERENCES = [
  { label: "x", latex: "x" },
  { label: "α", latex: "\\alpha" },
  { label: "∂", latex: "\\partial" },
  { label: "=", latex: "=" },
  { label: "∫", latex: "\\displaystyle\\int" },
  { label: "Σ", latex: "\\displaystyle\\sum" },
] as const;

type SideTab = "properties" | "layers" | "register";

function isTypingTarget(target: EventTarget | null) {
  return target instanceof HTMLElement && Boolean(target.closest("input, textarea, select, [contenteditable='true']"));
}

function clone<T>(value: T): T {
  return typeof structuredClone === "function" ? structuredClone(value) : (JSON.parse(JSON.stringify(value)) as T);
}

/** A new layer on the baseline, centred over what is already there (a shape: on the math axis). */
function placeNewLayer(layer: CustomSymbolDesignerLayer, document: CustomSymbolDesignerDocument) {
  const baseline = baselineOf(document);
  const content = contentBounds(document);
  if (layer.kind === "glyph") {
    const width = layer.asset.metrics.widthEm * 1000;
    const x = content ? content.x + content.width / 2 - width / 2 : 0;
    return withAnchorPosition(layer, baseline, { xEm: x / 1000, raiseEm: 0 });
  }
  const box = layerWorldBounds(layer);
  const cx = content ? content.x + content.width / 2 : box.width / 2;
  return translated(layer, cx - (box.x + box.width / 2), baseline - MATH_AXIS - (box.y + box.height / 2));
}

export function CustomSymbolDesigner({ open, language, platform, onClose }: CustomSymbolDesignerProps) {
  const isEn = language === "en";
  const t = (zh: string, en: string) => (isEn ? en : zh);
  const history = useDesignerHistory(createEmptyCustomSymbolDesignerDocument);
  const { document: doc, change } = history;
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [tool, setTool] = useState<DesignerTool>("select");
  const [eraserSize, setEraserSize] = useState(80);
  const [sideTab, setSideTab] = useState<SideTab>("properties");
  const [referenceLatex, setReferenceLatex] = useState<string>("x");
  const [showReference, setShowReference] = useState(true);
  const [referenceAsset, setReferenceAsset] = useState<CustomSymbolGlyphAsset | null>(null);
  const [busyKey, setBusyKey] = useState<string | null>(null);
  const [materialError, setMaterialError] = useState("");
  const [materialStatus, setMaterialStatus] = useState("");
  const [registerStatus, setRegisterStatus] = useState<{ kind: "idle" | "success" | "error"; message: string }>({ kind: "idle", message: "" });
  const [savedFingerprint, setSavedFingerprint] = useState("");
  const [libraryOpen, setLibraryOpen] = useState(false);
  const [pendingDelete, setPendingDelete] = useState<string | null>(null);
  const [fitNonce, setFitNonce] = useState(0);
  const revision = platform.library.useRevision();
  const symbols = useMemo(() => platform.library.read().symbols, [platform, revision]);
  const baseline = baselineOf(doc);
  const selected = doc.layers.find((layer) => layer.id === selectedId) ?? null;
  const libraryRef = useRef<HTMLDivElement>(null);

  // the registered box, measured from the drawing (as registration crops it)
  // each LaTeX material's italic correction and, for a display operator, its
  // text size, measured once with the engine (an operator's limits use them)
  const [materials, setMaterials] = useState<Record<string, { italicEm: number; text?: TextStyleMaterial } | null>>({});
  useEffect(() => {
    const sources = new Set(
      doc.layers.filter((layer) => layer.kind === "glyph" && isLatexMaterial(layer)).map((layer) => (layer as { asset: { sourceLatex: string } }).asset.sourceLatex),
    );
    for (const source of sources) {
      if (source in materials) continue;
      setMaterials((current) => (source in current ? current : { ...current, [source]: null }));
      void (async () => {
        try {
          const italicEm = await measureMaterialItalic(platform.drawMaterial, source);
          const textSource = textStyleVariant(source);
          let text: TextStyleMaterial | undefined;
          if (textSource) {
            const drawing = await platform.drawMaterial(textSource);
            text = {
              metrics: { widthEm: drawing.right / 1000, ascentEm: drawing.ascent / 1000, descentEm: drawing.descent / 1000 },
              italicEm: await measureMaterialItalic(platform.drawMaterial, textSource),
            };
          }
          setMaterials((current) => ({ ...current, [source]: { italicEm, ...(text ? { text } : {}) } }));
        } catch {
          setMaterials((current) => ({ ...current, [source]: { italicEm: 0 } }));
        }
      })();
    }
  }, [doc.layers, materials, platform]);
  const italicOf = useCallback((source: string) => materials[source]?.italicEm, [materials]);
  const crop = useMemo(() => (open ? measureCustomSymbolBox(doc, italicOf) : null), [doc, open, italicOf]);
  const [mathFont, setMathFont] = useState<"cm" | "ntx" | undefined>(undefined);
  useEffect(() => {
    if (!open || !platform.mathFont) return;
    let cancelled = false;
    void platform.mathFont().then((font) => {
      if (!cancelled) setMathFont(font);
    });
    return () => {
      cancelled = true;
    };
  }, [open, platform]);
  const tex = useMemo(
    () => customSymbolTexExport(doc, crop, { textStyle: (source) => materials[source]?.text, mathFont }),
    [doc, crop, materials, mathFont],
  );
  const svg = useMemo(() => customSymbolSvgMarkup(doc, crop), [doc, crop]);
  const paintedShapes = useMemo(
    () => compileCustomSymbolDesignerArtwork(doc).filter((shape) => shape.operation !== "erase").length,
    [doc],
  );
  const fingerprint = useMemo(() => JSON.stringify(doc), [doc]);
  const registeredSymbol = doc.symbolId ? symbols.find((symbol) => symbol.id === doc.symbolId) ?? null : null;
  const dirty = Boolean(registeredSymbol && fingerprint !== savedFingerprint);
  const commandName = doc.command.trim().replace(/^\\/, "");
  const takenBy = symbols.find((symbol) => symbol.command === commandName && symbol.id !== doc.symbolId)?.name ?? null;
  const problem = commandProblem(doc.command, takenBy, isEn);
  const canRegister = !problem && paintedShapes > 0 && (!registeredSymbol || dirty);

  useEffect(() => {
    let cancelled = false;
    setReferenceAsset(null);
    void platform
      .drawMaterial(referenceLatex)
      .then((drawing) => {
        if (!cancelled) setReferenceAsset(glyphAssetFromEngineDrawing(referenceLatex, drawing));
      })
      .catch(() => undefined);
    return () => {
      cancelled = true;
    };
  }, [platform, referenceLatex]);

  useEffect(() => {
    if (open) setFitNonce((value) => value + 1);
  }, [open]);

  const updateLayer = useCallback(
    (id: string, update: (layer: CustomSymbolDesignerLayer) => CustomSymbolDesignerLayer, gesture?: string) => {
      change((current) => ({ ...current, layers: current.layers.map((layer) => (layer.id === id ? update(layer) : layer)) }), gesture);
    },
    [change],
  );

  const addLayer = useCallback(
    (layer: CustomSymbolDesignerLayer) => {
      const first = !doc.layers.some((item) => item.visible);
      change((current) => ({ ...current, layers: [...current.layers, placeNewLayer(layer, current)] }));
      setSelectedId(layer.id);
      setSideTab("properties");
      setTool("select");
      // the first material: frame it (later ones keep the view where the user put it)
      if (first) setFitNonce((value) => value + 1);
    },
    [change, doc.layers],
  );

  const addLatex = async (source: string) => {
    const latex = normalizeMaterialLatex(source);
    if (!latex) return;
    setBusyKey(`latex:${latex}`);
    setMaterialError("");
    setMaterialStatus("");
    try {
      const asset = glyphAssetFromEngineDrawing(latex, await platform.drawMaterial(latex));
      addLayer(glyphLayerFromAsset(asset, { id: createDesignerId(), name: source.trim() }));
    } catch (error) {
      setMaterialError(error instanceof Error ? error.message : String(error));
    } finally {
      setBusyKey(null);
    }
  };

  const addSystemGlyph = async (character: string, label: string, fontId: SystemMathFontId, italic: boolean) => {
    const font = systemFontPresetById(fontId);
    setBusyKey(`system:${font.id}:${character.codePointAt(0) ?? 0}`);
    setMaterialError("");
    setMaterialStatus("");
    try {
      const created = await createSystemFontGlyphAssetAsync({ character, font, italic }, platform.systemGlyphs);
      addLayer(glyphLayerFromAsset(created.asset, { id: createDesignerId(), name: `${character} · ${created.resolvedFamily}` }));
      setMaterialStatus(
        created.vectorOutline
          ? created.fallbackUsed
            ? `${created.requestedFamily} → ${created.resolvedFamily}`
            : created.resolvedFamily
          : t(`${label}：系统字体没有矢量轮廓，使用浏览器字形`, `${label}: no outline from the system, drawn by the browser`),
      );
    } catch (error) {
      setMaterialError(error instanceof Error ? error.message : String(error));
    } finally {
      setBusyKey(null);
    }
  };

  const addShape = (preset: Exclude<CustomSymbolGeometryPreset, "eraser">) => {
    addLayer(createCustomSymbolGeometryLayer(preset, doc.metrics, { name: shapeLabel(preset, isEn) }));
  };

  const addEraserStroke = (points: Array<{ x: number; y: number }>) => {
    const layer = createCustomSymbolEraserLayer(points, eraserSize, { name: t("橡皮擦轨迹", "Eraser stroke") });
    if (!layer) return;
    change((current) => ({ ...current, layers: [...current.layers, layer] }));
  };

  const deleteLayer = (id: string) => {
    change((current) => ({ ...current, layers: current.layers.filter((layer) => layer.id !== id) }));
    if (selectedId === id) setSelectedId(null);
  };

  const duplicateLayer = (id: string) => {
    const source = doc.layers.find((layer) => layer.id === id);
    if (!source) return;
    const copy = translated({ ...clone(source), id: createDesignerId(), name: `${source.name}${t(" 副本", " copy")}` }, 60, 0);
    change((current) => ({ ...current, layers: [...current.layers, copy] }));
    setSelectedId(copy.id);
  };

  const runAction = (action: LayerAction) => {
    if (!selected) return;
    const id = selected.id;
    switch (action) {
      case "duplicate":
        duplicateLayer(id);
        return;
      case "delete":
        deleteLayer(id);
        return;
      case "forward":
      case "backward":
        change((current) => {
          const index = current.layers.findIndex((layer) => layer.id === id);
          const next = index + (action === "forward" ? 1 : -1);
          if (index < 0 || next < 0 || next >= current.layers.length) return current;
          const layers = [...current.layers];
          [layers[index], layers[next]] = [layers[next], layers[index]];
          return { ...current, layers };
        });
        return;
      case "baseline":
        updateLayer(id, (layer) => {
          // the box's bottom (a glyph: its baseline) on the baseline
          if (layer.kind === "glyph") return withAnchorPosition(layer, baseline, { raiseEm: 0 });
          const box = layerWorldBounds(layer);
          return translated(layer, 0, baseline - (box.y + box.height));
        });
        return;
      case "center": {
        const content = contentBounds(doc, id);
        if (!content) return;
        updateLayer(id, (layer) => {
          const box = layerWorldBounds(layer);
          return translated(layer, content.x + content.width / 2 - (box.x + box.width / 2), 0);
        });
        return;
      }
      case "axis":
        updateLayer(id, (layer) => {
          const box = layerWorldBounds(layer);
          return translated(layer, 0, baseline - MATH_AXIS - (box.y + box.height / 2));
        });
        return;
      case "split-horizontal":
      case "split-vertical": {
        if (selected.kind !== "glyph") return;
        const slices = createGlyphLayerSlices(selected, action === "split-horizontal" ? "horizontal" : "vertical", 3).map(
          (slice) => withCenteredOrigin(slice),
        );
        change((current) => ({
          ...current,
          layers: [...current.layers.map((layer) => (layer.id === id ? { ...layer, visible: false } : layer)), ...slices],
        }));
        setSelectedId(slices[0]?.id ?? null);
        return;
      }
    }
  };

  const startDocument = (document: CustomSymbolDesignerDocument, saved: boolean) => {
    history.replace(document);
    setSelectedId(null);
    setSavedFingerprint(saved ? JSON.stringify(document) : "");
    setRegisterStatus({ kind: "idle", message: "" });
    setMaterialError("");
    setPendingDelete(null);
    setLibraryOpen(false);
    setTool("select");
    setFitNonce((value) => value + 1);
  };

  const openSymbol = (symbol: CustomSymbolDefinition) => {
    startDocument(restoreCustomSymbolDesignerDocument(symbol).document, true);
    setSideTab("register");
  };

  const duplicateSymbol = (symbol: CustomSymbolDefinition) => {
    const restored = restoreCustomSymbolDesignerDocument(symbol).document;
    const commands = new Set(symbols.map((item) => item.command));
    let command = `${symbol.command}copy`;
    while (commands.has(command)) command += "copy";
    startDocument({ ...restored, symbolId: null, name: `${symbol.name}${t(" 副本", " copy")}`, command }, false);
    setSideTab("register");
  };

  const deleteSymbol = (symbol: CustomSymbolDefinition) => {
    if (pendingDelete !== symbol.id) {
      setPendingDelete(symbol.id);
      return;
    }
    platform.library.remove(symbol.id);
    setPendingDelete(null);
    if (doc.symbolId === symbol.id) startDocument(createEmptyCustomSymbolDesignerDocument(), false);
  };

  const register = () => {
    try {
      // the command names a symbol that has no name of its own
      const name = doc.name.trim() || commandName;
      const fallback = doc.ommlFallback?.trim() ?? "";
      if (fallback) {
        if (platform.library.usesCustomSymbol(fallback)) {
          throw new Error(t("Word 替代不能使用其他自定义字符。", "The Word fallback cannot use another custom symbol."));
        }
        platform.validateWordFallback(fallback);
      }
      const existing = registeredSymbol;
      const id = existing?.id ?? doc.symbolId ?? createDesignerId();
      const next: CustomSymbolDesignerDocument = { ...doc, symbolId: id, command: commandName, name, ommlFallback: fallback || null };
      const definition = customSymbolDefinitionFromDesignerDocument(next, { id, createdAt: existing?.createdAt, italicOf });
      const library = existing ? platform.library.update(id, definition) : platform.library.register(definition);
      const saved = library.symbols.find((symbol) => symbol.id === id);
      if (!saved) throw new Error(t("无法读回刚注册的字符。", "The registered symbol could not be read back."));
      history.replace(next);
      setSavedFingerprint(JSON.stringify(next));
      setRegisterStatus({
        kind: "success",
        message: existing ? t(`已更新 \\${saved.command}`, `Updated \\${saved.command}`) : t(`已注册 \\${saved.command}`, `Registered \\${saved.command}`),
      });
    } catch (error) {
      setRegisterStatus({ kind: "error", message: error instanceof Error ? error.message : String(error) });
    }
  };

  // keyboard: tools, undo/redo, nudge, delete, duplicate; Esc steps back, then closes
  useEffect(() => {
    if (!open) return;
    const onKeyDown = (event: KeyboardEvent) => {
      const mod = event.metaKey || event.ctrlKey;
      if (event.key === "Escape") {
        event.preventDefault();
        if (libraryOpen) setLibraryOpen(false);
        else if (isTypingTarget(event.target)) (event.target as HTMLElement).blur();
        else if (tool !== "select") setTool("select");
        else if (selectedId) setSelectedId(null);
        else onClose();
        return;
      }
      if (isTypingTarget(event.target)) return;
      const key = event.key.toLowerCase();
      if (mod && key === "z") {
        event.preventDefault();
        if (event.shiftKey) history.redo();
        else history.undo();
        return;
      }
      if (mod && key === "y") {
        event.preventDefault();
        history.redo();
        return;
      }
      if (mod && key === "d" && selectedId) {
        event.preventDefault();
        duplicateLayer(selectedId);
        return;
      }
      if (mod || event.altKey) return;
      if ((event.key === "Delete" || event.key === "Backspace") && selectedId) {
        event.preventDefault();
        deleteLayer(selectedId);
        return;
      }
      if (event.key.startsWith("Arrow") && selected && !selected.locked) {
        event.preventDefault();
        const step = event.shiftKey ? 100 : 10;
        const dx = event.key === "ArrowLeft" ? -step : event.key === "ArrowRight" ? step : 0;
        const dy = event.key === "ArrowUp" ? -step : event.key === "ArrowDown" ? step : 0;
        updateLayer(selected.id, (layer) => translated(layer, dx, dy), `nudge:${selected.id}`);
        return;
      }
      if (key === "v") setTool("select");
      else if (key === "h") setTool("pan");
      else if (key === "e") setTool("eraser");
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  });

  useEffect(() => {
    if (!libraryOpen) return;
    const onPointerDown = (event: PointerEvent) => {
      if (!libraryRef.current?.contains(event.target as Node)) setLibraryOpen(false);
    };
    window.addEventListener("pointerdown", onPointerDown);
    return () => window.removeEventListener("pointerdown", onPointerDown);
  }, [libraryOpen]);

  if (!open) return null;
  const { MathPreview } = platform;
  const reference = showReference && referenceAsset ? { asset: referenceAsset, label: referenceLatex } : null;

  return createPortal(
    <div className="csd-backdrop" data-csd-backdrop onMouseDown={(event) => event.target === event.currentTarget && onClose()}>
      <section className="csd-dialog" role="dialog" aria-modal="true" aria-label={t("自定义字符设计器", "Custom symbol designer")} data-csd-dialog>
        <header className="csd-header">
          <strong>{t("自定义字符设计器", "Custom Symbol Designer")}</strong>
          <span className="csd-doc-title" data-csd-doc-title>
            {doc.command.trim() ? `\\${commandName}` : t("未命名草稿", "Untitled draft")}
            {registeredSymbol ? <em>{dirty ? t("有未保存修改", "unsaved changes") : t("已注册", "registered")}</em> : null}
          </span>
          <div className="csd-header-actions">
            <button type="button" title={t("撤销 (⌘/Ctrl+Z)", "Undo (Cmd/Ctrl+Z)")} disabled={!history.canUndo} onClick={history.undo} data-csd-undo>
              <Undo2 size={15} />
            </button>
            <button type="button" title={t("重做 (⌘/Ctrl+Shift+Z)", "Redo (Cmd/Ctrl+Shift+Z)")} disabled={!history.canRedo} onClick={history.redo} data-csd-redo>
              <Redo2 size={15} />
            </button>
            <span className="csd-divider" />
            <button type="button" title={t("新建空白字符", "New symbol")} onClick={() => startDocument(createEmptyCustomSymbolDesignerDocument(), false)} data-csd-new>
              <FilePlus2 size={15} />
            </button>
            <div className="csd-library" ref={libraryRef}>
              <button type="button" className="csd-library-button" aria-expanded={libraryOpen} onClick={() => setLibraryOpen(!libraryOpen)} data-csd-library-button>
                {t("已注册字符", "Registered")} <b>{symbols.length}</b> <ChevronDown size={13} />
              </button>
              {libraryOpen ? (
                <div className="csd-library-menu" data-csd-library>
                  {symbols.length ? (
                    symbols.map((symbol) => (
                      <div key={symbol.id} className={`csd-library-item${symbol.id === doc.symbolId ? " is-active" : ""}`} data-csd-symbol={symbol.command}>
                        <button type="button" className="csd-library-open" onClick={() => openSymbol(symbol)} data-csd-open-symbol>
                          <span className="csd-library-glyph"><MathPreview latex={`\\${symbol.command}`} fit /></span>
                          <span>
                            <strong>{symbol.name}</strong>
                            <code>{`\\${symbol.command}`}</code>
                          </span>
                        </button>
                        <button type="button" title={t("复制为新草稿", "Duplicate as a new draft")} onClick={() => duplicateSymbol(symbol)}>
                          <Copy size={13} />
                        </button>
                        <button
                          type="button"
                          className={pendingDelete === symbol.id ? "is-danger is-confirm" : "is-danger"}
                          title={pendingDelete === symbol.id ? t("再点一次删除", "Click again to delete") : t("删除", "Delete")}
                          onClick={() => deleteSymbol(symbol)}
                          data-csd-delete-symbol
                        >
                          {pendingDelete === symbol.id ? t("确认删除", "Delete") : <Trash2 size={13} />}
                        </button>
                      </div>
                    ))
                  ) : (
                    <p className="csd-empty">{t("还没有注册的字符。", "No registered symbols yet.")}</p>
                  )}
                  {pendingDelete ? (
                    <p className="csd-message is-error">
                      {t("删除后，已有公式中的该命令将无法显示。", "Formulas that use a deleted command can no longer show it.")}
                    </p>
                  ) : null}
                </div>
              ) : null}
            </div>
            <span className="csd-divider" />
            <button type="button" title={t("关闭 (Esc)", "Close (Esc)")} onClick={onClose} data-csd-close>
              <X size={16} />
            </button>
          </div>
        </header>

        <div className="csd-body">
          <MaterialsPanel
            platform={platform}
            isEn={isEn}
            revision={revision}
            busyKey={busyKey}
            error={materialError}
            status={materialStatus}
            onAddLatex={(source) => void addLatex(source)}
            onAddSystemGlyph={(character, label, fontId, italic) => void addSystemGlyph(character, label, fontId, italic)}
            onAddShape={addShape}
          />

          <main className="csd-stage">
            <div className="csd-toolbar" data-csd-toolbar>
              <div className="csd-segmented" role="toolbar">
                <button type="button" className={tool === "select" ? "is-active" : ""} aria-pressed={tool === "select"} title={t("选择与变换 (V)", "Select & transform (V)")} onClick={() => setTool("select")} data-csd-tool="select">
                  <MousePointer2 size={14} />
                </button>
                <button type="button" className={tool === "pan" ? "is-active" : ""} aria-pressed={tool === "pan"} title={t("平移画布 (H，或按住空格)", "Pan (H, or hold Space)")} onClick={() => setTool("pan")} data-csd-tool="pan">
                  <Hand size={14} />
                </button>
                <button type="button" className={tool === "eraser" ? "is-active" : ""} aria-pressed={tool === "eraser"} title={t("橡皮擦 (E)", "Eraser (E)")} onClick={() => setTool("eraser")} data-csd-tool="eraser">
                  <Eraser size={14} />
                </button>
              </div>
              {tool === "eraser" ? (
                <label className="csd-inline-field">
                  <span>{t("粗细", "Size")}</span>
                  <input type="range" min={10} max={400} step={5} value={eraserSize} data-csd-eraser-size onChange={(event) => setEraserSize(Number(event.currentTarget.value))} />
                  <output>{(eraserSize / 1000).toFixed(2)}em</output>
                </label>
              ) : null}
              <span className="csd-toolbar-spacer" />
              <label className="csd-inline-field">
                <span>{t("参考字符", "Reference")}</span>
                <select value={referenceLatex} data-csd-reference-select onChange={(event) => { setReferenceLatex(event.currentTarget.value); setShowReference(true); }}>
                  {REFERENCES.map((item) => (
                    <option key={item.latex} value={item.latex}>{item.label}</option>
                  ))}
                </select>
              </label>
              <button
                type="button"
                className={`csd-icon-button${showReference ? " is-active" : ""}`}
                aria-pressed={showReference}
                title={showReference ? t("隐藏参考字符", "Hide the reference") : t("显示参考字符", "Show the reference")}
                onClick={() => setShowReference(!showReference)}
                data-csd-reference-toggle
              >
                {showReference ? <Eye size={14} /> : <EyeOff size={14} />}
              </button>
            </div>
            <DesignerCanvas
              document={doc}
              selectedId={selectedId}
              tool={tool}
              eraserSize={eraserSize}
              reference={reference}
              crop={crop}
              fitKey={`${doc.symbolId ?? "draft"}:${fitNonce}`}
              isEn={isEn}
              onSelect={(id) => {
                setSelectedId(id);
                if (id) setSideTab("properties");
              }}
              onTransform={(id, transform, gesture) => updateLayer(id, (layer) => ({ ...layer, transform }), gesture)}
              onGestureEnd={history.endGesture}
              onEraserStroke={addEraserStroke}
            />
            <CodePanel tex={crop ? tex.code : ""} svg={svg} omissions={tex.omissions} isEn={isEn} />
          </main>

          <aside className="csd-panel csd-side">
            <div className="csd-tabs" role="tablist">
              {([
                ["properties", t("属性", "Properties")],
                ["layers", `${t("图层", "Layers")} ${doc.layers.length}`],
                ["register", t("注册", "Register")],
              ] as Array<[SideTab, string]>).map(([id, label]) => (
                <button key={id} type="button" role="tab" aria-selected={sideTab === id} className={sideTab === id ? "is-active" : ""} onClick={() => setSideTab(id)} data-csd-side-tab={id}>
                  {label}
                </button>
              ))}
            </div>
            <div className="csd-side-body">
              {sideTab === "properties" ? (
                selected ? (
                  <InspectorPanel
                    layer={selected}
                    baseline={baseline}
                    isEn={isEn}
                    onChange={(update, gesture) => updateLayer(selected.id, update, gesture)}
                    onAction={runAction}
                  />
                ) : (
                  <p className="csd-empty">
                    {doc.layers.length
                      ? t("在画布或图层列表中选择一个图层。", "Select a layer on the canvas or in the layer list.")
                      : t("从左侧加入符号、字体字符或图形，在画布上组合成新字符。", "Add symbols, font glyphs or shapes from the left and combine them on the canvas.")}
                  </p>
                )
              ) : null}
              {sideTab === "layers" ? (
                <LayersPanel
                  document={doc}
                  selectedId={selectedId}
                  isEn={isEn}
                  platform={platform}
                  onSelect={(id) => setSelectedId(id)}
                  onChange={(id, update) => updateLayer(id, update)}
                />
              ) : null}
              {sideTab === "register" ? (
                <RegistrationPanel
                  document={doc}
                  isEn={isEn}
                  registered={Boolean(registeredSymbol)}
                  dirty={dirty}
                  canRegister={canRegister}
                  problem={problem}
                  box={crop}
                  status={registerStatus}
                  onChange={(patch) => change((current) => ({ ...current, ...patch }), `meta:${Object.keys(patch).join(",")}`)}
                  onRegister={register}
                />
              ) : null}
            </div>
          </aside>
        </div>
      </section>
    </div>,
    document.body,
  );
}
