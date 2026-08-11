import { useEffect, useMemo, useState } from "react";
import { createPortal } from "react-dom";
import {
  ArrowDown,
  ArrowUp,
  Copy,
  Eye,
  EyeOff,
  Eraser,
  FileCode2,
  Grip,
  Lock,
  Plus,
  RotateCcw,
  Trash2,
  Unlock,
  X,
} from "lucide-react";
import { MathPreview } from "./MathPreview";
import { CustomSymbolDesignerCanvas } from "./CustomSymbolDesignerCanvas";
import { compileLatexGlyphAsset } from "../math/customSymbolGlyphCompiler";
import {
  compileCustomSymbolDesignerArtwork,
  customSymbolDefinitionFromDesignerDocument,
  glyphLayerFromAsset,
} from "../math/customSymbolDesignerCompiler";
import {
  CUSTOM_SYMBOL_GEOMETRY_PRESETS,
  createCustomSymbolEraserLayer,
  createCustomSymbolGeometryLayer,
  createGlyphLayerSlices,
  customSymbolGeometryProperties,
  glyphLayerFullClip,
  updateCustomSymbolGeometryLayer,
  type CustomSymbolGeometryPreset,
} from "../math/customSymbolDesignerGeometry";
import {
  createEmptyCustomSymbolDesignerDocument,
  type CustomSymbolDesignerDocument,
  type CustomSymbolDesignerLayer,
} from "../math/customSymbolDesignerTypes";
import { createUuid } from "../runtime/browserCompatibility";
import {
  registerCustomSymbolSafely,
  updateCustomSymbolSafely,
} from "../math/customSymbolRegistration";
import {
  deleteCustomSymbol,
  readCustomSymbolLibrary,
} from "../math/customSymbolRegistry";
import { useCustomSymbolRevision } from "../math/customSymbolReact";
import { restoreCustomSymbolDesignerDocument } from "../math/customSymbolDesignerArchive";
import type {
  CustomSymbolDefinition,
  CustomSymbolLimitsBehavior,
  CustomSymbolMathRole,
  CustomSymbolMetrics,
} from "../math/customSymbolTypes";

interface Props {
  open: boolean;
  language: "cn" | "en";
  onClose: () => void;
}

const quickMaterials = [
  "\\partial",
  "\\alpha",
  "\\int",
  "\\sum",
  "\\rightarrow",
  "\\nabla",
  "\\infty",
  "\\hbar",
] as const;

const customSymbolMathRoles: readonly CustomSymbolMathRole[] = [
  "ordinary",
  "binary",
  "relation",
  "operator",
  "open",
  "close",
  "punctuation",
];

const customSymbolLimitsBehaviors: readonly CustomSymbolLimitsBehavior[] = [
  "auto",
  "limits",
  "nolimits",
];

const referencePresets = [
  { label: "α", latex: "\\alpha" },
  { label: "x", latex: "x" },
  { label: "∫", latex: "\\displaystyle\\int" },
  { label: "∮", latex: "\\displaystyle\\oint" },
  { label: "Σ", latex: "\\displaystyle\\sum" },
  { label: "Π", latex: "\\displaystyle\\prod" },
  { label: "⋃", latex: "\\displaystyle\\bigcup" },
] as const;

const designerLargeOperatorCommands = new Set([
  "int",
  "iint",
  "iiint",
  "oint",
  "oiint",
  "oiiint",
  "sum",
  "prod",
  "coprod",
  "bigcup",
  "bigcap",
  "bigvee",
  "bigwedge",
]);

function normalizeDesignerMaterialLatex(source: string) {
  const trimmed = source.trim();
  if (/^\\(?:display|text|script|scriptscript)style\b/.test(trimmed)) return trimmed;
  const match = /^\\([A-Za-z]+)$/.exec(trimmed);
  if (!match || !designerLargeOperatorCommands.has(match[1])) return trimmed;
  return `\\displaystyle${trimmed}`;
}

const standardDesignerMetrics: CustomSymbolMetrics = {
  widthEm: 3.2,
  ascentEm: 3,
  descentEm: 1.5,
};

function paddedGlyphMetrics(
  metrics: CustomSymbolMetrics,
  marginEm = 0.65,
): CustomSymbolMetrics {
  return {
    widthEm: Math.max(standardDesignerMetrics.widthEm, metrics.widthEm + marginEm * 2),
    ascentEm: Math.max(standardDesignerMetrics.ascentEm, metrics.ascentEm + marginEm),
    descentEm: Math.max(standardDesignerMetrics.descentEm, metrics.descentEm + marginEm),
  };
}

function expandMetrics(
  current: CustomSymbolMetrics,
  required: CustomSymbolMetrics,
): CustomSymbolMetrics {
  return {
    widthEm: Math.max(current.widthEm, required.widthEm),
    ascentEm: Math.max(current.ascentEm, required.ascentEm),
    descentEm: Math.max(current.descentEm, required.descentEm),
  };
}

function cloneLayer(layer: CustomSymbolDesignerLayer): CustomSymbolDesignerLayer {
  if (typeof structuredClone === "function") return structuredClone(layer);
  return JSON.parse(JSON.stringify(layer)) as CustomSymbolDesignerLayer;
}

function NumericField({
  label,
  value,
  step = 1,
  min,
  max,
  field,
  onChange,
}: {
  label: string;
  value: number;
  step?: number;
  min?: number;
  max?: number;
  field: string;
  onChange: (value: number) => void;
}) {
  return (
    <label className="custom-symbol-designer-number-field">
      <span>{label}</span>
      <input
        type="number"
        value={Number(value.toFixed(3))}
        step={step}
        min={min}
        max={max}
        data-designer-field={field}
        onChange={(event) => {
          const next = Number(event.currentTarget.value);
          if (Number.isFinite(next)) onChange(next);
        }}
      />
    </label>
  );
}

export function CustomSymbolDesignerDialog({ open, language, onClose }: Props) {
  const isEn = language === "en";
  const [documentState, setDocumentState] = useState<CustomSymbolDesignerDocument>(
    createEmptyCustomSymbolDesignerDocument,
  );
  const [selectedLayerId, setSelectedLayerId] = useState<string | null>(null);
  const [materialLatex, setMaterialLatex] = useState("\\partial");
  const [materialError, setMaterialError] = useState("");
  const [showReference, setShowReference] = useState(true);
  const [referenceLatex, setReferenceLatex] = useState<string>("\\alpha");
  const [eraserMode, setEraserMode] = useState(false);
  const [eraserSize, setEraserSize] = useState(40);
  const referencePreset =
    referencePresets.find((preset) => preset.latex === referenceLatex) ??
    referencePresets[0];
  const referenceAsset = useMemo(() => {
    try {
      return compileLatexGlyphAsset(referenceLatex);
    } catch {
      return null;
    }
  }, [referenceLatex]);
  const [registrationState, setRegistrationState] = useState<{
    kind: "idle" | "success" | "error";
    message: string;
  }>({ kind: "idle", message: "" });
  const [savedRegistrationFingerprint, setSavedRegistrationFingerprint] =
    useState("");
  const [designerSourceMode, setDesignerSourceMode] = useState<
    "editable" | "flattened-legacy" | null
  >(null);
  const [pendingDeleteSymbolId, setPendingDeleteSymbolId] = useState<string | null>(
    null,
  );
  const customSymbolRevision = useCustomSymbolRevision();
  const registeredSymbols = useMemo(
    () => readCustomSymbolLibrary().symbols,
    [customSymbolRevision],
  );
  const selectedLayer =
    documentState.layers.find((layer) => layer.id === selectedLayerId) ?? null;
  const selectedGeometryProperties =
    selectedLayer?.kind === "geometry"
      ? customSymbolGeometryProperties(selectedLayer)
      : null;
  const shapeCount = useMemo(
    () =>
      compileCustomSymbolDesignerArtwork(documentState).filter(
        (shape) => shape.operation !== "erase",
      ).length,
    [documentState],
  );
  const registrationFingerprint = useMemo(
    () => JSON.stringify(documentState),
    [documentState],
  );
  const registrationDirty = Boolean(
    documentState.symbolId &&
      savedRegistrationFingerprint &&
      registrationFingerprint !== savedRegistrationFingerprint,
  );
  const registeredSymbol = documentState.symbolId
    ? registeredSymbols.find((symbol) => symbol.id === documentState.symbolId) ?? null
    : null;

  useEffect(() => {
    if (!open) return;
    const keydown = (event: KeyboardEvent) => {
      if (event.key === "Escape") onClose();
    };
    window.addEventListener("keydown", keydown);
    return () => window.removeEventListener("keydown", keydown);
  }, [onClose, open]);

  const updateLayer = (
    id: string,
    update: (layer: CustomSymbolDesignerLayer) => CustomSymbolDesignerLayer,
  ) => {
    setDocumentState((current) => ({
      ...current,
      layers: current.layers.map((layer) => (layer.id === id ? update(layer) : layer)),
    }));
  };

  const addMaterial = (source = materialLatex) => {
    try {
      const normalizedSource = normalizeDesignerMaterialLatex(source);
      const asset = compileLatexGlyphAsset(normalizedSource);
      const layer = glyphLayerFromAsset(asset, { id: createUuid(), name: source.trim() });
      setDocumentState((current) => ({
        ...current,
        metrics: expandMetrics(current.metrics, paddedGlyphMetrics(asset.metrics, 0.5)),
        layers: [...current.layers, layer],
      }));
      setSelectedLayerId(layer.id);
      setMaterialLatex(source);
      setMaterialError("");
    } catch (error) {
      setMaterialError(error instanceof Error ? error.message : String(error));
    }
  };

  const geometryLabel = (preset: CustomSymbolGeometryPreset) => {
    const labels = isEn
      ? {
          line: "Line",
          circle: "Circle",
          ellipse: "Ellipse",
          rect: "Rectangle",
          triangle: "Triangle",
          arrow: "Arrow",
          arc: "Arc",
          eraser: "Eraser",
        }
      : {
          line: "线段",
          circle: "圆",
          ellipse: "椭圆",
          rect: "矩形",
          triangle: "三角形",
          arrow: "箭头",
          arc: "圆弧",
          eraser: "橡皮擦",
        };
    return labels[preset];
  };

  const addGeometry = (
    preset: Exclude<CustomSymbolGeometryPreset, "eraser">,
  ) => {
    const layer = createCustomSymbolGeometryLayer(preset, documentState.metrics, {
      name: geometryLabel(preset),
    });
    setDocumentState((current) => ({
      ...current,
      layers: [...current.layers, layer],
    }));
    setSelectedLayerId(layer.id);
    setEraserMode(false);
  };

  const addEraserStroke = (points: Array<{ x: number; y: number }>) => {
    const layer = createCustomSymbolEraserLayer(points, eraserSize, {
      name: isEn ? "Eraser stroke" : "橡皮擦轨迹",
    });
    if (!layer) return;
    setDocumentState((current) => ({
      ...current,
      layers: [...current.layers, layer],
    }));
    setSelectedLayerId(layer.id);
  };

  const updateSelectedGeometry = (
    patch: Parameters<typeof updateCustomSymbolGeometryLayer>[1],
  ) => {
    if (!selectedLayer || selectedLayer.kind !== "geometry") return;
    updateLayer(selectedLayer.id, (layer) =>
      layer.kind === "geometry"
        ? updateCustomSymbolGeometryLayer(layer, patch)
        : layer,
    );
  };

  const applyMetricPreset = (
    preset: "standard" | "large" | "extra" | "reference",
  ) => {
    const metrics =
      preset === "reference" && referenceAsset
        ? paddedGlyphMetrics(referenceAsset.metrics)
        : preset === "standard"
          ? { ...standardDesignerMetrics }
          : preset === "large"
            ? { widthEm: 4.5, ascentEm: 4, descentEm: 2 }
            : { widthEm: 6.5, ascentEm: 5.5, descentEm: 3 };
    setDocumentState((current) => ({ ...current, metrics }));
  };

  const updateSelectedClip = (
    patch: Partial<{ x: number; y: number; width: number; height: number }>,
  ) => {
    if (!selectedLayer || selectedLayer.kind !== "glyph") return;
    updateLayer(selectedLayer.id, (layer) => {
      if (layer.kind !== "glyph") return layer;
      const current = layer.clipRect ?? glyphLayerFullClip(layer);
      const next = { ...current, ...patch };
      return {
        ...layer,
        clipRect: {
          x: next.x,
          y: next.y,
          width: Math.max(1, next.width),
          height: Math.max(1, next.height),
        },
      };
    });
  };

  const applyClipPreset = (
    preset: "full" | "top" | "middle" | "bottom" | "left" | "center" | "right",
  ) => {
    if (!selectedLayer || selectedLayer.kind !== "glyph") return;
    const full = glyphLayerFullClip(selectedLayer);
    const thirdWidth = full.width / 3;
    const thirdHeight = full.height / 3;
    const clip =
      preset === "full"
        ? null
        : preset === "top"
          ? { ...full, height: thirdHeight }
          : preset === "middle"
            ? { ...full, y: thirdHeight, height: thirdHeight }
            : preset === "bottom"
              ? { ...full, y: thirdHeight * 2, height: thirdHeight }
              : preset === "left"
                ? { ...full, width: thirdWidth }
                : preset === "center"
                  ? { ...full, x: thirdWidth, width: thirdWidth }
                  : { ...full, x: thirdWidth * 2, width: thirdWidth };
    updateLayer(selectedLayer.id, (layer) =>
      layer.kind === "glyph" ? { ...layer, clipRect: clip } : layer,
    );
  };

  const splitSelectedGlyph = (orientation: "horizontal" | "vertical") => {
    if (!selectedLayer || selectedLayer.kind !== "glyph") return;
    const slices = createGlyphLayerSlices(selectedLayer, orientation, 3);
    setDocumentState((current) => ({
      ...current,
      layers: [
        ...current.layers.map((layer) =>
          layer.id === selectedLayer.id ? { ...layer, visible: false } : layer,
        ),
        ...slices,
      ],
    }));
    setSelectedLayerId(slices[0]?.id ?? null);
  };

  const mathRoleLabel = (role: CustomSymbolMathRole) => {
    const labels = isEn
      ? {
          ordinary: "Ordinary",
          binary: "Binary operator",
          relation: "Relation",
          operator: "Large operator",
          open: "Opening delimiter",
          close: "Closing delimiter",
          punctuation: "Punctuation",
        }
      : {
          ordinary: "普通字符",
          binary: "二元运算符",
          relation: "关系符号",
          operator: "大型算子",
          open: "左定界符",
          close: "右定界符",
          punctuation: "标点",
        };
    return labels[role];
  };

  const limitsLabel = (behavior: CustomSymbolLimitsBehavior) => {
    if (isEn) {
      return behavior === "auto"
        ? "Automatic"
        : behavior === "limits"
          ? "Limits above/below"
          : "Limits beside";
    }
    return behavior === "auto"
      ? "自动"
      : behavior === "limits"
        ? "上下放置上下限"
        : "侧边放置上下限";
  };

  const saveRegisteredSymbol = () => {
    try {
      const name = documentState.name.trim();
      const command = documentState.command.trim().replace(/^\\/, "");
      if (!name) throw new Error(isEn ? "Enter a symbol name." : "请输入字符名称。");
      if (!command) throw new Error(isEn ? "Enter a LaTeX command name." : "请输入 LaTeX 命令名。");
      if (shapeCount === 0) {
        throw new Error(
          isEn
            ? "The symbol must contain at least one visible vector shape."
            : "字符至少需要一个可见的矢量单元。",
        );
      }

      const existing = documentState.symbolId
        ? readCustomSymbolLibrary().symbols.find(
            (symbol) => symbol.id === documentState.symbolId,
          ) ?? null
        : null;
      const id = existing?.id ?? documentState.symbolId ?? createUuid();
      const nextDocument: CustomSymbolDesignerDocument = {
        ...documentState,
        symbolId: id,
        command,
        name,
        ommlFallback: null,
      };
      const definition = customSymbolDefinitionFromDesignerDocument(nextDocument, {
        id,
        createdAt: existing?.createdAt,
      });
      const library = existing
        ? updateCustomSymbolSafely(id, definition)
        : registerCustomSymbolSafely(definition);
      const saved = library.symbols.find((symbol) => symbol.id === id);
      if (!saved) {
        throw new Error(
          isEn
            ? "VisualTeX could not read the registered symbol back."
            : "VisualTeX 无法重新读取刚注册的字符。",
        );
      }
      const savedDocument: CustomSymbolDesignerDocument = {
        ...nextDocument,
        symbolId: saved.id,
        command: saved.command,
        name: saved.name,
        role: saved.role,
        limitsBehavior: saved.limitsBehavior,
        metrics: { ...nextDocument.metrics },
        ommlFallback: null,
      };
      setDocumentState(savedDocument);
      setSavedRegistrationFingerprint(JSON.stringify(savedDocument));
      setDesignerSourceMode("editable");
      setRegistrationState({
        kind: "success",
        message: existing
          ? isEn
            ? `Updated \\${saved.command}`
            : `已更新 \\${saved.command}`
          : isEn
            ? `Registered \\${saved.command}`
            : `已注册 \\${saved.command}`,
      });
    } catch (error) {
      setRegistrationState({
        kind: "error",
        message: error instanceof Error ? error.message : String(error),
      });
    }
  };

  const moveLayerOrder = (id: string, delta: -1 | 1) => {
    setDocumentState((current) => {
      const index = current.layers.findIndex((layer) => layer.id === id);
      const next = index + delta;
      if (index < 0 || next < 0 || next >= current.layers.length) return current;
      const layers = [...current.layers];
      [layers[index], layers[next]] = [layers[next], layers[index]];
      return { ...current, layers };
    });
  };

  const duplicateLayer = (id: string) => {
    const source = documentState.layers.find((layer) => layer.id === id);
    if (!source) return;
    const duplicate = cloneLayer(source);
    duplicate.id = createUuid();
    duplicate.name += isEn ? " copy" : " 副本";
    duplicate.transform = {
      ...duplicate.transform,
      translateX: (duplicate.transform.translateX ?? 0) + 45,
      translateY: (duplicate.transform.translateY ?? 0) + 45,
    };
    setDocumentState((current) => ({
      ...current,
      layers: [...current.layers, duplicate],
    }));
    setSelectedLayerId(duplicate.id);
  };

  const deleteLayer = (id: string) => {
    setDocumentState((current) => ({
      ...current,
      layers: current.layers.filter((layer) => layer.id !== id),
    }));
    if (selectedLayerId === id) setSelectedLayerId(null);
  };

  useEffect(() => {
    if (!open || !selectedLayerId) return;
    const handleDeleteKey = (event: KeyboardEvent) => {
      if (event.key !== "Delete" && event.key !== "Backspace") return;
      const target = event.target;
      if (
        target instanceof HTMLInputElement ||
        target instanceof HTMLTextAreaElement ||
        target instanceof HTMLSelectElement ||
        (target instanceof HTMLElement && target.isContentEditable)
      ) {
        return;
      }
      event.preventDefault();
      deleteLayer(selectedLayerId);
    };
    window.addEventListener("keydown", handleDeleteKey);
    return () => window.removeEventListener("keydown", handleDeleteKey);
  }, [open, selectedLayerId]);

  const updateSelectedTransform = (key: string, value: number) => {
    if (!selectedLayer) return;
    updateLayer(selectedLayer.id, (layer) => ({
      ...layer,
      transform: { ...layer.transform, [key]: value },
    }));
  };

  const openRegisteredSymbol = (symbol: CustomSymbolDefinition) => {
    const restored = restoreCustomSymbolDesignerDocument(symbol);
    const webDocument = { ...restored.document, ommlFallback: null };
    setDocumentState(webDocument);
    setSelectedLayerId(webDocument.layers[0]?.id ?? null);
    setMaterialError("");
    setRegistrationState({ kind: "idle", message: "" });
    setSavedRegistrationFingerprint(JSON.stringify(webDocument));
    setDesignerSourceMode(restored.sourceMode);
    setPendingDeleteSymbolId(null);
    setEraserMode(false);
  };

  const duplicateRegisteredSymbol = (symbol: CustomSymbolDefinition) => {
    const restored = restoreCustomSymbolDesignerDocument(symbol);
    const commands = new Set(registeredSymbols.map((item) => item.command));
    let command = `${symbol.command}copy`;
    while (commands.has(command)) command += "copy";
    const duplicate: CustomSymbolDesignerDocument = {
      ...restored.document,
      symbolId: null,
      name: `${symbol.name}${isEn ? " copy" : " 副本"}`,
      command,
      ommlFallback: null,
    };
    setDocumentState(duplicate);
    setSelectedLayerId(duplicate.layers[0]?.id ?? null);
    setMaterialError("");
    setRegistrationState({ kind: "idle", message: "" });
    setSavedRegistrationFingerprint("");
    setDesignerSourceMode(restored.sourceMode);
    setPendingDeleteSymbolId(null);
    setEraserMode(false);
  };

  const requestDeleteRegisteredSymbol = (symbol: CustomSymbolDefinition) => {
    if (pendingDeleteSymbolId !== symbol.id) {
      setPendingDeleteSymbolId(symbol.id);
      return;
    }
    deleteCustomSymbol(symbol.id);
    setPendingDeleteSymbolId(null);
    if (documentState.symbolId === symbol.id) {
      setDocumentState(createEmptyCustomSymbolDesignerDocument());
      setSelectedLayerId(null);
      setSavedRegistrationFingerprint("");
      setDesignerSourceMode(null);
      setRegistrationState({ kind: "idle", message: "" });
    }
  };

  const reset = () => {
    setDocumentState(createEmptyCustomSymbolDesignerDocument());
    setSelectedLayerId(null);
    setMaterialLatex("\\partial");
    setMaterialError("");
    setRegistrationState({ kind: "idle", message: "" });
    setSavedRegistrationFingerprint("");
    setDesignerSourceMode(null);
    setPendingDeleteSymbolId(null);
    setShowReference(true);
    setReferenceLatex("\\alpha");
    setEraserMode(false);
    setEraserSize(90);
  };

  if (!open) return null;
  const widthEm = documentState.metrics.widthEm;
  const heightEm = documentState.metrics.ascentEm + documentState.metrics.descentEm;

  return createPortal(
    <div
      className="modal-backdrop custom-symbol-designer-backdrop"
      data-custom-symbol-designer-backdrop
      onMouseDown={(event) => {
        if (event.target === event.currentTarget) onClose();
      }}
    >
      <section
        className="custom-symbol-designer-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={isEn ? "Custom symbol designer" : "自定义字符设计器"}
        data-custom-symbol-designer
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="custom-symbol-designer-header">
          <div>
            <strong>{isEn ? "Custom Symbol Designer" : "自定义字符设计器"}</strong>
            <span>
              {isEn
                ? "Compose editable LaTeX vector glyphs on a mathematical baseline."
                : "在数学基线上重新组合可编辑的 LaTeX 矢量字形。"}
            </span>
          </div>
          <div>
            <button
              className="icon-button compact"
              type="button"
              data-reset-custom-symbol-designer
              onClick={reset}
              aria-label={isEn ? "Reset" : "重置"}
            >
              <RotateCcw size={15} />
            </button>
            <button
              className="icon-button compact"
              type="button"
              onClick={onClose}
              aria-label={isEn ? "Close" : "关闭"}
            >
              <X size={16} />
            </button>
          </div>
        </header>

        <div className="custom-symbol-designer-body">
          <aside className="custom-symbol-designer-sidebar is-materials">
            <section className="custom-symbol-designer-panel is-registered-symbols">
              <header>
                <strong>{isEn ? "Registered symbols" : "已注册字符"}</strong>
                <span>{registeredSymbols.length}</span>
              </header>
              <div className="custom-symbol-registered-list" data-custom-symbol-registered-list>
                {registeredSymbols.map((symbol) => {
                  const active = documentState.symbolId === symbol.id;
                  const pendingDelete = pendingDeleteSymbolId === symbol.id;
                  return (
                    <div
                      key={symbol.id}
                      className={`custom-symbol-registered-item${active ? " is-active" : ""}`}
                      data-registered-custom-symbol={symbol.id}
                      data-registered-custom-symbol-command={symbol.command}
                    >
                      <button
                        type="button"
                        className="custom-symbol-registered-main"
                        data-edit-registered-custom-symbol
                        onClick={() => openRegisteredSymbol(symbol)}
                      >
                        <span className="custom-symbol-registered-glyph">
                          <MathPreview latex={`\\${symbol.command}`} staticLayout />
                        </span>
                        <span className="custom-symbol-registered-copy">
                          <strong>{symbol.name}</strong>
                          <code>{`\\${symbol.command}`}</code>
                        </span>
                      </button>
                      <div className="custom-symbol-registered-actions">
                        <button
                          type="button"
                          data-duplicate-registered-custom-symbol
                          onClick={() => duplicateRegisteredSymbol(symbol)}
                          title={isEn ? "Duplicate as draft" : "复制为新草稿"}
                        >
                          <Copy size={12} />
                        </button>
                        <button
                          type="button"
                          data-delete-registered-custom-symbol
                          className={pendingDelete ? "is-pending-delete" : ""}
                          onClick={() => requestDeleteRegisteredSymbol(symbol)}
                          title={
                            pendingDelete
                              ? isEn
                                ? "Click again below to confirm"
                                : "请在下方再次确认"
                              : isEn
                                ? "Delete registered symbol"
                                : "删除已注册字符"
                          }
                        >
                          <Trash2 size={12} />
                        </button>
                      </div>
                      {pendingDelete ? (
                        <div className="custom-symbol-delete-warning" data-custom-symbol-delete-warning>
                          <span>
                            {isEn
                              ? `Deleting \\${symbol.command} makes existing formulas using it unresolved.`
                              : `删除 \\${symbol.command} 后，已有公式中的这个命令会变成未解析状态。`}
                          </span>
                          <div>
                            <button
                              type="button"
                              data-confirm-delete-registered-custom-symbol
                              onClick={() => requestDeleteRegisteredSymbol(symbol)}
                            >
                              {isEn ? "Delete" : "确认删除"}
                            </button>
                            <button
                              type="button"
                              data-cancel-delete-registered-custom-symbol
                              onClick={() => setPendingDeleteSymbolId(null)}
                            >
                              {isEn ? "Cancel" : "取消"}
                            </button>
                          </div>
                        </div>
                      ) : null}
                    </div>
                  );
                })}
                {!registeredSymbols.length ? (
                  <div className="custom-symbol-registered-empty">
                    {isEn
                      ? "Registered symbols will appear here for reuse and editing."
                      : "注册后的字符会出现在这里，可再次编辑或复用。"}
                  </div>
                ) : null}
              </div>
            </section>

            <section className="custom-symbol-designer-panel">
              <header>
                <FileCode2 size={15} />
                <strong>{isEn ? "LaTeX material" : "LaTeX 素材"}</strong>
              </header>
              <div className="custom-symbol-material-input-row">
                <input
                  value={materialLatex}
                  data-custom-symbol-material-input
                  spellCheck={false}
                  onChange={(event) => setMaterialLatex(event.currentTarget.value)}
                  onKeyDown={(event) => event.key === "Enter" && addMaterial()}
                />
                <button
                  type="button"
                  data-add-custom-symbol-material
                  onClick={() => addMaterial()}
                  disabled={!materialLatex.trim()}
                >
                  <Plus size={14} />
                  {isEn ? "Add" : "加入"}
                </button>
              </div>
              <div className="custom-symbol-quick-materials">
                {quickMaterials.map((latex) => (
                  <button type="button" key={latex} onClick={() => addMaterial(latex)} title={latex}>
                    <MathPreview latex={latex} staticLayout />
                  </button>
                ))}
              </div>
              <div className="custom-symbol-geometry-heading">
                {isEn ? "Geometry" : "几何图形"}
              </div>
              <div className="custom-symbol-geometry-materials">
                {CUSTOM_SYMBOL_GEOMETRY_PRESETS.map((preset) => (
                  <button
                    type="button"
                    key={preset}
                    data-add-custom-symbol-geometry={preset}
                    onClick={() => addGeometry(preset)}
                  >
                    <span className={`custom-symbol-geometry-icon is-${preset}`} aria-hidden="true" />
                    <span>{geometryLabel(preset)}</span>
                  </button>
                ))}
              </div>
              {materialError ? (
                <div className="custom-symbol-designer-error">{materialError}</div>
              ) : null}
            </section>

            <section className="custom-symbol-designer-panel is-layers">
              <header>
                <Grip size={15} />
                <strong>{isEn ? "Layers" : "图层"}</strong>
                <span>{documentState.layers.length}</span>
              </header>
              <div className="custom-symbol-layer-list" data-custom-symbol-layer-list>
                {[...documentState.layers].reverse().map((layer) => {
                  const index = documentState.layers.findIndex((item) => item.id === layer.id);
                  return (
                    <div
                      key={layer.id}
                      className={`custom-symbol-layer-item${selectedLayerId === layer.id ? " is-selected" : ""}`}
                      data-custom-symbol-layer={layer.id}
                      data-layer-kind={layer.kind}
                      data-layer-geometry-preset={
                        layer.kind === "geometry" ? layer.geometryPreset ?? "" : ""
                      }
                      data-layer-source-latex={layer.kind === "glyph" ? layer.asset.sourceLatex : ""}
                      data-layer-visible={layer.visible ? "true" : "false"}
                      data-layer-locked={layer.locked ? "true" : "false"}
                      onClick={() => setSelectedLayerId(layer.id)}
                    >
                      <button
                        type="button"
                        data-toggle-custom-symbol-layer-visibility
                        onClick={(event) => {
                          event.stopPropagation();
                          updateLayer(layer.id, (item) => ({ ...item, visible: !item.visible }));
                        }}
                      >
                        {layer.visible ? <Eye size={13} /> : <EyeOff size={13} />}
                      </button>
                      <div className="custom-symbol-layer-preview">
                        {layer.kind === "glyph" ? (
                          <MathPreview latex={layer.asset.sourceLatex} staticLayout />
                        ) : (
                          <span className={`custom-symbol-geometry-icon is-${layer.shape.kind}`} aria-hidden="true" />
                        )}
                      </div>
                      <div className="custom-symbol-layer-label">
                        <strong>{layer.name}</strong>
                        <span>{layer.kind === "glyph" ? layer.asset.sourceLatex : "Geometry"}</span>
                      </div>
                      <div className="custom-symbol-layer-actions">
                        <button
                          type="button"
                          data-toggle-custom-symbol-layer-lock
                          onClick={(event) => {
                            event.stopPropagation();
                            updateLayer(layer.id, (item) => ({ ...item, locked: !item.locked }));
                          }}
                        >
                          {layer.locked ? <Lock size={12} /> : <Unlock size={12} />}
                        </button>
                        <button
                          type="button"
                          data-move-custom-symbol-layer-up
                          disabled={index >= documentState.layers.length - 1}
                          onClick={(event) => {
                            event.stopPropagation();
                            moveLayerOrder(layer.id, 1);
                          }}
                        >
                          <ArrowUp size={12} />
                        </button>
                        <button
                          type="button"
                          data-move-custom-symbol-layer-down
                          disabled={index <= 0}
                          onClick={(event) => {
                            event.stopPropagation();
                            moveLayerOrder(layer.id, -1);
                          }}
                        >
                          <ArrowDown size={12} />
                        </button>
                      </div>
                    </div>
                  );
                })}
                {!documentState.layers.length ? (
                  <div className="custom-symbol-layer-empty">
                    {isEn ? "Add a LaTeX material to begin." : "先加入一个 LaTeX 素材。"}
                  </div>
                ) : null}
              </div>
            </section>
          </aside>

          <main className="custom-symbol-designer-stage-column">
            <div className="custom-symbol-designer-stage-toolbar">
              <span>
                {isEn ? "Output box" : "输出范围"} · {widthEm.toFixed(3)}em × {heightEm.toFixed(3)}em
              </span>
              <div className="custom-symbol-designer-stage-toolbar-actions">
                <span>{isEn ? "Dashed line = baseline" : "虚线 = 数学基线"}</span>
                <label className="custom-symbol-reference-selector">
                  <span>{isEn ? "Reference" : "参考"}</span>
                  <select
                    value={referenceLatex}
                    data-custom-symbol-reference-select
                    onChange={(event) => {
                      const nextLatex = event.currentTarget.value;
                      setReferenceLatex(nextLatex);
                      setShowReference(true);
                      try {
                        const nextAsset = compileLatexGlyphAsset(nextLatex);
                        const required = paddedGlyphMetrics(nextAsset.metrics);
                        setDocumentState((current) => ({
                          ...current,
                          metrics: expandMetrics(current.metrics, required),
                        }));
                      } catch {
                        // Built-in reference presets are validated by the renderer.
                      }
                    }}
                  >
                    {referencePresets.map((preset) => (
                      <option key={preset.latex} value={preset.latex}>
                        {preset.label}
                      </option>
                    ))}
                  </select>
                </label>
                <button
                  type="button"
                  className={showReference ? "is-active" : ""}
                  data-toggle-custom-symbol-reference
                  data-toggle-custom-symbol-reference-alpha
                  aria-pressed={showReference}
                  onClick={() => setShowReference((current) => !current)}
                  title={
                    isEn
                      ? "Show or hide the selected mathematical reference on the same baseline"
                      : "显示或隐藏与当前字符共用数学基线的参考符号"
                  }
                >
                  {showReference ? <Eye size={12} /> : <EyeOff size={12} />}
                  {showReference ? referencePreset.label : isEn ? "Reference" : "参考"}
                </button>
                <button
                  type="button"
                  className={`custom-symbol-eraser-tool${eraserMode ? " is-active" : ""}`}
                  data-custom-symbol-eraser-tool
                  aria-pressed={eraserMode}
                  onClick={() => setEraserMode((current) => !current)}
                  title={
                    isEn
                      ? "Erase unwanted portions with a true transparent vector cutout"
                      : "用真正的透明矢量擦除去掉不需要的部分"
                  }
                >
                  <Eraser size={12} />
                  {isEn ? "Eraser" : "橡皮擦"}
                </button>
                {eraserMode ? (
                  <label className="custom-symbol-eraser-size">
                    <span>{isEn ? "Size" : "粗细"}</span>
                    <input
                      type="range"
                      min="4"
                      max="240"
                      step="2"
                      value={Math.min(240, eraserSize)}
                      data-custom-symbol-eraser-size
                      onChange={(event) => setEraserSize(Number(event.currentTarget.value))}
                    />
                    <input
                      type="number"
                      min="4"
                      max="600"
                      step="2"
                      value={eraserSize}
                      data-custom-symbol-eraser-size-number
                      aria-label={isEn ? "Precise eraser size" : "精确橡皮擦粗细"}
                      onChange={(event) => {
                        const value = Number(event.currentTarget.value);
                        if (Number.isFinite(value)) {
                          setEraserSize(Math.max(4, Math.min(600, value)));
                        }
                      }}
                    />
                    <output>{(eraserSize / 1000).toFixed(3)}em</output>
                  </label>
                ) : null}
              </div>
            </div>
            <div className="custom-symbol-designer-stage" data-custom-symbol-stage>
              <CustomSymbolDesignerCanvas
                documentState={documentState}
                selectedLayerId={selectedLayerId}
                referenceAsset={referenceAsset}
                showReference={showReference}
                referenceLabel={referencePreset.label}
                eraserMode={eraserMode}
                eraserSize={eraserSize}
                isEn={isEn}
                onSelectLayer={setSelectedLayerId}
                onMoveLayer={(id, x, y) =>
                  updateLayer(id, (layer) => ({
                    ...layer,
                    transform: { ...layer.transform, translateX: x, translateY: y },
                  }))
                }
                onResizeLayer={(id, scaleX, scaleY) =>
                  updateLayer(id, (layer) => ({
                    ...layer,
                    transform: { ...layer.transform, scaleX, scaleY },
                  }))
                }
                onAddEraserStroke={addEraserStroke}
              />
            </div>
          </main>

          <aside className="custom-symbol-designer-sidebar is-inspector">
            {designerSourceMode === "flattened-legacy" ? (
              <div className="custom-symbol-designer-legacy-warning" data-custom-symbol-legacy-warning>
                {isEn
                  ? "This older symbol has no editable source archive. VisualTeX restored its compiled vector shapes, but the original LaTeX materials and non-destructive slice relationships cannot be recovered. Saving an update will create a new editable source archive from this flattened state."
                  : "这个旧字符没有可编辑源档。VisualTeX 已恢复其编译后的矢量图形，但原始 LaTeX 素材和非破坏分片关系无法恢复。再次保存后，会从当前扁平状态建立新的可编辑源档。"}
              </div>
            ) : null}
            <section className="custom-symbol-designer-panel">
              <header>
                <strong>{isEn ? "Canvas metrics" : "画布指标"}</strong>
              </header>
              <div className="custom-symbol-designer-number-grid">
                <NumericField
                  label={isEn ? "Width em" : "宽度 em"}
                  field="canvas-width"
                  value={widthEm}
                  step={0.05}
                  min={0.02}
                  max={12}
                  onChange={(value) =>
                    setDocumentState((current) => ({
                      ...current,
                      metrics: { ...current.metrics, widthEm: Math.max(0.02, value) },
                    }))
                  }
                />
                <NumericField
                  label={isEn ? "Ascent em" : "基线上方 em"}
                  field="canvas-ascent"
                  value={documentState.metrics.ascentEm}
                  step={0.05}
                  min={0.02}
                  max={12}
                  onChange={(value) =>
                    setDocumentState((current) => ({
                      ...current,
                      metrics: { ...current.metrics, ascentEm: Math.max(0.02, value) },
                    }))
                  }
                />
                <NumericField
                  label={isEn ? "Descent em" : "基线下方 em"}
                  field="canvas-descent"
                  value={documentState.metrics.descentEm}
                  step={0.05}
                  min={0}
                  max={12}
                  onChange={(value) =>
                    setDocumentState((current) => ({
                      ...current,
                      metrics: { ...current.metrics, descentEm: Math.max(0, value) },
                    }))
                  }
                />
              </div>
              <div className="custom-symbol-metric-presets" data-custom-symbol-metric-presets>
                <button type="button" data-metric-preset="standard" onClick={() => applyMetricPreset("standard")}>
                  {isEn ? "Standard" : "标准"}
                </button>
                <button
                  type="button"
                  data-metric-preset="reference"
                  onClick={() => applyMetricPreset("reference")}
                  disabled={!referenceAsset}
                >
                  {isEn ? `Match ${referencePreset.label}` : `匹配 ${referencePreset.label}`}
                </button>
                <button type="button" data-metric-preset="large" onClick={() => applyMetricPreset("large")}>
                  {isEn ? "Large operator" : "大型算子"}
                </button>
                <button type="button" data-metric-preset="extra" onClick={() => applyMetricPreset("extra")}>
                  {isEn ? "Extra large" : "超大型"}
                </button>
              </div>
              <div className="custom-symbol-metric-hint">
                {isEn
                  ? "The large workspace is only for editing. This output box controls the final TeX size and clipping bounds."
                  : "大工作区只用于设计；这里的输出范围才决定最终 TeX 字符尺寸与裁剪边界。"}
              </div>
            </section>

            <section className="custom-symbol-designer-panel">
              <header>
                <strong>{isEn ? "Selected layer" : "所选图层"}</strong>
              </header>
              {selectedLayer ? (
                <>
                  <div className="custom-symbol-designer-number-grid">
                    <NumericField
                      label="X"
                      field="layer-x"
                      value={selectedLayer.transform.translateX ?? 0}
                      onChange={(value) => updateSelectedTransform("translateX", value)}
                    />
                    <NumericField
                      label="Y"
                      field="layer-y"
                      value={selectedLayer.transform.translateY ?? 0}
                      onChange={(value) => updateSelectedTransform("translateY", value)}
                    />
                    <NumericField
                      label={isEn ? "Scale X" : "横向缩放"}
                      field="layer-scale-x"
                      value={selectedLayer.transform.scaleX ?? 1}
                      step={0.05}
                      onChange={(value) => updateSelectedTransform("scaleX", Math.max(0.02, value))}
                    />
                    <NumericField
                      label={isEn ? "Scale Y" : "纵向缩放"}
                      field="layer-scale-y"
                      value={selectedLayer.transform.scaleY ?? 1}
                      step={0.05}
                      onChange={(value) => updateSelectedTransform("scaleY", Math.max(0.02, value))}
                    />
                    <NumericField
                      label={isEn ? "Rotation °" : "旋转 °"}
                      field="layer-rotation"
                      value={selectedLayer.transform.rotateDeg ?? 0}
                      onChange={(value) => updateSelectedTransform("rotateDeg", value)}
                    />
                  </div>
                  <div className="custom-symbol-designer-inspector-actions">
                    <button
                      type="button"
                      data-duplicate-custom-symbol-layer
                      onClick={() => duplicateLayer(selectedLayer.id)}
                    >
                      <Copy size={13} />
                      {isEn ? "Duplicate" : "复制"}
                    </button>
                    <button
                      type="button"
                      className="is-danger"
                      data-delete-custom-symbol-layer
                      onClick={() => deleteLayer(selectedLayer.id)}
                    >
                      <Trash2 size={13} />
                      {isEn ? "Delete" : "删除"}
                    </button>
                  </div>
                  {selectedLayer.kind === "geometry" && selectedGeometryProperties ? (
                    <div className="custom-symbol-geometry-properties" data-custom-symbol-geometry-properties>
                      <div className="custom-symbol-crop-heading">
                        <strong>
                          {selectedLayer.geometryPreset === "eraser"
                            ? isEn
                              ? "Eraser stroke"
                              : "橡皮擦轨迹"
                            : isEn
                              ? "Geometry"
                              : "几何属性"}
                        </strong>
                        <span>{selectedLayer.geometryPreset ?? selectedLayer.shape.kind}</span>
                      </div>
                      <div className="custom-symbol-designer-number-grid">
                        {selectedLayer.geometryPreset !== "eraser" ? (
                          <NumericField
                            label={
                              selectedLayer.geometryPreset === "line"
                                ? isEn
                                  ? "Length"
                                  : "长度"
                                : selectedLayer.geometryPreset === "circle"
                                  ? isEn
                                    ? "Diameter"
                                    : "直径"
                                  : isEn
                                    ? "Width"
                                    : "宽度"
                            }
                            field="geometry-width"
                            value={selectedGeometryProperties.width}
                            step={10}
                            min={10}
                            max={8000}
                            onChange={(value) => updateSelectedGeometry({ width: value })}
                          />
                        ) : null}
                        {selectedLayer.geometryPreset !== "line" &&
                        selectedLayer.geometryPreset !== "circle" &&
                        selectedLayer.geometryPreset !== "eraser" ? (
                          <NumericField
                            label={isEn ? "Height" : "高度"}
                            field="geometry-height"
                            value={selectedGeometryProperties.height}
                            step={10}
                            min={10}
                            max={8000}
                            onChange={(value) => updateSelectedGeometry({ height: value })}
                          />
                        ) : null}
                        <NumericField
                          label={
                            selectedLayer.geometryPreset === "eraser"
                              ? isEn
                                ? "Eraser width"
                                : "擦除宽度"
                              : isEn
                                ? "Stroke width"
                                : "线宽"
                          }
                          field="geometry-stroke-width"
                          value={selectedGeometryProperties.strokeWidth}
                          step={2}
                          min={0}
                          max={600}
                          onChange={(value) => updateSelectedGeometry({ strokeWidth: value })}
                        />
                        {selectedLayer.geometryPreset === "rect" ? (
                          <NumericField
                            label={isEn ? "Corner radius" : "圆角"}
                            field="geometry-corner-radius"
                            value={selectedGeometryProperties.cornerRadius}
                            step={2}
                            min={0}
                            max={4000}
                            onChange={(value) => updateSelectedGeometry({ cornerRadius: value })}
                          />
                        ) : null}
                      </div>
                      {selectedLayer.geometryPreset !== "line" &&
                      selectedLayer.geometryPreset !== "arrow" &&
                      selectedLayer.geometryPreset !== "arc" &&
                      selectedLayer.geometryPreset !== "eraser" ? (
                        <label className="custom-symbol-geometry-fill-toggle">
                          <input
                            type="checkbox"
                            data-geometry-fill
                            checked={selectedGeometryProperties.fill}
                            onChange={(event) =>
                              updateSelectedGeometry({ fill: event.currentTarget.checked })
                            }
                          />
                          <span>{isEn ? "Filled shape" : "填充图形"}</span>
                        </label>
                      ) : null}
                    </div>
                  ) : null}
                  {selectedLayer.kind === "glyph" ? (
                    <div className="custom-symbol-crop-controls" data-custom-symbol-crop-controls>
                      <div className="custom-symbol-crop-heading">
                        <strong>{isEn ? "Non-destructive crop" : "非破坏裁剪"}</strong>
                        <span>
                          {selectedLayer.clipRect
                            ? isEn
                              ? "Enabled"
                              : "已启用"
                            : isEn
                              ? "Full glyph"
                              : "完整字形"}
                        </span>
                      </div>
                      <div className="custom-symbol-crop-presets">
                        <button type="button" data-crop-preset="full" onClick={() => applyClipPreset("full")}>
                          {isEn ? "Full" : "完整"}
                        </button>
                        <button type="button" data-crop-preset="top" onClick={() => applyClipPreset("top")}>
                          {isEn ? "Top" : "上"}
                        </button>
                        <button type="button" data-crop-preset="middle" onClick={() => applyClipPreset("middle")}>
                          {isEn ? "Middle" : "中"}
                        </button>
                        <button type="button" data-crop-preset="bottom" onClick={() => applyClipPreset("bottom")}>
                          {isEn ? "Bottom" : "下"}
                        </button>
                        <button type="button" data-crop-preset="left" onClick={() => applyClipPreset("left")}>
                          {isEn ? "Left" : "左"}
                        </button>
                        <button type="button" data-crop-preset="center" onClick={() => applyClipPreset("center")}>
                          {isEn ? "Center" : "中列"}
                        </button>
                        <button type="button" data-crop-preset="right" onClick={() => applyClipPreset("right")}>
                          {isEn ? "Right" : "右"}
                        </button>
                      </div>
                      {selectedLayer.clipRect ? (
                        <div className="custom-symbol-designer-number-grid is-crop-grid">
                          <NumericField
                            label="Crop X"
                            field="crop-x"
                            value={selectedLayer.clipRect.x}
                            onChange={(value) => updateSelectedClip({ x: value })}
                          />
                          <NumericField
                            label="Crop Y"
                            field="crop-y"
                            value={selectedLayer.clipRect.y}
                            onChange={(value) => updateSelectedClip({ y: value })}
                          />
                          <NumericField
                            label="Crop W"
                            field="crop-width"
                            value={selectedLayer.clipRect.width}
                            min={1}
                            onChange={(value) => updateSelectedClip({ width: value })}
                          />
                          <NumericField
                            label="Crop H"
                            field="crop-height"
                            value={selectedLayer.clipRect.height}
                            min={1}
                            onChange={(value) => updateSelectedClip({ height: value })}
                          />
                        </div>
                      ) : null}
                      <div className="custom-symbol-split-actions">
                        <button
                          type="button"
                          data-split-custom-symbol-glyph="horizontal"
                          onClick={() => splitSelectedGlyph("horizontal")}
                        >
                          {isEn ? "Split top / middle / bottom" : "上 / 中 / 下三分"}
                        </button>
                        <button
                          type="button"
                          data-split-custom-symbol-glyph="vertical"
                          onClick={() => splitSelectedGlyph("vertical")}
                        >
                          {isEn ? "Split left / center / right" : "左 / 中 / 右三分"}
                        </button>
                      </div>
                    </div>
                  ) : null}
                </>
              ) : (
                <div className="custom-symbol-designer-inspector-empty">
                  {isEn ? "Select a layer." : "选择一个图层。"}
                </div>
              )}
            </section>

            <section
              className="custom-symbol-designer-panel is-registration"
              data-custom-symbol-registration-panel
              data-registration-dirty={registrationDirty ? "true" : "false"}
              data-registration-symbol-id={registeredSymbol?.id ?? ""}
            >
              <header>
                <strong>{isEn ? "Register as a LaTeX command" : "注册为 LaTeX 命令"}</strong>
                <span>
                  {registeredSymbol
                    ? registrationDirty
                      ? isEn
                        ? "Unsaved changes"
                        : "有未保存修改"
                      : isEn
                        ? "Registered"
                        : "已注册"
                    : isEn
                      ? "Draft"
                      : "草稿"}
                </span>
              </header>
              <label className="custom-symbol-registration-field">
                <span>{isEn ? "Symbol name" : "字符名称"}</span>
                <input
                  type="text"
                  value={documentState.name}
                  data-custom-symbol-name-input
                  onChange={(event) => {
                    const value = event.currentTarget.value;
                    setDocumentState((current) => ({ ...current, name: value }));
                  }}
                  placeholder={isEn ? "My custom symbol" : "我的自定义字符"}
                />
              </label>
              <label className="custom-symbol-registration-field">
                <span>{isEn ? "Command" : "命令"}</span>
                <input
                  type="text"
                  value={documentState.command}
                  data-custom-symbol-command-input
                  spellCheck={false}
                  onChange={(event) => {
                    const value = event.currentTarget.value;
                    setDocumentState((current) => ({ ...current, command: value }));
                  }}
                  placeholder="\\selfdefa"
                />
              </label>
              <label className="custom-symbol-registration-field">
                <span>{isEn ? "Math role" : "数学类型"}</span>
                <select
                  value={documentState.role}
                  data-custom-symbol-role-select
                  onChange={(event) => {
                    const role = event.currentTarget.value as CustomSymbolMathRole;
                    setDocumentState((current) => ({ ...current, role }));
                  }}
                >
                  {customSymbolMathRoles.map((role) => (
                    <option key={role} value={role}>
                      {mathRoleLabel(role)}
                    </option>
                  ))}
                </select>
              </label>
              {documentState.role === "operator" ? (
                <label className="custom-symbol-registration-field">
                  <span>{isEn ? "Limits" : "上下限行为"}</span>
                  <select
                    value={documentState.limitsBehavior}
                    data-custom-symbol-limits-select
                    onChange={(event) => {
                      const limitsBehavior = event.currentTarget.value as CustomSymbolLimitsBehavior;
                      setDocumentState((current) => ({ ...current, limitsBehavior }));
                    }}
                  >
                    {customSymbolLimitsBehaviors.map((behavior) => (
                      <option key={behavior} value={behavior}>
                        {limitsLabel(behavior)}
                      </option>
                    ))}
                  </select>
                </label>
              ) : null}
              <button
                type="button"
                className="custom-symbol-register-button"
                data-register-custom-symbol
                onClick={saveRegisteredSymbol}
                disabled={!documentState.name.trim() || !documentState.command.trim() || shapeCount === 0}
              >
                {registeredSymbol
                  ? isEn
                    ? "Update registered symbol"
                    : "更新已注册字符"
                  : isEn
                    ? "Register symbol"
                    : "注册字符"}
              </button>
              {registrationState.kind !== "idle" ? (
                <div
                  className={`custom-symbol-registration-status is-${registrationState.kind}`}
                  data-custom-symbol-registration-status={registrationState.kind}
                  role={registrationState.kind === "error" ? "alert" : "status"}
                >
                  {registrationState.message}
                </div>
              ) : null}
              {registeredSymbol ? (
                <div className="custom-symbol-registered-preview" data-custom-symbol-registered-preview>
                  <MathPreview latex={`\\${registeredSymbol.command}`} staticLayout />
                  <code>{`\\${registeredSymbol.command}`}</code>
                </div>
              ) : null}
            </section>
          </aside>
        </div>

        <footer className="custom-symbol-designer-footer">
          <span>
            {documentState.layers.length} {isEn ? "layers" : "图层"} · {shapeCount}{" "}
            {isEn ? "shapes" : "矢量单元"}
          </span>
          <button type="button" onClick={onClose}>
            {isEn ? "Close" : "关闭"}
          </button>
        </footer>
      </section>
    </div>,
    document.body,
  );
}
