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
  FlipHorizontal2,
  FlipVertical2,
  Grip,
  Italic,
  Lock,
  Plus,
  RotateCcw,
  RotateCw,
  Search,
  Trash2,
  Unlock,
  X,
} from "lucide-react";
import { MathPreview } from "./MathPreview";
import {
  categoryLabels,
  categoryLabelsEn,
  commandRegistry,
} from "../autocomplete/commandRegistry";
import { customSymbolCommands } from "../autocomplete/runtimeCommandRegistry";
import type { CommandCategory, LatexCommand } from "../types/command";
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
import { containsCustomSymbolCommand } from "../math/customSymbolRendering";
import {
  SYSTEM_GLYPH_CATEGORIES,
  SYSTEM_GLYPH_CATEGORY_LABELS,
  SYSTEM_MATH_FONT_PRESETS,
  createSystemFontGlyphAssetAsync,
  detectSystemMathFontAvailability,
  searchSystemGlyphs,
  systemFontPresetById,
  type SystemGlyphCategory,
  type SystemMathFontId,
} from "../math/customSymbolSystemGlyphs";
import { latexToMathMl } from "../export/runtime";
import type {
  CustomSymbolDefinition,
  CustomSymbolLayerEffects,
  CustomSymbolLimitsBehavior,
  CustomSymbolMathRole,
  CustomSymbolMetrics,
} from "../math/customSymbolTypes";

interface Props {
  open: boolean;
  language: "cn" | "en";
  onClose: () => void;
}

const designerMaterialCategories: readonly CommandCategory[] = [
  "common",
  "calculus",
  "greek",
  "relation",
  "set",
  "arrow",
  "physics",
];

const designerCommonBareCommandIds = new Set([
  "intplain",
  "partial",
  "nabla",
  "infty",
  "alpha",
  "beta",
  "gamma",
  "theta",
  "lambda",
  "pi",
  "sigma",
  "omega",
  "equal",
  "neq",
  "approx",
  "leq",
  "geq",
  "in",
  "notin",
  "subset",
  "forall",
  "exists",
  "rightarrow",
  "leftarrow",
  "hbar",
]);

const designerBareOperatorCommands = new Set([
  "\\int",
  "\\iint",
  "\\iiint",
  "\\oint",
  "\\oiint",
  "\\oiiint",
  "\\sum",
  "\\prod",
  "\\coprod",
  "\\bigcup",
  "\\bigcap",
  "\\bigvee",
  "\\bigwedge",
  "\\partial",
]);

interface DesignerMaterialEntry {
  command: LatexCommand;
  source: string;
}

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

function designerBareMaterialSource(command: LatexCommand) {
  if (!command.supportedInMathMode) return null;
  if (!designerMaterialCategories.includes(command.category)) return null;
  const raw = command.command.trim();
  const simpleLiteral = Array.from(raw).length === 1 && raw !== "\\";
  const simpleControlWord = /^\\[A-Za-z]+$/.test(raw);
  if (!simpleLiteral && !simpleControlWord) return null;
  if (command.id.startsWith("custom-symbol:")) {
    return normalizeDesignerMaterialLatex(raw);
  }
  if (designerBareOperatorCommands.has(raw)) {
    return normalizeDesignerMaterialLatex(raw);
  }
  if (command.insertTemplate.includes("\\placeholder")) return null;
  if (["\\begin", "\\end", "\\left", "\\right", "\\middle"].includes(raw)) {
    return null;
  }
  return normalizeDesignerMaterialLatex(raw);
}

function isSingleUnicodeMaterial(source: string) {
  const characters = Array.from(source.trim());
  return characters.length === 1 && characters[0] !== "\\" && !/\s/u.test(characters[0]);
}

function cloneLayer(layer: CustomSymbolDesignerLayer): CustomSymbolDesignerLayer {
  if (typeof structuredClone === "function") return structuredClone(layer);
  return JSON.parse(JSON.stringify(layer)) as CustomSymbolDesignerLayer;
}

function designerLayerCenter(layer: CustomSymbolDesignerLayer) {
  if (layer.kind === "glyph") {
    const bounds = layer.clipRect ?? {
      x: 0,
      y: 0,
      width: layer.asset.metrics.widthEm * 1000,
      height:
        (layer.asset.metrics.ascentEm + layer.asset.metrics.descentEm) * 1000,
    };
    return {
      x: bounds.x + bounds.width / 2,
      y: bounds.y + bounds.height / 2,
    };
  }
  return {
    x: layer.bounds.x + layer.bounds.width / 2,
    y: layer.bounds.y + layer.bounds.height / 2,
  };
}

function centerGlyphLayerOnDesignerBaseline(
  layer: Extract<CustomSymbolDesignerLayer, { kind: "glyph" }>,
  metrics: CustomSymbolMetrics,
): CustomSymbolDesignerLayer {
  return {
    ...layer,
    transform: {
      ...layer.transform,
      translateX:
        ((metrics.widthEm - layer.asset.metrics.widthEm) * 1000) / 2,
      translateY:
        (metrics.ascentEm - layer.asset.metrics.ascentEm) * 1000,
    },
  };
}

function centerDesignerLayerTransform(
  layer: CustomSymbolDesignerLayer,
): CustomSymbolDesignerLayer {
  const center = designerLayerCenter(layer);
  return {
    ...layer,
    transform: {
      ...layer.transform,
      originX: center.x,
      originY: center.y,
    },
  };
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
  const [materialSearch, setMaterialSearch] = useState("");
  const [materialCategory, setMaterialCategory] =
    useState<CommandCategory>("common");
  const [systemGlyphSearch, setSystemGlyphSearch] = useState("");
  const [systemGlyphCategory, setSystemGlyphCategory] =
    useState<SystemGlyphCategory>("basic-italic");
  const [systemGlyphFontId, setSystemGlyphFontId] =
    useState<SystemMathFontId>("cambria-math");
  const [systemGlyphItalic, setSystemGlyphItalic] = useState(true);
  const [systemGlyphBusyKey, setSystemGlyphBusyKey] = useState<string | null>(null);
  const [systemGlyphStatus, setSystemGlyphStatus] = useState("");
  const [systemFontAvailability, setSystemFontAvailability] = useState<
    Partial<Record<SystemMathFontId, boolean>>
  >({});
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
  const designerMaterialEntries = useMemo<DesignerMaterialEntry[]>(() => {
    const sorted = [...commandRegistry, ...customSymbolCommands()]
      .map((command) => ({
        command,
        source: designerBareMaterialSource(command),
      }))
      .filter(
        (entry): entry is DesignerMaterialEntry => typeof entry.source === "string",
      )
      .sort(
        (left, right) =>
          right.command.defaultPriority - left.command.defaultPriority ||
          left.command.labelZh.localeCompare(right.command.labelZh),
      );
    const bySource = new Map<string, DesignerMaterialEntry>();
    for (const entry of sorted) {
      if (!bySource.has(entry.source)) bySource.set(entry.source, entry);
    }
    return [...bySource.values()];
  }, [customSymbolRevision]);
  const designerMaterialCommands = useMemo(() => {
    const query = materialSearch.trim().toLocaleLowerCase();
    const normalizedQuery = query.replace(/^\\/, "");
    return designerMaterialEntries.filter(({ command }) => {
      if (!query) {
        if (materialCategory === "common") {
          return (
            command.id.startsWith("custom-symbol:") ||
            designerCommonBareCommandIds.has(command.id)
          );
        }
        return command.category === materialCategory;
      }
      const searchText = [
        command.command,
        command.labelZh,
        command.labelEn,
        ...command.aliases,
        ...command.keywords,
      ]
        .join(" ")
        .toLocaleLowerCase();
      return searchText.includes(query) || searchText.includes(normalizedQuery);
    });
  }, [designerMaterialEntries, materialCategory, materialSearch]);
  const systemGlyphFont = systemFontPresetById(systemGlyphFontId);
  const systemGlyphDefinitions = useMemo(
    () => searchSystemGlyphs(systemGlyphCategory, systemGlyphSearch),
    [systemGlyphCategory, systemGlyphSearch],
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

  useEffect(() => {
    if (!open) return;
    let cancelled = false;
    void detectSystemMathFontAvailability().then((availability) => {
      if (!cancelled) setSystemFontAvailability(availability);
    });
    return () => {
      cancelled = true;
    };
  }, [open]);

  useEffect(() => {
    if (!open || !pendingDeleteSymbolId) return;
    const frame = requestAnimationFrame(() => {
      const item = document.querySelector<HTMLElement>(
        `[data-registered-custom-symbol="${pendingDeleteSymbolId}"]`,
      );
      item?.scrollIntoView({ block: "nearest", inline: "nearest" });
    });
    return () => cancelAnimationFrame(frame);
  }, [open, pendingDeleteSymbolId]);

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
      const requestedSource = source.trim();
      const normalizedSource = normalizeDesignerMaterialLatex(requestedSource);
      const knownBareMaterial = designerMaterialEntries.some(
        (entry) => entry.source === normalizedSource,
      );
      if (!isSingleUnicodeMaterial(requestedSource) && !knownBareMaterial) {
        throw new Error(
          isEn
            ? "Only one bare character or symbol command can be added."
            : "这里只能加入单个裸字符或符号命令。",
        );
      }
      const asset = compileLatexGlyphAsset(normalizedSource);
      const layer = glyphLayerFromAsset(asset, { id: createUuid(), name: source.trim() });
      setDocumentState((current) => ({
        ...current,
        layers: [
          ...current.layers,
          centerGlyphLayerOnDesignerBaseline(layer, current.metrics),
        ],
      }));
      setSelectedLayerId(layer.id);
      setMaterialLatex(source);
      setMaterialError("");
    } catch (error) {
      setMaterialError(error instanceof Error ? error.message : String(error));
    }
  };

  const addSystemGlyph = async (character: string, label: string) => {
    const busyKey = `${systemGlyphFont.id}:${character.codePointAt(0) ?? 0}`;
    setSystemGlyphBusyKey(busyKey);
    setSystemGlyphStatus("");
    try {
      const created = await createSystemFontGlyphAssetAsync({
        character,
        font: systemGlyphFont,
        italic: systemGlyphItalic,
      });
      const asset = created.asset;
      const layer = glyphLayerFromAsset(asset, {
        id: createUuid(),
        name: `${character} · ${label} · ${created.resolvedFamily}`,
      });
      setDocumentState((current) => ({
        ...current,
        layers: [
          ...current.layers,
          centerGlyphLayerOnDesignerBaseline(layer, current.metrics),
        ],
      }));
      setSelectedLayerId(layer.id);
      setMaterialError("");
      if (created.vectorOutline) {
        setSystemGlyphStatus(
          created.fallbackUsed
            ? `${created.requestedFamily} → ${created.resolvedFamily} · ${isEn ? "vector" : "矢量"}`
            : `${created.resolvedFamily} · ${isEn ? "vector" : "矢量"}`,
        );
      } else {
        setSystemGlyphStatus(
          created.warning
            ? isEn
              ? "System-font fallback"
              : "系统字体回退"
            : isEn
              ? "Browser preview"
              : "浏览器预览",
        );
      }
    } catch (error) {
      setMaterialError(error instanceof Error ? error.message : String(error));
    } finally {
      setSystemGlyphBusyKey(null);
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
        ? centerDesignerLayerTransform(
            updateCustomSymbolGeometryLayer(layer, patch),
          )
        : layer,
    );
  };

  const updateSelectedClip = (
    patch: Partial<{ x: number; y: number; width: number; height: number }>,
  ) => {
    if (!selectedLayer || selectedLayer.kind !== "glyph") return;
    updateLayer(selectedLayer.id, (layer) => {
      if (layer.kind !== "glyph") return layer;
      const current = layer.clipRect ?? glyphLayerFullClip(layer);
      const next = { ...current, ...patch };
      return centerDesignerLayerTransform({
        ...layer,
        clipRect: {
          x: next.x,
          y: next.y,
          width: Math.max(1, next.width),
          height: Math.max(1, next.height),
        },
      });
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
      layer.kind === "glyph"
        ? centerDesignerLayerTransform({ ...layer, clipRect: clip })
        : layer,
    );
  };

  const splitSelectedGlyph = (orientation: "horizontal" | "vertical") => {
    if (!selectedLayer || selectedLayer.kind !== "glyph") return;
    const slices = createGlyphLayerSlices(selectedLayer, orientation, 3).map(
      centerDesignerLayerTransform,
    );
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
      if (!name) {
        throw new Error(isEn ? "Enter a symbol name." : "请输入字符名称。");
      }
      if (!command) {
        throw new Error(isEn ? "Enter a LaTeX command name." : "请输入 LaTeX 命令名。");
      }
      if (shapeCount === 0) {
        throw new Error(
          isEn
            ? "The symbol must contain at least one visible vector shape."
            : "字符至少需要一个可见的矢量单元。",
        );
      }
      const fallback = documentState.ommlFallback?.trim() ?? "";
      if (fallback) {
        if (containsCustomSymbolCommand(fallback)) {
          throw new Error(
            isEn
              ? "Word fallback cannot depend on another VisualTeX custom symbol."
              : "Word fallback 不能依赖其他 VisualTeX 自定义字符。",
          );
        }
        latexToMathMl(fallback, false);
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
        ommlFallback: fallback || null,
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
        ommlFallback: saved.ommlFallback ?? null,
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
    setDocumentState((current) => ({ ...current, layers: [...current.layers, duplicate] }));
    setSelectedLayerId(duplicate.id);
  };

  const deleteLayer = (id: string) => {
    setDocumentState((current) => ({
      ...current,
      layers: current.layers.filter((layer) => layer.id !== id),
    }));
    if (selectedLayerId === id) setSelectedLayerId(null);
  };

  const updateSelectedTransform = (key: string, value: number) => {
    if (!selectedLayer) return;
    updateLayer(selectedLayer.id, (layer) => {
      const centered = centerDesignerLayerTransform(layer);
      return {
        ...centered,
        transform: {
          ...centered.transform,
          [key]: value,
        },
      };
    });
  };

  const flipSelectedLayer = (axis: "horizontal" | "vertical") => {
    if (!selectedLayer) return;
    updateLayer(selectedLayer.id, (layer) => {
      const centered = centerDesignerLayerTransform(layer);
      return {
        ...centered,
        transform: {
          ...centered.transform,
          ...(axis === "horizontal"
            ? { scaleX: -(centered.transform.scaleX ?? 1) }
            : { scaleY: -(centered.transform.scaleY ?? 1) }),
        },
      };
    });
  };

  const updateSelectedEffects = (
    kind: keyof CustomSymbolLayerEffects,
    patch: Record<string, number | boolean>,
  ) => {
    if (!selectedLayer) return;
    updateLayer(selectedLayer.id, (layer) => {
      const defaults =
        kind === "outline"
          ? { enabled: false, width: 30 }
          : { enabled: false, depth: 240, angleDeg: 35, steps: 8 };
      return {
        ...layer,
        effects: {
          ...layer.effects,
          [kind]: {
            ...defaults,
            ...(layer.effects?.[kind] ?? {}),
            ...patch,
          },
        },
      };
    });
  };

  const openRegisteredSymbol = (symbol: CustomSymbolDefinition) => {
    const restored = restoreCustomSymbolDesignerDocument(symbol);
    setDocumentState(restored.document);
    setSelectedLayerId(restored.document.layers[0]?.id ?? null);
    setMaterialError("");
    setRegistrationState({ kind: "idle", message: "" });
    setSavedRegistrationFingerprint(JSON.stringify(restored.document));
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
      setRegistrationState({
        kind: "idle",
        message: isEn
          ? `Confirm deletion of \\${symbol.command} below.`
          : `请在下方确认删除 \\${symbol.command}。`,
      });
      return;
    }
    deleteCustomSymbol(symbol.id);
    setPendingDeleteSymbolId(null);
    if (documentState.symbolId === symbol.id) {
      setDocumentState(createEmptyCustomSymbolDesignerDocument());
      setSelectedLayerId(null);
      setSavedRegistrationFingerprint("");
      setDesignerSourceMode(null);
    }
    setRegistrationState({
      kind: "success",
      message: isEn
        ? `Deleted \\${symbol.command}`
        : `已删除 \\${symbol.command}`,
    });
  };

  const reset = () => {
    setDocumentState(createEmptyCustomSymbolDesignerDocument());
    setSelectedLayerId(null);
    setMaterialLatex("\\partial");
    setMaterialSearch("");
    setMaterialCategory("common");
    setSystemGlyphSearch("");
    setSystemGlyphCategory("basic-italic");
    setSystemGlyphFontId("cambria-math");
    setSystemGlyphItalic(true);
    setSystemGlyphBusyKey(null);
    setSystemGlyphStatus("");
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
          </div>
          <div>
            <button className="icon-button compact" type="button" data-reset-custom-symbol-designer onClick={reset} aria-label={isEn ? "Reset" : "重置"}>
              <RotateCcw size={15} />
            </button>
            <button className="icon-button compact" type="button" onClick={onClose} aria-label={isEn ? "Close" : "关闭"}>
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
                      className={`custom-symbol-registered-item${active ? " is-active" : ""}${pendingDelete ? " is-pending-delete" : ""}`}
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
                              ? `Existing formulas using \\${symbol.command} will become unresolved.`
                              : `已有公式中的 \\${symbol.command} 将无法解析。`}
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
                    {isEn ? "No registered symbols" : "暂无已注册字符"}
                  </div>
                ) : null}
              </div>
            </section>

            <section className="custom-symbol-designer-panel">
              <header><FileCode2 size={16} /><strong>{isEn ? "Character materials" : "字符素材"}</strong></header>
              <div className="custom-symbol-material-input-row">
                <input
                  value={materialLatex}
                  data-custom-symbol-material-input
                  spellCheck={false}
                  onChange={(event) => setMaterialLatex(event.currentTarget.value)}
                  onKeyDown={(event) => event.key === "Enter" && addMaterial()}
                />
                <button type="button" data-add-custom-symbol-material onClick={() => addMaterial()} disabled={!materialLatex.trim()}>
                  <Plus size={14} />{isEn ? "Add" : "加入"}
                </button>
              </div>
              <label className="custom-symbol-material-search">
                <Search size={13} aria-hidden="true" />
                <input
                  type="search"
                  value={materialSearch}
                  data-custom-symbol-material-search
                  placeholder={isEn ? "Search characters" : "搜索字符"}
                  onChange={(event) => setMaterialSearch(event.currentTarget.value)}
                />
              </label>
              <div
                className="custom-symbol-material-categories"
                role="tablist"
                aria-label={isEn ? "Material categories" : "素材分类"}
              >
                {designerMaterialCategories.map((category) => (
                  <button
                    type="button"
                    role="tab"
                    key={category}
                    className={materialCategory === category ? "is-active" : ""}
                    aria-selected={materialCategory === category}
                    data-custom-symbol-material-category={category}
                    onClick={() => setMaterialCategory(category)}
                  >
                    {(isEn ? categoryLabelsEn : categoryLabels)[category]}
                  </button>
                ))}
              </div>
              <div
                className="custom-symbol-material-library"
                data-custom-symbol-material-library
                data-material-count={designerMaterialEntries.length}
                data-visible-material-count={designerMaterialCommands.length}
                data-bare-materials-only="true"
              >
                {designerMaterialCommands.map(({ command, source }) => (
                  <button
                    type="button"
                    key={`${command.id}-${source}`}
                    data-custom-symbol-material-command={command.id}
                    data-custom-symbol-material-latex={source}
                    onClick={() => addMaterial(source)}
                    title={`${isEn ? command.labelEn : command.labelZh} · ${command.command}`}
                    aria-label={isEn ? command.labelEn : command.labelZh}
                  >
                    <MathPreview
                      latex={source}
                      fit
                      maximumFitScale={1.18}
                      fitInsetRatio={0.72}
                    />
                  </button>
                ))}
                {!designerMaterialCommands.length ? (
                  <div className="custom-symbol-material-empty">
                    {isEn ? "No matching material." : "没有匹配的素材。"}
                  </div>
                ) : null}
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
              {materialError ? <div className="custom-symbol-designer-error">{materialError}</div> : null}
            </section>

            <section className="custom-symbol-designer-panel is-system-glyphs">
              <header>
                <strong>{isEn ? "System math glyphs" : "扩展字体字符"}</strong>
                <span>{systemGlyphDefinitions.length}</span>
              </header>
              <label className="custom-symbol-system-font-field">
                <span>{isEn ? "Font" : "字体"}</span>
                <select
                  value={systemGlyphFontId}
                  data-custom-symbol-system-font-select
                  onChange={(event) =>
                    setSystemGlyphFontId(
                      event.currentTarget.value as SystemMathFontId,
                    )
                  }
                >
                  {SYSTEM_MATH_FONT_PRESETS.map((font) => {
                    const available = systemFontAvailability[font.id];
                    const status =
                      available === true
                        ? isEn
                          ? " · detected"
                          : " · 已检测"
                        : available === false
                          ? isEn
                            ? " · fallback"
                            : " · 将回退"
                          : "";
                    return (
                      <option key={font.id} value={font.id}>
                        {isEn ? font.labelEn : font.labelZh}
                        {status}
                      </option>
                    );
                  })}
                </select>
              </label>
              <div
                className="custom-symbol-system-font-style"
                data-custom-symbol-system-font-style
              >
                <button
                  type="button"
                  className={!systemGlyphItalic ? "is-active" : ""}
                  aria-pressed={!systemGlyphItalic}
                  data-custom-symbol-system-font-upright
                  onClick={() => setSystemGlyphItalic(false)}
                >
                  {isEn ? "Upright" : "正体"}
                </button>
                <button
                  type="button"
                  className={systemGlyphItalic ? "is-active" : ""}
                  aria-pressed={systemGlyphItalic}
                  data-custom-symbol-system-font-italic
                  onClick={() => setSystemGlyphItalic(true)}
                >
                  {isEn ? "Math italic" : "数学斜体"}
                </button>
              </div>
              <label className="custom-symbol-material-search is-system-glyph-search">
                <Search size={13} aria-hidden="true" />
                <input
                  type="search"
                  value={systemGlyphSearch}
                  data-custom-symbol-system-glyph-search
                  placeholder={isEn ? "Character, U+ code or decimal code" : "字符、U+ 编码或十进制编码"}
                  onChange={(event) =>
                    setSystemGlyphSearch(event.currentTarget.value)
                  }
                />
              </label>
              <div
                className="custom-symbol-system-glyph-categories"
                role="tablist"
                aria-label={isEn ? "System glyph categories" : "扩展字符分类"}
              >
                {SYSTEM_GLYPH_CATEGORIES.map((category) => (
                  <button
                    type="button"
                    role="tab"
                    key={category}
                    className={systemGlyphCategory === category ? "is-active" : ""}
                    aria-selected={systemGlyphCategory === category}
                    data-custom-symbol-system-glyph-category={category}
                    onClick={() => setSystemGlyphCategory(category)}
                  >
                    {
                      SYSTEM_GLYPH_CATEGORY_LABELS[category][
                        isEn ? "en" : "zh"
                      ]
                    }
                  </button>
                ))}
              </div>
              <div
                className="custom-symbol-system-glyph-grid"
                data-custom-symbol-system-glyph-grid
                data-system-glyph-font={systemGlyphFont.id}
              >
                {systemGlyphDefinitions.map((glyph) => {
                  const glyphBusyKey = `${systemGlyphFont.id}:${glyph.codePoint}`;
                  const busy = systemGlyphBusyKey === glyphBusyKey;
                  return (
                    <button
                      type="button"
                      key={`${glyph.category}-${glyph.codePoint}`}
                      className={busy ? "is-loading" : ""}
                      disabled={Boolean(systemGlyphBusyKey)}
                      data-custom-symbol-system-glyph={glyph.label}
                      data-system-glyph-character={glyph.character}
                      data-system-glyph-vector-target="true"
                      title={`${glyph.character} · ${glyph.label}`}
                      aria-label={`${glyph.character} ${glyph.label}`}
                      aria-busy={busy}
                      style={{
                        fontFamily: systemGlyphFont.family,
                        fontStyle: systemGlyphItalic ? "italic" : "normal",
                      }}
                      onClick={() => {
                        void addSystemGlyph(glyph.character, glyph.label);
                      }}
                    >
                      {glyph.character}
                    </button>
                  );
                })}
              </div>
              {systemGlyphStatus ? (
                <div
                  className="custom-symbol-system-font-status"
                  data-custom-symbol-system-font-status
                  role="status"
                >
                  {systemGlyphStatus}
                </div>
              ) : null}
            </section>

            <section className="custom-symbol-designer-panel is-layers">
              <header><Grip size={15} /><strong>{isEn ? "Layers" : "图层"}</strong><span>{documentState.layers.length}</span></header>
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
                      data-layer-source-latex={
                        layer.kind === "glyph" ? layer.asset.sourceLatex : ""
                      }
                      data-layer-visible={layer.visible ? "true" : "false"}
                      data-layer-locked={layer.locked ? "true" : "false"}
                      onClick={() => setSelectedLayerId(layer.id)}
                    >
                      <button type="button" data-toggle-custom-symbol-layer-visibility onClick={(event) => { event.stopPropagation(); updateLayer(layer.id, (item) => ({ ...item, visible: !item.visible })); }}>
                        {layer.visible ? <Eye size={13} /> : <EyeOff size={13} />}
                      </button>
                      <div className="custom-symbol-layer-preview">
                        {layer.kind === "glyph" ? (
                          <MathPreview latex={layer.asset.sourceLatex} staticLayout />
                        ) : (
                          <span className={`custom-symbol-geometry-icon is-${layer.shape.kind}`} aria-hidden="true" />
                        )}
                      </div>
                      <div className="custom-symbol-layer-label"><strong>{layer.name}</strong><span>{layer.kind === "glyph" ? layer.asset.sourceLatex : "Geometry"}</span></div>
                      <div className="custom-symbol-layer-actions">
                        <button type="button" data-toggle-custom-symbol-layer-lock onClick={(event) => { event.stopPropagation(); updateLayer(layer.id, (item) => ({ ...item, locked: !item.locked })); }}>{layer.locked ? <Lock size={12} /> : <Unlock size={12} />}</button>
                        <button type="button" data-move-custom-symbol-layer-up disabled={index >= documentState.layers.length - 1} onClick={(event) => { event.stopPropagation(); moveLayerOrder(layer.id, 1); }}><ArrowUp size={12} /></button>
                        <button type="button" data-move-custom-symbol-layer-down disabled={index <= 0} onClick={(event) => { event.stopPropagation(); moveLayerOrder(layer.id, -1); }}><ArrowDown size={12} /></button>
                      </div>
                    </div>
                  );
                })}
                {!documentState.layers.length ? <div className="custom-symbol-layer-empty">{isEn ? "No layers" : "暂无图层"}</div> : null}
              </div>
            </section>
          </aside>

          <main className="custom-symbol-designer-stage-column">
            <div className="custom-symbol-designer-stage-toolbar">
              <div className="custom-symbol-designer-stage-toolbar-actions">
                <label className="custom-symbol-reference-selector">
                  <span>{isEn ? "Reference" : "参考"}</span>
                  <select
                    value={referenceLatex}
                    data-custom-symbol-reference-select
                    onChange={(event) => {
                      setReferenceLatex(event.currentTarget.value);
                      setShowReference(true);
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
                        if (Number.isFinite(value)) setEraserSize(Math.max(4, Math.min(600, value)));
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
                onResizeLayer={(id, scaleX, scaleY, translateX, translateY) =>
                  updateLayer(id, (layer) => ({
                    ...layer,
                    transform: {
                      ...layer.transform,
                      scaleX,
                      scaleY,
                      translateX,
                      translateY,
                    },
                  }))
                }
                onRotateLayer={(id, rotateDeg) =>
                  updateLayer(id, (layer) => {
                    const centered = centerDesignerLayerTransform(layer);
                    return {
                      ...centered,
                      transform: { ...centered.transform, rotateDeg },
                    };
                  })
                }
                onDeleteLayer={deleteLayer}
                onAddEraserStroke={addEraserStroke}
              />
            </div>
          </main>

          <aside className="custom-symbol-designer-sidebar is-inspector">
            {designerSourceMode === "flattened-legacy" ? (
              <div className="custom-symbol-designer-legacy-warning" data-custom-symbol-legacy-warning>
                {isEn ? "Legacy flattened layers" : "旧字符：扁平图层"}
              </div>
            ) : null}
            <section className="custom-symbol-designer-panel">
              <header><strong>{isEn ? "Selected layer" : "所选图层"}</strong></header>
              {selectedLayer ? (
                <>
                  <div className="custom-symbol-designer-number-grid">
                    <NumericField label="X" field="layer-x" value={selectedLayer.transform.translateX ?? 0} onChange={(value) => updateSelectedTransform("translateX", value)} />
                    <NumericField label="Y" field="layer-y" value={selectedLayer.transform.translateY ?? 0} onChange={(value) => updateSelectedTransform("translateY", value)} />
                    <NumericField
                      label={isEn ? "Scale X" : "横向缩放"}
                      field="layer-scale-x"
                      value={Math.abs(selectedLayer.transform.scaleX ?? 1)}
                      step={0.05}
                      min={0.02}
                      onChange={(value) =>
                        updateSelectedTransform(
                          "scaleX",
                          (selectedLayer.transform.scaleX ?? 1) < 0
                            ? -Math.max(0.02, value)
                            : Math.max(0.02, value),
                        )
                      }
                    />
                    <NumericField
                      label={isEn ? "Scale Y" : "纵向缩放"}
                      field="layer-scale-y"
                      value={Math.abs(selectedLayer.transform.scaleY ?? 1)}
                      step={0.05}
                      min={0.02}
                      onChange={(value) =>
                        updateSelectedTransform(
                          "scaleY",
                          (selectedLayer.transform.scaleY ?? 1) < 0
                            ? -Math.max(0.02, value)
                            : Math.max(0.02, value),
                        )
                      }
                    />
                    <NumericField label={isEn ? "Italic shear °" : "横向倾斜 °"} field="layer-skew-x" value={selectedLayer.transform.skewXDeg ?? 0} step={1} min={-45} max={45} onChange={(value) => updateSelectedTransform("skewXDeg", Math.max(-45, Math.min(45, value)))} />
                    <NumericField label={isEn ? "Vertical shear °" : "纵向倾斜 °"} field="layer-skew-y" value={selectedLayer.transform.skewYDeg ?? 0} step={1} min={-45} max={45} onChange={(value) => updateSelectedTransform("skewYDeg", Math.max(-45, Math.min(45, value)))} />
                    <NumericField label={isEn ? "Rotation °" : "中心旋转 °"} field="layer-rotation" value={selectedLayer.transform.rotateDeg ?? 0} onChange={(value) => updateSelectedTransform("rotateDeg", value)} />
                  </div>
                  <div className="custom-symbol-transform-actions" data-custom-symbol-transform-actions>
                    <button type="button" data-flip-custom-symbol-layer="horizontal" onClick={() => flipSelectedLayer("horizontal")}>
                      <FlipHorizontal2 size={15} />
                      {isEn ? "Horizontal" : "水平翻转"}
                    </button>
                    <button type="button" data-flip-custom-symbol-layer="vertical" onClick={() => flipSelectedLayer("vertical")}>
                      <FlipVertical2 size={15} />
                      {isEn ? "Vertical" : "垂直翻转"}
                    </button>
                    <button type="button" data-custom-symbol-math-italic onClick={() => updateSelectedTransform("skewXDeg", -12)}>
                      <Italic size={15} />
                      {isEn ? "Math italic" : "数学斜体"}
                    </button>
                    <button type="button" data-custom-symbol-original-slant onClick={() => updateSelectedTransform("skewXDeg", 0)}>
                      {isEn ? "Original slant" : "恢复斜度"}
                    </button>
                    <button type="button" data-rotate-custom-symbol-layer="-90" onClick={() => updateSelectedTransform("rotateDeg", (selectedLayer.transform.rotateDeg ?? 0) - 90)}>
                      <RotateCcw size={15} />
                      −90°
                    </button>
                    <button type="button" data-rotate-custom-symbol-layer="90" onClick={() => updateSelectedTransform("rotateDeg", (selectedLayer.transform.rotateDeg ?? 0) + 90)}>
                      <RotateCw size={15} />
                      +90°
                    </button>
                  </div>
                  <div className="custom-symbol-designer-inspector-actions">
                    <button type="button" data-duplicate-custom-symbol-layer onClick={() => duplicateLayer(selectedLayer.id)}><Copy size={13} />{isEn ? "Duplicate" : "复制"}</button>
                    <button type="button" className="is-danger" data-delete-custom-symbol-layer onClick={() => deleteLayer(selectedLayer.id)}><Trash2 size={13} />{isEn ? "Delete" : "删除"}</button>
                  </div>
                  {!(selectedLayer.kind === "geometry" && selectedLayer.geometryPreset === "eraser") ? (
                    <div className="custom-symbol-layer-effects" data-custom-symbol-layer-effects>
                      <div className="custom-symbol-crop-heading">
                        <strong>{isEn ? "Appearance" : "外观"}</strong>
                      </div>
                      <label className="custom-symbol-effect-toggle">
                        <input
                          type="checkbox"
                          data-custom-symbol-outline-toggle
                          checked={selectedLayer.effects?.outline?.enabled ?? false}
                          onChange={(event) =>
                            updateSelectedEffects("outline", {
                              enabled: event.currentTarget.checked,
                            })
                          }
                        />
                        <span>{isEn ? "Hollow vector outline" : "智能矢量空心"}</span>
                      </label>
                      {selectedLayer.effects?.outline?.enabled ? (
                        <div className="custom-symbol-designer-number-grid">
                          <NumericField
                            label={isEn ? "Outline width" : "空心描边粗细"}
                            field="outline-width"
                            value={selectedLayer.effects.outline.width}
                            step={2}
                            min={1}
                            max={1200}
                            onChange={(value) =>
                              updateSelectedEffects("outline", {
                                width: Math.max(1, Math.min(1200, value)),
                              })
                            }
                          />
                        </div>
                      ) : null}
                      <label className="custom-symbol-effect-toggle">
                        <input
                          type="checkbox"
                          data-custom-symbol-perspective-toggle
                          checked={selectedLayer.effects?.perspective?.enabled ?? false}
                          onChange={(event) =>
                            updateSelectedEffects("perspective", {
                              enabled: event.currentTarget.checked,
                            })
                          }
                        />
                        <span>{isEn ? "Vector extrusion / perspective" : "三维挤出透视"}</span>
                      </label>
                      {selectedLayer.effects?.perspective?.enabled ? (
                        <div className="custom-symbol-designer-number-grid">
                          <NumericField label={isEn ? "Depth" : "透视深度"} field="perspective-depth" value={selectedLayer.effects.perspective.depth} step={10} min={0} max={4000} onChange={(value) => updateSelectedEffects("perspective", { depth: Math.max(0, Math.min(4000, value)) })} />
                          <NumericField label={isEn ? "Direction °" : "透视方向 °"} field="perspective-angle" value={selectedLayer.effects.perspective.angleDeg} step={1} min={-720} max={720} onChange={(value) => updateSelectedEffects("perspective", { angleDeg: Math.max(-720, Math.min(720, value)) })} />
                          <NumericField label={isEn ? "Smoothness" : "挤出层数"} field="perspective-steps" value={selectedLayer.effects.perspective.steps} step={1} min={1} max={24} onChange={(value) => updateSelectedEffects("perspective", { steps: Math.max(1, Math.min(24, Math.round(value))) })} />
                        </div>
                      ) : null}
                    </div>
                  ) : null}
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
                            onChange={(event) => updateSelectedGeometry({ fill: event.currentTarget.checked })}
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
                        <button type="button" data-crop-preset="full" onClick={() => applyClipPreset("full")}>{isEn ? "Full" : "完整"}</button>
                        <button type="button" data-crop-preset="top" onClick={() => applyClipPreset("top")}>{isEn ? "Top" : "上"}</button>
                        <button type="button" data-crop-preset="middle" onClick={() => applyClipPreset("middle")}>{isEn ? "Middle" : "中"}</button>
                        <button type="button" data-crop-preset="bottom" onClick={() => applyClipPreset("bottom")}>{isEn ? "Bottom" : "下"}</button>
                        <button type="button" data-crop-preset="left" onClick={() => applyClipPreset("left")}>{isEn ? "Left" : "左"}</button>
                        <button type="button" data-crop-preset="center" onClick={() => applyClipPreset("center")}>{isEn ? "Center" : "中列"}</button>
                        <button type="button" data-crop-preset="right" onClick={() => applyClipPreset("right")}>{isEn ? "Right" : "右"}</button>
                      </div>
                      {selectedLayer.clipRect ? (
                        <div className="custom-symbol-designer-number-grid is-crop-grid">
                          <NumericField label="Crop X" field="crop-x" value={selectedLayer.clipRect.x} onChange={(value) => updateSelectedClip({ x: value })} />
                          <NumericField label="Crop Y" field="crop-y" value={selectedLayer.clipRect.y} onChange={(value) => updateSelectedClip({ y: value })} />
                          <NumericField label="Crop W" field="crop-width" value={selectedLayer.clipRect.width} min={1} onChange={(value) => updateSelectedClip({ width: value })} />
                          <NumericField label="Crop H" field="crop-height" value={selectedLayer.clipRect.height} min={1} onChange={(value) => updateSelectedClip({ height: value })} />
                        </div>
                      ) : null}
                      <div className="custom-symbol-split-actions">
                        <button type="button" data-split-custom-symbol-glyph="horizontal" onClick={() => splitSelectedGlyph("horizontal")}>
                          {isEn ? "Split top / middle / bottom" : "上 / 中 / 下三分"}
                        </button>
                        <button type="button" data-split-custom-symbol-glyph="vertical" onClick={() => splitSelectedGlyph("vertical")}>
                          {isEn ? "Split left / center / right" : "左 / 中 / 右三分"}
                        </button>
                      </div>
                    </div>
                  ) : null}
                </>
              ) : <div className="custom-symbol-designer-inspector-empty">{isEn ? "Select a layer." : "选择一个图层。"}</div>}
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
                    setDocumentState((current) => ({
                      ...current,
                      name: value,
                    }));
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
                    setDocumentState((current) => ({
                      ...current,
                      command: value,
                    }));
                  }}
                  placeholder={"\\selfdefa"}
                />
              </label>

              <label className="custom-symbol-registration-field">
                <span>{isEn ? "Math role" : "数学类型"}</span>
                <select
                  value={documentState.role}
                  data-custom-symbol-role-select
                  onChange={(event) => {
                    const role = event.currentTarget.value as CustomSymbolMathRole;
                    setDocumentState((current) => ({
                      ...current,
                      role,
                    }));
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
                      const limitsBehavior =
                        event.currentTarget.value as CustomSymbolLimitsBehavior;
                      setDocumentState((current) => ({
                        ...current,
                        limitsBehavior,
                      }));
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

              <label className="custom-symbol-registration-field">
                <span>{isEn ? "Word OMML fallback (optional)" : "Word OMML fallback（可选）"}</span>
                <input
                  type="text"
                  value={documentState.ommlFallback ?? ""}
                  data-custom-symbol-omml-fallback-input
                  spellCheck={false}
                  onChange={(event) => {
                    const value = event.currentTarget.value;
                    setDocumentState((current) => ({
                      ...current,
                      ommlFallback: value || null,
                    }));
                  }}
                  placeholder={"\\approx"}
                />
              </label>

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
          <span>{documentState.layers.length} {isEn ? "layers" : "图层"} · {shapeCount} {isEn ? "shapes" : "矢量单元"}</span>
          <button type="button" onClick={onClose}>{isEn ? "Close" : "关闭"}</button>
        </footer>
      </section>
    </div>,
    document.body,
  );
}
