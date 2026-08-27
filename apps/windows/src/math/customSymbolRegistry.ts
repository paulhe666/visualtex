import type { MacroDictionary, MathfieldElement } from "mathlive";
import { commandRegistry } from "../autocomplete/commandRegistry";
import { safeStorage } from "../runtime/safeStorage";
import type {
  CustomSymbolDefinition,
  CustomSymbolDesignerSourceArchive,
  CustomSymbolLayerEffects,
  CustomSymbolLibrary,
  CustomSymbolLimitsBehavior,
  CustomSymbolMathRole,
  CustomSymbolMetrics,
  CustomSymbolVectorShape,
  CustomSymbolVectorTransform,
} from "./customSymbolTypes";

export const CUSTOM_SYMBOL_STORAGE_KEY = "visualtex.custom-symbols.v1";
export const CUSTOM_SYMBOL_STORAGE_INDEX_KEY = "visualtex.custom-symbols.v2.index";
export const CUSTOM_SYMBOL_STORAGE_BACKUP_INDEX_KEY =
  "visualtex.custom-symbols.v2.backup-index";
export const CUSTOM_SYMBOL_STORAGE_RECORD_PREFIX =
  "visualtex.custom-symbols.v2.record.";
export const CUSTOM_SYMBOL_LIBRARY_VERSION = 1 as const;
export const CUSTOM_SYMBOL_LIBRARY_CHANGED_EVENT =
  "visualtex-custom-symbol-library-changed";

const CUSTOM_SYMBOL_STORAGE_INDEX_VERSION = 2 as const;
const broadcastChannelName = "visualtex-custom-symbols-v2";
const maximumShapesPerDesignerAsset = 256;
const maximumCompiledArtworkShapes = 4_096;
const maximumDesignerAssets = 96;
const maximumDesignerLayers = 192;
const maximumDesignerSourceLatexLength = 2_000;
const maximumPathLength = 24_000;
const maximumNameLength = 64;
const maximumFallbackLength = 2_000;
const legacySnapshotMaximumLength = 1_800_000;
const minimumMetric = 0.02;
const maximumWidthEm = 64;
const maximumVerticalMetricEm = 64;
const finiteCoordinateLimit = 100_000;

const mathRoles = new Set<CustomSymbolMathRole>([
  "ordinary",
  "binary",
  "relation",
  "operator",
  "open",
  "close",
  "punctuation",
]);
const limitsBehaviors = new Set<CustomSymbolLimitsBehavior>([
  "auto",
  "limits",
  "nolimits",
]);
const lineCaps = new Set(["butt", "round", "square"] as const);
const lineJoins = new Set(["miter", "round", "bevel"] as const);
const registeredVisualTexControlWords = new Set(
  commandRegistry.flatMap((command) => {
    const match = command.command.match(/^\\([A-Za-z]+)/);
    return match?.[1] ? [match[1]] : [];
  }),
);
const reservedCoreControlWords = new Set([
  "begin", "end", "left", "right", "middle",
  "class", "htmlClass", "cssId", "htmlId", "href",
  "color", "textcolor", "colorbox", "fcolorbox",
  "rule", "raise", "raisebox", "phantom", "hphantom", "vphantom",
  "mathord", "mathbin", "mathrel", "mathop", "mathopen", "mathclose", "mathpunct",
  "limits", "nolimits", "displaylimits",
  "over", "atop", "above", "choose",
  "def", "gdef", "let", "newcommand", "renewcommand", "providecommand",
  "operatorname", "operatornamewithlimits",
  "text", "mbox", "hbox", "vbox",
  "kern", "mkern", "hskip", "vskip", "hspace", "vspace",
  "style", "bbox", "unicode", "char",
]);

const prototypePartialPath =
  "M202 508Q179 508 169 520T158 547Q158 557 164 577T185 624T230 675T301 710L333 715H345Q378 715 384 714Q447 703 489 661T549 568T566 457Q566 362 519 240T402 53Q321 -22 223 -22Q123 -22 73 56Q42 102 42 148V159Q42 276 129 370T322 465Q383 465 414 434T455 367L458 378Q478 461 478 515Q478 603 437 639T344 676Q266 676 223 612Q264 606 264 572Q264 547 246 528T202 508ZM430 306Q430 372 401 400T333 428Q270 428 222 382Q197 354 183 323T150 221Q132 149 132 116Q132 21 232 21Q244 21 250 22Q327 35 374 112Q389 137 409 196T430 306Z";

/** Internal phase-2 acceptance symbol. It never enters persisted user data. */
export const CUSTOM_SYMBOL_PROTOTYPE_DEFINITION: CustomSymbolDefinition = {
  id: "vtxtestsymbol",
  command: "vtxtestsymbol",
  name: "VisualTeX prototype symbol",
  role: "ordinary",
  limitsBehavior: "auto",
  metrics: {
    widthEm: 0.566,
    ascentEm: 0.715,
    descentEm: 0.022,
  },
  artwork: {
    shapes: [
      {
        kind: "path",
        d: prototypePartialPath,
        fill: true,
        transform: {
          translateY: 715,
          scaleY: -1,
        },
      },
      {
        kind: "circle",
        cx: 290,
        cy: 365,
        r: 245,
        fill: false,
        strokeWidth: 72,
      },
      {
        kind: "line",
        x1: 65,
        y1: 365,
        x2: 515,
        y2: 365,
        fill: false,
        strokeWidth: 72,
        lineCap: "round",
      },
    ],
  },
  ommlFallback: null,
  createdAt: 0,
  updatedAt: 0,
};

interface CustomSymbolStorageIndexEntry {
  id: string;
  updatedAt: number;
  recordKey: string;
}

interface CustomSymbolStorageIndex {
  version: typeof CUSTOM_SYMBOL_STORAGE_INDEX_VERSION;
  libraryVersion: typeof CUSTOM_SYMBOL_LIBRARY_VERSION;
  revision: number;
  symbols: CustomSymbolStorageIndexEntry[];
}

let cachedStorageSignature: string | null | undefined;
let cachedUserLibrary: CustomSymbolLibrary | null = null;
let revision = 0;
const listeners = new Set<() => void>();
let broadcastChannel: BroadcastChannel | null = null;
let browserEventsInstalled = false;
let runtimeCommandAvailabilityValidator: ((command: string) => boolean) | null = null;
const appliedMathfieldCommands = new WeakMap<MathfieldElement, Set<string>>();

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function finiteNumber(value: unknown, fallback = 0) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : fallback;
}

function finiteCoordinate(value: unknown) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || Math.abs(numeric) > finiteCoordinateLimit) {
    throw new Error("Custom symbol geometry contains an invalid coordinate.");
  }
  return numeric;
}

function normalizeTransform(value: unknown): CustomSymbolVectorTransform | undefined {
  if (!isRecord(value)) return undefined;
  const result: CustomSymbolVectorTransform = {};
  const numericKeys = [
    "translateX",
    "translateY",
    "scaleX",
    "scaleY",
    "skewXDeg",
    "skewYDeg",
    "rotateDeg",
    "originX",
    "originY",
  ] as const;
  for (const key of numericKeys) {
    if (value[key] === undefined) continue;
    result[key] = finiteCoordinate(value[key]);
  }
  if (result.scaleX === 0 || result.scaleY === 0) {
    throw new Error("Custom symbol geometry cannot use a zero scale.");
  }
  if (value.matrix !== undefined) {
    if (!Array.isArray(value.matrix) || value.matrix.length !== 6) {
      throw new Error("Custom symbol geometry contains an invalid transform matrix.");
    }
    const matrix = value.matrix.map(finiteCoordinate) as [
      number,
      number,
      number,
      number,
      number,
      number,
    ];
    const determinant = matrix[0] * matrix[3] - matrix[1] * matrix[2];
    if (Math.abs(determinant) < 0.0000001) {
      throw new Error("Custom symbol geometry cannot use a singular transform matrix.");
    }
    result.matrix = matrix;
  }
  return Object.keys(result).length ? result : undefined;
}

function normalizeClipRect(value: unknown) {
  if (!isRecord(value)) return undefined;
  const x = finiteCoordinate(value.x);
  const y = finiteCoordinate(value.y);
  const width = finiteCoordinate(value.width);
  const height = finiteCoordinate(value.height);
  if (width <= 0 || height <= 0) {
    throw new Error("Custom symbol clip rectangles must have positive size.");
  }
  return { x, y, width, height };
}

function normalizeShape(value: unknown): CustomSymbolVectorShape {
  if (!isRecord(value) || typeof value.kind !== "string") {
    throw new Error("Custom symbol artwork contains an invalid shape.");
  }
  const common = {
    ...(value.operation === "erase" ? { operation: "erase" as const } : {}),
    ...(typeof value.fill === "boolean" ? { fill: value.fill } : {}),
    ...(value.strokeWidth === undefined
      ? {}
      : { strokeWidth: Math.max(0, finiteCoordinate(value.strokeWidth)) }),
    ...(lineCaps.has(value.lineCap as "butt" | "round" | "square")
      ? { lineCap: value.lineCap as "butt" | "round" | "square" }
      : {}),
    ...(lineJoins.has(value.lineJoin as "miter" | "round" | "bevel")
      ? { lineJoin: value.lineJoin as "miter" | "round" | "bevel" }
      : {}),
    ...(value.transform === undefined
      ? {}
      : { transform: normalizeTransform(value.transform) }),
    ...(value.clipRect === undefined
      ? {}
      : { clipRect: normalizeClipRect(value.clipRect) }),
  };

  switch (value.kind) {
    case "path": {
      const d = typeof value.d === "string" ? value.d.trim() : "";
      if (!d || d.length > maximumPathLength) {
        throw new Error("Custom symbol path data is empty or too large.");
      }
      if (!/^[MmZzLlHhVvCcSsQqTtAaEe0-9+.,\-\s]+$/.test(d)) {
        throw new Error("Custom symbol path contains unsupported SVG syntax.");
      }
      return { kind: "path", d, ...common };
    }
    case "circle":
      return {
        kind: "circle",
        cx: finiteCoordinate(value.cx),
        cy: finiteCoordinate(value.cy),
        r: Math.max(0, finiteCoordinate(value.r)),
        ...common,
      };
    case "line":
      return {
        kind: "line",
        x1: finiteCoordinate(value.x1),
        y1: finiteCoordinate(value.y1),
        x2: finiteCoordinate(value.x2),
        y2: finiteCoordinate(value.y2),
        ...common,
      };
    case "rect":
      return {
        kind: "rect",
        x: finiteCoordinate(value.x),
        y: finiteCoordinate(value.y),
        width: Math.max(0, finiteCoordinate(value.width)),
        height: Math.max(0, finiteCoordinate(value.height)),
        ...(value.rx === undefined ? {} : { rx: Math.max(0, finiteCoordinate(value.rx)) }),
        ...(value.ry === undefined ? {} : { ry: Math.max(0, finiteCoordinate(value.ry)) }),
        ...common,
      };
    case "ellipse":
      return {
        kind: "ellipse",
        cx: finiteCoordinate(value.cx),
        cy: finiteCoordinate(value.cy),
        rx: Math.max(0, finiteCoordinate(value.rx)),
        ry: Math.max(0, finiteCoordinate(value.ry)),
        ...common,
      };
    case "polygon": {
      if (!Array.isArray(value.points) || value.points.length < 2 || value.points.length > 256) {
        throw new Error("Custom symbol polygon has an invalid number of points.");
      }
      return {
        kind: "polygon",
        points: value.points.map((point) => {
          if (!Array.isArray(point) || point.length !== 2) {
            throw new Error("Custom symbol polygon contains an invalid point.");
          }
          return [finiteCoordinate(point[0]), finiteCoordinate(point[1])];
        }),
        ...common,
      };
    }
    case "text": {
      const text = typeof value.text === "string" ? value.text.normalize("NFC") : "";
      const fontFamily =
        typeof value.fontFamily === "string" ? value.fontFamily.trim() : "";
      const fontSize = finiteNumber(value.fontSize, NaN);
      const fontWeight =
        value.fontWeight === undefined
          ? undefined
          : Math.round(finiteNumber(value.fontWeight, NaN));
      if (
        !text ||
        Array.from(text).length > 16 ||
        /[\u0000-\u001f\u007f]/.test(text) ||
        !fontFamily ||
        fontFamily.length > 128 ||
        /[\u0000-\u001f\u007f"'<>]/.test(fontFamily) ||
        !Number.isFinite(fontSize) ||
        fontSize < 20 ||
        fontSize > 8_000 ||
        (fontWeight !== undefined &&
          (!Number.isFinite(fontWeight) || fontWeight < 100 || fontWeight > 900))
      ) {
        throw new Error("Custom symbol system-font glyph is invalid or too large.");
      }
      return {
        kind: "text",
        text,
        x: finiteCoordinate(value.x),
        y: finiteCoordinate(value.y),
        fontFamily,
        fontSize,
        ...(value.fontStyle === "italic" || value.fontStyle === "normal"
          ? { fontStyle: value.fontStyle }
          : {}),
        ...(fontWeight === undefined ? {} : { fontWeight }),
        ...common,
      };
    }
    default:
      throw new Error(`Unsupported custom symbol shape: ${value.kind}`);
  }
}

function normalizeMetrics(value: unknown): CustomSymbolMetrics {
  if (!isRecord(value)) throw new Error("Custom symbol metrics are missing.");
  const widthEm = finiteNumber(value.widthEm, NaN);
  const ascentEm = finiteNumber(value.ascentEm, NaN);
  const descentEm = finiteNumber(value.descentEm, NaN);
  if (
    !Number.isFinite(widthEm) ||
    !Number.isFinite(ascentEm) ||
    !Number.isFinite(descentEm) ||
    widthEm < minimumMetric ||
    widthEm > maximumWidthEm ||
    ascentEm < minimumMetric ||
    ascentEm > maximumVerticalMetricEm ||
    descentEm < 0 ||
    descentEm > maximumVerticalMetricEm
  ) {
    throw new Error("Custom symbol metrics are outside the supported range.");
  }
  return { widthEm, ascentEm, descentEm };
}

function normalizeCommandName(value: unknown) {
  const command = typeof value === "string" ? value.trim().replace(/^\\/, "") : "";
  if (!/^[A-Za-z]+$/.test(command)) {
    throw new Error(
      "Custom symbol commands must be TeX control words containing letters only.",
    );
  }
  if (command.length > 48) {
    throw new Error("Custom symbol command names are too long.");
  }
  return command;
}

function isReservedControlWord(command: string) {
  return (
    registeredVisualTexControlWords.has(command) ||
    reservedCoreControlWords.has(command)
  );
}

function normalizeDesignerLayerId(value: unknown, kind: string) {
  const id = typeof value === "string" ? value.trim() : "";
  if (!/^[A-Za-z0-9][A-Za-z0-9._:-]{0,95}$/.test(id)) {
    throw new Error(`Custom symbol ${kind} has an invalid identifier.`);
  }
  return id;
}

function normalizeDesignerLayerEffects(
  value: unknown,
): CustomSymbolLayerEffects | undefined {
  if (!isRecord(value)) return undefined;
  const effects: CustomSymbolLayerEffects = {};
  if (isRecord(value.outline)) {
    const width = finiteNumber(value.outline.width, 30);
    if (width < 1 || width > 1_200) {
      throw new Error("Custom symbol outline width is outside the supported range.");
    }
    effects.outline = {
      enabled: value.outline.enabled === true,
      width,
    };
  }
  if (isRecord(value.perspective)) {
    const depth = finiteNumber(value.perspective.depth, 240);
    const angleDeg = finiteNumber(value.perspective.angleDeg, 35);
    const steps = Math.round(finiteNumber(value.perspective.steps, 8));
    if (
      depth < 0 ||
      depth > 4_000 ||
      angleDeg < -720 ||
      angleDeg > 720 ||
      steps < 1 ||
      steps > 24
    ) {
      throw new Error("Custom symbol perspective parameters are outside the supported range.");
    }
    effects.perspective = {
      enabled: value.perspective.enabled === true,
      depth,
      angleDeg,
      steps,
    };
  }
  return Object.keys(effects).length ? effects : undefined;
}

function normalizeDesignerSource(
  value: unknown,
): CustomSymbolDesignerSourceArchive | null {
  if (value == null) return null;
  if (!isRecord(value) || value.version !== 1) {
    throw new Error("Custom symbol designer source has an unsupported version.");
  }
  if (
    !Array.isArray(value.assets) ||
    !Array.isArray(value.layers) ||
    value.assets.length > maximumDesignerAssets ||
    value.layers.length > maximumDesignerLayers
  ) {
    throw new Error("Custom symbol designer source is too large.");
  }

  const assetIds = new Set<string>();
  const assets = value.assets.map((candidate) => {
    if (!isRecord(candidate)) {
      throw new Error("Custom symbol designer asset must be an object.");
    }
    const id = normalizeDesignerLayerId(candidate.id, "designer asset");
    if (assetIds.has(id)) {
      throw new Error("Custom symbol designer source has duplicate asset IDs.");
    }
    assetIds.add(id);
    const sourceLatex =
      typeof candidate.sourceLatex === "string" ? candidate.sourceLatex.trim() : "";
    if (
      !sourceLatex ||
      sourceLatex.length > maximumDesignerSourceLatexLength ||
      !Array.isArray(candidate.shapes) ||
      candidate.shapes.length === 0 ||
      candidate.shapes.length > maximumShapesPerDesignerAsset
    ) {
      throw new Error("Custom symbol designer asset is invalid or too large.");
    }
    return {
      id,
      sourceLatex,
      metrics: normalizeMetrics(candidate.metrics),
      shapes: candidate.shapes.map(normalizeShape),
    };
  });

  const layerIds = new Set<string>();
  const layers = value.layers.map((candidate) => {
    if (!isRecord(candidate)) {
      throw new Error("Custom symbol designer layer must be an object.");
    }
    const id = normalizeDesignerLayerId(candidate.id, "designer layer");
    if (layerIds.has(id)) {
      throw new Error("Custom symbol designer source has duplicate layer IDs.");
    }
    layerIds.add(id);
    const name = typeof candidate.name === "string" ? candidate.name.trim() : "";
    if (!name || name.length > maximumNameLength) {
      throw new Error("Custom symbol designer layer name is invalid.");
    }
    const base = {
      id,
      name,
      visible: candidate.visible !== false,
      locked: candidate.locked === true,
      transform: normalizeTransform(candidate.transform) ?? {},
      ...(normalizeDesignerLayerEffects(candidate.effects)
        ? { effects: normalizeDesignerLayerEffects(candidate.effects) }
        : {}),
      ...(candidate.clipRect == null
        ? {}
        : { clipRect: normalizeClipRect(candidate.clipRect) }),
    };
    if (candidate.kind === "glyph") {
      const assetId = normalizeDesignerLayerId(candidate.assetId, "designer asset reference");
      if (!assetIds.has(assetId)) {
        throw new Error("Custom symbol designer layer references a missing glyph asset.");
      }
      return { ...base, kind: "glyph" as const, assetId };
    }
    if (candidate.kind === "geometry") {
      if (!isRecord(candidate.bounds)) {
        throw new Error("Custom symbol geometry layer is missing designer bounds.");
      }
      const bounds = {
        x: finiteCoordinate(candidate.bounds.x),
        y: finiteCoordinate(candidate.bounds.y),
        width: finiteCoordinate(candidate.bounds.width),
        height: finiteCoordinate(candidate.bounds.height),
      };
      if (bounds.width <= 0 || bounds.height <= 0) {
        throw new Error("Custom symbol geometry bounds must have positive size.");
      }
      const geometryPresets = new Set([
        "line",
        "circle",
        "ellipse",
        "rect",
        "triangle",
        "arrow",
        "arc",
        "eraser",
      ]);
      const geometryPreset =
        typeof candidate.geometryPreset === "string" &&
        geometryPresets.has(candidate.geometryPreset)
          ? (candidate.geometryPreset as
              | "line"
              | "circle"
              | "ellipse"
              | "rect"
              | "triangle"
              | "arrow"
              | "arc"
              | "eraser")
          : undefined;
      return {
        ...base,
        kind: "geometry" as const,
        ...(geometryPreset ? { geometryPreset } : {}),
        shape: normalizeShape(candidate.shape),
        bounds,
      };
    }
    throw new Error("Custom symbol designer layer has an invalid kind.");
  });

  return {
    version: 1,
    ...(value.metrics == null ? {} : { metrics: normalizeMetrics(value.metrics) }),
    assets,
    layers,
  };
}

function normalizeDefinition(
  value: unknown,
  options: { allowPrototypeCommand?: boolean } = {},
): CustomSymbolDefinition {
  if (!isRecord(value)) throw new Error("Custom symbol definition must be an object.");
  const id = typeof value.id === "string" ? value.id.trim() : "";
  if (!/^[A-Za-z0-9][A-Za-z0-9._:-]{0,95}$/.test(id)) {
    throw new Error("Custom symbol has an invalid identifier.");
  }
  const command = normalizeCommandName(value.command);
  if (
    !options.allowPrototypeCommand &&
    command === CUSTOM_SYMBOL_PROTOTYPE_DEFINITION.command
  ) {
    throw new Error("The VisualTeX prototype command is reserved.");
  }
  if (!options.allowPrototypeCommand && isReservedControlWord(command)) {
    throw new Error(`\\${command} is already reserved by VisualTeX/LaTeX.`);
  }
  if (
    !options.allowPrototypeCommand &&
    runtimeCommandAvailabilityValidator &&
    !runtimeCommandAvailabilityValidator(command)
  ) {
    throw new Error(`\\${command} is already defined by MathLive/LaTeX.`);
  }
  const name = typeof value.name === "string" ? value.name.trim() : "";
  if (!name || name.length > maximumNameLength) {
    throw new Error("Custom symbol name is empty or too long.");
  }
  if (!mathRoles.has(value.role as CustomSymbolMathRole)) {
    throw new Error("Custom symbol has an invalid mathematical role.");
  }
  if (!limitsBehaviors.has(value.limitsBehavior as CustomSymbolLimitsBehavior)) {
    throw new Error("Custom symbol has an invalid limits behavior.");
  }
  if (!isRecord(value.artwork) || !Array.isArray(value.artwork.shapes)) {
    throw new Error("Custom symbol artwork is missing.");
  }
  if (
    value.artwork.shapes.length === 0 ||
    value.artwork.shapes.length > maximumCompiledArtworkShapes
  ) {
    throw new Error("Custom symbol artwork has an invalid number of shapes.");
  }
  const ommlFallback =
    value.ommlFallback == null
      ? null
      : typeof value.ommlFallback === "string"
        ? value.ommlFallback.trim()
        : null;
  if (ommlFallback && ommlFallback.length > maximumFallbackLength) {
    throw new Error("Custom symbol Word fallback is too large.");
  }
  const now = Date.now();
  return {
    id,
    command,
    name,
    role: value.role as CustomSymbolMathRole,
    limitsBehavior: value.limitsBehavior as CustomSymbolLimitsBehavior,
    metrics: normalizeMetrics(value.metrics),
    artwork: {
      shapes: value.artwork.shapes.map(normalizeShape),
    },
    ommlFallback,
    designerSource: normalizeDesignerSource(value.designerSource),
    createdAt: Number.isFinite(Number(value.createdAt)) ? Number(value.createdAt) : now,
    updatedAt: Number.isFinite(Number(value.updatedAt)) ? Number(value.updatedAt) : now,
  };
}

function emptyUserLibrary(): CustomSymbolLibrary {
  return { version: CUSTOM_SYMBOL_LIBRARY_VERSION, symbols: [] };
}

export function normalizeCustomSymbolLibrary(value: unknown): CustomSymbolLibrary {
  if (!isRecord(value) || value.version !== CUSTOM_SYMBOL_LIBRARY_VERSION) {
    return emptyUserLibrary();
  }
  if (!Array.isArray(value.symbols)) return emptyUserLibrary();
  const ids = new Set<string>();
  const commands = new Set<string>();
  const symbols: CustomSymbolDefinition[] = [];
  for (const candidate of value.symbols) {
    try {
      const symbol = normalizeDefinition(candidate);
      if (ids.has(symbol.id) || commands.has(symbol.command)) continue;
      ids.add(symbol.id);
      commands.add(symbol.command);
      symbols.push(symbol);
    } catch {
      // Invalid persisted records are ignored without making startup fail.
    }
  }
  return { version: CUSTOM_SYMBOL_LIBRARY_VERSION, symbols };
}

function parseStorageIndex(raw: string | null): CustomSymbolStorageIndex | null {
  if (!raw) return null;
  try {
    const value: unknown = JSON.parse(raw);
    if (
      !isRecord(value) ||
      value.version !== CUSTOM_SYMBOL_STORAGE_INDEX_VERSION ||
      value.libraryVersion !== CUSTOM_SYMBOL_LIBRARY_VERSION ||
      !Array.isArray(value.symbols)
    ) {
      return null;
    }
    const ids = new Set<string>();
    const recordKeys = new Set<string>();
    const symbols: CustomSymbolStorageIndexEntry[] = [];
    for (const candidate of value.symbols) {
      if (!isRecord(candidate)) continue;
      const id = typeof candidate.id === "string" ? candidate.id.trim() : "";
      const updatedAt = Number(candidate.updatedAt);
      const recordKey =
        typeof candidate.recordKey === "string" ? candidate.recordKey : "";
      if (
        !/^[A-Za-z0-9][A-Za-z0-9._:-]{0,95}$/.test(id) ||
        !Number.isFinite(updatedAt) ||
        !recordKey.startsWith(CUSTOM_SYMBOL_STORAGE_RECORD_PREFIX) ||
        recordKey.length > 320 ||
        ids.has(id) ||
        recordKeys.has(recordKey)
      ) {
        return null;
      }
      ids.add(id);
      recordKeys.add(recordKey);
      symbols.push({ id, updatedAt, recordKey });
    }
    return {
      version: CUSTOM_SYMBOL_STORAGE_INDEX_VERSION,
      libraryVersion: CUSTOM_SYMBOL_LIBRARY_VERSION,
      revision: Number.isFinite(Number(value.revision))
        ? Number(value.revision)
        : 0,
      symbols,
    };
  } catch {
    return null;
  }
}

function readLegacyLibrary(raw: string | null) {
  if (!raw) return emptyUserLibrary();
  try {
    return normalizeCustomSymbolLibrary(JSON.parse(raw));
  } catch {
    return emptyUserLibrary();
  }
}

function readIndexedLibrary(
  index: CustomSymbolStorageIndex,
): CustomSymbolLibrary | null {
  const candidates: unknown[] = [];
  for (const entry of index.symbols) {
    const raw = safeStorage.getItem(entry.recordKey);
    if (!raw) return null;
    try {
      candidates.push(JSON.parse(raw));
    } catch {
      return null;
    }
  }
  const library = normalizeCustomSymbolLibrary({
    version: CUSTOM_SYMBOL_LIBRARY_VERSION,
    symbols: candidates,
  });
  if (library.symbols.length !== index.symbols.length) return null;
  for (let position = 0; position < library.symbols.length; position += 1) {
    const symbol = library.symbols[position];
    const entry = index.symbols[position];
    if (symbol.id !== entry.id || symbol.updatedAt !== entry.updatedAt) {
      return null;
    }
  }
  return library;
}

function storageSignature(
  indexRaw: string | null,
  backupIndexRaw: string | null,
  legacyRaw: string | null,
) {
  return `index:${indexRaw ?? ""}|backup:${backupIndexRaw ?? ""}|legacy:${
    legacyRaw ?? ""
  }`;
}

function readStoredUserLibrary() {
  const indexRaw = safeStorage.getItem(CUSTOM_SYMBOL_STORAGE_INDEX_KEY);
  const backupIndexRaw = safeStorage.getItem(
    CUSTOM_SYMBOL_STORAGE_BACKUP_INDEX_KEY,
  );
  const legacyRaw = safeStorage.getItem(CUSTOM_SYMBOL_STORAGE_KEY);
  const signature = storageSignature(indexRaw, backupIndexRaw, legacyRaw);
  if (cachedUserLibrary && signature === cachedStorageSignature) {
    return cachedUserLibrary;
  }
  cachedStorageSignature = signature;

  const index = parseStorageIndex(indexRaw);
  const indexed = index ? readIndexedLibrary(index) : null;
  if (indexed) {
    cachedUserLibrary = indexed;
    return cachedUserLibrary;
  }

  const backupIndex = parseStorageIndex(backupIndexRaw);
  const backup = backupIndex ? readIndexedLibrary(backupIndex) : null;
  if (backup) {
    if (indexRaw) {
      console.warn(
        "VisualTeX recovered the previous complete custom-symbol library because the current indexed generation was incomplete.",
      );
    }
    cachedUserLibrary = backup;
    return cachedUserLibrary;
  }

  cachedUserLibrary = readLegacyLibrary(legacyRaw);
  return cachedUserLibrary;
}

function notifyLocalChange() {
  revision += 1;
  listeners.forEach((listener) => listener());
  if (typeof window !== "undefined") {
    window.dispatchEvent(new CustomEvent(CUSTOM_SYMBOL_LIBRARY_CHANGED_EVENT));
  }
}

function customSymbolStorageEventKey(key: string | null) {
  return (
    key === CUSTOM_SYMBOL_STORAGE_KEY ||
    key === CUSTOM_SYMBOL_STORAGE_INDEX_KEY ||
    key === CUSTOM_SYMBOL_STORAGE_BACKUP_INDEX_KEY ||
    Boolean(key?.startsWith(CUSTOM_SYMBOL_STORAGE_RECORD_PREFIX))
  );
}

function ensureBrowserEvents() {
  if (browserEventsInstalled || typeof window === "undefined") return;
  browserEventsInstalled = true;
  window.addEventListener("storage", (event) => {
    if (!customSymbolStorageEventKey(event.key)) return;
    cachedStorageSignature = undefined;
    cachedUserLibrary = null;
    notifyLocalChange();
  });
  if (typeof BroadcastChannel !== "undefined") {
    broadcastChannel = new BroadcastChannel(broadcastChannelName);
    broadcastChannel.addEventListener("message", () => {
      cachedStorageSignature = undefined;
      cachedUserLibrary = null;
      notifyLocalChange();
    });
  }
}

function nextRecordKey(symbol: CustomSymbolDefinition, token: string) {
  return `${CUSTOM_SYMBOL_STORAGE_RECORD_PREFIX}${encodeURIComponent(symbol.id)}.${token}`;
}

function persistUserLibrary(library: CustomSymbolLibrary) {
  const normalized = normalizeCustomSymbolLibrary(library);
  const currentIndexRaw = safeStorage.getItem(CUSTOM_SYMBOL_STORAGE_INDEX_KEY);
  const backupIndexRaw = safeStorage.getItem(
    CUSTOM_SYMBOL_STORAGE_BACKUP_INDEX_KEY,
  );
  const currentIndex = parseStorageIndex(currentIndexRaw);
  const backupIndex = parseStorageIndex(backupIndexRaw);
  const currentComplete = currentIndex ? readIndexedLibrary(currentIndex) : null;
  const backupComplete = backupIndex ? readIndexedLibrary(backupIndex) : null;
  const baseIndex = currentComplete
    ? currentIndex
    : backupComplete
      ? backupIndex
      : null;
  const baseIndexRaw = currentComplete
    ? currentIndexRaw
    : backupComplete
      ? backupIndexRaw
      : null;
  const previousEntries = new Map(
    (baseIndex?.symbols ?? []).map((entry) => [entry.id, entry]),
  );
  const writeToken = `${Date.now().toString(36)}-${Math.random()
    .toString(36)
    .slice(2, 10)}`;
  const nextEntries: CustomSymbolStorageIndexEntry[] = [];
  const nextRecordKeys = new Set<string>();
  const createdRecordKeys: string[] = [];
  let index: CustomSymbolStorageIndex;
  let indexSerialized: string;

  try {
    for (const symbol of normalized.symbols) {
      const previous = previousEntries.get(symbol.id);
      const serializedSymbol = JSON.stringify(symbol);
      const previousRecord = previous
        ? safeStorage.getItem(previous.recordKey)
        : null;
      const canReuse =
        previous?.updatedAt === symbol.updatedAt &&
        previousRecord === serializedSymbol;
      const recordKey = canReuse
        ? previous.recordKey
        : nextRecordKey(symbol, writeToken);
      if (!canReuse) {
        safeStorage.setItemStrict(recordKey, serializedSymbol);
        createdRecordKeys.push(recordKey);
      }
      nextEntries.push({
        id: symbol.id,
        updatedAt: symbol.updatedAt,
        recordKey,
      });
      nextRecordKeys.add(recordKey);
    }

    index = {
      version: CUSTOM_SYMBOL_STORAGE_INDEX_VERSION,
      libraryVersion: CUSTOM_SYMBOL_LIBRARY_VERSION,
      revision: Math.max(Date.now(), (baseIndex?.revision ?? 0) + 1),
      symbols: nextEntries,
    };
    indexSerialized = JSON.stringify(index);
    if (baseIndex && baseIndexRaw) {
      safeStorage.setItemStrict(
        CUSTOM_SYMBOL_STORAGE_BACKUP_INDEX_KEY,
        baseIndexRaw,
      );
    } else {
      safeStorage.removeItem(CUSTOM_SYMBOL_STORAGE_BACKUP_INDEX_KEY);
    }
    safeStorage.setItemStrict(CUSTOM_SYMBOL_STORAGE_INDEX_KEY, indexSerialized);
  } catch (error) {
    for (const recordKey of createdRecordKeys) {
      safeStorage.removeItem(recordKey);
    }
    throw error;
  }

  const retainedRecordKeys = new Set(nextRecordKeys);
  for (const entry of baseIndex?.symbols ?? []) {
    retainedRecordKeys.add(entry.recordKey);
  }
  const oldEntries = [
    ...(currentIndex?.symbols ?? []),
    ...(backupIndex?.symbols ?? []),
  ];
  for (const previous of oldEntries) {
    if (!retainedRecordKeys.has(previous.recordKey)) {
      safeStorage.removeItem(previous.recordKey);
    }
  }

  const legacySnapshot = JSON.stringify(normalized);
  const storedLegacySnapshot =
    legacySnapshot.length <= legacySnapshotMaximumLength
      ? legacySnapshot
      : null;
  if (storedLegacySnapshot) {
    safeStorage.setItem(CUSTOM_SYMBOL_STORAGE_KEY, storedLegacySnapshot);
  } else {
    safeStorage.removeItem(CUSTOM_SYMBOL_STORAGE_KEY);
  }

  cachedStorageSignature = storageSignature(
    indexSerialized,
    baseIndexRaw,
    storedLegacySnapshot,
  );
  cachedUserLibrary = normalized;
  notifyLocalChange();
  broadcastChannel?.postMessage({ revision: index.revision });
  return normalized;
}

export function setCustomSymbolCommandAvailabilityValidator(
  validator: ((command: string) => boolean) | null,
) {
  runtimeCommandAvailabilityValidator = validator;
  cachedStorageSignature = undefined;
  cachedUserLibrary = null;
}

export function readCustomSymbolLibrary() {
  ensureBrowserEvents();
  return readStoredUserLibrary();
}

export function getActiveCustomSymbols() {
  return [CUSTOM_SYMBOL_PROTOTYPE_DEFINITION, ...readCustomSymbolLibrary().symbols];
}

export function getCustomSymbolRevision() {
  ensureBrowserEvents();
  return revision;
}

export function subscribeCustomSymbols(listener: () => void) {
  ensureBrowserEvents();
  listeners.add(listener);
  return () => listeners.delete(listener);
}

export function findCustomSymbolByCommand(command: string) {
  const normalized = command.trim().replace(/^\\/, "");
  return getActiveCustomSymbols().find((symbol) => symbol.command === normalized) ?? null;
}

export function registerCustomSymbol(value: unknown) {
  const symbol = normalizeDefinition(value);
  const library = readCustomSymbolLibrary();
  if (library.symbols.some((candidate) => candidate.id === symbol.id)) {
    throw new Error("A custom symbol with this identifier already exists.");
  }
  if (library.symbols.some((candidate) => candidate.command === symbol.command)) {
    throw new Error(`\\${symbol.command} is already registered as a custom symbol.`);
  }
  return persistUserLibrary({
    version: CUSTOM_SYMBOL_LIBRARY_VERSION,
    symbols: [...library.symbols, symbol],
  });
}

export function updateCustomSymbol(id: string, patch: unknown) {
  const library = readCustomSymbolLibrary();
  const index = library.symbols.findIndex((symbol) => symbol.id === id);
  if (index < 0) throw new Error("Custom symbol does not exist.");
  const current = library.symbols[index];
  const merged = normalizeDefinition({
    ...current,
    ...(isRecord(patch) ? patch : {}),
    id: current.id,
    createdAt: current.createdAt,
    updatedAt: Date.now(),
  });
  if (
    library.symbols.some(
      (symbol, candidateIndex) =>
        candidateIndex !== index && symbol.command === merged.command,
    )
  ) {
    throw new Error(`\\${merged.command} is already registered as a custom symbol.`);
  }
  const symbols = [...library.symbols];
  symbols[index] = merged;
  return persistUserLibrary({ version: CUSTOM_SYMBOL_LIBRARY_VERSION, symbols });
}

export function deleteCustomSymbol(id: string) {
  const library = readCustomSymbolLibrary();
  const symbols = library.symbols.filter((symbol) => symbol.id !== id);
  if (symbols.length === library.symbols.length) return library;
  return persistUserLibrary({ version: CUSTOM_SYMBOL_LIBRARY_VERSION, symbols });
}

export function replaceCustomSymbolLibrary(value: unknown) {
  return persistUserLibrary(normalizeCustomSymbolLibrary(value));
}

export function refreshCustomSymbolLibraryFromStorage() {
  cachedStorageSignature = undefined;
  cachedUserLibrary = null;
  notifyLocalChange();
  broadcastChannel?.postMessage({ revision: Date.now() });
}

function formatEm(value: number) {
  const normalized = Math.abs(value) < 0.000001 ? 0 : value;
  return Number(normalized.toFixed(4)).toString();
}

function roleWrapper(role: CustomSymbolMathRole) {
  switch (role) {
    case "binary":
      return "mathbin";
    case "relation":
      return "mathrel";
    case "operator":
      return "mathop";
    case "open":
      return "mathopen";
    case "close":
      return "mathclose";
    case "punctuation":
      return "mathpunct";
    default:
      return "mathord";
  }
}

export function customSymbolCssClass(symbol: CustomSymbolDefinition) {
  return `visualtex-custom-symbol-${symbol.id.replace(/[^A-Za-z0-9_-]/g, "-")}`;
}

export function customSymbolSvgMarkerClass(symbol: CustomSymbolDefinition) {
  return `visualtex-custom-symbol-export-${symbol.id.replace(/[^A-Za-z0-9_-]/g, "-")}`;
}

function metricProbeLatex(symbol: CustomSymbolDefinition, className: string) {
  const { widthEm, ascentEm, descentEm } = symbol.metrics;
  const thickness = ascentEm + descentEm;
  return (
    `\\class{${className}}{` +
    `\\phantom{\\rule[-${formatEm(descentEm)}em]{${formatEm(widthEm)}em}{${formatEm(thickness)}em}}` +
    "}"
  );
}

export function customSymbolMathLiveMacroDefinition(symbol: CustomSymbolDefinition) {
  const wrapper = roleWrapper(symbol.role);
  const body = `\\${wrapper}{${metricProbeLatex(symbol, customSymbolCssClass(symbol))}}`;
  const limits =
    symbol.role === "operator" && symbol.limitsBehavior !== "auto"
      ? `\\${symbol.limitsBehavior}`
      : "";
  return {
    def: body + limits,
    args: 0,
    expand: false,
    captureSelection: true,
  } as const;
}

export function customSymbolSvgMacro(symbol: CustomSymbolDefinition) {
  const wrapper = roleWrapper(symbol.role);
  const body = `\\${wrapper}{${metricProbeLatex(symbol, customSymbolSvgMarkerClass(symbol))}}`;
  const limits =
    symbol.role === "operator" && symbol.limitsBehavior !== "auto"
      ? `\\${symbol.limitsBehavior}`
      : "";
  return body + limits;
}

export function customSymbolMathLiveMacros() {
  return Object.fromEntries(
    getActiveCustomSymbols().map((symbol) => [
      symbol.command,
      customSymbolMathLiveMacroDefinition(symbol),
    ]),
  );
}

export function addCustomSymbolMacros<T extends Record<string, unknown>>(macros: T) {
  return {
    ...macros,
    ...customSymbolMathLiveMacros(),
  };
}

export function getAppliedCustomSymbolCommandsForMathfield(
  field: MathfieldElement,
) {
  return new Set(appliedMathfieldCommands.get(field) ?? []);
}

export function composeCustomSymbolMacrosForMathfield(
  field: MathfieldElement,
  baseMacros: MacroDictionary,
  additionalMacros: MacroDictionary = {},
) {
  const macros: MacroDictionary = { ...baseMacros };
  const previous = appliedMathfieldCommands.get(field);
  previous?.forEach((command) => delete macros[command]);
  Object.assign(macros, additionalMacros);
  const customMacros = customSymbolMathLiveMacros();
  Object.assign(macros, customMacros);
  appliedMathfieldCommands.set(field, new Set(Object.keys(customMacros)));
  return macros;
}

export function applyCustomSymbolMacrosToMathfield(field: MathfieldElement) {
  field.macros = composeCustomSymbolMacrosForMathfield(field, field.macros);
}

export function customSymbolSvgMacros() {
  return Object.fromEntries(
    getActiveCustomSymbols().map((symbol) => [symbol.command, customSymbolSvgMacro(symbol)]),
  ) as Record<string, string>;
}
