import {
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent as ReactKeyboardEvent,
  type MouseEvent,
  type PointerEvent as ReactPointerEvent,
} from "react";
import {
  Check,
  ChevronRight,
  FolderPlus,
  Keyboard,
  MoreHorizontal,
  Palette,
  PanelRightClose,
  Pencil,
  Plus,
  RotateCcw,
  SlidersHorizontal,
  Trash2,
  X,
} from "lucide-react";
import { createPortal } from "react-dom";
import { create } from "zustand";
import type { LatexCommand } from "../../../shared/visualtex-app/src/types/command";
import {
  categoryLabels,
  categoryLabelsEn,
  commandRegistry,
} from "../autocomplete/commandRegistry";
import { MathPreview } from "../components/MathPreview";
import { FormulaTileLibraryDialog } from "./FormulaTileLibraryDialog";
import {
  defaultCommonFormulaTileIds,
  formulaTileLibraryById,
} from "../../../shared/visualtex-app/src/toolbar/formulaTileLibrary";
import { FormulaHotkeyRecorderDialog } from "../components/FormulaHotkeyRecorderDialog";
import { customSymbolCommands } from "../autocomplete/runtimeCommandRegistry";
import {
  deleteCustomSymbol,
  readCustomSymbolLibrary,
} from "../math/customSymbolRegistry";
import { useCustomSymbolRevision } from "../math/customSymbolReact";
import {
  createFormulaHotkeyTarget,
  formatFormulaHotkeyChord,
  formulaHotkeyTargetIdForCommand,
  formulaHotkeyTargetIdForTile,
  formulaHotkeyTargetLabel,
  type FormulaHotkeyTarget,
} from "../shortcuts/formulaHotkeys";
import { useEditorStore } from "../stores/editorStore";
import { writeLocalStorage } from "../../../shared/visualtex-app/src/runtime/safeStorage";
import { createLayoutOscillationGuard } from "../../../shared/visualtex-app/src/runtime/layoutOscillationGuard";
import { useFormulaHotkeyStore } from "../stores/formulaHotkeyStore";
import { useCustomSymbolDesignerStore } from "../../../shared/visualtex-app/src/stores/customSymbolDesignerStore";
import {
  BUILT_IN_GROUP_IDS,
  groupItems,
  sameToolbarItem,
  useToolbarLayoutStore,
  type BuiltInGroupId,
  type ToolbarGroup,
  type ToolbarItem,
} from "./toolbarLayout";
import { armDrag, END_INDEX, type DragHandlers, type DragSource, type DropTarget } from "../../../shared/visualtex-app/src/toolbar/toolbarDrag";
import { ToolbarItemPickerDialog } from "./ToolbarItemPickerDialog";

type MatrixDelimiter =
  | "matrix"
  | "pmatrix"
  | "bmatrix"
  | "Bmatrix"
  | "vmatrix"
  | "Vmatrix";
type ToolbarView = "tools" | "tiles";
type ToolbarLayout = "sidebar" | "horizontal";
type TileCategory = "custom" | "common";

interface Props {
  onInsert: (command: LatexCommand) => void;
  view?: ToolbarView;
  layout?: ToolbarLayout;
  className?: string;
  stabilizeTileLayout?: boolean;
  compactDensity?: boolean;
  onCollapseTiles?: () => void;
}

interface FormulaTileDefinition {
  id: string;
  latex: string;
  labelZh: string;
  labelEn: string;
  sectionId?: string;
  color?: string | null;
}

interface CustomFormulaTileRecord {
  id: string;
  latex: string;
  sectionId: string;
  color: string | null;
  createdAt: number;
}

interface CustomFormulaTileSection {
  id: string;
  name: string;
  createdAt: number;
}

interface CustomFormulaTileLibrary {
  version: 3;
  sections: CustomFormulaTileSection[];
  tiles: CustomFormulaTileRecord[];
}

interface FormulaContextMenuState {
  target: FormulaHotkeyTarget;
  customTileId?: string;
  commonTileId?: string;
  /** a card of a group (right-clicked in the tools view) */
  toolbarItem?: { groupId: string; index: number; item: ToolbarItem };
  x: number;
  y: number;
}

interface SectionEditorState {
  mode: "create" | "rename";
  sectionId?: string;
  value: string;
}

const customFormulaTilesStorageKey = "visualtex-custom-formula-tiles";
const commonFormulaTileIdsStorageKey = "visualtex-common-formula-tile-ids-v1";
const defaultCustomSectionId = "default";
const defaultCustomSectionName = "未命名分区";
const maxCustomFormulaTiles = 30;
const maxCustomFormulaSections = 12;
const customTileColorPresets = [
  "#6f8fbf",
  "#ca6f7b",
  "#b180ca",
  "#4d9c8d",
  "#d18b45",
  "#78955b",
  "#7f8794",
] as const;
const customTileRowUnits = 12;
const customTileRowGap = 3;
const customTileHorizontalChrome = 6;
const customTileMinimumScale = 0.9;
const customTileMaximumItemsPerRow = 7;

function compactCustomTileWidth(latex: string) {
  const normalized = latex.replace(/\s+/g, "");
  const commandBody = normalized.charCodeAt(0) === 92 ? normalized.slice(1) : "";
  if (commandBody && /^[A-Za-z]+$/.test(commandBody)) {
    return 24;
  }
  if (
    commandBody &&
    /^[A-Za-z]+(?:(?:_|\^)(?:\{[^{}]+\}|[A-Za-z0-9]+)){1,2}$/.test(
      commandBody,
    )
  ) {
    return 42;
  }
  if (/^(?:\\[A-Za-z]+|[A-Za-z0-9])$/.test(normalized)) {
    return 24;
  }
  if (
    /^(?:\\[A-Za-z]+|[A-Za-z0-9])(?:_[{]?[A-Za-z0-9\\]+[}]?)?(?:\^[{]?[A-Za-z0-9\\]+[}]?)?$/.test(
      normalized,
    )
  ) {
    return 42;
  }
  return null;
}

function estimateCustomTileNaturalWidth(latex: string) {
  const commandCount = latex.match(/\\[A-Za-z]+/g)?.length ?? 0;
  const visibleCount = latex
    .replace(/\\[A-Za-z]+/g, "x")
    .replace(/[{}_^()[\]\\|\s]/g, "").length;
  const structuralBonus = /\\(?:frac|dfrac|tfrac|sum|prod|int|iint|iiint|lim|binom)/.test(
    latex,
  )
    ? 24
    : 0;
  return Math.min(
    260,
    Math.max(24, visibleCount * 10 + commandCount * 5 + structuralBonus),
  );
}

interface CustomTileRowItem {
  tile: FormulaTileDefinition;
  naturalWidth: number;
  minimumWidth: number;
  weight: number;
}

interface CustomTileLayoutRow {
  id: string;
  items: CustomTileRowItem[];
  fill: boolean;
}

function customTileBaseWeight(naturalWidth: number) {
  if (naturalWidth <= 28) return 1;
  if (naturalWidth <= 50) return 2;
  if (naturalWidth <= 76) return 3;
  if (naturalWidth <= 104) return 4;
  if (naturalWidth <= 138) return 5;
  if (naturalWidth <= 176) return 6;
  if (naturalWidth <= 220) return 8;
  return 12;
}

function normalizeCustomTileRowWeights(items: CustomTileRowItem[]) {
  if (items.length === 0) return items;
  const widths = items.map((item) => item.naturalWidth);
  const minimum = Math.min(...widths);
  const maximum = Math.max(...widths);
  const similarlySized = maximum <= Math.max(1, minimum) * 1.15;

  if (similarlySized) {
    if (items.length === 5 || items.length === 7) {
      return items.map((item) => ({ ...item, weight: 1 }));
    }
    if (customTileRowUnits % items.length === 0) {
      const equalWeight = customTileRowUnits / items.length;
      return items.map((item) => ({ ...item, weight: equalWeight }));
    }
  }

  const normalized = items.map((item) => ({ ...item }));
  let total = normalized.reduce((sum, item) => sum + item.weight, 0);
  if (total > customTileRowUnits) {
    const reducible = [...normalized].sort(
      (left, right) => right.weight - left.weight,
    );
    let index = 0;
    while (total > customTileRowUnits && reducible.length > 0) {
      const item = reducible[index % reducible.length];
      if (item.weight > 1) {
        item.weight -= 1;
        total -= 1;
      }
      index += 1;
      if (index > 100) break;
    }
  } else if (total < customTileRowUnits) {
    // Keep genuinely compact one-symbol tiles compact. Filling spare row units
    // should widen the formulas that actually need room rather than turning a
    // single \beta/letter into a full two-unit tile on Windows font metrics.
    const widerItems = normalized
      .filter((item) => item.naturalWidth > 28)
      .sort((left, right) => right.naturalWidth - left.naturalWidth);
    const expandable = widerItems.length > 0
      ? widerItems
      : [...normalized].sort(
          (left, right) => right.naturalWidth - left.naturalWidth,
        );
    let index = 0;
    while (total < customTileRowUnits && expandable.length > 0) {
      expandable[index % expandable.length].weight += 1;
      total += 1;
      index += 1;
    }
  }
  return normalized;
}

function buildCustomTileRows(
  tiles: readonly FormulaTileDefinition[],
  naturalWidths: Readonly<Record<string, number>>,
  availableWidth: number,
): CustomTileLayoutRow[] {
  const safeWidth = Math.max(120, availableWidth);
  const rows: CustomTileRowItem[][] = [];
  let current: CustomTileRowItem[] = [];
  let currentMinimumWidth = 0;
  let currentWeight = 0;

  const commit = () => {
    if (current.length > 0) rows.push(current);
    current = [];
    currentMinimumWidth = 0;
    currentWeight = 0;
  };

  for (const tile of tiles) {
    const compactWidth = compactCustomTileWidth(tile.latex);
    const measuredWidth =
      naturalWidths[tile.id] ?? estimateCustomTileNaturalWidth(tile.latex);
    const naturalWidth = compactWidth ?? measuredWidth;
    const minimumWidth = Math.max(
      22,
      Math.ceil(naturalWidth * customTileMinimumScale + customTileHorizontalChrome),
    );
    const weight = customTileBaseWeight(naturalWidth);
    const nextCount = current.length + 1;
    const nextMinimumWidth =
      currentMinimumWidth +
      minimumWidth +
      (current.length > 0 ? customTileRowGap : 0);
    const nextWeight = currentWeight + weight;
    const fits =
      nextCount <= customTileMaximumItemsPerRow &&
      nextMinimumWidth <= safeWidth &&
      nextWeight <= customTileRowUnits + 1;

    if (!fits && current.length > 0) commit();
    current.push({ tile, naturalWidth, minimumWidth, weight });
    currentMinimumWidth = minimumWidth;
    currentWeight = weight;
    if (current.length > 1) {
      currentMinimumWidth = current.reduce(
        (sum, item, index) =>
          sum + item.minimumWidth + (index > 0 ? customTileRowGap : 0),
        0,
      );
      currentWeight = current.reduce((sum, item) => sum + item.weight, 0);
    }
  }
  commit();

  return rows.map((items, index) => {
    const totalMinimumWidth = items.reduce(
      (sum, item, itemIndex) =>
        sum + item.minimumWidth + (itemIndex > 0 ? customTileRowGap : 0),
      0,
    );
    const fill =
      items.length > 1 || totalMinimumWidth >= safeWidth * 0.58;
    const normalizedItems = fill
      ? normalizeCustomTileRowWeights(items)
      : items;
    const hasNonCompactItem = normalizedItems.some(
      (item) => compactCustomTileWidth(item.tile.latex) !== 24,
    );
    if (fill && hasNonCompactItem) {
      for (const item of normalizedItems) {
        if (compactCustomTileWidth(item.tile.latex) === 24) {
          item.weight = 1;
        }
      }
    }
    return {
      id: `${index}-${items.map((item) => item.tile.id).join("-")}`,
      items: normalizedItems,
      fill,
    };
  });
}

function persistentCustomId(prefix: string) {
  return globalThis.crypto?.randomUUID?.() ??
    `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

function validCustomTileColor(value: unknown): string | null {
  return typeof value === "string" && /^#[0-9a-f]{6}$/i.test(value)
    ? value.toLowerCase()
    : null;
}

// The Common tab shows built-in library tiles picked by the user, in the order
// they were added. Unknown ids (from a newer or older build) are dropped.
function normalizeCommonFormulaTileIds(value: unknown) {
  if (!Array.isArray(value)) return null;
  const ids: string[] = [];
  for (const item of value) {
    if (
      typeof item === "string" &&
      formulaTileLibraryById.has(item) &&
      !ids.includes(item)
    ) {
      ids.push(item);
    }
  }
  return ids;
}

function loadCommonFormulaTileIds() {
  try {
    const stored = localStorage.getItem(commonFormulaTileIdsStorageKey);
    if (stored !== null) {
      const ids = normalizeCommonFormulaTileIds(JSON.parse(stored));
      if (ids) return ids;
    }
  } catch {
    // Fall back to the defaults below.
  }
  return [...defaultCommonFormulaTileIds];
}

function emptyCustomFormulaTileLibrary(): CustomFormulaTileLibrary {
  return {
    version: 3,
    sections: [
      {
        id: defaultCustomSectionId,
        name: defaultCustomSectionName,
        createdAt: 0,
      },
    ],
    tiles: [],
  };
}

function loadCustomFormulaTiles(): CustomFormulaTileLibrary {
  try {
    const stored = JSON.parse(
      localStorage.getItem(customFormulaTilesStorageKey) ?? "[]",
    );

    // Migrate the original string-array storage without losing any formula or
    // its existing hotkey, whose stable target ID is derived from the LaTeX.
    if (Array.isArray(stored)) {
      const seen = new Set<string>();
      return {
        version: 3,
        sections: [
          {
            id: defaultCustomSectionId,
            name: defaultCustomSectionName,
            createdAt: 0,
          },
        ],
        tiles: stored
          .filter((value): value is string => typeof value === "string")
          .map((value) => value.trim())
          .filter((value) => Boolean(value) && !seen.has(value) && seen.add(value))
          .slice(0, maxCustomFormulaTiles)
          .map((latex, index) => ({
            id: persistentCustomId("tile"),
            latex,
            sectionId: defaultCustomSectionId,
            color: null,
            createdAt: Date.now() - index,
          })),
      };
    }

    if (!stored || typeof stored !== "object") {
      return emptyCustomFormulaTileLibrary();
    }
    const candidate = stored as Partial<CustomFormulaTileLibrary> & {
      version?: number;
    };
    const isCurrentVersion = candidate.version === 3;
    const sectionIds = new Set<string>();
    const parsedSections = Array.isArray(candidate.sections)
      ? candidate.sections
          .filter((section): section is CustomFormulaTileSection => {
            if (!section || typeof section !== "object") return false;
            const id = typeof section.id === "string" ? section.id.trim() : "";
            const name = typeof section.name === "string" ? section.name.trim() : "";
            if (!id || !name || sectionIds.has(id)) return false;
            sectionIds.add(id);
            return true;
          })
          .slice(0, maxCustomFormulaSections)
          .map((section, index) => ({
            id: section.id,
            name: section.name.slice(0, 24),
            createdAt:
              typeof section.createdAt === "number" && Number.isFinite(section.createdAt)
                ? section.createdAt
                : Date.now() + index,
          }))
      : [];
    const sections = isCurrentVersion
      ? parsedSections
      : [
          {
            id: defaultCustomSectionId,
            name: defaultCustomSectionName,
            createdAt: 0,
          },
          ...parsedSections.filter(
            (section) => section.id !== defaultCustomSectionId,
          ),
        ].slice(0, maxCustomFormulaSections);
    const validSectionIds = new Set(sections.map((section) => section.id));
    const fallbackSectionId = sections[0]?.id ?? defaultCustomSectionId;
    const tileIds = new Set<string>();
    const latexValues = new Set<string>();
    const tiles = Array.isArray(candidate.tiles)
      ? candidate.tiles
          .filter((tile): tile is CustomFormulaTileRecord => {
            if (!tile || typeof tile !== "object") return false;
            const id = typeof tile.id === "string" ? tile.id.trim() : "";
            const latex = typeof tile.latex === "string" ? tile.latex.trim() : "";
            if (!id || !latex || tileIds.has(id) || latexValues.has(latex)) {
              return false;
            }
            tileIds.add(id);
            latexValues.add(latex);
            return true;
          })
          .slice(0, maxCustomFormulaTiles)
          .map((tile, index) => ({
            id: tile.id,
            latex: tile.latex.trim(),
            sectionId: validSectionIds.has(tile.sectionId)
              ? tile.sectionId
              : fallbackSectionId,
            color: validCustomTileColor(tile.color),
            createdAt:
              typeof tile.createdAt === "number" && Number.isFinite(tile.createdAt)
                ? tile.createdAt
                : Date.now() - index,
          }))
      : [];
    if (sections.length === 0 && tiles.length > 0) {
      return {
        version: 3,
        sections: [
          {
            id: defaultCustomSectionId,
            name: defaultCustomSectionName,
            createdAt: 0,
          },
        ],
        tiles: tiles.map((tile) => ({
          ...tile,
          sectionId: defaultCustomSectionId,
        })),
      };
    }
    return { version: 3, sections, tiles };
  } catch {
    return emptyCustomFormulaTileLibrary();
  }
}

function persistCustomFormulaTiles(library: CustomFormulaTileLibrary) {
  try {
    localStorage.setItem(customFormulaTilesStorageKey, JSON.stringify(library));
  } catch {
    // Keep the current session usable even when storage is unavailable.
  }
}

/** a group of the tools view (its id) */
type ToolbarCategory = string;
type ToolbarPreviewMode = "full" | "static";

// The tiles (Common and custom) are one set for every toolbar on the page: an edit in one shows in the others, and
// a drag from the classic bottom toolbar to the tiles panel lands in the same set.
interface FormulaTilesState {
  customTileLibrary: CustomFormulaTileLibrary;
  commonFormulaTileIds: string[];
  setCustomTileLibrary: (updater: (current: CustomFormulaTileLibrary) => CustomFormulaTileLibrary) => void;
  setCommonFormulaTileIds: (updater: (current: string[]) => string[]) => void;
}

const useFormulaTilesStore = create<FormulaTilesState>()((set) => ({
  customTileLibrary: loadCustomFormulaTiles(),
  commonFormulaTileIds: loadCommonFormulaTileIds(),
  setCustomTileLibrary: (updater) =>
    set((state) => {
      const customTileLibrary = updater(state.customTileLibrary);
      if (customTileLibrary === state.customTileLibrary) return state;
      persistCustomFormulaTiles(customTileLibrary);
      return { customTileLibrary };
    }),
  setCommonFormulaTileIds: (updater) =>
    set((state) => {
      const commonFormulaTileIds = updater(state.commonFormulaTileIds);
      if (commonFormulaTileIds === state.commonFormulaTileIds) return state;
      writeLocalStorage(commonFormulaTileIdsStorageKey, JSON.stringify(commonFormulaTileIds));
      return { commonFormulaTileIds };
    }),
}));

if (typeof window !== "undefined") {
  window.addEventListener("storage", (event) => {
    if (event.key === customFormulaTilesStorageKey) useFormulaTilesStore.setState({ customTileLibrary: loadCustomFormulaTiles() });
    if (event.key === commonFormulaTileIdsStorageKey) useFormulaTilesStore.setState({ commonFormulaTileIds: loadCommonFormulaTileIds() });
  });
}

/** The tint of a group of the user's: the built-in tints in turn. */
const groupTintVariables = BUILT_IN_GROUP_IDS.map((id) => `var(--toolbar-category-${id})`);

const registryCommandById = new Map(commandRegistry.map((command) => [command.id, command]));

const toolbarPreviewOverscanRatio = 1.35;
const toolbarScrollIdleDelayMs = 110;

function sameToolbarCategoryList(
  left: readonly string[],
  right: readonly string[],
) {
  return (
    left.length === right.length &&
    left.every((category, index) => category === right[index])
  );
}

const matrixGridSize = 5;
// Horizontal-dock matrix builder geometry used to wrap the delimiter rail;
// keep in sync with styles-windows-responsive.css.
const matrixBuilderVerticalPadding = 8;
const matrixDelimiterLabelBlockHeight = 15;
const matrixDelimiterMinButtonHeight = 28;
const matrixDelimiterGap = 1;
const matrixGridCells = Array.from(
  { length: matrixGridSize * matrixGridSize },
  (_, index) => ({
    row: Math.floor(index / matrixGridSize) + 1,
    column: (index % matrixGridSize) + 1,
  }),
);
const matrixDelimiterOptions: Array<{
  id: MatrixDelimiter;
  preview: string;
  labelZh: string;
  labelEn: string;
}> = [
  {
    id: "matrix",
    preview: "\\begin{matrix}a\\end{matrix}",
    labelZh: "无边界",
    labelEn: "Plain",
  },
  {
    id: "pmatrix",
    preview: "\\begin{pmatrix}a\\end{pmatrix}",
    labelZh: "圆括号",
    labelEn: "Parentheses",
  },
  {
    id: "bmatrix",
    preview: "\\begin{bmatrix}a\\end{bmatrix}",
    labelZh: "方括号",
    labelEn: "Brackets",
  },
  {
    id: "Bmatrix",
    preview: "\\begin{Bmatrix}a\\end{Bmatrix}",
    labelZh: "花括号",
    labelEn: "Braces",
  },
  {
    id: "vmatrix",
    preview: "\\begin{vmatrix}a\\end{vmatrix}",
    labelZh: "单竖线",
    labelEn: "Bars",
  },
  {
    id: "Vmatrix",
    preview: "\\begin{Vmatrix}a\\end{Vmatrix}",
    labelZh: "双竖线",
    labelEn: "Double bars",
  },
];

const wideToolbarCommandIds = new Set([
  "matrixelement",
  "expectation-operator",
  "matrix-dots",
  "diagonal-matrix",
  "augmented-matrix",
  "block-matrix",
]);
const minimumToolbarPreviewInsetRatio = 0.58;
const maximumToolbarPreviewInsetRatio = 0.94;
const compactToolbarPreviewInsetRatio = 0.8;
const compactToolbarPreviewMaximumScale = 0.82;
const compactCasesPreviewMaximumScale = 0.92;

const calculusPreviewById: Record<string, string> = {
  intplain: "\\int",
  int: "\\int_a^b",
  "iint-bounds": "\\iint_D^S",
  "iiint-bounds": "\\iiint_V^W",
  "oint-bounds": "\\oint_C^D",
  lineintegral: "\\int_C",
  iint: "\\iint_D",
  surfaceintegral: "\\iint_S",
  iiint: "\\iiint_V",
  volumeintegral: "\\iiint_V",
  oint: "\\oint_C",
  "closed-surface-integral": "\\oiint_S",
  "closed-volume-integral": "\\oiiint_V",
  sum: "\\sum_{i=1}^{n}",
  "sum-finite": "\\sum_{k=1}^{n}",
  series: "\\sum_{n=0}^{\\infty}",
  prod: "\\prod_{i=1}^{n}",
  "prod-finite": "\\prod_{k=1}^{n}",
  productseries: "\\prod_{n=1}^{\\infty}",
  coproduct: "\\coprod_{i=1}^{n}",
  "bigcap-limits": "\\bigcap_{i=1}^{n}",
  "bigcup-limits": "\\bigcup_{i=1}^{n}",
  "bigsqcup-limits": "\\bigsqcup_{i=1}^{n}",
  "bigvee-limits": "\\bigvee_{i=1}^{n}",
  "bigwedge-limits": "\\bigwedge_{i=1}^{n}",
  "bigodot-limits": "\\bigodot_{i=1}^{n}",
  "bigoplus-limits": "\\bigoplus_{i=1}^{n}",
  "bigotimes-limits": "\\bigotimes_{i=1}^{n}",
  "biguplus-limits": "\\biguplus_{i=1}^{n}",
  lim: "\\lim_{x\\to0}",
  "lim-infty": "\\lim_{x\\to\\infty}",
  "lim-left": "\\lim_{x\\to a^-}",
  "lim-right": "\\lim_{x\\to a^+}",
  derivative: "\\frac{\\mathrm{d}}{\\mathrm{d}x}",
  secondderivative: "\\frac{\\mathrm{d}^{2}}{\\mathrm{d}x^{2}}",
  "partialderivative-frac": "\\frac{\\partial}{\\partial x}",
  dd: "\\mathrm{d}",
  partialsecond: "\\frac{\\partial^{2}}{\\partial x^{2}}",
  mixedpartial: "\\frac{\\partial^{2}}{\\partial x\\partial y}",
  evalbar: "\\left.\\vphantom{F}\\right|_a^b",
  nabla: "\\nabla",
  ln: "\\ln",
  log: "\\log_a",
  exp: "\\exp",
  sin: "\\sin",
  cos: "\\cos",
  tan: "\\tan",
};

const toolbarPreviewById: Record<string, string> = {
  ...calculusPreviewById,
  "skewed-fraction": "{}^{a}\\!/\\!{}_{b}",
  cases: "\\begin{cases}a\\\\b\\end{cases}",
  "cases-three": "\\left\\{\\begin{array}{l}a\\\\b\\\\c\\end{array}\\right.",
  overbrace: "\\overbrace{a+b}",
  underbrace: "\\underbrace{a+b}",
  rowvector: "\\begin{bmatrix}a&b\\end{bmatrix}",
  colvector: "\\begin{bmatrix}a\\\\b\\end{bmatrix}",
  det: "\\det",
  trace: "\\operatorname{tr}",
  rank: "\\operatorname{rank}",
  transpose: "A^{\\mathsf{T}}",
  inverse: "A^{-1}",
  dotproduct: "\\bullet",
  "matrix-dots": "\\begin{bmatrix}a&\\cdots&b\\\\\\vdots&\\ddots&\\vdots\\\\c&\\cdots&d\\end{bmatrix}",
  "diagonal-matrix": "\\begin{bmatrix}a_1&0&\\cdots\\\\0&a_2&\\ddots\\\\\\vdots&\\ddots&a_n\\end{bmatrix}",
  "augmented-matrix": "\\left[\\begin{array}{cc|c}1&0&a\\\\0&1&b\\end{array}\\right]",
  "block-matrix": "\\begin{bmatrix}A&B\\\\C&D\\end{bmatrix}",
};

const toolbarPreviewLatex = (command: LatexCommand) =>
  toolbarPreviewById[command.id] ?? command.previewLatex;

function createTileCommand(tile: FormulaTileDefinition): LatexCommand {
  return {
    id: `formula-tile-${tile.id}`,
    command: tile.latex,
    insertTemplate: tile.latex,
    previewLatex: tile.latex,
    labelZh: tile.labelZh,
    labelEn: tile.labelEn,
    aliases: ["tile", "formula"],
    keywords: ["磁贴", "公式"],
    category: "structure",
    defaultPriority: 120,
    supportedInMathMode: true,
  };
}

function createMatrixCommand(
  rows: number,
  columns: number,
  delimiter: MatrixDelimiter,
): LatexCommand {
  const matrixBody = Array.from({ length: rows }, () =>
    Array.from({ length: columns }, () => "\\placeholder{}").join(" & "),
  ).join(" \\\\ ");
  const delimiterCopy = matrixDelimiterOptions.find(
    (option) => option.id === delimiter,
  ) ?? matrixDelimiterOptions[1];

  return {
    id: `custom-${delimiter}-${rows}x${columns}`,
    command: `\\begin{${delimiter}}`,
    insertTemplate: `\\begin{${delimiter}}${matrixBody}\\end{${delimiter}}`,
    previewLatex: delimiterCopy.preview,
    labelZh: `${rows}×${columns} ${delimiterCopy.labelZh}矩阵`,
    labelEn: `${rows}×${columns} ${delimiterCopy.labelEn.toLowerCase()} matrix`,
    aliases: ["matrix", delimiter],
    keywords: ["矩阵", "自定义矩阵", `${rows}x${columns}`],
    category: "matrix",
    defaultPriority: 120,
    supportedInMathMode: true,
  };
}

/** The tiles with one tile put at a place among its section's tiles (index past them: after the last). */
function placeTileInSection(
  tiles: readonly CustomFormulaTileRecord[],
  tile: CustomFormulaTileRecord,
  index: number,
) {
  const rest = tiles.filter((candidate) => candidate.id !== tile.id);
  const sectionTiles = rest.filter((candidate) => candidate.sectionId === tile.sectionId);
  const fromIndex = tiles
    .filter((candidate) => candidate.sectionId === tile.sectionId)
    .findIndex((candidate) => candidate.id === tile.id);
  const at = fromIndex >= 0 && fromIndex < index ? index - 1 : index;
  const before = sectionTiles[Math.max(0, at)];
  if (before && at < sectionTiles.length) {
    rest.splice(rest.indexOf(before), 0, tile);
  } else {
    const last = sectionTiles[sectionTiles.length - 1];
    rest.splice(last ? rest.indexOf(last) + 1 : rest.length, 0, tile);
  }
  return rest.slice(0, maxCustomFormulaTiles);
}

export function FormulaToolbar({
  onInsert,
  view: fixedView,
  layout = "sidebar",
  className = "",
  stabilizeTileLayout = false,
  compactDensity = false,
  onCollapseTiles,
}: Props) {
  const [internalActiveView, setInternalActiveView] =
    useState<ToolbarView>("tools");
  const activeView = fixedView ?? internalActiveView;
  const [activeTileCategory, setActiveTileCategory] =
    useState<TileCategory>("common");
  const customTileLibrary = useFormulaTilesStore((state) => state.customTileLibrary);
  const setCustomTileLibrary = useFormulaTilesStore((state) => state.setCustomTileLibrary);
  const commonFormulaTileIds = useFormulaTilesStore((state) => state.commonFormulaTileIds);
  const setCommonFormulaTileIds = useFormulaTilesStore((state) => state.setCommonFormulaTileIds);
  const [tileLibraryOpen, setTileLibraryOpen] = useState(false);
  const toolbarGroups = useToolbarLayoutStore((state) => state.layout.groups);
  const editing = useToolbarLayoutStore((state) => state.editing);
  const setEditing = useToolbarLayoutStore((state) => state.setEditing);
  const [pickerGroupId, setPickerGroupId] = useState<string | null>(null);
  const [renamingGroupId, setRenamingGroupId] = useState<string | null>(null);
  const [renameValue, setRenameValue] = useState("");
  const [layoutMenuOpen, setLayoutMenuOpen] = useState(false);
  /** where the Add to group list stands (beside its menu item) */
  const [addToGroupOpen, setAddToGroupOpen] = useState<CSSProperties | null>(null);
  const [activeCustomSectionId, setActiveCustomSectionId] = useState(
    defaultCustomSectionId,
  );
  const [sectionEditor, setSectionEditor] =
    useState<SectionEditorState | null>(null);
  const [pendingSectionDeleteId, setPendingSectionDeleteId] =
    useState<string | null>(null);
  // Result of the last "save as tile" click, shown in place of the hint.
  const [tileSaveNotice, setTileSaveNotice] = useState<string | null>(null);
  const [pendingCustomSymbolDeleteId, setPendingCustomSymbolDeleteId] =
    useState<string | null>(null);
  const [contextMenu, setContextMenu] =
    useState<FormulaContextMenuState | null>(null);
  const [hotkeyTarget, setHotkeyTarget] =
    useState<FormulaHotkeyTarget | null>(null);
  const openCustomSymbolDesigner = useCustomSymbolDesignerStore((state) => state.openDesigner);
  const [activeCategory, setActiveCategory] =
    useState<ToolbarCategory>("common");
  const [fullPreviewCategories, setFullPreviewCategories] = useState<
    ToolbarCategory[]
  >(["common", "structure"]);
  const [matrixRows, setMatrixRows] = useState(2);
  const [matrixColumns, setMatrixColumns] = useState(2);
  const [matrixHover, setMatrixHover] = useState<{
    rows: number;
    columns: number;
  } | null>(null);
  const [matrixDelimiter, setMatrixDelimiter] =
    useState<MatrixDelimiter>("bmatrix");
  const toolbarRef = useRef<HTMLElement>(null);
  // The category the user picked from a tab, and the scroll offset that pick
  // produced. Until the strip scrolls away from that offset, the pick wins over
  // the scroll-position heuristic, which cannot tell the first/last categories
  // apart from their neighbours on very wide strips.
  const pickedToolbarCategoryRef = useRef<{
    category: ToolbarCategory;
    scrollLeft: number;
  } | null>(null);
  const [horizontalRowCount, setHorizontalRowCount] = useState(3);
  const [customTileGridWidths, setCustomTileGridWidths] = useState<
    Record<string, number>
  >({});
  const [customTileNaturalWidths, setCustomTileNaturalWidths] = useState<
    Record<string, number>
  >({});
  const language = useEditorStore((state) => state.language);
  const customSymbolRevision = useCustomSymbolRevision();
  const registeredCustomSymbols = useMemo(
    () => readCustomSymbolLibrary().symbols,
    [customSymbolRevision],
  );
  const registeredCustomSymbolCommands = useMemo(
    () =>
      new Map(
        customSymbolCommands().map((command) => [command.command, command]),
      ),
    [customSymbolRevision],
  );
  const formulaToolButtonSize = useEditorStore(
    (state) => state.formulaToolButtonSize,
  );
  const formulaToolButtonPadding = useEditorStore(
    (state) => state.formulaToolButtonPadding,
  );
  const effectiveFormulaToolButtonSize = compactDensity
    ? Math.max(34, formulaToolButtonSize - 8)
    : formulaToolButtonSize;
  const effectiveFormulaToolButtonPadding = compactDensity
    ? Math.max(0, formulaToolButtonPadding - 1)
    : formulaToolButtonPadding;
  const fullSizeToolbarPreviewInsetRatio = Math.max(
    minimumToolbarPreviewInsetRatio,
    Math.min(
      maximumToolbarPreviewInsetRatio,
      maximumToolbarPreviewInsetRatio - effectiveFormulaToolButtonPadding * 0.03,
    ),
  );
  const toolbarPreviewInsetRatio = compactDensity
    ? Math.min(compactToolbarPreviewInsetRatio, fullSizeToolbarPreviewInsetRatio)
    : fullSizeToolbarPreviewInsetRatio;
  const toolbarPreviewMaximumScale = compactDensity
    ? compactToolbarPreviewMaximumScale
    : Math.min(1.55, Math.max(1, effectiveFormulaToolButtonSize / 42));
  const lines = useEditorStore((state) => state.lines);
  const activeLineId = useEditorStore((state) => state.activeLineId);
  const hotkeyBindings = useFormulaHotkeyStore((state) => state.bindings);
  const removeBindingsForTarget = useFormulaHotkeyStore(
    (state) => state.removeBindingsForTarget,
  );
  const isEn = language === "en";
  const activeLineLatex = useMemo(
    () => lines.find((line) => line.id === activeLineId)?.latex.trim() ?? "",
    [activeLineId, lines],
  );
  const customSections = customTileLibrary.sections;
  const activeCustomSection = customSections.find(
    (section) => section.id === activeCustomSectionId,
  );

  useEffect(() => {
    if (
      !customTileLibrary.sections.some(
        (section) => section.id === activeCustomSectionId,
      )
    ) {
      setActiveCustomSectionId(customTileLibrary.sections[0]?.id ?? "");
    }
    if (
      pendingSectionDeleteId &&
      !customTileLibrary.sections.some(
        (section) => section.id === pendingSectionDeleteId,
      )
    ) {
      setPendingSectionDeleteId(null);
    }
  }, [
    activeCustomSectionId,
    customTileLibrary.sections,
    pendingSectionDeleteId,
  ]);

  useEffect(() => {
    if (activeView !== "tiles" || activeTileCategory !== "custom") return;
    const root = toolbarRef.current;
    if (!root) return;

    let frame = 0;
    const quantizeWidth = (width: number) =>
      Math.max(120, Math.floor(width / 2) * 2);
    // Re-packing custom tile rows changes the list height; when that toggles
    // a classic scroll bar the measured width flips with it. One guard per
    // section settles such a flip-flop.
    const widthGuards = new Map<string, (width: number) => number>();
    const commitWidths = (rawEntries: Array<[string, number]>) => {
      const entries = rawEntries.map(([sectionId, width]): [string, number] => {
        let guard = widthGuards.get(sectionId);
        if (!guard) {
          guard = createLayoutOscillationGuard();
          widthGuards.set(sectionId, guard);
        }
        return [sectionId, guard(width)];
      });
      if (!entries.length) return;
      setCustomTileGridWidths((current) => {
        let changed = false;
        const next = { ...current };
        for (const [sectionId, width] of entries) {
          if (Math.abs((current[sectionId] ?? 0) - width) < 2) continue;
          next[sectionId] = width;
          changed = true;
        }
        return changed ? next : current;
      });
    };
    const schedule = (callback: () => void) => {
      window.cancelAnimationFrame(frame);
      frame = window.requestAnimationFrame(callback);
    };

    if (stabilizeTileLayout) {
      const recordStableRootWidth = () =>
        schedule(() => {
          const availableWidth = quantizeWidth(root.clientWidth - 16);
          commitWidths(
            customTileLibrary.sections.map((section) => [
              section.id,
              availableWidth,
            ]),
          );
        });
      const observer = new ResizeObserver(recordStableRootWidth);
      observer.observe(root);
      recordStableRootWidth();
      return () => {
        window.cancelAnimationFrame(frame);
        observer.disconnect();
      };
    }

    const grids = Array.from(
      root.querySelectorAll<HTMLElement>(".custom-formula-tile-grid"),
    );
    const recordGridWidths = () =>
      schedule(() => {
        commitWidths(
          grids.flatMap((grid) => {
            const sectionId = grid.dataset.customTileGridSection;
            const width = quantizeWidth(grid.getBoundingClientRect().width);
            return sectionId && width > 0 ? [[sectionId, width]] : [];
          }),
        );
      });
    const observer = new ResizeObserver(recordGridWidths);
    grids.forEach((grid) => observer.observe(grid));
    recordGridWidths();
    return () => {
      window.cancelAnimationFrame(frame);
      observer.disconnect();
    };
  }, [
    activeTileCategory,
    activeView,
    customTileLibrary.sections,
    stabilizeTileLayout,
  ]);

  const scrollToToolbarCategory = (
    category: ToolbarCategory,
    edge: "start" | "end" = "start",
  ) => {
    const root = toolbarRef.current;
    const strip = root?.querySelector<HTMLElement>(
      ".template-strip.is-continuous-categories",
    );
    const section = strip?.querySelector<HTMLElement>(
      `[data-toolbar-category-section="${category}"]`,
    );
    if (!root || !strip || !section) {
      setActiveCategory(category);
      return false;
    }

    const sectionStart = section.offsetLeft;
    const sectionEnd = Math.max(
      sectionStart,
      section.offsetLeft + section.offsetWidth - strip.clientWidth,
    );
    strip.scrollLeft = edge === "end" ? sectionEnd : sectionStart;
    pickedToolbarCategoryRef.current = {
      category,
      scrollLeft: strip.scrollLeft,
    };
    setActiveCategory(category);
    root
      .querySelector<HTMLElement>(`.toolbar-tab[data-category="${category}"]`)
      ?.scrollIntoView({
        block: "nearest",
        inline: "nearest",
        behavior: "auto",
      });
    return true;
  };

  useLayoutEffect(() => {
    const root = toolbarRef.current;
    const activeTab = root?.querySelector<HTMLElement>(
      `.toolbar-tab[data-category="${activeCategory}"]`,
    );
    activeTab?.scrollIntoView({
      block: "nearest",
      inline: "nearest",
      behavior: "auto",
    });

    if (layout !== "horizontal" || activeView !== "tools") return;
    const strip = root?.querySelector<HTMLElement>(
      ".template-strip.is-continuous-categories",
    );
    const section = strip?.querySelector<HTMLElement>(
      `[data-toolbar-category-section="${activeCategory}"]`,
    );
    if (!strip || !section) return;

    // Horizontal row-count changes reflow every category section. Keep the
    // active category anchored to its new geometry immediately; otherwise the
    // old scrollLeft can leave the selected section far outside the viewport.
    const targetLeft = Math.max(
      0,
      Math.min(
        section.offsetLeft,
        Math.max(0, strip.scrollWidth - strip.clientWidth),
      ),
    );
    if (Math.abs(strip.scrollLeft - targetLeft) > 1) {
      strip.scrollTo({ left: targetLeft, top: 0, behavior: "auto" });
    }
    const picked = pickedToolbarCategoryRef.current;
    if (picked?.category === activeCategory) {
      picked.scrollLeft = strip.scrollLeft;
    }
  }, [activeCategory, activeView, horizontalRowCount, layout]);

  useEffect(() => {
    if (layout !== "horizontal" || activeView !== "tools") return;
    const root = toolbarRef.current;
    const strip = root?.querySelector<HTMLElement>(
      ".template-strip.is-continuous-categories",
    );
    const tabs = root?.querySelector<HTMLElement>(".toolbar-tabs");
    if (!root || !strip || !tabs) return;

    const sections = Array.from(
      strip.querySelectorAll<HTMLElement>(
        ":scope > [data-toolbar-category-section]",
      ),
    );
    const tabByCategory = new Map<ToolbarCategory, HTMLElement>();
    tabs.querySelectorAll<HTMLElement>(".toolbar-tab").forEach((tab) => {
      const category = tab.dataset.category as ToolbarCategory | undefined;
      if (category) tabByCategory.set(category, tab);
    });
    if (!sections.length) return;

    let idleTimer = 0;
    let lastScrollAt = 0;
    let currentDomCategory =
      (strip.dataset.activeCategory as ToolbarCategory | undefined) ?? "common";
    let geometry: Array<{
      category: ToolbarCategory;
      start: number;
      end: number;
    }> = [];

    const measureGeometry = () => {
      geometry = sections.flatMap((section) => {
        const category = section.dataset
          .toolbarCategorySection as ToolbarCategory | undefined;
        if (!category) return [];
        const start = section.offsetLeft;
        return [{ category, start, end: start + section.offsetWidth }];
      });
    };

    const resolveCategory = () => {
      const picked = pickedToolbarCategoryRef.current;
      if (picked) {
        if (Math.abs(strip.scrollLeft - picked.scrollLeft) <= 2) {
          return picked.category;
        }
        pickedToolbarCategoryRef.current = null;
      }
      const viewportWidth = Math.max(1, strip.clientWidth);
      const maxScrollLeft = Math.max(0, strip.scrollWidth - viewportWidth);
      // The first/last category cannot always be aligned with the viewport
      // marker because the browser clamps scrollLeft at the strip edges. In
      // particular, after Physics became wider, clicking the final Sets tab
      // could land at maxScrollLeft while the 38% marker still sat inside the
      // Physics section, immediately switching the active tab back to Physics.
      // Treat the physical scroll edges as authoritative category boundaries.
      if (strip.scrollLeft <= 1) {
        return geometry[0]?.category ?? "common";
      }
      if (maxScrollLeft - strip.scrollLeft <= 1) {
        return geometry[geometry.length - 1]?.category ?? "common";
      }
      const marker = strip.scrollLeft + viewportWidth * 0.38;
      let resolved: ToolbarCategory = geometry[0]?.category ?? "common";
      let nearestDistance = Number.POSITIVE_INFINITY;
      for (const item of geometry) {
        const distance = marker < item.start
          ? item.start - marker
          : marker > item.end
            ? marker - item.end
            : 0;
        if (distance < nearestDistance) {
          resolved = item.category;
          nearestDistance = distance;
        }
      }
      return resolved;
    };

    const updateActiveCategoryDom = (resolved: ToolbarCategory) => {
      if (resolved === currentDomCategory) return;
      tabByCategory.get(currentDomCategory)?.classList.remove("is-active");
      tabByCategory
        .get(currentDomCategory)
        ?.setAttribute("aria-pressed", "false");
      tabByCategory.get(resolved)?.classList.add("is-active");
      tabByCategory.get(resolved)?.setAttribute("aria-pressed", "true");
      currentDomCategory = resolved;
      strip.dataset.activeCategory = resolved;
    };

    const commitIdleState = () => {
      const elapsed = performance.now() - lastScrollAt;
      if (lastScrollAt > 0 && elapsed < toolbarScrollIdleDelayMs) {
        idleTimer = window.setTimeout(
          commitIdleState,
          toolbarScrollIdleDelayMs - elapsed,
        );
        return;
      }
      idleTimer = 0;
      strip.dataset.scrolling = "false";
      const resolved = resolveCategory();
      updateActiveCategoryDom(resolved);

      const viewportStart = strip.scrollLeft;
      const viewportWidth = Math.max(1, strip.clientWidth);
      const viewportEnd = viewportStart + viewportWidth;
      const overscan = viewportWidth * toolbarPreviewOverscanRatio;
      const renderStart = viewportStart - overscan;
      const renderEnd = viewportEnd + overscan;
      const nextFullPreviewCategories = geometry.flatMap((item) =>
        item.end >= renderStart && item.start <= renderEnd
          ? [item.category]
          : [],
      );
      if (!nextFullPreviewCategories.includes(resolved)) {
        nextFullPreviewCategories.push(resolved);
      }

      setActiveCategory((current) =>
        current === resolved ? current : resolved,
      );
      setFullPreviewCategories((current) =>
        sameToolbarCategoryList(current, nextFullPreviewCategories)
          ? current
          : nextFullPreviewCategories,
      );
    };

    const ensureIdleCommit = () => {
      if (!idleTimer) {
        idleTimer = window.setTimeout(
          commitIdleState,
          toolbarScrollIdleDelayMs,
        );
      }
    };
    const handleScroll = () => {
      lastScrollAt = performance.now();
      if (strip.dataset.scrolling !== "true") {
        strip.dataset.scrolling = "true";
      }
      ensureIdleCommit();
    };
    const handleResize = () => {
      measureGeometry();
      updateActiveCategoryDom(resolveCategory());
      lastScrollAt = 0;
      window.clearTimeout(idleTimer);
      idleTimer = window.setTimeout(commitIdleState, 0);
    };

    const resizeObserver = new ResizeObserver(handleResize);
    resizeObserver.observe(strip);
    strip.addEventListener("scroll", handleScroll, { passive: true });
    measureGeometry();
    updateActiveCategoryDom(resolveCategory());
    idleTimer = window.setTimeout(commitIdleState, 0);

    return () => {
      window.clearTimeout(idleTimer);
      resizeObserver.disconnect();
      strip.removeEventListener("scroll", handleScroll);
      delete strip.dataset.scrolling;
    };
  }, [activeView, horizontalRowCount, layout]);

  useEffect(() => {
    if (layout !== "horizontal") return;
    const root = toolbarRef.current;
    const strip = root?.querySelector<HTMLElement>(
      ".template-strip.is-continuous-categories",
    );
    const tabs = root?.querySelector<HTMLElement>(".toolbar-tabs");
    if (!root || !strip || !tabs) return;

    const handleHorizontalWheel = (event: WheelEvent) => {
      if (event.ctrlKey || event.metaKey || event.altKey) return;
      const eventTarget = event.target;
      if (!(eventTarget instanceof Node)) return;
      const target = eventTarget === strip || strip.contains(eventTarget)
        ? strip
        : eventTarget === tabs || tabs.contains(eventTarget)
          ? tabs
          : null;
      if (!target) return;

      const unit = event.deltaMode === WheelEvent.DOM_DELTA_LINE
        ? 24
        : event.deltaMode === WheelEvent.DOM_DELTA_PAGE
          ? Math.max(120, target.clientWidth * 0.82)
          : 1;
      const rawDelta = Math.abs(event.deltaY) >= Math.abs(event.deltaX)
        ? event.deltaY
        : event.deltaX;
      const pixelDelta = rawDelta * unit;
      if (!pixelDelta) return;

      event.preventDefault();
      target.scrollBy({ left: pixelDelta, top: 0, behavior: "auto" });
    };

    root.addEventListener("wheel", handleHorizontalWheel, { passive: false });
    return () => root.removeEventListener("wheel", handleHorizontalWheel);
  }, [layout]);

  useLayoutEffect(() => {
    if (layout !== "horizontal" || activeView !== "tools") return;
    const root = toolbarRef.current;
    const tabs = root?.querySelector<HTMLElement>(".toolbar-tabs");
    const strip = root?.querySelector<HTMLElement>(".template-strip");
    if (!root || !tabs || !strip) return;

    let frame = 0;
    // The row count changes how wide the strip's content is, which can show
    // or hide its horizontal scroll bar and so change clientHeight.
    const stableStripHeight = createLayoutOscillationGuard();
    const measureRows = () => {
      window.cancelAnimationFrame(frame);
      frame = window.requestAnimationFrame(() => {
        const styles = window.getComputedStyle(strip);
        const paddingTop = Number.parseFloat(styles.paddingTop) || 0;
        const paddingBottom = Number.parseFloat(styles.paddingBottom) || 0;
        const rowGap = Number.parseFloat(styles.rowGap) || 0;
        const availableHeight = Math.max(
          0,
          stableStripHeight(strip.clientHeight) - paddingTop - paddingBottom,
        );
        const targetRowHeight = Math.max(26, effectiveFormulaToolButtonSize - 8);
        const nextRowCount = Math.max(
          1,
          Math.floor(
            (availableHeight + rowGap) / (targetRowHeight + rowGap),
          ),
        );
        const matrixMainWidth = Math.max(
          40,
          effectiveFormulaToolButtonSize * 5 - 44,
        );
        const matrixPickerSize = Math.max(
          32,
          Math.min(
            matrixMainWidth,
            Math.floor(availableHeight - 32),
          ),
        );

        // The delimiter previews need ~28px of height each. Stacking all six in
        // one column squeezed them to a few pixels in short docks, so wrap them
        // into as many columns as the dock height requires.
        const delimiterHeight = availableHeight -
          matrixBuilderVerticalPadding -
          matrixDelimiterLabelBlockHeight;
        const delimiterRowsThatFit = Math.max(
          1,
          Math.min(
            matrixDelimiterOptions.length,
            Math.floor(
              (delimiterHeight + matrixDelimiterGap) /
                (matrixDelimiterMinButtonHeight + matrixDelimiterGap),
            ),
          ),
        );
        const delimiterColumns = Math.ceil(
          matrixDelimiterOptions.length / delimiterRowsThatFit,
        );
        const delimiterRows = Math.ceil(
          matrixDelimiterOptions.length / delimiterColumns,
        );

        root.style.setProperty("--toolbar-row-count", String(nextRowCount));
        root.style.setProperty("--matrix-picker-size", `${matrixPickerSize}px`);
        root.style.setProperty(
          "--matrix-delimiter-columns",
          String(delimiterColumns),
        );
        root.style.setProperty("--matrix-delimiter-rows", String(delimiterRows));
        root.style.setProperty(
          "--matrix-panel-height",
          `${Math.max(56, Math.floor(availableHeight))}px`,
        );
        root.dataset.toolbarRowCount = String(nextRowCount);
        setHorizontalRowCount((current) =>
          current === nextRowCount ? current : nextRowCount,
        );
      });
    };

    const observer = new ResizeObserver(measureRows);
    observer.observe(root);
    observer.observe(strip);
    measureRows();
    return () => {
      window.cancelAnimationFrame(frame);
      observer.disconnect();
    };
  }, [activeCategory, activeView, effectiveFormulaToolButtonSize, layout]);

  useLayoutEffect(() => {
    if (layout === "horizontal" || activeView !== "tools") return;
    const root = toolbarRef.current;
    const strip = root?.querySelector<HTMLElement>(".template-strip");
    if (!root || !strip) return;

    let frame = 0;
    // The matrix picker size feeds back into the strip's scroll bars.
    const stableStripHeight = createLayoutOscillationGuard();
    const stableStripWidth = createLayoutOscillationGuard();
    const measureMatrix = () => {
      window.cancelAnimationFrame(frame);
      frame = window.requestAnimationFrame(() => {
        const styles = window.getComputedStyle(strip);
        const paddingTop = Number.parseFloat(styles.paddingTop) || 0;
        const paddingBottom = Number.parseFloat(styles.paddingBottom) || 0;
        const paddingLeft = Number.parseFloat(styles.paddingLeft) || 0;
        const paddingRight = Number.parseFloat(styles.paddingRight) || 0;
        const availableHeight = Math.max(
          0,
          stableStripHeight(strip.clientHeight) - paddingTop - paddingBottom,
        );
        const availableWidth = Math.max(
          0,
          stableStripWidth(strip.clientWidth) - paddingLeft - paddingRight,
        );
        const railWidth = 32;
        const builderGap = 6;
        const builderHorizontalPadding = 12;
        const headingHeight = 20;
        const insertHeight = 28;
        const verticalGaps = 10;
        const builderVerticalPadding = 12;
        const widthBound = Math.max(
          44,
          availableWidth - railWidth - builderGap - builderHorizontalPadding,
        );
        const heightBound = Math.max(
          44,
          availableHeight -
            headingHeight -
            insertHeight -
            verticalGaps -
            builderVerticalPadding,
        );
        const matrixPickerSize = Math.max(
          44,
          Math.min(220, Math.floor(widthBound), Math.floor(heightBound)),
        );
        const matrixPanelHeight = Math.max(
          112,
          Math.min(
            Math.floor(availableHeight),
            matrixPickerSize +
              headingHeight +
              insertHeight +
              verticalGaps +
              builderVerticalPadding,
          ),
        );

        root.style.setProperty("--matrix-picker-size", `${matrixPickerSize}px`);
        root.style.setProperty("--matrix-panel-height", `${matrixPanelHeight}px`);
        root.dataset.matrixPickerSize = String(matrixPickerSize);
      });
    };

    const observer = new ResizeObserver(measureMatrix);
    observer.observe(root);
    observer.observe(strip);
    measureMatrix();
    return () => {
      window.cancelAnimationFrame(frame);
      observer.disconnect();
    };
  }, [activeCategory, activeView, layout]);

  useEffect(() => {
    if (!contextMenu) return;

    const closeFromPointer = (event: PointerEvent) => {
      const target = event.target;
      if (
        target instanceof Element &&
        target.closest(".formula-tile-context-menu")
      ) {
        return;
      }
      setContextMenu(null);
    };
    const closeFromKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") setContextMenu(null);
    };
    const closeMenu = () => setContextMenu(null);
    const closeFromScroll = (event: Event) => {
      const target = event.target;
      if (target instanceof Element) {
        const isTileMenu =
          contextMenu.target.kind === "common-tile" ||
          contextMenu.target.kind === "custom-tile";
        // Switching between the tools and tiles views can leave the hidden
        // horizontal tools strip finishing a smooth-scroll animation. That
        // unrelated scroll must not instantly dismiss a freshly opened tile
        // context menu. Apply the same protection in reverse for tool menus.
        if (
          (isTileMenu &&
            target.closest(".template-strip, .toolbar-tabs")) ||
          (!isTileMenu &&
            target.closest(".formula-tile-list, .formula-tiles-panel"))
        ) {
          return;
        }
      }
      setContextMenu(null);
    };

    document.addEventListener("pointerdown", closeFromPointer);
    document.addEventListener("keydown", closeFromKey);
    window.addEventListener("blur", closeMenu);
    window.addEventListener("resize", closeMenu);
    window.addEventListener("scroll", closeFromScroll, true);

    return () => {
      document.removeEventListener("pointerdown", closeFromPointer);
      document.removeEventListener("keydown", closeFromKey);
      window.removeEventListener("blur", closeMenu);
      window.removeEventListener("resize", closeMenu);
      window.removeEventListener("scroll", closeFromScroll, true);
    };
  }, [contextMenu]);

  // a command of a group: the registry's, or a custom symbol's (gone once the symbol is deleted)
  const customSymbolCommandById = useMemo(
    () => new Map(customSymbolCommands().map((command) => [command.id, command])),
    [customSymbolRevision],
  );
  const resolveToolbarCommand = (id: string) =>
    registryCommandById.get(id) ?? customSymbolCommandById.get(id);
  const groupLabel = (group: ToolbarGroup) =>
    group.name ??
    (group.builtIn
      ? (isEn ? categoryLabelsEn : categoryLabels)[group.builtIn]
      : "");
  const groupTintStyle = (group: ToolbarGroup) =>
    group.builtIn
      ? undefined
      : ({
          "--toolbar-category-tint":
            groupTintVariables[
              Math.max(0, toolbarGroups.indexOf(group)) % groupTintVariables.length
            ],
        } as CSSProperties);
  const itemsByGroup = useMemo(
    () => new Map(toolbarGroups.map((group) => [group.id, groupItems(group)])),
    [toolbarGroups],
  );
  const activeGroup = toolbarGroups.find((group) => group.id === activeCategory);
  const removedBuiltInGroups = BUILT_IN_GROUP_IDS.filter(
    (id) => !toolbarGroups.some((group) => group.id === id),
  );

  useEffect(() => {
    if (toolbarGroups.some((group) => group.id === activeCategory)) return;
    setActiveCategory(toolbarGroups[0]?.id ?? "");
  }, [activeCategory, toolbarGroups]);

  const toolbarPreviewModeForCategory = (
    category: ToolbarCategory,
  ): ToolbarPreviewMode =>
    layout !== "horizontal" ||
    category === activeCategory ||
    fullPreviewCategories.includes(category)
      ? "full"
      : "static";

  const customTileDefinitions = useMemo<FormulaTileDefinition[]>(
    () =>
      customTileLibrary.tiles.map((tile, index) => ({
        id: tile.id,
        latex: tile.latex,
        labelZh: `自定义公式 ${index + 1}`,
        labelEn: `Custom formula ${index + 1}`,
        sectionId: tile.sectionId,
        color: tile.color,
      })),
    [customTileLibrary.tiles],
  );
  const commonFormulaTiles = useMemo<FormulaTileDefinition[]>(
    () =>
      commonFormulaTileIds.flatMap((id) => {
        const tile = formulaTileLibraryById.get(id);
        return tile ? [tile] : [];
      }),
    [commonFormulaTileIds],
  );
  const visibleFormulaTiles =
    activeTileCategory === "common"
      ? commonFormulaTiles
      : customTileDefinitions;
  const toggleCommonFormulaTile = (tileId: string) =>
    setCommonFormulaTileIds((current) =>
      current.includes(tileId)
        ? current.filter((id) => id !== tileId)
        : [...current, tileId],
    );
  const removeCommonFormulaTile = (tileId: string) => {
    setCommonFormulaTileIds((current) =>
      current.filter((id) => id !== tileId),
    );
    setContextMenu(null);
  };

  const recordCustomTileNaturalWidth = (tileId: string, width: number) => {
    const stableWidth = Math.max(1, Math.round(width));
    setCustomTileNaturalWidths((current) =>
      Math.abs((current[tileId] ?? 0) - stableWidth) < 2
        ? current
        : { ...current, [tileId]: stableWidth },
    );
  };
  const customTileRowsForSection = (
    tiles: readonly FormulaTileDefinition[],
    sectionId: string,
  ) =>
    buildCustomTileRows(
      tiles,
      customTileNaturalWidths,
      customTileGridWidths[sectionId] ?? 420,
    );

  const commandHotkeyTarget = (command: LatexCommand) =>
    createFormulaHotkeyTarget(
      formulaHotkeyTargetIdForCommand(command.id),
      "command",
      command,
    );
  const tileHotkeyTarget = (
    tile: FormulaTileDefinition,
    kind: "common-tile" | "custom-tile",
  ) =>
    createFormulaHotkeyTarget(
      formulaHotkeyTargetIdForTile(kind, tile.id, tile.latex),
      kind,
      createTileCommand(tile),
    );
  const matrixHotkeyTarget = () => {
    const command = createMatrixCommand(
      matrixRows,
      matrixColumns,
      matrixDelimiter,
    );
    return createFormulaHotkeyTarget(
      `matrix:${command.id}`,
      "matrix",
      command,
    );
  };
  const insertCustomMatrix = () => {
    onInsert(matrixHotkeyTarget().command);
  };
  const insertFormulaTile = (tile: FormulaTileDefinition) => {
    onInsert(createTileCommand(tile));
  };
  const updateCustomTileLibrary = (
    updater: (current: CustomFormulaTileLibrary) => CustomFormulaTileLibrary,
  ) => setCustomTileLibrary((current) => updater(current));

  useEffect(() => {
    if (!tileSaveNotice) return;
    const timeout = window.setTimeout(() => setTileSaveNotice(null), 3000);
    return () => window.clearTimeout(timeout);
  }, [tileSaveNotice]);

  const saveActiveFormulaAsTile = () => {
    if (!activeLineLatex || !activeCustomSection) return;
    const duplicate = customTileLibrary.tiles.find(
      (tile) => tile.latex === activeLineLatex,
    );
    const sectionName = activeCustomSection.name;
    setTileSaveNotice(
      !duplicate
        ? isEn
          ? `Saved to “${sectionName}”`
          : `已保存到「${sectionName}」`
        : duplicate.sectionId === activeCustomSection.id
          ? isEn
            ? "Already a tile; moved to the front"
            : "磁贴已存在，已移到最前"
          : isEn
            ? `Already a tile; moved to “${sectionName}”`
            : `磁贴已存在，已移到「${sectionName}」`,
    );
    updateCustomTileLibrary((current) => {
      const existing = current.tiles.find(
        (tile) => tile.latex === activeLineLatex,
      );
      const tile: CustomFormulaTileRecord = existing
        ? {
            ...existing,
            sectionId: activeCustomSection.id,
          }
        : {
            id: persistentCustomId("tile"),
            latex: activeLineLatex,
            sectionId: activeCustomSection.id,
            color: null,
            createdAt: Date.now(),
          };
      return {
        ...current,
        tiles: [
          tile,
          ...current.tiles.filter((item) => item.id !== tile.id),
        ].slice(0, maxCustomFormulaTiles),
      };
    });
    setActiveTileCategory("custom");
  };

  const beginCreateSection = () => {
    if (customTileLibrary.sections.length >= maxCustomFormulaSections) return;
    setPendingSectionDeleteId(null);
    setSectionEditor({ mode: "create", value: "" });
  };
  const beginRenameSection = (section: CustomFormulaTileSection) => {
    setPendingSectionDeleteId(null);
    setActiveCustomSectionId(section.id);
    setSectionEditor({
      mode: "rename",
      sectionId: section.id,
      value: section.name,
    });
  };
  const commitSectionEditor = () => {
    if (!sectionEditor) return;
    const name = sectionEditor.value.trim().slice(0, 24);
    if (!name) return;
    if (sectionEditor.mode === "create") {
      const section: CustomFormulaTileSection = {
        id: persistentCustomId("section"),
        name,
        createdAt: Date.now(),
      };
      updateCustomTileLibrary((current) => ({
        ...current,
        sections: [...current.sections, section].slice(0, maxCustomFormulaSections),
      }));
      setActiveCustomSectionId(section.id);
    } else if (sectionEditor.sectionId) {
      updateCustomTileLibrary((current) => ({
        ...current,
        sections: current.sections.map((section) =>
          section.id === sectionEditor.sectionId
            ? { ...section, name }
            : section,
        ),
      }));
    }
    setSectionEditor(null);
  };
  const deleteCustomSection = (sectionId: string) => {
    const deletedTiles = customTileLibrary.tiles.filter(
      (tile) => tile.sectionId === sectionId,
    );
    deletedTiles.forEach((tile) =>
      removeBindingsForTarget(
        formulaHotkeyTargetIdForTile("custom-tile", tile.id, tile.latex),
      ),
    );
    const remainingSections = customTileLibrary.sections.filter(
      (section) => section.id !== sectionId,
    );
    updateCustomTileLibrary((current) => ({
      ...current,
      sections: current.sections.filter((section) => section.id !== sectionId),
      tiles: current.tiles.filter((tile) => tile.sectionId !== sectionId),
    }));
    setActiveCustomSectionId(remainingSections[0]?.id ?? "");
    setPendingSectionDeleteId(null);
    setSectionEditor(null);
  };
  const updateCustomTile = (
    tileId: string,
    patch: Partial<Pick<CustomFormulaTileRecord, "sectionId" | "color">>,
  ) => {
    updateCustomTileLibrary((current) => ({
      ...current,
      tiles: current.tiles.map((tile) =>
        tile.id === tileId ? { ...tile, ...patch } : tile,
      ),
    }));
  };
  // ---- the groups as the user arranges them (edit mode) ----
  const layoutStore = useToolbarLayoutStore.getState;
  const itemAt = (groupId: string, index: number) =>
    itemsByGroup.get(groupId)?.[index];
  const libraryTileIdOf = (item: ToolbarItem | undefined) => {
    if (item?.kind !== "formula" || !item.id.startsWith("library-")) return null;
    const id = item.id.slice("library-".length);
    return formulaTileLibraryById.has(id) ? id : null;
  };
  /** what a toolbar item puts in a custom tile (the matrix builder makes none) */
  const tileLatexOf = (item: ToolbarItem | undefined) => {
    if (!item || item.kind === "matrix") return null;
    if (item.kind === "formula") return item.latex;
    return resolveToolbarCommand(item.id)?.insertTemplate ?? null;
  };
  const formulaItemForTile = (
    kind: "common" | "custom",
    tileId: string,
    latex: string,
  ): ToolbarItem => ({
    kind: "formula",
    id: kind === "common" ? `library-${tileId}` : `tile-${tileId}`,
    latex,
  });
  const formulaItemCommand = (item: Extract<ToolbarItem, { kind: "formula" }>) => {
    const libraryId = libraryTileIdOf(item);
    const libraryTile = libraryId ? formulaTileLibraryById.get(libraryId) : undefined;
    return createTileCommand({
      id: item.id,
      latex: item.latex,
      labelZh: libraryTile?.labelZh ?? item.latex,
      labelEn: libraryTile?.labelEn ?? item.latex,
    });
  };
  const formulaItemHotkeyTarget = (
    item: Extract<ToolbarItem, { kind: "formula" }>,
  ) =>
    createFormulaHotkeyTarget(
      formulaHotkeyTargetIdForTile("custom-tile", item.id, item.latex),
      "custom-tile",
      formulaItemCommand(item),
    );
  /** a removed formula's hotkey goes with it, unless the same formula is still somewhere */
  const forgetFormulaHotkey = (latex: string) => {
    const trimmed = latex.trim();
    const stillUsed =
      useFormulaTilesStore
        .getState()
        .customTileLibrary.tiles.some((tile) => tile.latex === trimmed) ||
      layoutStore().layout.groups.some((group) =>
        groupItems(group).some(
          (item) => item.kind === "formula" && item.latex.trim() === trimmed,
        ),
      );
    if (!stillUsed) {
      removeBindingsForTarget(formulaHotkeyTargetIdForTile("custom-tile", "", trimmed));
    }
  };
  const removeToolbarItem = (groupId: string, index: number) => {
    const item = itemAt(groupId, index);
    layoutStore().removeItem(groupId, index);
    if (item?.kind === "formula") forgetFormulaHotkey(item.latex);
  };
  const toggleItemInGroup = (groupId: string, item: ToolbarItem) => {
    const items = itemsByGroup.get(groupId) ?? [];
    const index = items.findIndex((candidate) => sameToolbarItem(candidate, item));
    if (index >= 0) removeToolbarItem(groupId, index);
    else layoutStore().insertItem(groupId, END_INDEX, item);
  };
  const addCustomTileAt = (latex: string, sectionId: string, index: number) => {
    const trimmed = latex.trim();
    if (!trimmed) return;
    updateCustomTileLibrary((current) => {
      const existing = current.tiles.find((tile) => tile.latex === trimmed);
      if (!existing && current.tiles.length >= maxCustomFormulaTiles) return current;
      const tile: CustomFormulaTileRecord = existing
        ? { ...existing, sectionId }
        : {
            id: persistentCustomId("tile"),
            latex: trimmed,
            sectionId,
            color: null,
            createdAt: Date.now(),
          };
      return { ...current, tiles: placeTileInSection(current.tiles, tile, index) };
    });
  };
  const beginRenameGroup = (group: ToolbarGroup) => {
    setRenamingGroupId(group.id);
    setRenameValue(groupLabel(group));
  };
  const commitRenameGroup = () => {
    if (renamingGroupId) layoutStore().renameGroup(renamingGroupId, renameValue);
    setRenamingGroupId(null);
  };
  const addGroup = () => {
    const id = layoutStore().addGroup(isEn ? "New group" : "新分组");
    setActiveCategory(id);
    setRenamingGroupId(id);
    setRenameValue(isEn ? "New group" : "新分组");
    if (!editing) setEditing(true);
  };
  const removeGroup = (group: ToolbarGroup) => {
    groupItems(group).forEach((item) => {
      if (item.kind === "formula") window.setTimeout(() => forgetFormulaHotkey(item.latex), 0);
    });
    layoutStore().removeGroup(group.id);
  };

  const dragHandlers: DragHandlers = {
    accepts: (source: DragSource, target: DropTarget) => {
      if (source.type === "group-tab") return target.type === "tabs";
      if (target.type === "tabs") return false;
      if (target.type === "group") return true;
      if (target.type === "common-tiles") {
        if (source.type === "common-tile") return true;
        return source.type === "toolbar" && libraryTileIdOf(itemAt(source.groupId, source.index)) !== null;
      }
      if (source.type === "toolbar") return tileLatexOf(itemAt(source.groupId, source.index)) !== null;
      return true;
    },
    drop: (source: DragSource, target: DropTarget) => {
      const store = layoutStore();
      if (source.type === "group-tab") {
        if (target.type === "tabs") store.moveGroup(source.groupId, target.index);
        return;
      }
      if (target.type === "group") {
        if (source.type === "toolbar") {
          store.moveItem(source.groupId, source.index, target.groupId, target.index);
        } else if (source.type === "common-tile") {
          const tile = formulaTileLibraryById.get(source.tileId);
          if (tile) store.insertItem(target.groupId, target.index, formulaItemForTile("common", tile.id, tile.latex));
        } else {
          const tile = customTileLibrary.tiles.find((candidate) => candidate.id === source.tileId);
          if (tile) store.insertItem(target.groupId, target.index, formulaItemForTile("custom", tile.id, tile.latex));
        }
        setActiveCategory(target.groupId);
        return;
      }
      if (target.type === "common-tiles") {
        const tileId =
          source.type === "common-tile"
            ? source.tileId
            : source.type === "toolbar"
              ? libraryTileIdOf(itemAt(source.groupId, source.index))
              : null;
        if (!tileId) return;
        setCommonFormulaTileIds((current) => {
          const from = current.indexOf(tileId);
          const next = current.filter((id) => id !== tileId);
          const at = from >= 0 && from < target.index ? target.index - 1 : target.index;
          next.splice(Math.max(0, Math.min(next.length, at)), 0, tileId);
          return next;
        });
        return;
      }
      if (target.type === "custom-section") {
        if (source.type === "custom-tile") {
          updateCustomTileLibrary((current) => {
            const tile = current.tiles.find((candidate) => candidate.id === source.tileId);
            if (!tile) return current;
            return {
              ...current,
              tiles: placeTileInSection(current.tiles, { ...tile, sectionId: target.sectionId }, target.index),
            };
          });
        } else {
          const latex =
            source.type === "common-tile"
              ? formulaTileLibraryById.get(source.tileId)?.latex
              : tileLatexOf(itemAt(source.groupId, source.index));
          if (latex) addCustomTileAt(latex, target.sectionId, target.index);
        }
        setActiveCustomSectionId(target.sectionId);
      }
    },
  };
  /** a press on a card in edit mode: a drag (a press on its × or a field inside is not) */
  const armCardDrag = (
    event: ReactPointerEvent<HTMLElement>,
    source: DragSource,
  ) => {
    if (!editing) return;
    if ((event.target as Element).closest("[data-toolbar-remove], input")) return;
    armDrag(event.nativeEvent, source, event.currentTarget, dragHandlers);
  };
  /** Alt+arrow moves the focused card (or tab) one place; Delete removes it */
  const moveByKeyboard = (
    event: ReactKeyboardEvent<HTMLElement>,
    count: number,
    index: number,
    move: (toIndex: number) => void,
    remove: (() => void) | null,
    focusSelector: (toIndex: number) => string,
  ) => {
    if (!editing) return;
    if ((event.key === "Delete" || event.key === "Backspace") && remove) {
      event.preventDefault();
      remove();
      return;
    }
    if (!event.altKey) return;
    const back = event.key === "ArrowLeft" || event.key === "ArrowUp";
    const forward = event.key === "ArrowRight" || event.key === "ArrowDown";
    if (!back && !forward) return;
    event.preventDefault();
    const toIndex = back ? index - 1 : index + 1;
    if (toIndex < 0 || toIndex >= count) return;
    // insertion before the card that is there now (forward: past it)
    move(back ? toIndex : toIndex + 1);
    window.requestAnimationFrame(() =>
      toolbarRef.current?.querySelector<HTMLElement>(focusSelector(toIndex))?.focus(),
    );
  };

  useEffect(() => {
    if (!editing) return;
    const leave = (event: KeyboardEvent) => {
      if (event.key !== "Escape" || pickerGroupId || renamingGroupId) return;
      const target = event.target as Element | null;
      if (target?.closest("input, textarea, [contenteditable='true'], math-field, .modal-backdrop")) return;
      if (document.body.classList.contains("vtx-toolbar-dragging")) return;
      setEditing(false);
    };
    window.addEventListener("keydown", leave);
    return () => window.removeEventListener("keydown", leave);
  }, [editing, pickerGroupId, renamingGroupId, setEditing]);

  // the name field opens with the name selected (autoFocus + select on focus leaves WebKit's caret at the end)
  useEffect(() => {
    if (!renamingGroupId) return;
    const frame = window.requestAnimationFrame(() => {
      const input = toolbarRef.current?.querySelector<HTMLInputElement>("[data-toolbar-tab-rename]");
      input?.focus();
      input?.select();
    });
    return () => window.cancelAnimationFrame(frame);
  }, [renamingGroupId]);

  useEffect(() => {
    if (!editing) {
      setLayoutMenuOpen(false);
      setRenamingGroupId(null);
    }
  }, [editing]);

  useEffect(() => {
    if (!layoutMenuOpen) return;
    const close = (event: PointerEvent) => {
      if ((event.target as Element | null)?.closest(".toolbar-layout-menu, [data-toolbar-layout-menu-button]")) return;
      setLayoutMenuOpen(false);
    };
    document.addEventListener("pointerdown", close);
    return () => document.removeEventListener("pointerdown", close);
  }, [layoutMenuOpen]);

  // the menu's real height (its items vary): it stays inside the window
  const contextMenuRef = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    const menu = contextMenuRef.current;
    if (!menu || !contextMenu) return;
    const box = menu.getBoundingClientRect();
    const overflow = box.bottom - (window.innerHeight - 8);
    if (overflow > 0) menu.style.top = `${Math.max(8, contextMenu.y - overflow)}px`;
  }, [contextMenu]);
  const addToGroupPlace = (item: HTMLElement): CSSProperties => {
    const box = item.getBoundingClientRect();
    const width = 188;
    const right = box.right + width + 4 <= window.innerWidth - 8;
    const height = Math.min(toolbarGroups.length * 32 + 10, window.innerHeight - 16);
    return {
      position: "fixed",
      left: right ? box.right + 2 : Math.max(8, box.left - width - 2),
      top: Math.max(8, Math.min(box.top - 5, window.innerHeight - height - 8)),
      width,
      maxHeight: height,
    };
  };
  /** what a right-clicked thing is as an item of a group */
  const contextMenuItem = (menu: FormulaContextMenuState): ToolbarItem | null => {
    if (menu.toolbarItem) return menu.toolbarItem.item;
    const { target } = menu;
    if (target.kind === "command") return { kind: "command", id: target.command.id };
    if (target.kind === "matrix") return { kind: "matrix" };
    if (menu.commonTileId) return formulaItemForTile("common", menu.commonTileId, target.command.insertTemplate);
    if (menu.customTileId) return formulaItemForTile("custom", menu.customTileId, target.command.insertTemplate);
    return null;
  };
  const openFormulaContextMenu = (
    event: MouseEvent<HTMLButtonElement>,
    target: FormulaHotkeyTarget,
    customTileId?: string,
    commonTileId?: string,
    toolbarItem?: FormulaContextMenuState["toolbarItem"],
  ) => {
    event.preventDefault();
    event.stopPropagation();
    const menuWidth = customTileId ? 252 : 224;
    const menuHeight =
      (customTileId ? 330 : 132 + (commonTileId ? 38 : 0)) +
      76 +
      (toolbarItem ? 38 : 0);
    setAddToGroupOpen(null);
    setContextMenu({
      target,
      customTileId,
      commonTileId,
      toolbarItem,
      x: Math.max(8, Math.min(event.clientX, window.innerWidth - menuWidth - 8)),
      y: Math.max(8, Math.min(event.clientY, window.innerHeight - menuHeight - 8)),
    });
  };
  const deleteCustomFormulaTile = (tileId: string, targetId: string) => {
    updateCustomTileLibrary((current) => ({
      ...current,
      tiles: current.tiles.filter((item) => item.id !== tileId),
    }));
    removeBindingsForTarget(targetId);
    setContextMenu(null);
  };
  const requestDeleteRegisteredCustomSymbol = (symbolId: string) => {
    if (pendingCustomSymbolDeleteId !== symbolId) {
      setPendingCustomSymbolDeleteId(symbolId);
      return;
    }
    deleteCustomSymbol(symbolId);
    removeBindingsForTarget(
      formulaHotkeyTargetIdForCommand(`custom-symbol:${symbolId}`),
    );
    setPendingCustomSymbolDeleteId(null);
  };
  const contextBinding = contextMenu
    ? hotkeyBindings.find(
        (binding) => binding.target.id === contextMenu.target.id,
      ) ?? null
    : null;
  const contextCustomTile = contextMenu?.customTileId
    ? customTileLibrary.tiles.find(
        (tile) => tile.id === contextMenu.customTileId,
      ) ?? null
    : null;
  const previewRows = matrixHover?.rows ?? matrixRows;
  const previewColumns = matrixHover?.columns ?? matrixColumns;

  const renderMatrixBuilder = (
    previewMode: ToolbarPreviewMode = "full",
    edit?: { group: ToolbarGroup; index: number },
  ) => (
    <section
      key="matrix-builder"
      className={"matrix-builder" + (edit ? " is-toolbar-edit-item" : "")}
      aria-label={isEn ? "Custom matrix" : "自定义矩阵"}
      {...(edit ? editItemProps(edit.group, edit.index, { kind: "matrix" }) : {})}
    >
      {edit && renderRemoveBadge(edit.group, edit.index)}
      <div className="matrix-main-picker">
        <div className="matrix-builder-heading">
          <strong>{isEn ? "Matrix" : "矩阵"}</strong>
          <span className="matrix-size-badge" aria-live="polite">
            {previewRows} × {previewColumns}
          </span>
        </div>
        <div className="matrix-size-picker">
          <div
            className="matrix-size-grid"
            role="grid"
            aria-label={
              isEn
                ? "Select matrix rows and columns"
                : "选择矩阵行数和列数"
            }
            aria-rowcount={matrixGridSize}
            aria-colcount={matrixGridSize}
            onPointerLeave={() => setMatrixHover(null)}
            onBlur={(event) => {
              if (!event.currentTarget.contains(event.relatedTarget as Node | null)) {
                setMatrixHover(null);
              }
            }}
          >
            {matrixGridCells.map(({ row, column }) => {
              const previewed = row <= previewRows && column <= previewColumns;
              const selectedCorner = row === matrixRows && column === matrixColumns;
              return (
                <button
                  key={`${row}-${column}`}
                  type="button"
                  role="gridcell"
                  className={
                    "matrix-size-cell" +
                    (previewed ? " is-previewed" : "") +
                    (selectedCorner ? " is-selected-corner" : "")
                  }
                  aria-label={
                    isEn
                      ? `${row} rows by ${column} columns`
                      : `${row} 行 ${column} 列`
                  }
                  aria-selected={selectedCorner}
                  data-matrix-rows={row}
                  data-matrix-columns={column}
                  onPointerEnter={() => setMatrixHover({ rows: row, columns: column })}
                  onFocus={() => setMatrixHover({ rows: row, columns: column })}
                  onClick={() => {
                    setMatrixRows(row);
                    setMatrixColumns(column);
                    setMatrixHover(null);
                  }}
                />
              );
            })}
          </div>
        </div>
        <button
          type="button"
          className="matrix-insert-button"
          data-command-id="custom-matrix"
          onClick={insertCustomMatrix}
          onContextMenu={(event) =>
            openFormulaContextMenu(event, matrixHotkeyTarget())
          }
          title={
            isEn
              ? "Insert matrix"
              : "插入矩阵"
          }
        >
          {isEn
            ? `Insert ${matrixRows} × ${matrixColumns}`
            : `插入 ${matrixRows} × ${matrixColumns}`}
        </button>
      </div>

      <div className="matrix-compact-options">
        <span className="matrix-control-label">
          {isEn ? "Border" : "边界"}
        </span>
        <div
          className="matrix-delimiter-options"
          role="group"
          aria-label={isEn ? "Matrix delimiter" : "矩阵边界"}
        >
          {matrixDelimiterOptions.map((option) => (
            <button
              key={option.id}
              type="button"
              className={matrixDelimiter === option.id ? "is-active" : ""}
              aria-pressed={matrixDelimiter === option.id}
              onClick={() => setMatrixDelimiter(option.id)}
              title={isEn ? option.labelEn : option.labelZh}
              aria-label={isEn ? option.labelEn : option.labelZh}
            >
              <MathPreview
                latex={option.preview}
                className={
                  previewMode === "static" ? "toolbar-static-preview" : ""
                }
                fit
                staticLayout={previewMode === "static"}
                maximumFitScale={1.15}
                fitInsetRatio={0.72}
              />
            </button>
          ))}
        </div>
      </div>
    </section>
  );

  /** a card's part in edit mode: dragged, moved with Alt+arrows, removed with Delete */
  const editItemProps = (group: ToolbarGroup, index: number, item: ToolbarItem) => {
    const count = itemsByGroup.get(group.id)?.length ?? 0;
    return {
      "data-drag-index": index,
      "data-toolbar-item-kind": item.kind,
      tabIndex: 0,
      onPointerDown: (event: ReactPointerEvent<HTMLElement>) =>
        armCardDrag(event, { type: "toolbar", groupId: group.id, index }),
      onKeyDown: (event: ReactKeyboardEvent<HTMLElement>) =>
        moveByKeyboard(
          event,
          count,
          index,
          (toIndex) => layoutStore().moveItem(group.id, index, group.id, toIndex),
          () => removeToolbarItem(group.id, index),
          (toIndex) => `[data-drop-kind="group"][data-group-id="${group.id}"] [data-drag-index="${toIndex}"]`,
        ),
    };
  };
  const renderRemoveBadge = (group: ToolbarGroup, index: number) => (
    <span
      role="button"
      className="toolbar-item-remove"
      data-toolbar-remove
      aria-label={isEn ? "Remove" : "移除"}
      title={isEn ? "Remove" : "移除"}
      onPointerDown={(event) => event.stopPropagation()}
      onClick={(event) => {
        event.preventDefault();
        event.stopPropagation();
        removeToolbarItem(group.id, index);
      }}
    >
      <X size={10} strokeWidth={2.6} />
    </span>
  );
  const renderToolbarItem = (
    group: ToolbarGroup,
    item: ToolbarItem,
    index: number,
    previewMode: ToolbarPreviewMode,
  ) => {
    if (item.kind === "matrix") {
      return renderMatrixBuilder(previewMode, editing ? { group, index } : undefined);
    }
    const command =
      item.kind === "command" ? resolveToolbarCommand(item.id) : formulaItemCommand(item);
    if (!command) return null;
    return renderToolbarCommandButton(command, previewMode, { group, index, item });
  };
  const renderAddItemCard = (group: ToolbarGroup) => (
    <button
      type="button"
      key="add-item"
      className="template-button toolbar-add-item"
      data-toolbar-add-item={group.id}
      aria-label={isEn ? "Add" : "添加"}
      title={isEn ? "Add" : "添加"}
      onClick={() => setPickerGroupId(group.id)}
    >
      <Plus size={16} />
    </button>
  );

  const renderToolbarCommandButton = (
    command: LatexCommand,
    previewMode: ToolbarPreviewMode = "full",
    place?: { group: ToolbarGroup; index: number; item: ToolbarItem },
  ) => {
    const formulaItem = place?.item.kind === "formula" ? place.item : null;
    const previewLatex = formulaItem ? formulaItem.latex : toolbarPreviewLatex(command);
    const widePreview = wideToolbarCommandIds.has(command.id);
    const enlargedCasesPreview =
      command.id === "cases" || command.id === "cases-three";
    return (
      <button
        type="button"
        className={
          "template-button is-unified-fit" +
          (widePreview ? " is-wide-preview" : "") +
          (enlargedCasesPreview ? " is-enlarged-cases-preview" : "")
        }
        style={
          formulaItem
            ? {
                gridColumn: `span ${Math.max(1, Math.min(4, Math.ceil(estimateCustomTileNaturalWidth(formulaItem.latex) / 46)))}`,
              }
            : undefined
        }
        data-command-id={formulaItem ? `formula:${formulaItem.id}` : command.id}
        data-preview-latex={previewLatex}
        key={formulaItem ? `formula:${formulaItem.id}` : command.id}
        {...(editing && place ? editItemProps(place.group, place.index, place.item) : {})}
        onClick={() => {
          if (!editing) onInsert(command);
        }}
        onContextMenu={(event) =>
          openFormulaContextMenu(
            event,
            formulaItem ? formulaItemHotkeyTarget(formulaItem) : commandHotkeyTarget(command),
            undefined,
            undefined,
            place ? { groupId: place.group.id, index: place.index, item: place.item } : undefined,
          )
        }
        aria-label={isEn ? command.labelEn : command.labelZh}
        title={
          (isEn ? command.labelEn : command.labelZh) +
          " · " +
          command.command
        }
      >
        <MathPreview
          latex={previewLatex}
          className={
            previewMode === "static" ? "toolbar-static-preview" : ""
          }
          fit
          showPlaceholders={Boolean(formulaItem)}
          staticLayout={previewMode === "static"}
          maximumFitScale={
            enlargedCasesPreview
              ? compactDensity
                ? compactCasesPreviewMaximumScale
                : Math.max(1.85, toolbarPreviewMaximumScale)
              : toolbarPreviewMaximumScale
          }
          fitInsetRatio={
            enlargedCasesPreview && !compactDensity
              ? 0.98
              : toolbarPreviewInsetRatio
          }
        />
        {editing && place && renderRemoveBadge(place.group, place.index)}
      </button>
    );
  };

  return (
    <aside
      ref={toolbarRef}
      className={
        "formula-toolbar" +
        (fixedView ? " is-view-fixed" : "") +
        (layout === "horizontal" ? " is-horizontal" : "") +
        (className ? ` ${className}` : "")
      }
      style={
        {
          "--formula-toolbar-button-size": `${effectiveFormulaToolButtonSize}px`,
          "--formula-toolbar-button-padding": `${effectiveFormulaToolButtonPadding}px`,
          ...(layout === "horizontal"
            ? { "--toolbar-row-count": horizontalRowCount }
            : {}),
        } as CSSProperties
      }
      data-toolbar-layout={layout}
      data-toolbar-fixed-view={fixedView ?? ""}
      data-toolbar-row-count={
        layout === "horizontal" ? horizontalRowCount : undefined
      }
      aria-label={isEn ? "Formula toolbar" : "公式工具栏"}
    >
      {!fixedView && <header className="formula-toolbar-header">
        <div
          className="formula-toolbar-view-tabs"
          role="tablist"
          aria-label={isEn ? "Sidebar view" : "侧栏视图"}
        >
          <button
            type="button"
            role="tab"
            className={activeView === "tools" ? "is-active" : ""}
            aria-selected={activeView === "tools"}
            data-toolbar-view="tools"
            onClick={() => {
              setContextMenu(null);
              setInternalActiveView("tools");
            }}
          >
            {isEn ? "Formula tools" : "公式工具"}
          </button>
          <button
            type="button"
            role="tab"
            className={activeView === "tiles" ? "is-active" : ""}
            aria-selected={activeView === "tiles"}
            data-toolbar-view="tiles"
            onClick={() => {
              setContextMenu(null);
              setInternalActiveView("tiles");
            }}
          >
            {isEn ? "Tiles" : "磁贴"}
          </button>
        </div>
      </header>}

      {activeView === "tools" ? (
        <>
          <div className={"toolbar-tabs-row" + (editing ? " is-editing" : "")}>
            <nav
              className="toolbar-tabs"
              aria-label={isEn ? "Formula categories" : "公式分类"}
              {...(editing ? { "data-drop-kind": "tabs" } : {})}
            >
              {toolbarGroups.map((group, index) => {
                const label = groupLabel(group);
                const active = activeCategory === group.id;
                if (renamingGroupId === group.id) {
                  return (
                    <input
                      key={group.id}
                      className="toolbar-tab toolbar-tab-rename is-active"
                      data-category={group.id}
                      data-toolbar-tab-rename
                      style={groupTintStyle(group)}
                      maxLength={24}
                      value={renameValue}
                      aria-label={isEn ? "Group name" : "分组名称"}
                      onChange={(event) => setRenameValue(event.currentTarget.value)}
                      onBlur={commitRenameGroup}
                      onKeyDown={(event) => {
                        if (event.key === "Enter") commitRenameGroup();
                        if (event.key === "Escape") {
                          event.stopPropagation();
                          setRenamingGroupId(null);
                        }
                      }}
                    />
                  );
                }
                return (
                  <button
                    key={group.id}
                    type="button"
                    className={"toolbar-tab " + (active ? "is-active" : "")}
                    data-category={group.id}
                    data-group-id={group.id}
                    style={groupTintStyle(group)}
                    aria-pressed={active}
                    {...(editing
                      ? {
                          "data-drop-kind": "group-tab",
                          "data-drag-index": index,
                          title: isEn ? "Double-click to rename" : "双击重命名",
                          onPointerDown: (event: ReactPointerEvent<HTMLElement>) =>
                            armCardDrag(event, { type: "group-tab", groupId: group.id }),
                          onDoubleClick: () => beginRenameGroup(group),
                          onKeyDown: (event: ReactKeyboardEvent<HTMLElement>) => {
                            if (event.key === "F2") {
                              event.preventDefault();
                              beginRenameGroup(group);
                              return;
                            }
                            moveByKeyboard(
                              event,
                              toolbarGroups.length,
                              index,
                              (toIndex) => layoutStore().moveGroup(group.id, toIndex),
                              null,
                              (toIndex) => `.toolbar-tabs [data-drag-index="${toIndex}"]`,
                            );
                          },
                        }
                      : {})}
                    onClick={() => {
                      if (layout === "horizontal") {
                        scrollToToolbarCategory(group.id, "start");
                      } else {
                        setActiveCategory(group.id);
                      }
                    }}
                  >
                    {label}
                    {editing && (
                      <span
                        role="button"
                        className="toolbar-tab-remove"
                        data-toolbar-remove
                        data-toolbar-remove-group={group.id}
                        aria-label={isEn ? `Remove ${label}` : `移除${label}`}
                        title={isEn ? "Remove group" : "移除分组"}
                        onPointerDown={(event) => event.stopPropagation()}
                        onClick={(event) => {
                          event.preventDefault();
                          event.stopPropagation();
                          removeGroup(group);
                        }}
                      >
                        <X size={10} strokeWidth={2.6} />
                      </span>
                    )}
                  </button>
                );
              })}
              {(editing || toolbarGroups.length === 0) && (
                <button
                  type="button"
                  className="toolbar-tab toolbar-tab-add"
                  data-toolbar-add-group
                  aria-label={isEn ? "New group" : "新建分组"}
                  title={isEn ? "New group" : "新建分组"}
                  onClick={addGroup}
                >
                  <Plus size={14} />
                </button>
              )}
            </nav>
            <div className="toolbar-tabs-actions">
              {editing && (
                <button
                  type="button"
                  className="toolbar-tabs-action"
                  data-toolbar-layout-menu-button
                  aria-label={isEn ? "More" : "更多"}
                  title={isEn ? "More" : "更多"}
                  aria-expanded={layoutMenuOpen}
                  onClick={() => setLayoutMenuOpen((open) => !open)}
                >
                  <MoreHorizontal size={15} />
                </button>
              )}
              <button
                type="button"
                className={"toolbar-tabs-action" + (editing ? " is-done" : "")}
                data-toolbar-edit-toggle
                aria-pressed={editing}
                aria-label={editing ? (isEn ? "Done" : "完成") : isEn ? "Edit toolbar" : "编辑工具栏"}
                title={editing ? (isEn ? "Done (Esc)" : "完成 (Esc)") : isEn ? "Edit toolbar" : "编辑工具栏"}
                onClick={() => setEditing(!editing)}
              >
                {editing ? <Check size={15} /> : <SlidersHorizontal size={14} />}
              </button>
              {layoutMenuOpen && (
                <div className="toolbar-layout-menu formula-tile-context-menu" role="menu">
                  {removedBuiltInGroups.length > 0 && (
                    <>
                      <div className="toolbar-layout-menu-label">{isEn ? "Restore group" : "恢复分组"}</div>
                      {removedBuiltInGroups.map((id: BuiltInGroupId) => (
                        <button
                          key={id}
                          type="button"
                          role="menuitem"
                          className="formula-hotkey-context-action"
                          data-toolbar-restore-group={id}
                          onClick={() => {
                            layoutStore().restoreGroup(id);
                            setActiveCategory(id);
                            setLayoutMenuOpen(false);
                          }}
                        >
                          <Plus size={14} />
                          {(isEn ? categoryLabelsEn : categoryLabels)[id]}
                        </button>
                      ))}
                      <div className="formula-hotkey-context-divider" />
                    </>
                  )}
                  {activeGroup?.builtIn && activeGroup.items && (
                    <button
                      type="button"
                      role="menuitem"
                      className="formula-hotkey-context-action"
                      data-toolbar-reset-group
                      onClick={() => {
                        layoutStore().resetGroupItems(activeGroup.id);
                        setLayoutMenuOpen(false);
                      }}
                    >
                      <RotateCcw size={14} />
                      {isEn ? `Reset “${groupLabel(activeGroup)}”` : `重置「${groupLabel(activeGroup)}」`}
                    </button>
                  )}
                  <button
                    type="button"
                    role="menuitem"
                    className="formula-hotkey-context-action"
                    data-toolbar-reset-all
                    onClick={() => {
                      layoutStore().resetAll();
                      setLayoutMenuOpen(false);
                    }}
                  >
                    <RotateCcw size={14} />
                    {isEn ? "Restore defaults" : "恢复默认"}
                  </button>
                  {toolbarGroups.length > 0 && <button
                    type="button"
                    role="menuitem"
                    className="formula-hotkey-context-action is-danger"
                    data-toolbar-clear-all
                    onClick={() => {
                      layoutStore().clearAll();
                      setLayoutMenuOpen(false);
                    }}
                  >
                    <Trash2 size={14} />
                    {isEn ? "Start from blank" : "从空白开始"}
                  </button>}
                </div>
              )}
            </div>
          </div>

          <div
            className={
              "template-strip" +
              (layout === "horizontal" ? " is-continuous-categories" : "") +
              (editing ? " is-toolbar-editing" : "")
            }
            data-active-category={activeCategory}
            style={layout !== "horizontal" && activeGroup ? groupTintStyle(activeGroup) : undefined}
            {...(editing && layout !== "horizontal" && activeGroup
              ? { "data-drop-kind": "group", "data-group-id": activeGroup.id }
              : {})}
            aria-label={isEn ? "Formula templates" : "公式模板"}
          >
            {layout === "horizontal" ? (
              toolbarGroups.map((group) => {
                const previewMode = toolbarPreviewModeForCategory(group.id);
                return (
                  <section
                    key={group.id}
                    className="toolbar-category-section"
                    data-toolbar-category-section={group.id}
                    data-preview-mode={previewMode}
                    style={groupTintStyle(group)}
                    {...(editing
                      ? { "data-drop-kind": "group", "data-group-id": group.id, "data-drop-flow": "column" }
                      : {})}
                    aria-label={groupLabel(group)}
                  >
                    {(itemsByGroup.get(group.id) ?? []).map((item, index) =>
                      renderToolbarItem(group, item, index, previewMode),
                    )}
                    {editing && renderAddItemCard(group)}
                  </section>
                );
              })
            ) : activeGroup ? (
              <>
                {(itemsByGroup.get(activeGroup.id) ?? []).map((item, index) =>
                  renderToolbarItem(activeGroup, item, index, "full"),
                )}
                {editing && renderAddItemCard(activeGroup)}
              </>
            ) : null}
          </div>
        </>
      ) : (
        <section
          className="formula-tiles-panel"
          aria-label={isEn ? "Formula tiles" : "公式磁贴"}
        >
          <nav
            className={`formula-tile-tabs${onCollapseTiles ? " has-collapse" : ""}`}
            aria-label={isEn ? "Tile categories" : "磁贴分类"}
          >
            <button
              type="button"
              className={activeTileCategory === "custom" ? "is-active" : ""}
              aria-pressed={activeTileCategory === "custom"}
              data-tile-category="custom"
              onClick={() => {
                setContextMenu(null);
                setActiveTileCategory("custom");
              }}
            >
              {isEn ? "Custom" : "自定义"}
            </button>
            <button
              type="button"
              className={activeTileCategory === "common" ? "is-active" : ""}
              aria-pressed={activeTileCategory === "common"}
              data-tile-category="common"
              onClick={() => {
                setContextMenu(null);
                setActiveTileCategory("common");
              }}
            >
              {isEn ? "Common" : "常用"}
            </button>
            <button
              type="button"
              className={"formula-tile-edit-toggle" + (editing ? " is-done" : "")}
              data-toolbar-edit-toggle
              aria-pressed={editing}
              aria-label={editing ? (isEn ? "Done" : "完成") : isEn ? "Edit tiles" : "编辑磁贴"}
              title={editing ? (isEn ? "Done (Esc)" : "完成 (Esc)") : isEn ? "Edit tiles" : "编辑磁贴"}
              onClick={() => {
                setContextMenu(null);
                setEditing(!editing);
              }}
            >
              {editing ? <Check size={15} /> : <SlidersHorizontal size={14} />}
            </button>
            {onCollapseTiles && (
              <button
                type="button"
                className="formula-tile-collapse-button"
                data-formula-tile-collapse
                aria-label={isEn ? "Collapse formula tiles" : "收起公式磁贴"}
                title={isEn ? "Collapse formula tiles" : "收起公式磁贴"}
                onClick={() => {
                  setContextMenu(null);
                  onCollapseTiles();
                }}
              >
                <PanelRightClose size={15} />
              </button>
            )}
          </nav>

          {activeTileCategory === "custom" && (
            <div className="custom-formula-tile-controls">
              <div className="custom-formula-tile-actions">
                <button
                  type="button"
                  className="save-current-formula-tile"
                  disabled={!activeLineLatex || !activeCustomSection}
                  onClick={saveActiveFormulaAsTile}
                >
                  <Plus size={14} />
                  {activeCustomSection
                    ? isEn
                      ? `Save to “${activeCustomSection.name}”`
                      : `保存到「${activeCustomSection.name}」`
                    : isEn
                      ? "Create a section first"
                      : "请先新建分区"}
                </button>
                <button
                  type="button"
                  className="create-formula-tile-section"
                  disabled={
                    customTileLibrary.sections.length >= maxCustomFormulaSections
                  }
                  onClick={beginCreateSection}
                  title={isEn ? "Create section" : "新建分区"}
                >
                  <FolderPlus size={14} />
                  {isEn ? "Section" : "分区"}
                </button>
                <button
                  type="button"
                  className="open-custom-symbol-designer"
                  data-open-custom-symbol-designer
                  onClick={() => {
                    setContextMenu(null);
                    openCustomSymbolDesigner();
                  }}
                  title={isEn ? "Open custom symbol designer" : "打开自定义字符设计器"}
                >
                  <Pencil size={14} />
                  {isEn ? "Symbol designer" : "字符设计器"}
                </button>
              </div>
              <span role="status" data-tile-save-notice={tileSaveNotice ? "" : undefined}>
                {tileSaveNotice ?? (!activeCustomSection
                  ? isEn
                    ? "No section"
                    : "无分区"
                  : activeLineLatex
                    ? isEn
                      ? ""
                      : ""
                    : isEn
                      ? "Select a non-empty formula line"
                      : "请先选择非空公式行")}
              </span>
              {sectionEditor && (
                <div className="custom-formula-section-editor">
                  <input
                    autoFocus
                    type="text"
                    maxLength={24}
                    value={sectionEditor.value}
                    placeholder={isEn ? "Section name" : "分区名称"}
                    onChange={(event) =>
                      setSectionEditor({
                        ...sectionEditor,
                        value: event.currentTarget.value,
                      })
                    }
                    onKeyDown={(event) => {
                      if (event.key === "Enter") commitSectionEditor();
                      if (event.key === "Escape") setSectionEditor(null);
                    }}
                  />
                  <button
                    type="button"
                    className="icon-button compact"
                    disabled={!sectionEditor.value.trim()}
                    onClick={commitSectionEditor}
                    aria-label={isEn ? "Confirm section" : "确认分区"}
                  >
                    <Check size={14} />
                  </button>
                  <button
                    type="button"
                    className="icon-button compact"
                    onClick={() => setSectionEditor(null)}
                    aria-label={isEn ? "Cancel" : "取消"}
                  >
                    <X size={14} />
                  </button>
                </div>
              )}
            </div>
          )}

          {activeTileCategory === "common" ? (
            <div
              className={"formula-tile-list is-common" + (editing ? " is-toolbar-editing" : "")}
              {...(editing ? { "data-drop-kind": "common-tiles" } : {})}
            >
              {visibleFormulaTiles.map((tile, tileIndex) => (
                <div
                  className="formula-tile-item"
                  key={tile.id}
                  {...(editing
                    ? {
                        "data-drag-index": tileIndex,
                        onPointerDown: (event: ReactPointerEvent<HTMLElement>) =>
                          armCardDrag(event, { type: "common-tile", index: tileIndex, tileId: tile.id }),
                      }
                    : {})}
                >
                  <button
                    type="button"
                    className="formula-tile-button"
                    data-formula-tile-id={tile.id}
                    data-formula-tile-latex={tile.latex}
                    onClick={() => {
                      if (!editing) insertFormulaTile(tile);
                    }}
                    onKeyDown={(event) =>
                      moveByKeyboard(
                        event,
                        visibleFormulaTiles.length,
                        tileIndex,
                        (toIndex) =>
                          dragHandlers.drop(
                            { type: "common-tile", index: tileIndex, tileId: tile.id },
                            { type: "common-tiles", index: toIndex },
                          ),
                        () => removeCommonFormulaTile(tile.id),
                        (toIndex) => `[data-drop-kind="common-tiles"] [data-drag-index="${toIndex}"] .formula-tile-button`,
                      )
                    }
                    onContextMenu={(event) =>
                      openFormulaContextMenu(
                        event,
                        tileHotkeyTarget(tile, "common-tile"),
                        undefined,
                        tile.id,
                      )
                    }
                    aria-label={isEn ? tile.labelEn : tile.labelZh}
                    title={
                      isEn
                        ? tile.labelEn
                        : tile.labelZh
                    }
                  >
                    <MathPreview
                      latex={tile.latex}
                      className="formula-tile-preview"
                      fit
                      fluidHeight
                    />
                  </button>
                  <button
                    type="button"
                    className="formula-tile-remove"
                    data-formula-tile-remove={tile.id}
                    data-toolbar-remove
                    aria-label={
                      isEn
                        ? `Remove ${tile.labelEn} from Common`
                        : `从常用中移除${tile.labelZh}`
                    }
                    title={isEn ? "Remove from Common" : "从常用中移除"}
                    onClick={() => removeCommonFormulaTile(tile.id)}
                  >
                    <X size={11} />
                  </button>
                </div>
              ))}
              <button
                type="button"
                className="formula-tile-add"
                data-formula-tile-add
                onClick={() => {
                  setContextMenu(null);
                  setTileLibraryOpen(true);
                }}
                title={
                  isEn
                    ? "Add tiles from the formula library"
                    : "从公式库添加磁贴"
                }
              >
                <Plus size={13} />
                {isEn ? "Add tile" : "添加磁贴"}
              </button>
            </div>
          ) : (
            <div className="formula-tile-list is-custom">
              {registeredCustomSymbols.length > 0 ? (
                <section
                  className="registered-custom-symbol-toolbar-section"
                  data-registered-custom-symbol-toolbar-section
                >
                  <header>
                    <strong>{isEn ? "Registered symbols" : "已注册字符"}</strong>
                    <span>{registeredCustomSymbols.length}</span>
                  </header>
                  <div className="registered-custom-symbol-toolbar-grid">
                    {registeredCustomSymbols.map((symbol) => {
                      const latex = `\\${symbol.command}`;
                      const command = registeredCustomSymbolCommands.get(latex);
                      const pendingDelete = pendingCustomSymbolDeleteId === symbol.id;
                      return (
                        <div
                          className="registered-custom-symbol-toolbar-item"
                          key={symbol.id}
                          data-registered-custom-symbol-toolbar-item={symbol.id}
                        >
                          <button
                            type="button"
                            className="formula-tile-button is-registered-custom-symbol"
                            data-registered-custom-symbol-toolbar={symbol.id}
                            data-registered-custom-symbol-command={symbol.command}
                            onClick={() => {
                              if (command) onInsert(command);
                            }}
                            aria-label={`${symbol.name} · ${latex}`}
                            title={`${symbol.name} · ${latex}`}
                          >
                            <MathPreview
                              latex={latex}
                              className="formula-tile-preview"
                              fit
                              fluidHeight
                            />
                            <span className="registered-custom-symbol-toolbar-command">
                              {latex}
                            </span>
                          </button>
                          <button
                            type="button"
                            className={`registered-custom-symbol-toolbar-delete${pendingDelete ? " is-confirming" : ""}`}
                            data-delete-registered-custom-symbol-toolbar={symbol.id}
                            onClick={(event) => {
                              event.stopPropagation();
                              requestDeleteRegisteredCustomSymbol(symbol.id);
                            }}
                            onBlur={() => {
                              if (pendingDelete) setPendingCustomSymbolDeleteId(null);
                            }}
                            aria-label={
                              pendingDelete
                                ? isEn
                                  ? `Confirm deletion of ${symbol.name}`
                                  : `确认删除${symbol.name}`
                                : isEn
                                  ? `Delete ${symbol.name}`
                                  : `删除${symbol.name}`
                            }
                            title={
                              pendingDelete
                                ? isEn
                                  ? "Click again to confirm"
                                  : "再次点击确认删除"
                                : isEn
                                  ? "Delete custom symbol"
                                  : "删除自定义字符"
                            }
                          >
                            <Trash2 size={11} />
                            <span>{pendingDelete ? (isEn ? "Confirm" : "确认") : isEn ? "Delete" : "删除"}</span>
                          </button>
                        </div>
                      );
                    })}
                  </div>
                </section>
              ) : null}
              {customTileDefinitions.length === 0 &&
                customTileLibrary.sections.length === 0 && (
                  <div className="formula-tile-empty">
                    <strong>{isEn ? "No custom tiles" : "无自定义磁贴"}</strong>
                  </div>
                )}
              <div
                className={
                  "custom-formula-sections" +
                  (customTileDefinitions.length === 0 &&
                  customTileLibrary.sections.length === 0
                    ? " is-empty-hidden"
                    : "")
                }
              >
                {customSections.map((section) => {
                  const sectionTiles = customTileDefinitions.filter(
                    (tile) => tile.sectionId === section.id,
                  );
                  const storedSection = customTileLibrary.sections.find(
                    (item) => item.id === section.id,
                  );
                  const sectionRows = customTileRowsForSection(
                    sectionTiles,
                    section.id,
                  );
                  return (
                    <section
                      className={
                        "custom-formula-section" +
                        (activeCustomSection?.id === section.id
                          ? " is-active"
                          : "")
                      }
                      data-custom-section-id={section.id}
                      key={section.id}
                    >
                      <header className="custom-formula-section-header">
                        <button
                          type="button"
                          className="custom-formula-section-select"
                          onClick={() => setActiveCustomSectionId(section.id)}
                          aria-pressed={activeCustomSection?.id === section.id}
                        >
                          <strong>{section.name}</strong>
                          <span>{sectionTiles.length}</span>
                        </button>
                        {storedSection && (
                          <div className="custom-formula-section-actions">
                            {pendingSectionDeleteId === storedSection.id ? (
                              <>
                                <span className="custom-formula-section-delete-copy">
                                  {isEn
                                    ? `Delete ${sectionTiles.length} tile${sectionTiles.length === 1 ? "" : "s"}?`
                                    : `确认删除分区和其中 ${sectionTiles.length} 个磁贴？`}
                                </span>
                                <button
                                  type="button"
                                  className="icon-button compact is-danger"
                                  onClick={() => deleteCustomSection(storedSection.id)}
                                  aria-label={isEn ? "Confirm delete" : "确认删除"}
                                  title={isEn ? "Confirm delete" : "确认删除"}
                                >
                                  <Check size={13} />
                                </button>
                                <button
                                  type="button"
                                  className="icon-button compact"
                                  onClick={() => setPendingSectionDeleteId(null)}
                                  aria-label={isEn ? "Cancel delete" : "取消删除"}
                                  title={isEn ? "Cancel" : "取消"}
                                >
                                  <X size={13} />
                                </button>
                              </>
                            ) : (
                              <>
                                <button
                                  type="button"
                                  className="icon-button compact"
                                  onClick={() => beginRenameSection(storedSection)}
                                  aria-label={
                                    isEn
                                      ? `Rename ${storedSection.name}`
                                      : `重命名${storedSection.name}`
                                  }
                                  title={isEn ? "Rename section" : "重命名分区"}
                                >
                                  <Pencil size={13} />
                                </button>
                                <button
                                  type="button"
                                  className="icon-button compact is-danger"
                                  onClick={() => {
                                    setSectionEditor(null);
                                    setPendingSectionDeleteId(storedSection.id);
                                  }}
                                  aria-label={
                                    isEn
                                      ? `Delete ${storedSection.name}`
                                      : `删除${storedSection.name}`
                                  }
                                  title={
                                    isEn
                                      ? "Delete section"
                                      : "删除分区"
                                  }
                                >
                                  <Trash2 size={13} />
                                </button>
                              </>
                            )}
                          </div>
                        )}
                      </header>
                      <div
                        className="custom-formula-tile-grid"
                        data-custom-tile-grid-section={section.id}
                        {...(editing
                          ? { "data-drop-kind": "custom-section", "data-section-id": section.id }
                          : {})}
                      >
                        {sectionTiles.length === 0 && (
                          <button
                            type="button"
                            className="custom-formula-section-empty"
                            onClick={() => setActiveCustomSectionId(section.id)}
                          >
                            {isEn
                              ? "Empty section"
                              : "空分区"}
                          </button>
                        )}
                        {sectionRows.map((row) => (
                          <div
                            className={
                              "custom-formula-tile-row" +
                              (row.fill ? "" : " is-loose")
                            }
                            data-custom-tile-row={row.id}
                            key={row.id}
                          >
                            {row.items.map((item) => {
                              const tile = item.tile;
                              const tileIndex = sectionTiles.indexOf(tile);
                              const tileStyle = {
                                flex: row.fill
                                  ? `${item.weight} 1 0px`
                                  : `0 1 ${item.minimumWidth}px`,
                                minWidth: `${item.minimumWidth}px`,
                                ...(tile.color
                                  ? { "--custom-tile-color": tile.color }
                                  : {}),
                              } as CSSProperties;
                              return (
                                <button
                                  type="button"
                                  className={
                                    "formula-tile-button is-custom" +
                                    (tile.color ? " has-custom-color" : "")
                                  }
                                  style={tileStyle}
                                  key={tile.id}
                                  data-formula-tile-id={tile.id}
                                  data-formula-tile-latex={tile.latex}
                                  data-custom-tile-weight={item.weight}
                                  data-custom-tile-min-width={item.minimumWidth}
                                  {...(editing
                                    ? {
                                        "data-drag-index": tileIndex,
                                        onPointerDown: (event: ReactPointerEvent<HTMLElement>) =>
                                          armCardDrag(event, { type: "custom-tile", tileId: tile.id }),
                                        onKeyDown: (event: ReactKeyboardEvent<HTMLElement>) =>
                                          moveByKeyboard(
                                            event,
                                            sectionTiles.length,
                                            tileIndex,
                                            (toIndex) =>
                                              dragHandlers.drop(
                                                { type: "custom-tile", tileId: tile.id },
                                                { type: "custom-section", sectionId: section.id, index: toIndex },
                                              ),
                                            () =>
                                              deleteCustomFormulaTile(
                                                tile.id,
                                                formulaHotkeyTargetIdForTile("custom-tile", tile.id, tile.latex),
                                              ),
                                            (toIndex) =>
                                              `[data-section-id="${section.id}"] [data-drag-index="${toIndex}"]`,
                                          ),
                                      }
                                    : {})}
                                  onClick={() => {
                                    if (!editing) insertFormulaTile(tile);
                                  }}
                                  onContextMenu={(event) =>
                                    openFormulaContextMenu(
                                      event,
                                      tileHotkeyTarget(tile, "custom-tile"),
                                      tile.id,
                                    )
                                  }
                                  aria-label={isEn ? tile.labelEn : tile.labelZh}
                                  title={
                                    isEn
                                      ? tile.labelEn
                                      : tile.labelZh
                                  }
                                >
                                  <MathPreview
                                    latex={tile.latex}
                                    className="formula-tile-preview"
                                    showPlaceholders
                                    fit
                                    fluidHeight
                                    minimumFluidScale={0.8}
                                    maximumFluidScale={1.2}
                                    fitInsetRatio={0.84}
                                    minimumFluidHeight={44}
                                    fluidVerticalPadding={8}
                                    onMeasure={({ width }) =>
                                      recordCustomTileNaturalWidth(tile.id, width)
                                    }
                                  />
                                  {editing && (
                                    <span
                                      role="button"
                                      className="toolbar-item-remove"
                                      data-toolbar-remove
                                      aria-label={isEn ? "Delete tile" : "删除磁贴"}
                                      title={isEn ? "Delete tile" : "删除磁贴"}
                                      onPointerDown={(event) => event.stopPropagation()}
                                      onClick={(event) => {
                                        event.preventDefault();
                                        event.stopPropagation();
                                        deleteCustomFormulaTile(
                                          tile.id,
                                          formulaHotkeyTargetIdForTile("custom-tile", tile.id, tile.latex),
                                        );
                                      }}
                                    >
                                      <X size={10} strokeWidth={2.6} />
                                    </span>
                                  )}
                                </button>
                              );
                            })}
                          </div>
                        ))}
                      </div>
                    </section>
                  );
                })}
              </div>
            </div>
          )}
        </section>
      )}

      {contextMenu && (
        <div
          ref={contextMenuRef}
          className="formula-tile-context-menu formula-hotkey-context-menu"
          role="menu"
          aria-label={isEn ? "Formula item actions" : "公式项目操作"}
          style={{ left: contextMenu.x, top: contextMenu.y }}
        >
          <div className="formula-hotkey-context-heading">
            <strong>
              {formulaHotkeyTargetLabel(contextMenu.target, language)}
            </strong>
            <span>
              {contextBinding
                ? formatFormulaHotkeyChord(contextBinding.chord)
                : isEn
                  ? "No hotkey"
                  : "未设置快捷键"}
            </span>
          </div>
          {contextMenuItem(contextMenu) && toolbarGroups.length > 0 && (
            <>
              <button
                type="button"
                role="menuitem"
                className="formula-hotkey-context-action has-submenu"
                data-add-to-group-menu
                aria-expanded={Boolean(addToGroupOpen)}
                onPointerEnter={(event) => {
                  if (!addToGroupOpen) setAddToGroupOpen(addToGroupPlace(event.currentTarget));
                }}
                onClick={(event) =>
                  setAddToGroupOpen(addToGroupOpen ? null : addToGroupPlace(event.currentTarget))
                }
              >
                <FolderPlus size={14} />
                {isEn ? "Add to group" : "添加到分组"}
                <ChevronRight size={13} className="formula-hotkey-context-chevron" />
              </button>
              {addToGroupOpen && createPortal(
                <div className="formula-hotkey-context-submenu formula-tile-context-menu" role="group" style={addToGroupOpen}>
                  {toolbarGroups.map((group) => {
                    const item = contextMenuItem(contextMenu) as ToolbarItem;
                    const has = (itemsByGroup.get(group.id) ?? []).some((candidate) =>
                      sameToolbarItem(candidate, item),
                    );
                    return (
                      <button
                        key={group.id}
                        type="button"
                        role="menuitemcheckbox"
                        aria-checked={has}
                        className="formula-hotkey-context-action"
                        data-add-to-group={group.id}
                        {...(group.id === "common" && item.kind === "command"
                          ? { "data-add-to-common-command": item.id }
                          : {})}
                        disabled={has}
                        onClick={() => {
                          layoutStore().insertItem(group.id, END_INDEX, item);
                          setContextMenu(null);
                        }}
                      >
                        <span className="formula-hotkey-context-check">{has && <Check size={13} />}</span>
                        {groupLabel(group)}
                      </button>
                    );
                  })}
                </div>,
                document.body,
              )}
            </>
          )}
          {contextMenu.toolbarItem && (
            <button
              type="button"
              role="menuitem"
              className="formula-hotkey-context-action"
              data-remove-from-group={contextMenu.toolbarItem.groupId}
              onClick={() => {
                const place = contextMenu.toolbarItem;
                if (place) removeToolbarItem(place.groupId, place.index);
                setContextMenu(null);
              }}
            >
              <Trash2 size={14} />
              {isEn ? "Remove from group" : "从分组中移除"}
            </button>
          )}
          {!editing && (
            <button
              type="button"
              role="menuitem"
              className="formula-hotkey-context-action"
              data-edit-toolbar
              onClick={() => {
                setEditing(true);
                setContextMenu(null);
              }}
            >
              <SlidersHorizontal size={14} />
              {isEn ? "Edit toolbar" : "编辑工具栏"}
            </button>
          )}
          <button
            type="button"
            role="menuitem"
            className="formula-hotkey-context-action"
            onClick={() => {
              setHotkeyTarget(contextMenu.target);
              setContextMenu(null);
            }}
          >
            <Keyboard size={14} />
            {contextBinding
              ? isEn
                ? "Change hotkey…"
                : "修改快捷键…"
              : isEn
                ? "Set hotkey…"
                : "设置快捷键…"}
          </button>
          {contextBinding && (
            <button
              type="button"
              role="menuitem"
              className="formula-hotkey-context-action"
              onClick={() => {
                removeBindingsForTarget(contextMenu.target.id);
                setContextMenu(null);
              }}
            >
              {isEn ? "Clear hotkey" : "清除快捷键"}
            </button>
          )}
          {contextMenu.commonTileId && (
            <button
              type="button"
              role="menuitem"
              className="formula-hotkey-context-action"
              data-remove-common-tile={contextMenu.commonTileId}
              onClick={() =>
                removeCommonFormulaTile(contextMenu.commonTileId as string)
              }
            >
              <Trash2 size={14} />
              {isEn ? "Remove from Common" : "从常用中移除"}
            </button>
          )}
          {contextCustomTile && (
            <>
              <div className="formula-hotkey-context-divider" />
              <div className="custom-tile-context-options">
                <div className="custom-tile-context-label">
                  <Palette size={14} />
                  <span>{isEn ? "Tile color" : "磁贴颜色"}</span>
                </div>
                <div
                  className="custom-tile-color-palette"
                  role="group"
                  aria-label={isEn ? "Tile color" : "磁贴颜色"}
                >
                  <button
                    type="button"
                    className={
                      "custom-tile-color-swatch is-auto" +
                      (!contextCustomTile.color ? " is-active" : "")
                    }
                    data-custom-tile-color="auto"
                    onClick={() =>
                      updateCustomTile(contextCustomTile.id, { color: null })
                    }
                    aria-label={isEn ? "Automatic color" : "自动颜色"}
                    title={isEn ? "Automatic color" : "自动颜色"}
                  />
                  {customTileColorPresets.map((color) => (
                    <button
                      type="button"
                      className={
                        "custom-tile-color-swatch" +
                        (contextCustomTile.color === color ? " is-active" : "")
                      }
                      data-custom-tile-color={color}
                      key={color}
                      style={{ backgroundColor: color }}
                      onClick={() =>
                        updateCustomTile(contextCustomTile.id, { color })
                      }
                      aria-label={`${isEn ? "Use color" : "使用颜色"} ${color}`}
                      title={color}
                    />
                  ))}
                  <label
                    className="custom-tile-color-picker"
                    title={isEn ? "Choose any color" : "选择任意颜色"}
                  >
                    <input
                      type="color"
                      value={contextCustomTile.color ?? "#6f8fbf"}
                      onChange={(event) =>
                        updateCustomTile(contextCustomTile.id, {
                          color: validCustomTileColor(event.currentTarget.value),
                        })
                      }
                      aria-label={isEn ? "Choose any color" : "选择任意颜色"}
                    />
                    <Plus size={12} />
                  </label>
                </div>

                <label className="custom-tile-section-picker">
                  <span>{isEn ? "Section" : "所属分区"}</span>
                  <select
                    value={contextCustomTile.sectionId}
                    onChange={(event) => {
                      updateCustomTile(contextCustomTile.id, {
                        sectionId: event.currentTarget.value,
                      });
                      setActiveCustomSectionId(event.currentTarget.value);
                    }}
                  >
                    {customSections.map((section) => (
                      <option value={section.id} key={section.id}>
                        {section.name}
                      </option>
                    ))}
                  </select>
                </label>
              </div>
              <div className="formula-hotkey-context-divider" />
              <button
                type="button"
                role="menuitem"
                className="formula-hotkey-context-action is-danger"
                onClick={() =>
                  deleteCustomFormulaTile(
                    contextCustomTile.id,
                    contextMenu.target.id,
                  )
                }
              >
                <Trash2 size={14} />
                {isEn ? "Delete tile" : "删除磁贴"}
              </button>
            </>
          )}
        </div>
      )}

      <ToolbarItemPickerDialog
        open={pickerGroupId !== null}
        language={language}
        groupName={(() => {
          const group = toolbarGroups.find((candidate) => candidate.id === pickerGroupId);
          return group ? groupLabel(group) : "";
        })()}
        has={(item) =>
          (itemsByGroup.get(pickerGroupId ?? "") ?? []).some((candidate) =>
            sameToolbarItem(candidate, item),
          )
        }
        onToggle={(item) => {
          if (pickerGroupId) toggleItemInGroup(pickerGroupId, item);
        }}
        currentLatex={activeLineLatex}
        onClose={() => setPickerGroupId(null)}
      />
      <FormulaTileLibraryDialog
        open={tileLibraryOpen}
        language={language}
        selectedIds={commonFormulaTileIds}
        onToggle={toggleCommonFormulaTile}
        onClose={() => setTileLibraryOpen(false)}
      />
      <FormulaHotkeyRecorderDialog
        target={hotkeyTarget}
        onClose={() => setHotkeyTarget(null)}
      />
    </aside>
  );
}
