import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type MouseEvent,
} from "react";
import {
  Check,
  FolderPlus,
  Keyboard,
  Palette,
  Pencil,
  Plus,
  Trash2,
  X,
} from "lucide-react";
import type { LatexCommand } from "../types/command";
import {
  categoryLabels,
  categoryLabelsEn,
  calculusCommandIds,
  commandRegistry,
  commonCommandIds,
} from "../autocomplete/commandRegistry";
import { MathPreview } from "../components/MathPreview";
import { FormulaHotkeyRecorderDialog } from "../components/FormulaHotkeyRecorderDialog";
import {
  createFormulaHotkeyTarget,
  formatFormulaHotkeyChord,
  formulaHotkeyTargetIdForCommand,
  formulaHotkeyTargetIdForTile,
  formulaHotkeyTargetLabel,
  type FormulaHotkeyTarget,
} from "../shortcuts/formulaHotkeys";
import { useEditorStore } from "../stores/editorStore";
import { useFormulaHotkeyStore } from "../stores/formulaHotkeyStore";

type MatrixDelimiter = "bmatrix" | "pmatrix" | "vmatrix";
type ToolbarView = "tools" | "tiles";
type ToolbarLayout = "sidebar" | "horizontal";
type TileCategory = "custom" | "common";

interface Props {
  onInsert: (command: LatexCommand) => void;
  view?: ToolbarView;
  layout?: ToolbarLayout;
  className?: string;
  stabilizeTileLayout?: boolean;
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
  x: number;
  y: number;
}

interface SectionEditorState {
  mode: "create" | "rename";
  sectionId?: string;
  value: string;
}

const customFormulaTilesStorageKey = "visualtex-custom-formula-tiles";
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
    const expandable = [...normalized].sort(
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
    return {
      id: `${index}-${items.map((item) => item.tile.id).join("-")}`,
      items: fill ? normalizeCustomTileRowWeights(items) : items,
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

const commonFormulaTiles: FormulaTileDefinition[] = [
  {
    id: "quadratic-formula",
    latex: "x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}",
    labelZh: "一元二次方程求根公式",
    labelEn: "Quadratic formula",
  },
  {
    id: "euler-identity",
    latex: "e^{i\\pi}+1=0",
    labelZh: "欧拉恒等式",
    labelEn: "Euler identity",
  },
  {
    id: "pythagorean-theorem",
    latex: "a^2+b^2=c^2",
    labelZh: "勾股定理",
    labelEn: "Pythagorean theorem",
  },
  {
    id: "binomial-theorem",
    latex: "(a+b)^n=\\sum_{k=0}^{n}\\binom{n}{k}a^{n-k}b^k",
    labelZh: "二项式定理",
    labelEn: "Binomial theorem",
  },
  {
    id: "gaussian-integral",
    latex: "\\int_{-\\infty}^{\\infty}e^{-x^2}\\,\\mathrm{d}x=\\sqrt{\\pi}",
    labelZh: "高斯积分",
    labelEn: "Gaussian integral",
  },
  {
    id: "taylor-series",
    latex: "f(x)=\\sum_{n=0}^{\\infty}\\frac{f^{(n)}(a)}{n!}(x-a)^n",
    labelZh: "泰勒展开",
    labelEn: "Taylor series",
  },
  {
    id: "mass-energy",
    latex: "E=mc^2",
    labelZh: "质能方程",
    labelEn: "Mass-energy equivalence",
  },
  {
    id: "schrodinger-equation",
    latex: "i\\hbar\\frac{\\partial}{\\partial t}\\Psi=\\hat{H}\\Psi",
    labelZh: "含时薛定谔方程",
    labelEn: "Time-dependent Schrödinger equation",
  },
  {
    id: "gauss-law",
    latex: "\\nabla\\cdot\\mathbf{E}=\\frac{\\rho}{\\varepsilon_0}",
    labelZh: "高斯定律",
    labelEn: "Gauss's law",
  },
  {
    id: "characteristic-equation",
    latex: "\\det(A-\\lambda I)=0",
    labelZh: "矩阵特征方程",
    labelEn: "Matrix characteristic equation",
  },
];

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

const categories = [
  "common",
  "structure",
  "calculus",
  "matrix",
  "relation",
  "greek",
  "arrow",
  "physics",
  "set",
];

const matrixGridCells = Array.from({ length: 100 }, (_, index) => ({
  row: Math.floor(index / 10) + 1,
  column: (index % 10) + 1,
}));
const matrixDelimiterOptions: Array<{
  id: MatrixDelimiter;
  preview: string;
  labelZh: string;
  labelEn: string;
}> = [
  {
    id: "vmatrix",
    preview: "\\begin{vmatrix}a&b\\\\c&d\\end{vmatrix}",
    labelZh: "竖线",
    labelEn: "Bars",
  },
  {
    id: "bmatrix",
    preview: "\\begin{bmatrix}a&b\\\\c&d\\end{bmatrix}",
    labelZh: "方括号",
    labelEn: "Brackets",
  },
  {
    id: "pmatrix",
    preview: "\\begin{pmatrix}a&b\\\\c&d\\end{pmatrix}",
    labelZh: "圆括号",
    labelEn: "Parentheses",
  },
];

const hiddenToolbarCommandIds = new Set(["time-ordering"]);
const wideToolbarCommandIds = new Set([
  "matrixelement",
  "expectation-operator",
]);
const toolbarPreviewMaximumScale = 1;
const toolbarPreviewInsetRatio = 0.88;

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
  lim: "\\lim_{x\\to0}",
  "lim-infty": "\\lim_{x\\to\\infty}",
  "lim-left": "\\lim_{x\\to a^-}",
  "lim-right": "\\lim_{x\\to a^+}",
  derivative: "\\frac{\\mathrm{d}}{\\mathrm{d}x}",
  secondderivative: "\\frac{\\mathrm{d}^{2}}{\\mathrm{d}x^{2}}",
  partial: "\\frac{\\partial}{\\partial x}",
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
  cases: "\\begin{cases}a\\\\b\\end{cases}",
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

export function FormulaToolbar({
  onInsert,
  view: fixedView,
  layout = "sidebar",
  className = "",
  stabilizeTileLayout = false,
}: Props) {
  const [internalActiveView, setInternalActiveView] =
    useState<ToolbarView>("tools");
  const activeView = fixedView ?? internalActiveView;
  const [activeTileCategory, setActiveTileCategory] =
    useState<TileCategory>("common");
  const [customTileLibrary, setCustomTileLibrary] =
    useState<CustomFormulaTileLibrary>(loadCustomFormulaTiles);
  const [activeCustomSectionId, setActiveCustomSectionId] = useState(
    defaultCustomSectionId,
  );
  const [sectionEditor, setSectionEditor] =
    useState<SectionEditorState | null>(null);
  const [pendingSectionDeleteId, setPendingSectionDeleteId] =
    useState<string | null>(null);
  const [contextMenu, setContextMenu] =
    useState<FormulaContextMenuState | null>(null);
  const [hotkeyTarget, setHotkeyTarget] =
    useState<FormulaHotkeyTarget | null>(null);
  const [activeCategory, setActiveCategory] = useState("common");
  const [matrixRows, setMatrixRows] = useState(2);
  const [matrixColumns, setMatrixColumns] = useState(2);
  const [matrixHover, setMatrixHover] = useState<{
    rows: number;
    columns: number;
  } | null>(null);
  const [matrixDelimiter, setMatrixDelimiter] =
    useState<MatrixDelimiter>("bmatrix");
  const toolbarRef = useRef<HTMLElement>(null);
  const [customTileGridWidths, setCustomTileGridWidths] = useState<
    Record<string, number>
  >({});
  const [customTileNaturalWidths, setCustomTileNaturalWidths] = useState<
    Record<string, number>
  >({});
  const language = useEditorStore((state) => state.language);
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
    persistCustomFormulaTiles(customTileLibrary);
  }, [customTileLibrary]);

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
    const commitWidths = (entries: Array<[string, number]>) => {
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

    document.addEventListener("pointerdown", closeFromPointer);
    document.addEventListener("keydown", closeFromKey);
    window.addEventListener("blur", closeMenu);
    window.addEventListener("resize", closeMenu);
    window.addEventListener("scroll", closeMenu, true);

    return () => {
      document.removeEventListener("pointerdown", closeFromPointer);
      document.removeEventListener("keydown", closeFromKey);
      window.removeEventListener("blur", closeMenu);
      window.removeEventListener("resize", closeMenu);
      window.removeEventListener("scroll", closeMenu, true);
    };
  }, [contextMenu]);

  const visibleCommands = useMemo(() => {
    const preferredIds = activeCategory === "common"
      ? commonCommandIds
      : activeCategory === "calculus"
        ? calculusCommandIds
        : null;
    const candidates = preferredIds
      ? preferredIds
          .map((id) => commandRegistry.find((command) => command.id === id))
          .filter((command): command is LatexCommand => Boolean(command))
      : commandRegistry.filter(
          (command) =>
            command.category === activeCategory &&
            !hiddenToolbarCommandIds.has(command.id),
        );
    const seenCommandIds = new Set<string>();
    return candidates.filter((command) => {
      if (seenCommandIds.has(command.id)) return false;
      seenCommandIds.add(command.id);
      return true;
    });
  }, [activeCategory]);

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
  const visibleFormulaTiles =
    activeTileCategory === "common"
      ? commonFormulaTiles
      : customTileDefinitions;

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

  const saveActiveFormulaAsTile = () => {
    if (!activeLineLatex || !activeCustomSection) return;
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
  const openFormulaContextMenu = (
    event: MouseEvent<HTMLButtonElement>,
    target: FormulaHotkeyTarget,
    customTileId?: string,
  ) => {
    event.preventDefault();
    event.stopPropagation();
    const menuWidth = customTileId ? 252 : 224;
    const menuHeight = customTileId ? 330 : 132;
    setContextMenu({
      target,
      customTileId,
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

  return (
    <aside
      ref={toolbarRef}
      className={
        "formula-toolbar" +
        (fixedView ? " is-view-fixed" : "") +
        (layout === "horizontal" ? " is-horizontal" : "") +
        (className ? ` ${className}` : "")
      }
      data-toolbar-layout={layout}
      data-toolbar-fixed-view={fixedView ?? ""}
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
      <nav className="toolbar-tabs" aria-label={isEn ? "Formula categories" : "公式分类"}>
        {categories.map((category) => (
          <button
            key={category}
            type="button"
            className={
              "toolbar-tab " +
              (activeCategory === category ? "is-active" : "")
            }
            data-category={category}
            aria-pressed={activeCategory === category}
            onClick={() => setActiveCategory(category)}
          >
            {(isEn ? categoryLabelsEn : categoryLabels)[category]}
          </button>
        ))}
      </nav>

      <div
        key={`${layout}-${activeCategory}`}
        className="template-strip"
        data-active-category={activeCategory}
        aria-label={isEn ? "Formula templates" : "公式模板"}
      >
        {activeCategory === "matrix" && (
          <section className="matrix-builder" aria-label={isEn ? "Custom matrix" : "自定义矩阵"}>
            <div className="matrix-options-column">
              <div className="matrix-builder-heading">
                <strong>{isEn ? "Custom matrix" : "自定义矩阵"}</strong>
                <span className="matrix-size-badge" aria-live="polite">
                  {previewRows} × {previewColumns}
                </span>
              </div>

              <div className="matrix-delimiter-picker">
                <span className="matrix-control-label">
                  {isEn ? "Delimiter" : "边界样式"}
                </span>
                <div className="matrix-delimiter-options" role="group" aria-label={isEn ? "Matrix delimiter" : "矩阵边界"}>
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
                        fit
                        maximumFitScale={1.15}
                        fitInsetRatio={0.72}
                      />
                    </button>
                  ))}
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
                    ? "Insert the selected matrix · Right-click to set a hotkey"
                    : "插入当前矩阵 · 右键设置快捷键"
                }
              >
                {isEn
                  ? `Insert ${matrixRows} × ${matrixColumns}`
                  : `插入 ${matrixRows} × ${matrixColumns}`}
              </button>
            </div>

            <div className="matrix-size-picker">
              <span className="matrix-control-label">
                {isEn ? "Size" : "矩阵尺寸"}
              </span>
              <div
                className="matrix-size-grid"
                role="grid"
                aria-label={
                  isEn
                    ? "Select matrix rows and columns"
                    : "选择矩阵行数和列数"
                }
                aria-rowcount={10}
                aria-colcount={10}
                onPointerLeave={() => setMatrixHover(null)}
                onBlur={(event) => {
                  if (
                    !event.currentTarget.contains(
                      event.relatedTarget as Node | null,
                    )
                  ) {
                    setMatrixHover(null);
                  }
                }}
              >
                {matrixGridCells.map(({ row, column }) => {
                  const previewed =
                    row <= previewRows && column <= previewColumns;
                  const selectedCorner =
                    row === matrixRows && column === matrixColumns;
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
                      onPointerEnter={() =>
                        setMatrixHover({ rows: row, columns: column })
                      }
                      onFocus={() =>
                        setMatrixHover({ rows: row, columns: column })
                      }
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

          </section>
        )}

        {visibleCommands.map((command) => {
          const previewLatex = toolbarPreviewLatex(command);
          const widePreview = wideToolbarCommandIds.has(command.id);
          return (
            <button
              type="button"
              className={
                "template-button is-unified-fit" +
                (widePreview ? " is-wide-preview" : "")
              }
              data-command-id={command.id}
              data-preview-latex={previewLatex}
              key={command.id}
              onClick={() => onInsert(command)}
              onContextMenu={(event) =>
                openFormulaContextMenu(event, commandHotkeyTarget(command))
              }
              aria-label={isEn ? command.labelEn : command.labelZh}
              title={
                (isEn ? command.labelEn : command.labelZh) +
                " · " +
                command.command +
                (isEn
                  ? " · Right-click to set a hotkey"
                  : " · 右键设置快捷键")
              }
            >
              <MathPreview
                latex={previewLatex}
                fit
                maximumFitScale={toolbarPreviewMaximumScale}
                fitInsetRatio={toolbarPreviewInsetRatio}
              />
            </button>
          );
        })}
      </div>
        </>
      ) : (
        <section
          className="formula-tiles-panel"
          aria-label={isEn ? "Formula tiles" : "公式磁贴"}
        >
          <nav
            className="formula-tile-tabs"
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
              </div>
              <span>
                {!activeCustomSection
                  ? isEn
                    ? "Create a section before saving formula tiles."
                    : "请先新建一个分区，再保存公式磁贴。"
                  : activeLineLatex
                    ? isEn
                      ? "The selected formula line will be saved in the active section."
                      : "当前公式行将保存到选中的分区。"
                    : isEn
                      ? "Select a non-empty formula line first."
                      : "请先选择一个非空公式行。"}
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
            <div className="formula-tile-list">
              {visibleFormulaTiles.map((tile) => (
                <button
                  type="button"
                  className="formula-tile-button"
                  key={tile.id}
                  data-formula-tile-id={tile.id}
                  data-formula-tile-latex={tile.latex}
                  onClick={() => insertFormulaTile(tile)}
                  onContextMenu={(event) =>
                    openFormulaContextMenu(
                      event,
                      tileHotkeyTarget(tile, "common-tile"),
                    )
                  }
                  aria-label={isEn ? tile.labelEn : tile.labelZh}
                  title={
                    isEn
                      ? `${tile.labelEn} · Right-click for hotkey settings`
                      : `${tile.labelZh} · 右键设置快捷键`
                  }
                >
                  <MathPreview
                    latex={tile.latex}
                    className="formula-tile-preview"
                    fit
                    fluidHeight
                  />
                </button>
              ))}
            </div>
          ) : (
            <div className="formula-tile-list is-custom">
              {customTileDefinitions.length === 0 &&
                customTileLibrary.sections.length === 0 && (
                  <div className="formula-tile-empty">
                    <strong>{isEn ? "No custom tiles yet" : "还没有自定义磁贴"}</strong>
                    <span>
                      {isEn
                        ? "Select a formula line, then save it to a section."
                        : "选中一个公式行，然后保存到分区中。"}
                    </span>
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
                                      ? "Delete this section and all of its tiles"
                                      : "删除整个分区及其中所有磁贴"
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
                      >
                        {sectionTiles.length === 0 && (
                          <button
                            type="button"
                            className="custom-formula-section-empty"
                            onClick={() => setActiveCustomSectionId(section.id)}
                          >
                            {isEn
                              ? "Select this section, then save a formula here"
                              : "选择此分区后，可将公式保存到这里"}
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
                                  onClick={() => insertFormulaTile(tile)}
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
                                      ? `${tile.labelEn} · Right-click for hotkey, color and section`
                                      : `${tile.labelZh} · 右键设置快捷键、颜色和分区`
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

      <FormulaHotkeyRecorderDialog
        target={hotkeyTarget}
        onClose={() => setHotkeyTarget(null)}
      />
    </aside>
  );
}
