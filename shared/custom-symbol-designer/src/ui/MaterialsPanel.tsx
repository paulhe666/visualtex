import { useEffect, useMemo, useState } from "react";
import { Loader2, Plus, Search } from "lucide-react";
import type { CustomSymbolGeometryPreset } from "../designerTypes.ts";
import { CUSTOM_SYMBOL_GEOMETRY_PRESETS } from "../geometry.ts";
import type { CustomSymbolDesignerPlatform, DesignerCommand } from "../platform.ts";
import {
  SYSTEM_GLYPH_CATEGORIES,
  SYSTEM_GLYPH_CATEGORY_LABELS,
  SYSTEM_MATH_FONT_PRESETS,
  detectSystemMathFontAvailability,
  searchSystemGlyphs,
  systemFontPresetById,
  type SystemGlyphCategory,
  type SystemMathFontId,
} from "../systemGlyphs.ts";

// The materials: LaTeX symbols from the palette, characters from the system's
// math fonts, and shapes. A click adds one to the symbol.

const MATERIAL_CATEGORIES = ["common", "calculus", "greek", "relation", "set", "arrow", "physics"] as const;

const COMMON_COMMAND_IDS = new Set([
  "intplain", "partial", "nabla", "infty", "alpha", "beta", "gamma", "theta", "lambda", "pi", "sigma", "omega",
  "equal", "neq", "approx", "leq", "geq", "in", "notin", "subset", "forall", "exists", "rightarrow", "leftarrow", "hbar",
]);

const BARE_OPERATORS = new Set([
  "\\int", "\\iint", "\\iiint", "\\oint", "\\oiint", "\\oiiint", "\\sum", "\\prod", "\\coprod",
  "\\bigcup", "\\bigcap", "\\bigvee", "\\bigwedge", "\\partial",
]);

const LARGE_OPERATORS = new Set([
  "int", "iint", "iiint", "oint", "oiint", "oiiint", "sum", "prod", "coprod", "bigcup", "bigcap", "bigvee", "bigwedge",
]);

/** Large operators are added in their display size. */
export function normalizeMaterialLatex(source: string) {
  const trimmed = source.trim();
  if (/^\\(?:display|text|script|scriptscript)style\b/.test(trimmed)) return trimmed;
  const match = /^\\([A-Za-z]+)$/.exec(trimmed);
  if (!match || !LARGE_OPERATORS.has(match[1])) return trimmed;
  return `\\displaystyle${trimmed}`;
}

/** The single symbol a palette command typesets, or null when it is not one. */
function bareMaterialSource(command: DesignerCommand) {
  if (!command.supportedInMathMode) return null;
  if (!(MATERIAL_CATEGORIES as readonly string[]).includes(command.category)) return null;
  const raw = command.command.trim();
  const literal = Array.from(raw).length === 1 && raw !== "\\";
  const controlWord = /^\\[A-Za-z]+$/.test(raw);
  if (!literal && !controlWord) return null;
  if (command.id.startsWith("custom-symbol:") || BARE_OPERATORS.has(raw)) return normalizeMaterialLatex(raw);
  if (command.insertTemplate.includes("\\placeholder")) return null;
  if (["\\begin", "\\end", "\\left", "\\right", "\\middle"].includes(raw)) return null;
  return normalizeMaterialLatex(raw);
}

export interface MaterialEntry {
  command: DesignerCommand;
  source: string;
}

export function materialEntries(commands: readonly DesignerCommand[]): MaterialEntry[] {
  const sorted = commands
    .map((command) => ({ command, source: bareMaterialSource(command) }))
    .filter((entry): entry is MaterialEntry => typeof entry.source === "string")
    .sort(
      (left, right) =>
        right.command.defaultPriority - left.command.defaultPriority ||
        left.command.labelZh.localeCompare(right.command.labelZh),
    );
  const bySource = new Map<string, MaterialEntry>();
  for (const entry of sorted) if (!bySource.has(entry.source)) bySource.set(entry.source, entry);
  return [...bySource.values()];
}

type Tab = "latex" | "system" | "shapes";

interface Props {
  platform: CustomSymbolDesignerPlatform;
  isEn: boolean;
  revision: number;
  busyKey: string | null;
  error: string;
  status: string;
  onAddLatex: (source: string) => void;
  onAddSystemGlyph: (character: string, label: string, fontId: SystemMathFontId, italic: boolean) => void;
  onAddShape: (preset: Exclude<CustomSymbolGeometryPreset, "eraser">) => void;
}

export function shapeLabel(preset: CustomSymbolGeometryPreset, isEn: boolean) {
  const labels: Record<CustomSymbolGeometryPreset, [string, string]> = {
    line: ["线段", "Line"],
    circle: ["圆", "Circle"],
    ellipse: ["椭圆", "Ellipse"],
    rect: ["矩形", "Rectangle"],
    triangle: ["三角形", "Triangle"],
    arrow: ["箭头", "Arrow"],
    arc: ["圆弧", "Arc"],
    eraser: ["橡皮擦", "Eraser"],
  };
  return labels[preset][isEn ? 1 : 0];
}

export function MaterialsPanel({
  platform,
  isEn,
  revision,
  busyKey,
  error,
  status,
  onAddLatex,
  onAddSystemGlyph,
  onAddShape,
}: Props) {
  const [tab, setTab] = useState<Tab>("latex");
  const [source, setSource] = useState("");
  const [search, setSearch] = useState("");
  const [category, setCategory] = useState<string>("common");
  const [fontId, setFontId] = useState<SystemMathFontId>("cambria-math");
  const [italic, setItalic] = useState(true);
  const [glyphSearch, setGlyphSearch] = useState("");
  const [glyphCategory, setGlyphCategory] = useState<SystemGlyphCategory>("basic-italic");
  const [availability, setAvailability] = useState<Partial<Record<SystemMathFontId, boolean>>>({});
  const { MathPreview } = platform;

  const entries = useMemo(() => materialEntries(platform.commands()), [platform, revision]);
  const shown = useMemo(() => {
    const query = search.trim().toLocaleLowerCase();
    const bare = query.replace(/^\\/, "");
    return entries.filter(({ command }) => {
      if (!query) {
        if (category === "common") return command.id.startsWith("custom-symbol:") || COMMON_COMMAND_IDS.has(command.id);
        return command.category === category;
      }
      const text = [command.command, command.labelZh, command.labelEn, ...command.aliases, ...command.keywords]
        .join(" ")
        .toLocaleLowerCase();
      return text.includes(query) || text.includes(bare);
    });
  }, [category, entries, search]);
  const glyphs = useMemo(() => searchSystemGlyphs(glyphCategory, glyphSearch), [glyphCategory, glyphSearch]);
  const font = systemFontPresetById(fontId);

  useEffect(() => {
    if (tab !== "system") return;
    let cancelled = false;
    void detectSystemMathFontAvailability(platform.systemGlyphs).then((result) => {
      if (!cancelled) setAvailability(result);
    });
    return () => {
      cancelled = true;
    };
  }, [platform, tab]);

  const categoryLabels = isEn ? platform.categoryLabels.en : platform.categoryLabels.zh;
  const tabs: Array<[Tab, string]> = [
    ["latex", isEn ? "Symbols" : "符号"],
    ["system", isEn ? "Font glyphs" : "字体字符"],
    ["shapes", isEn ? "Shapes" : "图形"],
  ];

  return (
    <section className="csd-panel csd-materials" data-csd-materials>
      <div className="csd-tabs" role="tablist">
        {tabs.map(([id, label]) => (
          <button
            key={id}
            type="button"
            role="tab"
            aria-selected={tab === id}
            className={tab === id ? "is-active" : ""}
            data-csd-material-tab={id}
            onClick={() => setTab(id)}
          >
            {label}
          </button>
        ))}
      </div>

      {tab === "latex" ? (
        <div className="csd-materials-body">
          <form
            className="csd-source-row"
            onSubmit={(event) => {
              event.preventDefault();
              if (source.trim()) onAddLatex(source);
            }}
          >
            <input
              value={source}
              spellCheck={false}
              placeholder={isEn ? "LaTeX, e.g. \\partial" : "LaTeX，如 \\partial"}
              aria-label={isEn ? "LaTeX material" : "LaTeX 素材"}
              data-csd-source-input
              onChange={(event) => setSource(event.currentTarget.value)}
            />
            <button type="submit" disabled={!source.trim()} title={isEn ? "Add" : "加入"} data-csd-source-add>
              <Plus size={14} />
            </button>
          </form>
          <label className="csd-search">
            <Search size={13} aria-hidden="true" />
            <input
              type="search"
              value={search}
              placeholder={isEn ? "Search symbols" : "搜索符号"}
              data-csd-material-search
              onChange={(event) => setSearch(event.currentTarget.value)}
            />
          </label>
          {!search.trim() ? (
            <div className="csd-chips" role="tablist">
              {MATERIAL_CATEGORIES.map((id) => (
                <button
                  key={id}
                  type="button"
                  role="tab"
                  aria-selected={category === id}
                  className={category === id ? "is-active" : ""}
                  data-csd-material-category={id}
                  onClick={() => setCategory(id)}
                >
                  {categoryLabels[id] ?? id}
                </button>
              ))}
            </div>
          ) : null}
          <div className="csd-grid" data-csd-material-grid data-count={shown.length}>
            {shown.map(({ command, source: latex }) => {
              const busy = busyKey === `latex:${latex}`;
              return (
                <button
                  key={`${command.id}:${latex}`}
                  type="button"
                  className={busy ? "is-busy" : ""}
                  title={`${isEn ? command.labelEn : command.labelZh} · ${command.command}`}
                  aria-label={isEn ? command.labelEn : command.labelZh}
                  data-csd-material={latex}
                  disabled={Boolean(busyKey)}
                  onClick={() => onAddLatex(latex)}
                >
                  {busy ? <Loader2 size={14} className="csd-spin" /> : <MathPreview latex={latex} fit />}
                </button>
              );
            })}
            {!shown.length ? <p className="csd-empty">{isEn ? "No matching symbol." : "没有匹配的符号。"}</p> : null}
          </div>
        </div>
      ) : null}

      {tab === "system" ? (
        <div className="csd-materials-body">
          <select
            value={fontId}
            aria-label={isEn ? "Font" : "字体"}
            data-csd-system-font
            onChange={(event) => setFontId(event.currentTarget.value as SystemMathFontId)}
          >
            {SYSTEM_MATH_FONT_PRESETS.map((preset) => (
              <option key={preset.id} value={preset.id}>
                {isEn ? preset.labelEn : preset.labelZh}
                {availability[preset.id] === false ? (isEn ? " · not installed" : " · 未安装") : ""}
              </option>
            ))}
          </select>
          <div className="csd-segmented">
            <button type="button" className={!italic ? "is-active" : ""} aria-pressed={!italic} onClick={() => setItalic(false)}>
              {isEn ? "Upright" : "正体"}
            </button>
            <button type="button" className={italic ? "is-active" : ""} aria-pressed={italic} onClick={() => setItalic(true)}>
              {isEn ? "Math italic" : "数学斜体"}
            </button>
          </div>
          <label className="csd-search">
            <Search size={13} aria-hidden="true" />
            <input
              type="search"
              value={glyphSearch}
              placeholder={isEn ? "Character or U+ code" : "字符或 U+ 编码"}
              data-csd-system-search
              onChange={(event) => setGlyphSearch(event.currentTarget.value)}
            />
          </label>
          <div className="csd-chips">
            {SYSTEM_GLYPH_CATEGORIES.map((id) => (
              <button
                key={id}
                type="button"
                className={glyphCategory === id ? "is-active" : ""}
                aria-pressed={glyphCategory === id}
                onClick={() => setGlyphCategory(id)}
              >
                {SYSTEM_GLYPH_CATEGORY_LABELS[id][isEn ? "en" : "zh"]}
              </button>
            ))}
          </div>
          <div className="csd-grid is-text" data-csd-system-grid>
            {glyphs.map((glyph) => {
              const busy = busyKey === `system:${font.id}:${glyph.codePoint}`;
              return (
                <button
                  key={glyph.codePoint}
                  type="button"
                  className={busy ? "is-busy" : ""}
                  title={`${glyph.character} · ${glyph.label}`}
                  disabled={Boolean(busyKey)}
                  data-csd-system-glyph={glyph.label}
                  style={{ fontFamily: font.family, fontStyle: italic ? "italic" : "normal" }}
                  onClick={() => onAddSystemGlyph(glyph.character, glyph.label, font.id, italic)}
                >
                  {busy ? <Loader2 size={14} className="csd-spin" /> : glyph.character}
                </button>
              );
            })}
          </div>
        </div>
      ) : null}

      {tab === "shapes" ? (
        <div className="csd-materials-body">
          <div className="csd-shape-grid">
            {CUSTOM_SYMBOL_GEOMETRY_PRESETS.map((preset) => (
              <button key={preset} type="button" data-csd-shape={preset} onClick={() => onAddShape(preset)}>
                <span className={`csd-shape-icon is-${preset}`} aria-hidden="true" />
                <span>{shapeLabel(preset, isEn)}</span>
              </button>
            ))}
          </div>
          <p className="csd-hint">
            {isEn
              ? "Use the eraser in the canvas toolbar to cut parts away."
              : "要去掉某些部分，用画布工具栏中的橡皮擦。"}
          </p>
        </div>
      ) : null}

      {error ? <p className="csd-message is-error" role="alert">{error}</p> : null}
      {!error && status ? <p className="csd-message" role="status">{status}</p> : null}
    </section>
  );
}
