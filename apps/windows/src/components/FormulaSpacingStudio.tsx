import { useEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import { RotateCcw, X } from "lucide-react";
import { MathPreview } from "./MathPreview";
import { useEditorStore } from "../stores/editorStore";
import {
  SPACING_DEFINITIONS,
  spacingPercent,
  type SpacingDefinition,
  type SpacingKey,
} from "../../../shared/tex-layout/src/spacing.ts";
import {
  FORMULA_SPACING_MAX_CONFIGS,
  activeFormulaSpacingConfig,
  deleteSpacingConfig,
  renameSpacingConfig,
  resetActiveSpacing,
  saveSpacingConfigAs,
  selectSpacingConfig,
  setSpacingValue,
} from "../editor/formulaSpacingPreferences";
import { FORMULA_LETTER_FONT_OPTIONS } from "../editor/formulaFontPreferences";

export const SPACING_GROUP_LABELS: Record<SpacingDefinition["group"], [string, string]> = {
  scripts: ["Superscripts and subscripts", "上下标"],
  fractions: ["Fractions and radicals", "分式与根号"],
  limits: ["Large operators and delimiters", "大型运算符与定界符"],
  spaces: ["Spacing around symbols", "符号间距"],
  arrays: ["Matrices and multi-line formulas", "矩阵与多行公式"],
};

export const SPACING_LABELS: Record<SpacingKey, [string, string]> = {
  supRaise: ["Superscript height", "上标高度"],
  subDrop: ["Subscript depth", "下标深度"],
  scriptAttach: ["Scripts beside tall symbols", "高大符号旁的上下标位置"],
  scriptSpace: ["Space after scripts", "上下标后的空隙"],
  numRaise: ["Numerator height", "分子高度"],
  denomDrop: ["Denominator depth", "分母深度"],
  ruleThickness: ["Fraction and radical line thickness", "分数线与根号线粗细"],
  radicalGap: ["Gap inside radicals", "根号内间隙"],
  limitsGap: ["Gap between operator and limits", "运算符与上下限的间隙"],
  limitsDistance: ["Distance of limits", "上下限的距离"],
  limitsPadding: ["Space outside limits", "上下限外侧留白"],
  delimiterSize: ["Size of \\left \\right delimiters", "自动伸缩括号的大小"],
  thinSpace: ["Thin space (functions, punctuation)", "细间距（函数名、标点）"],
  medSpace: ["Medium space (+, −, ×)", "中间距（+ − × 等运算符）"],
  thickSpace: ["Thick space (=, <, →)", "粗间距（= < → 等关系符）"],
  arrayColSep: ["Matrix column spacing", "矩阵列间距"],
  arrayRowSep: ["Matrix row spacing", "矩阵行距"],
  alignRowSep: ["Multi-line formula row spacing", "多行公式行距"],
};

const GROUPS = Object.keys(SPACING_GROUP_LABELS) as SpacingDefinition["group"][];
const GROUP_OF = new Map(SPACING_DEFINITIONS.map((d) => [d.key, d.group]));

/**
 * Marks the parts of the specimen an adjustment moves: \class is transparent
 * to the layout (the content takes part in the list as if unwrapped) and
 * gives the part an element the studio outlines while that slider is used.
 */
function mark(keys: SpacingKey[], tex: string): string {
  return keys.reduceRight((inner, key) => `\\class{vtx-spec-${key}}{${inner}}`, tex);
}

/** The specimen: every structure the adjustments act on, in the editor's own layout. */
const SPECIMEN: { group: SpacingDefinition["group"]; caption: [string, string]; latex: string }[] = [
  {
    group: "scripts",
    caption: ["Scripts", "上下标"],
    latex:
      mark(["scriptSpace"], `x_{${mark(["subDrop"], "i")}}^{${mark(["supRaise"], "2")}}`) +
      `+e^{${mark(["supRaise"], "-x^{2}")}}+a_{${mark(["subDrop"], "n+1")}}+` +
      String.raw`\left(\frac{a}{b}\right)^{${mark(["scriptAttach"], "2")}}`,
  },
  {
    group: "fractions",
    caption: ["Fractions and radicals", "分式与根号"],
    latex:
      mark(["ruleThickness"], String.raw`\frac{${mark(["numRaise"], "a+b")}}{${mark(["denomDrop"], "c+d")}}`) +
      "+" +
      mark(["ruleThickness", "radicalGap"], String.raw`\sqrt{x^{2}+1}`) +
      "+" +
      mark(["radicalGap"], String.raw`\sqrt[3]{y}`),
  },
  {
    group: "limits",
    caption: ["Large operators and delimiters", "大型运算符与定界符"],
    latex:
      String.raw`\sum_{${mark(["limitsGap", "limitsDistance", "limitsPadding"], "k=1")}}^{${mark(["limitsGap", "limitsDistance", "limitsPadding"], "n")}}k^{2}` +
      String.raw`=\prod_{${mark(["limitsGap", "limitsDistance", "limitsPadding"], "i=1")}}^{${mark(["limitsGap", "limitsDistance", "limitsPadding"], "m")}}a_{i}` +
      "+" +
      mark(["delimiterSize"], String.raw`\left[\frac{1}{1+\frac{1}{x}}\right]`),
  },
  {
    group: "spaces",
    caption: ["Spacing around symbols", "符号间距"],
    latex:
      mark(["thinSpace"], String.raw`\sin x`) +
      "+" +
      mark(["medSpace"], String.raw`a\times b`) +
      mark(["thickSpace"], String.raw`=c\to d`) +
      mark(["thinSpace"], ",\\ f(x)"),
  },
  {
    group: "arrays",
    caption: ["Matrix", "矩阵"],
    latex:
      mark(["arrayColSep", "arrayRowSep"], String.raw`\begin{pmatrix}a&b&c\\d&e&f\end{pmatrix}`) +
      String.raw`,\quad ` +
      mark(["arrayColSep", "arrayRowSep"], String.raw`\begin{cases}1&x>0\\0&x\le 0\end{cases}`),
  },
  {
    group: "arrays",
    caption: ["Multi-line formula", "多行公式"],
    latex: mark(["alignRowSep"], String.raw`\begin{aligned}f(x)&=x^{2}+1\\g(x)&=\frac{1}{x}\end{aligned}`),
  },
];

/** The adjustments each specimen row shows (its marked parts). */
const ROW_KEYS = SPECIMEN.map((row) => new Set([...row.latex.matchAll(/vtx-spec-(\w+)/g)].map((m) => m[1] as SpacingKey)));

const HIGHLIGHT_CSS = SPACING_DEFINITIONS.map(
  (d) => `.formula-spacing-studio[data-active-key="${d.key}"] .vtx-spec-${d.key}`,
).join(",\n") + " { background: var(--formula-spacing-mark-bg); box-shadow: 0 0 0 1.5px var(--formula-spacing-mark-ring); border-radius: 3px; }";

/**
 * The spacing studio: the specimen in large type beside the sliders of the
 * chosen font's active configuration. Using a slider outlines the parts it
 * moves; every change applies at once to the editor, previews, images and
 * the Office window.
 */
export function FormulaSpacingStudio({ isEn, onClose }: { isEn: boolean; onClose: () => void }) {
  const font = useEditorStore((state) => state.formulaLetterFont);
  const spacing = useEditorStore((state) => state.formulaSpacing);
  const setFormulaSpacing = useEditorStore((state) => state.setFormulaSpacing);
  const [naming, setNaming] = useState<{ mode: "save" | "rename"; name: string } | null>(null);
  const [hovered, setHovered] = useState<SpacingKey | null>(null);
  const [focused, setFocused] = useState<SpacingKey | null>(null);
  const dialogRef = useRef<HTMLElement>(null);
  const specimenRef = useRef<HTMLDivElement>(null);
  const onCloseRef = useRef(onClose);
  onCloseRef.current = onClose;
  const fontLabel = FORMULA_LETTER_FONT_OPTIONS.find((o) => o.id === font)?.label ?? font;
  const state = spacing[font];
  const active = activeFormulaSpacingConfig(spacing, font);
  const nameOf = (name: string) => name || (isEn ? "Default" : "默认");
  const changed = Object.keys(active.values).length > 0;
  const activeKey = hovered ?? focused;
  const activeGroup = activeKey ? GROUP_OF.get(activeKey) : undefined;
  const t = (pair: [string, string]) => (isEn ? pair[0] : pair[1]);

  useEffect(() => {
    const previous = document.activeElement as HTMLElement | null;
    const frame = window.requestAnimationFrame(() => {
      dialogRef.current?.querySelector<HTMLElement>("[data-formula-spacing-studio-close]")?.focus();
    });
    // Escape and Tab stay in the studio (the settings dialogs beneath listen on the document)
    const onKeyDown = (event: KeyboardEvent) => {
      const dialog = dialogRef.current;
      if (!dialog) return;
      if (event.key === "Escape") {
        event.preventDefault();
        event.stopImmediatePropagation();
        onCloseRef.current();
        return;
      }
      if (event.key !== "Tab") return;
      event.stopImmediatePropagation();
      const focusable = Array.from(
        dialog.querySelectorAll<HTMLElement>('button:not(:disabled), input:not(:disabled), select:not(:disabled), [tabindex]:not([tabindex="-1"])'),
      );
      const first = focusable[0], last = focusable[focusable.length - 1];
      if (!first || !last) return;
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
      }
    };
    window.addEventListener("keydown", onKeyDown, true);
    return () => {
      window.cancelAnimationFrame(frame);
      window.removeEventListener("keydown", onKeyDown, true);
      previous?.focus({ preventScroll: true });
    };
  }, []);

  // the specimen row of the slider in use comes into view (below the fold, or in the narrow stacked layout)
  useEffect(() => {
    const specimen = specimenRef.current;
    const index = activeKey ? ROW_KEYS.findIndex((keys) => keys.has(activeKey)) : -1;
    const row = index >= 0 ? specimen?.querySelector<HTMLElement>(`[data-specimen-row="${index}"]`) : null;
    if (!specimen || !row) return;
    const view = specimen.getBoundingClientRect();
    const r = row.getBoundingClientRect();
    if (r.top >= view.top && r.bottom <= view.bottom) return;
    const top = specimen.scrollTop + (r.top - view.top) - Math.max(0, (view.height - r.height) / 2);
    specimen.scrollTo({ top: Math.max(0, top), behavior: "smooth" });
  }, [activeKey]);

  const commitName = () => {
    if (!naming) return;
    const name = naming.name.trim();
    if (!name) return;
    if (naming.mode === "save") {
      const next = saveSpacingConfigAs(spacing, font, name);
      if (next) setFormulaSpacing(next);
    } else {
      setFormulaSpacing(renameSpacingConfig(spacing, font, active.id, name));
    }
    setNaming(null);
  };

  return createPortal(
    <div
      className="settings-subdialog-backdrop formula-spacing-studio-backdrop"
      role="presentation"
      onMouseDown={(event) => {
        event.stopPropagation();
        onClose();
      }}
    >
      <section
        ref={dialogRef}
        className="formula-spacing-studio"
        role="dialog"
        aria-modal="true"
        aria-labelledby="formula-spacing-studio-title"
        data-formula-spacing-studio
        data-active-key={activeKey ?? undefined}
        onMouseDown={(event) => event.stopPropagation()}
      >
        <style>{HIGHLIGHT_CSS}</style>
        <header className="dialog-header">
          <div>
            <span className="eyebrow">SPACING</span>
            <h2 id="formula-spacing-studio-title">
              {isEn ? "Formula spacing" : "公式间距"}
              <small>{isEn ? `${fontLabel} · percentages of LaTeX's spacing` : `${fontLabel} · 以 LaTeX 标准间距为 100%`}</small>
            </h2>
          </div>
          <button
            type="button"
            className="icon-button"
            data-formula-spacing-studio-close
            onClick={onClose}
            aria-label={isEn ? "Close formula spacing" : "关闭公式间距"}
          >
            <X size={18} />
          </button>
        </header>

        <div className="formula-spacing-studio-toolbar formula-spacing-configs">
          {naming ? (
            <>
              <input
                type="text"
                autoFocus
                maxLength={60}
                value={naming.name}
                data-formula-spacing-name
                placeholder={isEn ? "Configuration name" : "配置名称"}
                onChange={(event) => setNaming({ ...naming, name: event.currentTarget.value })}
                onKeyDown={(event) => {
                  if (event.key === "Enter") commitName();
                  if (event.key === "Escape") {
                    event.stopPropagation();
                    event.nativeEvent.stopImmediatePropagation();
                    setNaming(null);
                  }
                }}
              />
              <button type="button" className="primary-button" data-formula-spacing-name-save disabled={!naming.name.trim()} onClick={commitName}>
                {isEn ? "Save" : "保存"}
              </button>
              <button type="button" className="secondary-button" onClick={() => setNaming(null)}>
                {isEn ? "Cancel" : "取消"}
              </button>
            </>
          ) : (
            <>
              <select
                value={active.id}
                data-formula-spacing-config
                aria-label={isEn ? "Spacing configuration" : "间距配置"}
                onChange={(event) => setFormulaSpacing(selectSpacingConfig(spacing, font, event.currentTarget.value))}
              >
                {state.configs.map((c) => (
                  <option key={c.id} value={c.id}>
                    {nameOf(c.name)}
                  </option>
                ))}
              </select>
              <button
                type="button"
                className="secondary-button"
                data-formula-spacing-save-as
                disabled={state.configs.length >= FORMULA_SPACING_MAX_CONFIGS}
                onClick={() => setNaming({ mode: "save", name: "" })}
              >
                {isEn ? "Save as new" : "另存为新配置"}
              </button>
              <button
                type="button"
                className="secondary-button"
                data-formula-spacing-rename
                onClick={() => setNaming({ mode: "rename", name: active.name })}
              >
                {isEn ? "Rename" : "重命名"}
              </button>
              <button
                type="button"
                className="secondary-button"
                data-formula-spacing-delete
                disabled={state.configs.length <= 1}
                onClick={() => setFormulaSpacing(deleteSpacingConfig(spacing, font, active.id))}
              >
                {isEn ? "Delete" : "删除"}
              </button>
              <button
                type="button"
                className="secondary-button formula-spacing-studio-reset"
                data-formula-spacing-reset
                disabled={!changed}
                onClick={() => setFormulaSpacing(resetActiveSpacing(spacing, font))}
              >
                <RotateCcw size={15} />
                {isEn ? "Restore defaults" : "恢复默认"}
              </button>
            </>
          )}
        </div>

        <div className="formula-spacing-studio-body">
          <div ref={specimenRef} className="formula-spacing-specimen" data-formula-spacing-preview aria-label={isEn ? "Specimen" : "效果样张"}>
            {SPECIMEN.map((row, index) => (
              <figure
                key={index}
                className={`formula-spacing-specimen-row${activeKey && ROW_KEYS[index].has(activeKey) ? " is-active" : ""}`}
                data-specimen-group={row.group}
                data-specimen-row={index}
              >
                <figcaption>{t(row.caption)}</figcaption>
                <MathPreview latex={row.latex} staticLayout />
              </figure>
            ))}
          </div>

          <div className="formula-spacing-controls">
            {GROUPS.map((group) => (
              <div
                key={group}
                className={`formula-spacing-group${activeGroup === group ? " is-active" : ""}`}
                role="group"
                aria-label={t(SPACING_GROUP_LABELS[group])}
              >
                <strong className="formula-spacing-group-title">{t(SPACING_GROUP_LABELS[group])}</strong>
                {SPACING_DEFINITIONS.filter((d) => d.group === group).map((d) => {
                  const value = spacingPercent(active.values, d.key);
                  const label = t(SPACING_LABELS[d.key]);
                  return (
                    <div
                      key={d.key}
                      className={`formula-spacing-slider${value !== 100 ? " is-adjusted" : ""}${activeKey === d.key ? " is-active" : ""}`}
                      onMouseEnter={() => setHovered(d.key)}
                      onMouseLeave={() => setHovered((k) => (k === d.key ? null : k))}
                    >
                      <span className="formula-spacing-slider-label">
                        <strong>{label}</strong>
                        <small>{value}%</small>
                      </span>
                      <input
                        type="range"
                        min={d.min}
                        max={d.max}
                        step="5"
                        value={value}
                        data-formula-spacing-setting={d.key}
                        aria-label={label}
                        aria-valuetext={`${value}%`}
                        onFocus={() => setFocused(d.key)}
                        onBlur={() => setFocused((k) => (k === d.key ? null : k))}
                        onChange={(event) => setFormulaSpacing(setSpacingValue(spacing, font, d.key, Number(event.target.value)))}
                        onDoubleClick={() => setFormulaSpacing(setSpacingValue(spacing, font, d.key, 100))}
                      />
                      <button
                        type="button"
                        className="icon-button formula-spacing-slider-reset"
                        data-formula-spacing-setting-reset={d.key}
                        disabled={value === 100}
                        title={isEn ? "Back to 100%" : "恢复为 100%"}
                        aria-label={isEn ? `${label}: back to 100%` : `${label}：恢复为 100%`}
                        onFocus={() => setFocused(d.key)}
                        onBlur={() => setFocused((k) => (k === d.key ? null : k))}
                        onClick={() => setFormulaSpacing(setSpacingValue(spacing, font, d.key, 100))}
                      >
                        <RotateCcw size={13} />
                      </button>
                    </div>
                  );
                })}
              </div>
            ))}
          </div>
        </div>

        <footer className="dialog-footer">
          <span>
            {isEn
              ? "Applies at once to the editor, previews, exported and copied images and the Office window."
              : "修改立即作用于编辑器、预览、导出与复制的图片以及 Office 窗口。"}
          </span>
          <button type="button" className="primary-button" onClick={onClose}>
            {isEn ? "Done" : "完成"}
          </button>
        </footer>
      </section>
    </div>,
    document.body,
  );
}
