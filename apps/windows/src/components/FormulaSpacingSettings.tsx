import { useState } from "react";
import { ChevronRight, RotateCcw, Ruler } from "lucide-react";
import { useEditorStore } from "../stores/editorStore";
import { activeFormulaSpacingConfig, resetActiveSpacing } from "../editor/formulaSpacingPreferences";
import { FORMULA_LETTER_FONT_OPTIONS } from "../editor/formulaFontPreferences";
import { FormulaSpacingStudio } from "./FormulaSpacingStudio";

/**
 * Settings > Interface customization: the chosen formula font's spacing
 * configuration, opened in the spacing studio (specimen beside the sliders),
 * and its one-click way back to LaTeX's spacing.
 */
export function FormulaSpacingSettings({ isEn }: { isEn: boolean }) {
  const font = useEditorStore((state) => state.formulaLetterFont);
  const spacing = useEditorStore((state) => state.formulaSpacing);
  const setFormulaSpacing = useEditorStore((state) => state.setFormulaSpacing);
  const [studioOpen, setStudioOpen] = useState(false);
  const fontLabel = FORMULA_LETTER_FONT_OPTIONS.find((o) => o.id === font)?.label ?? font;
  const active = activeFormulaSpacingConfig(spacing, font);
  const adjusted = Object.keys(active.values).length;
  const configName = active.name || (isEn ? "Default" : "默认");

  return (
    <section
      className="formula-spacing-customization"
      aria-labelledby="formula-spacing-customization-title"
      data-formula-spacing
    >
      <header className="formula-inset-customization-header">
        <div>
          <strong id="formula-spacing-customization-title">
            {isEn ? "Formula spacing" : "公式间距"}
          </strong>
          <small data-formula-spacing-summary>
            {isEn
              ? `${fontLabel} · ${configName} · ${adjusted ? `${adjusted} adjusted` : "LaTeX's spacing"}`
              : `${fontLabel} · ${configName} · ${adjusted ? `已调整 ${adjusted} 项` : "LaTeX 标准间距"}`}
          </small>
        </div>
        <button
          type="button"
          className="secondary-button formula-inset-reset"
          data-formula-spacing-reset
          disabled={!adjusted}
          onClick={() => setFormulaSpacing(resetActiveSpacing(spacing, font))}
        >
          <RotateCcw size={15} />
          {isEn ? "Restore defaults" : "恢复默认"}
        </button>
      </header>
      <button
        type="button"
        className="settings-subdialog-trigger"
        data-formula-spacing-studio-trigger
        onClick={() => setStudioOpen(true)}
      >
        <Ruler size={18} />
        <span>
          <strong>{isEn ? "Adjust spacing on a specimen" : "在样张上调整间距"}</strong>
          <small>
            {isEn
              ? "Scripts, fractions, radicals, limits, operators, matrices and multi-line formulas"
              : "上下标、分式、根号、上下限、运算符、矩阵与多行公式"}
          </small>
        </span>
        <ChevronRight size={18} aria-hidden="true" />
      </button>
      {studioOpen && <FormulaSpacingStudio isEn={isEn} onClose={() => setStudioOpen(false)} />}
    </section>
  );
}
