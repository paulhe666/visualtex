import { alignWordInlineSvg } from "./wordImageGeometry";
import { latexToSvg } from "../../export/latexToSvg";
import type { SvgExportResult } from "../../export/exportTypes";
import type { LatexCodeFormat } from "../../types/formula";
import type {
  FormulaChineseFont,
  FormulaLetterFont,
} from "../../editor/formulaFontPreferences";
import { inConversionStage } from "./conversionErrors";
import { useEditorStore } from "../../stores/editorStore";
import {
  latexLinesToOmmlArtifacts,
  type OmmlArtifacts,
} from "../omml/latexToOmml";
import { wordEquationMathMl } from "../wordMathMl/wordEquationMathMl";
import {
  normalizeFormulaEditorDocument,
  serializeFormulaEditorDocument,
  type FormulaEditorLine,
} from "./formulaEditorDocument";
import type { InlineImageMathStyle } from "./formulaMetadata";

export const OFFICE_FORMULA_REFERENCE_FONT_SIZE_PT = 14;

export interface RenderOfficeFormulaArtifactsInput {
  lines: FormulaEditorLine[];
  codeFormat: LatexCodeFormat;
  displayMode: "inline" | "block";
  inlineImageMathStyle?: InlineImageMathStyle;
  host?: "word" | "powerpoint";
  includeWordOmml?: boolean;
  numbered?: boolean;
  fontSizePt?: number;
  formulaLetterFont?: FormulaLetterFont;
  formulaChineseFont?: FormulaChineseFont;
}

export interface OfficeFormulaRenderArtifacts {
  lines: FormulaEditorLine[];
  codeFormat: LatexCodeFormat;
  canonicalLatex: string;
  svg: SvgExportResult;
  /**
   * PowerPoint native equations: the editor's own OMML, used only when the app
   * receives no wordMathMl.
   */
  omml: OmmlArtifacts | null;
  /**
   * Word and PowerPoint: the MathML Windows writes its OMML from. VisualTeX's
   * shared writer (src-tauri office/omml_write, the port of Windows' C# writer)
   * turns it into the equation, so both platforms start from identical OMML.
   */
  wordMathMl: string | null;
}

/**
 * Builds the canonical source and every vector/native artifact from the same
 * normalized editor document. Document import and edit replacement must use
 * this path so an align/align* formula cannot render differently after edit.
 */
export async function renderOfficeFormulaArtifacts({
  lines,
  codeFormat,
  displayMode,
  inlineImageMathStyle = "text",
  host,
  includeWordOmml = true,
  numbered = false,
  fontSizePt,
  formulaLetterFont,
  formulaChineseFont,
}: RenderOfficeFormulaArtifactsInput): Promise<OfficeFormulaRenderArtifacts> {
  const document = normalizeFormulaEditorDocument(lines, codeFormat);
  const canonicalLatex = serializeFormulaEditorDocument(document);
  if (!canonicalLatex.trim()) {
    throw new Error("Cannot render an empty Office formula.");
  }
  // A single raw formula may contain source-formatting newlines inside one
  // logical TeX expression (for example between \left[ and \right]). Passing
  // those newlines to latexToSvg makes its generic multi-row fallback wrap the
  // expression in `aligned`, where the inserted `&` markers become illegal.
  // Preserve canonicalLatex for editing/metadata, but render these internal
  // line breaks as ordinary TeX whitespace. Genuine multi-line documents and
  // explicit align/gather environments retain their structured rendering.
  const svgLatex =
    document.codeFormat === "raw" && document.lines.length === 1
      ? canonicalLatex.replace(/[ \t]*\n[ \t]*/g, " ")
      : canonicalLatex;
  const svgDisplayMode =
    displayMode === "block" ||
    (host === "word" &&
      displayMode === "inline" &&
      inlineImageMathStyle === "display");
  let svg: SvgExportResult;
  try {
    svg = await latexToSvg(svgLatex, {
      displayMode: svgDisplayMode,
      fontSizePt: OFFICE_FORMULA_REFERENCE_FONT_SIZE_PT,
      // Word uses the imported image bounds as part of its line box. A 10 px
      // display margin nearly doubled the apparent box height at 14 pt even
      // though the painted glyphs were unchanged. Keep Word exports tight while
      // retaining the wider PowerPoint margin used for slide selection.
      paddingPx:
        displayMode === "inline" ? 1 : host === "word" ? 2 : 10,
      background: "transparent",
      forceExplicitBlack: host === "word",
      formulaLetterFont,
      formulaChineseFont,
    });
  } catch (reason) {
    const isEn = useEditorStore.getState().language === "en";
    throw inConversionStage(isEn, "无法绘制公式图片", "The formula image could not be drawn", reason);
  }
  if (host === "word" && displayMode === "inline" && fontSizePt !== undefined) {
    svg = alignWordInlineSvg(svg, fontSizePt, formulaLetterFont);
  }
  let omml: OmmlArtifacts | null = null;
  let wordMathMl: string | null = null;
  if (includeWordOmml) {
    try {
      wordMathMl = wordEquationMathMl(document, displayMode);
    } catch (reason) {
      const isEn = useEditorStore.getState().language === "en";
      throw inConversionStage(isEn, "无法生成 Word 公式（OMML）", "The Word equation (OMML) could not be generated", reason);
    }
  }
  if (includeWordOmml && host !== "word") {
    try {
      omml = latexLinesToOmmlArtifacts(
        document.lines.map((line) => line.latex),
        displayMode,
        document.codeFormat,
        { formulaLetterFont, formulaChineseFont },
        numbered,
      );
    } catch (reason) {
      const isEn = useEditorStore.getState().language === "en";
      throw inConversionStage(isEn, "无法生成 Word 公式（OMML）", "The Word equation (OMML) could not be generated", reason);
    }
  }
  return {
    lines: document.lines,
    codeFormat: document.codeFormat,
    canonicalLatex,
    svg,
    omml,
    wordMathMl,
  };
}

/**
 * Autosave runs while a formula is still being composed. MathLive templates
 * deliberately expose temporary source such as `\\placeholder{}`, unclosed
 * groups/environments and a command that is only partly typed. Those states
 * must remain editable and persist their source, but they are not valid Office
 * artifacts yet. Keep explicit insert/apply strict by using
 * renderOfficeFormulaArtifacts there; only draft persistence uses this
 * best-effort wrapper.
 */
export async function tryRenderOfficeFormulaDraftArtifacts(
  input: RenderOfficeFormulaArtifactsInput,
): Promise<OfficeFormulaRenderArtifacts | null> {
  try {
    return await renderOfficeFormulaArtifacts(input);
  } catch {
    return null;
  }
}
