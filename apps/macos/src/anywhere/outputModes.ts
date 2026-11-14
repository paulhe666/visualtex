/**
 * VisualTeX Anywhere's output modes: what Return puts into the field.
 *
 * - Source: the formula's LaTeX in the user's LaTeX format (inline $…$ or
 *   \(…\), the main window's settings), for fields that render LaTeX.
 * - Markdown: math-enabled Markdown's conventions whatever those settings
 *   are ($…$ in a sentence, its text outside the math, $$…$$ for several
 *   rows, portable commands), for Markdown editors and sites (GitHub,
 *   Obsidian, Typora, Notion, Jupyter...).
 * - Math Characters: Unicode text (unicodeMath.ts), for plain-text fields.
 * - Image: the formula as a PNG (the export path), pasted into the field.
 *
 * The mode is remembered per application (a chat keeps Image, an AI chat
 * keeps Source). The export path (MathJax) loads when a mode needs it.
 */
import { formatFormulaLinesUniversal } from "../../../shared/visualtex-app/src/clipboard/LatexCopyService";
import { toPortableBoldLatex } from "../../../shared/visualtex-app/src/clipboard/portableBoldLatex";
import { useEditorStore } from "../../../shared/visualtex-app/src/stores/editorStore";
import type { UnicodeMath } from "../../../shared/visualtex-app/src/anywhere/unicodeMath";

export type AnywhereOutputMode = "source" | "markdown" | "unicode" | "image";

export const OUTPUT_MODES: readonly AnywhereOutputMode[] = ["source", "markdown", "unicode", "image"];

export const OUTPUT_MODE_LABELS: Record<AnywhereOutputMode, string> = {
  source: "LaTeX",
  markdown: "Markdown",
  unicode: "Unicode",
  image: "图片",
};

type FormulaRow = Parameters<typeof formatFormulaLinesUniversal>[0][number];

/** The formula as math-enabled Markdown, each row as it is: an inline row
 *  as $…$ with no space inside the dollars (Markdown renderers require that)
 *  and its text outside the math, a display row (or a multi-line one) as a
 *  $$…$$ block on lines of its own. */
export function formulaMarkdown(rows: readonly FormulaRow[]): string {
  const { latexFormatProfile } = useEditorStore.getState();
  const profile = {
    ...latexFormatProfile,
    inlineWrapper: "dollar" as const,
    inlineTextPolicy: "outside-math" as const,
    displayWrapper: "double-dollar" as const,
    numbered: false,
  };
  return rows
    .map((row) => {
      const text = formatFormulaLinesUniversal(
        [{ ...row, latex: toPortableBoldLatex(String(row.latex ?? "")), displayStyle: "default" }],
        profile,
      );
      // a multi-line environment is display math: its own $$ block
      return /^\\begin\{/.test(text) ? `$$\n${text}\n$$` : text;
    })
    .filter(Boolean)
    .join("\n\n");
}

const STORAGE_KEY = "visualtex.anywhere.output-modes";
const LAST_KEY = "*";

function isOutputMode(value: unknown): value is AnywhereOutputMode {
  return typeof value === "string" && (OUTPUT_MODES as readonly string[]).includes(value);
}

function readModes(): Record<string, AnywhereOutputMode> {
  try {
    const parsed: unknown = JSON.parse(window.localStorage.getItem(STORAGE_KEY) ?? "{}");
    if (!parsed || typeof parsed !== "object") return {};
    return Object.fromEntries(Object.entries(parsed).filter(([, mode]) => isOutputMode(mode))) as Record<
      string,
      AnywhereOutputMode
    >;
  } catch {
    return {};
  }
}

/** The mode last used in `app`, else the last one used anywhere, else Source. */
export function rememberedOutputMode(app: string): AnywhereOutputMode {
  const modes = readModes();
  return modes[app] ?? modes[LAST_KEY] ?? "source";
}

export function rememberOutputMode(app: string, mode: AnywhereOutputMode) {
  try {
    const modes = readModes();
    if (app) modes[app] = mode;
    modes[LAST_KEY] = mode;
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(modes));
  } catch {
    // The mode still applies to this formula.
  }
}

export function nextOutputMode(mode: AnywhereOutputMode): AnywhereOutputMode {
  return OUTPUT_MODES[(OUTPUT_MODES.indexOf(mode) + 1) % OUTPUT_MODES.length];
}

/** Loads what a mode needs before Return (the export path for Unicode and
 *  Image takes a moment the first time). */
export function warmOutputMode(mode: AnywhereOutputMode) {
  if (mode === "unicode") void import("../../../shared/visualtex-app/src/export/runtime");
  if (mode === "image") void import("../../../shared/visualtex-app/src/export/pngClipboard");
}

/** The formula (its rows' LaTeX) as Unicode text. */
export async function formulaUnicode(rows: readonly string[]): Promise<UnicodeMath> {
  const [{ latexToMathMlSync }, { mathMlToUnicode }] = await Promise.all([
    import("../../../shared/visualtex-app/src/export/runtime"),
    import("../../../shared/visualtex-app/src/anywhere/unicodeMath"),
  ]);
  const notes = new Set<string>();
  const texts = rows
    .map((row) => row.trim())
    .filter(Boolean)
    .map((row) => {
      const converted = mathMlToUnicode(latexToMathMlSync(row, false));
      for (const note of converted.notes) notes.add(note);
      return converted.text;
    });
  return { text: texts.join("\n"), notes: [...notes] };
}

/** The field's colours as read off the screen (#rrggbb). */
export interface FieldColors {
  ink: string;
  background: string;
}

/** #rrggbb from #rrggbb or #rrggbbaa. */
export function opaqueHex(color?: string | null) {
  return color && /^#[0-9a-f]{6}([0-9a-f]{2})?$/i.test(color) ? color.slice(0, 7).toLowerCase() : null;
}

/** The formula as a PNG (base64). In a field whose colours are known it
 *  matches the field (its background, ink that reads on it: a dark image in
 *  a dark chat), so it looks the same to whoever sees it in either theme;
 *  otherwise the main window's image settings. */
export async function formulaPng(rows: readonly string[], field?: FieldColors | null): Promise<string> {
  const { renderFormulaDocumentPng } = await import("../../../shared/visualtex-app/src/export/pngClipboard");
  const { pngExportBackground, formulaLetterFont, formulaChineseFont } = useEditorStore.getState();
  const png = await renderFormulaDocumentPng(rows, {
    background: field ? (field.background as `#${string}`) : pngExportBackground,
    color: field?.ink,
    formulaLetterFont,
    formulaChineseFont,
  });
  return png.base64;
}
