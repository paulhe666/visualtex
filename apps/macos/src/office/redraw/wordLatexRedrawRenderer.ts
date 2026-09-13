import { wordImageReferenceGeometry } from "../shared/wordImageGeometry";
import { createUuid } from "../../runtime/browserCompatibility";
import {
  normalizeFormulaEditorDocument,
  serializeFormulaEditorDocument,
} from "../shared/formulaEditorDocument";
import {
  createFormulaMetadata,
  type VisualTeXFormulaMetadata,
} from "../shared/formulaMetadata";
import { renderOfficeFormulaArtifacts } from "../shared/formulaRenderArtifacts";
import { latexLinesToOmmlArtifacts } from "../omml/latexToOmml";
import type {
  DocumentImportFormulaCommitItem,
} from "../documentImport/documentImportClient";
import type { WordLatexRedrawSpan } from "./wordLatexRedrawParser";
import { reusableImageGeometry } from "./wordLatexRedrawGeometry";
import { useEditorStore } from "../../stores/editorStore";

const WORD_LATEX_REDRAW_RENDER_CONCURRENCY = 4;

export type WordLatexRedrawOutputKind = "omml" | "image";
export type WordLatexRedrawRenderTarget = WordLatexRedrawSpan & {
  fontSizePt: number;
  formulaId?: string;
  numbered?: boolean;
  metadata?: VisualTeXFormulaMetadata;
  sourceKind?: "omml" | "image";
};

type RenderTemplate = {
  canonicalLatex: string;
  ommlBase64: string;
  ommlDocxBase64: string;
  svgBase64?: string;
  pngBase64?: string;
  width?: number;
  height?: number;
  baseline?: number;
  inkCenterYRatio?: number;
  renderWidthPx?: number;
  renderHeightPx?: number;
  referenceWidthPt?: number;
  referenceHeightPt?: number;
  referenceBaselinePt?: number;
};

function decodeUrlSafeBase64Utf8(value: string) {
  const normalized = value.replace(/-/g, "+").replace(/_/g, "/");
  const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
  const binary = atob(padded);
  const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
  return new TextDecoder().decode(bytes);
}

function ommlRetainsLiteralLatexCommand(ommlBase64: string, latex: string) {
  const commands = [...latex.matchAll(/\\([A-Za-z@]+)\b/g)]
    .map((match) => match[1])
    .filter((command, index, values) => values.indexOf(command) === index);
  if (!commands.length) return false;
  const omml = decodeUrlSafeBase64Utf8(ommlBase64);
  return commands.some((command) => omml.includes(`\\${command}`));
}

async function renderTemplate(
  span: WordLatexRedrawRenderTarget,
  outputKind: WordLatexRedrawOutputKind,
): Promise<RenderTemplate> {
  const line = { id: createUuid(), latex: span.latex };
  const editorDocument =
    span.sourceKind === "image" && span.metadata
      ? normalizeFormulaEditorDocument(span.metadata.lines, span.metadata.codeFormat)
      : normalizeFormulaEditorDocument([line], "raw");
  const canonicalLatex = serializeFormulaEditorDocument(editorDocument);

  const editorPreferences = useEditorStore.getState();
  const formulaLetterFont =
    span.metadata?.formulaLetterFont ?? editorPreferences.formulaLetterFont;
  const formulaChineseFont =
    span.metadata?.formulaChineseFont ?? editorPreferences.formulaChineseFont;

  if (outputKind === "omml") {
    const omml = latexLinesToOmmlArtifacts(
      editorDocument.lines.map((formulaLine) => formulaLine.latex),
      span.displayMode,
      editorDocument.codeFormat,
      { formulaLetterFont, formulaChineseFont },
      span.displayMode === "block" && Boolean(span.numbered),
    );
    if (ommlRetainsLiteralLatexCommand(omml.ommlBase64, canonicalLatex)) {
      throw new Error("The redraw formula contains a command unsupported by Word OMML.");
    }
    const cachedGeometry = reusableImageGeometry(span.metadata);
    if (cachedGeometry) {
      return {
        canonicalLatex,
        ommlBase64: omml.ommlBase64,
        ommlDocxBase64: omml.ommlDocxBase64,
        ...cachedGeometry,
      };
    }

    // OMML output does not need SVG or PNG data. Historical image formulas can
    // lack the geometry fields that current metadata carries, though, so render
    // only the inexpensive vector bounds for that compatibility case. The
    // backend must never require image bytes for an OMML-only commit.
    const geometryArtifacts = renderOfficeFormulaArtifacts({
      lines: editorDocument.lines,
      codeFormat: editorDocument.codeFormat,
      displayMode: span.displayMode,
      host: "word",
      fontSizePt: span.fontSizePt,
      numbered: span.displayMode === "block" && Boolean(span.numbered),
      formulaLetterFont,
      formulaChineseFont,
    });
    const geometry = geometryArtifacts.svg;
    return {
      canonicalLatex,
      ommlBase64: omml.ommlBase64,
      ommlDocxBase64: omml.ommlDocxBase64,
      width: geometry.width,
      height: geometry.height,
      baseline: geometry.baseline,
      renderWidthPx: geometry.width,
      renderHeightPx: geometry.height,
      ...wordImageReferenceGeometry(
        geometry.width,
        geometry.height,
        geometry.baseline,
        formulaLetterFont,
      ),
    };
  }

  const artifacts = renderOfficeFormulaArtifacts({
    lines: editorDocument.lines,
    codeFormat: editorDocument.codeFormat,
    displayMode: span.displayMode,
    host: "word",
      fontSizePt: span.fontSizePt,
    numbered: span.displayMode === "block" && Boolean(span.numbered),
    formulaLetterFont,
    formulaChineseFont,
  });
  if (!artifacts.omml) {
    throw new Error("Unable to generate Word OMML for the LaTeX redraw formula.");
  }
  const { omml, svg } = artifacts;
  if (ommlRetainsLiteralLatexCommand(omml.ommlBase64, canonicalLatex)) {
    throw new Error("The redraw formula contains a command unsupported by Word OMML.");
  }

  const { svgToPng } = await import("../../export/svgToPng");
  const png = await svgToPng(svg, { scale: 2, background: "transparent" });
  const pngBase64 = png.base64;
  const baseline = svg.baseline;
  return {
    canonicalLatex,
    ommlBase64: omml.ommlBase64,
    ommlDocxBase64: omml.ommlDocxBase64,
    svgBase64: svg.base64,
    pngBase64,
    width: svg.width,
    height: svg.height,
    baseline,
    inkCenterYRatio: png.inkCenterYRatio,
    renderWidthPx: svg.width,
    renderHeightPx: svg.height,
    ...wordImageReferenceGeometry(svg.width, svg.height, baseline, formulaLetterFont),
  };
}

function createMetadata(
  formulaId: string,
  span: WordLatexRedrawRenderTarget,
  template: RenderTemplate,
): VisualTeXFormulaMetadata {
  const original = span.metadata ?? null;
  const line = {
    id: original?.lines[0]?.id ?? createUuid(),
    latex: template.canonicalLatex,
  };
  const lines =
    span.sourceKind === "image" && original
      ? original.lines
      : [line];
  const editorPreferences = useEditorStore.getState();
  const formulaLetterFont =
    original?.formulaLetterFont ?? editorPreferences.formulaLetterFont;
  const formulaChineseFont =
    original?.formulaChineseFont ?? editorPreferences.formulaChineseFont;
  return createFormulaMetadata({
    formulaId,
    title:
      original?.title ??
      (span.displayMode === "inline"
        ? "Redrawn inline Word formula"
        : "Redrawn display Word formula"),
    lines,
    codeFormat:
      span.sourceKind === "image" && original
        ? original.codeFormat
        : "raw",
    sourceLatex: template.canonicalLatex,
    displayMode: span.displayMode,
    numbered: span.numbered ?? false,
    fontSizePt: span.fontSizePt,
    formulaLetterFont,
    formulaChineseFont,
    renderWidthPx: template.renderWidthPx,
    renderHeightPx: template.renderHeightPx,
    referenceWidthPt: template.referenceWidthPt,
    referenceHeightPt: template.referenceHeightPt,
    referenceBaselinePt: template.referenceBaselinePt,
    imageInkCenterYRatio: template.inkCenterYRatio,
    original,
  });
}

/**
 * Mirrors the Windows redraw renderer: scan targets are rendered directly,
 * templates are cached by output/display/font/source, and each Word target gets
 * an independent formula identity. No document-import blocks or preview UI are
 * constructed on this path.
 */
export async function prepareWindowsStyleWordLatexRedrawItems(
  spans: WordLatexRedrawRenderTarget[],
  outputKind: WordLatexRedrawOutputKind,
  onProgress?: (current: number, total: number) => void,
): Promise<DocumentImportFormulaCommitItem[]> {
  const templates = new Map<string, RenderTemplate>();
  const spanKeys = spans.map((span) =>
    [
      outputKind,
      span.displayMode,
      String(span.fontSizePt),
      span.metadata?.formulaLetterFont ?? "",
      span.metadata?.formulaChineseFont ?? "",
      span.sourceKind ?? "",
      span.metadata?.codeFormat ?? "",
      span.latex,
    ].join("\x1f"),
  );
  const uniqueTargets = new Map<string, WordLatexRedrawRenderTarget>();
  spans.forEach((span, index) => {
    if (!uniqueTargets.has(spanKeys[index])) uniqueTargets.set(spanKeys[index], span);
  });
  const pendingTemplates = [...uniqueTargets.entries()];
  let nextTemplateIndex = 0;
  const workerCount = Math.min(
    WORD_LATEX_REDRAW_RENDER_CONCURRENCY,
    pendingTemplates.length,
  );
  await Promise.all(
    Array.from({ length: workerCount }, async () => {
      while (true) {
        const templateIndex = nextTemplateIndex;
        nextTemplateIndex += 1;
        if (templateIndex >= pendingTemplates.length) return;
        const [key, span] = pendingTemplates[templateIndex];
        templates.set(key, await renderTemplate(span, outputKind));
      }
    }),
  );

  return spans.map((span, index) => {
    const template = templates.get(spanKeys[index]);
    if (!template) throw new Error("A cached Word redraw render is missing.");
    const formulaId = span.formulaId ?? createUuid();
    const numbered = span.numbered ?? false;
    const metadata = createMetadata(formulaId, span, template);
    const item: DocumentImportFormulaCommitItem = {
      kind: "formula",
      formulaId,
      latex: template.canonicalLatex,
      displayMode: span.displayMode,
      numbered,
      fontSizePt: span.fontSizePt,
      metadata,
      ommlBase64: template.ommlBase64,
      ommlDocxBase64: template.ommlDocxBase64,
      svgBase64: template.svgBase64,
      pngBase64: template.pngBase64,
      width: template.width,
      height: template.height,
      baseline: template.baseline,
      sourceStart: span.start,
      sourceEnd: span.end,
      sourceText: span.sourceText,
    };
    onProgress?.(index + 1, spans.length);
    return item;
  });
}
