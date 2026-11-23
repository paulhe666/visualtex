import { deflateSync, inflateSync, strFromU8, strToU8 } from "fflate";
import { APP_PLATFORM } from "@platform/appInfo";
import {
  FORMULA_CHINESE_FONT_OPTIONS,
  FORMULA_FACE_ROLES,
  LEGACY_FORMULA_LETTER_FONTS,
  normalizeFormulaChineseFont,
  normalizeFormulaFaces,
  normalizeFormulaLetterFont,
  readPersistedFormulaFaces,
  type FormulaChineseFont,
  type FormulaFaces,
  type LegacyFormulaLetterFont,
} from "../../editor/formulaFontPreferences";

// The formula metadata both apps store with an Office formula. Either app reads what both write; each still
// writes its own fields (the Windows add-ins: OLE extent, baseline, render size, OMML fingerprint; the macOS
// add-ins: reference size, inline image math style, ink centre), as createFormulaMetadata shows.

export type InlineImageMathStyle = "text" | "display";

export interface VisualTeXFormulaMetadata {
  schema: "visualtex-formula";
  schemaVersion: 1;
  formulaId: string;
  title: string;
  latex: string;
  lines: Array<{ id: string; latex: string }>;
  codeFormat: string;
  displayMode: "inline" | "block";
  /**
   * macOS: mathematical layout used by Word inline image formulas. This is separate
   * from displayMode: the picture can remain inline while MathJax renders it
   * with display-style operators/fractions. Missing means the historical text
   * style so old documents keep their appearance.
   */
  inlineImageMathStyle?: InlineImageMathStyle;
  /** Whether a Word display formula participates in document equation numbering. */
  numbered?: boolean;
  /** A visual equation tag rendered with the formula but excluded from MathLive editing. */
  equationTag?: string;
  /** Natural MathJax export bounds used to preserve PowerPoint's visual scale
   * when a formula is replaced with a longer or taller expression. */
  renderWidthPx?: number;
  renderHeightPx?: number;
  /** Mathematical baseline in natural render pixels. */
  baseline?: number;
  /** Office math size in points (on macOS: the size that scales Word images and PowerPoint SVG formulas). */
  fontSizePt?: number;
  /** Point size used to create the current cached SVG/EMF/PNG preview. */
  renderFontSizePt?: number;
  /** VisualTeX font preferences used for this formula's rendered/native Office form. */
  formulaLetterFont?: LegacyFormulaLetterFont;
  formulaChineseFont?: FormulaChineseFont;
  /**
   * The fonts for letters, Greek letters and digits the formula was drawn in over its math font ("" follows the
   * math font). A formula without the record was drawn without them (formulaFontsDiffer).
   */
  formulaFaces?: FormulaFaces;
  /** Physical Word inline OLE extent retained across OLE/OMML conversions. */
  wordInlineOleWidthPt?: number;
  wordInlineOleHeightPt?: number;
  /** Whole-point Word object-character offset retained for inline formulas. */
  wordInlineOlePositionPt?: number;
  /** Source MathType preview whitespace retained for sub-point baseline repair. */
  wordInlineSourceBottomWhitespacePt?: number;
  /** In-frame vertical preview compensation already applied, in render pixels. */
  wordInlinePreviewShiftPx?: number;
  /** Legacy 64-hex XML digest, or omml2:<64-hex> canonical-content digest. */
  nativeOmmlFingerprint?: string;
  /** macOS: vector/image dimensions at VisualTeX's 14 pt reference size. */
  referenceWidthPt?: number;
  referenceHeightPt?: number;
  /** macOS: Word inline baseline at the 14 pt reference size; omitted for PowerPoint. */
  referenceBaselinePt?: number;
  /** macOS: painted PNG ink center measured from the image top, normalized to [0, 1]. */
  imageInkCenterYRatio?: number;
  createdWithVersion: string;
  updatedWithVersion: string;
  createdAt: string;
  updatedAt: string;
}

export interface CreateFormulaMetadataInput {
  formulaId: string;
  title: string;
  lines: VisualTeXFormulaMetadata["lines"];
  codeFormat: string;
  /** Canonical serialized source for formats whose editable lines omit wrappers. */
  sourceLatex?: string;
  displayMode?: "inline" | "block";
  inlineImageMathStyle?: InlineImageMathStyle;
  numbered?: boolean;
  equationTag?: string;
  renderWidthPx?: number;
  renderHeightPx?: number;
  baseline?: number;
  fontSizePt?: number;
  renderFontSizePt?: number;
  formulaLetterFont?: LegacyFormulaLetterFont;
  formulaChineseFont?: FormulaChineseFont;
  /** macOS: the faces the formula is drawn in (default: the ones the app draws in now). */
  formulaFaces?: FormulaFaces;
  wordInlineOleWidthPt?: number;
  wordInlineOleHeightPt?: number;
  wordInlineOlePositionPt?: number;
  wordInlineSourceBottomWhitespacePt?: number;
  wordInlinePreviewShiftPx?: number;
  referenceWidthPt?: number;
  referenceHeightPt?: number;
  referenceBaselinePt?: number;
  imageInkCenterYRatio?: number;
  appVersion?: string;
  original?: VisualTeXFormulaMetadata | null;
}

export const VISUALTEX_FORMULA_SCHEMA = "visualtex-formula" as const;
export const VISUALTEX_FORMULA_SCHEMA_VERSION = 1 as const;
export const VISUALTEX_FORMULA_XML_NAMESPACE = "urn:visualtex:formula:1";
export const VISUALTEX_METADATA_PREFIX = "visualtex:v1:deflate:";
export const CURRENT_VISUALTEX_VERSION = "1.3.1";

function bytesToBase64Url(bytes: Uint8Array) {
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
  }
  return btoa(binary)
    .replaceAll("+", "-")
    .replaceAll("/", "_")
    .replace(/=+$/g, "");
}

function base64UrlToBytes(value: string) {
  const normalized = value.replaceAll("-", "+").replaceAll("_", "/");
  const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
  const binary = atob(padded);
  return Uint8Array.from(binary, (character) => character.charCodeAt(0));
}

function escapeXmlAttribute(value: string) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function unescapeXmlAttribute(value: string) {
  return value
    .replaceAll("&quot;", '"')
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&amp;", "&");
}

export function validFormulaId(value: string) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
    value,
  );
}

/** A recorded set of faces: each role a font name of at most 128 characters (Windows writes all three). */
function validFormulaFacesRecord(value: unknown) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const faces = value as Record<string, unknown>;
  return FORMULA_FACE_ROLES.every(
    (role) => faces[role] === undefined || faces[role] === null ||
      (typeof faces[role] === "string" && (faces[role] as string).length <= 128),
  );
}

export function isVisualTeXFormulaMetadata(
  value: unknown,
): value is VisualTeXFormulaMetadata {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<VisualTeXFormulaMetadata>;
  return (
    candidate.schema === VISUALTEX_FORMULA_SCHEMA &&
    candidate.schemaVersion === VISUALTEX_FORMULA_SCHEMA_VERSION &&
    typeof candidate.formulaId === "string" &&
    validFormulaId(candidate.formulaId) &&
    typeof candidate.title === "string" &&
    typeof candidate.latex === "string" &&
    Array.isArray(candidate.lines) &&
    candidate.lines.length > 0 &&
    candidate.lines.every(
      (line) =>
        Boolean(line) &&
        typeof line.id === "string" &&
        typeof line.latex === "string",
    ) &&
    typeof candidate.codeFormat === "string" &&
    (candidate.displayMode === "inline" || candidate.displayMode === "block") &&
    (candidate.inlineImageMathStyle === undefined ||
      candidate.inlineImageMathStyle === "text" ||
      candidate.inlineImageMathStyle === "display") &&
    (candidate.numbered === undefined || typeof candidate.numbered === "boolean") &&
    (candidate.equationTag === undefined ||
      (typeof candidate.equationTag === "string" &&
        candidate.equationTag.trim().length > 0 &&
        candidate.equationTag.length <= 256)) &&
    (candidate.renderWidthPx === undefined ||
      (typeof candidate.renderWidthPx === "number" &&
        Number.isFinite(candidate.renderWidthPx) &&
        candidate.renderWidthPx > 0)) &&
    (candidate.renderHeightPx === undefined ||
      (typeof candidate.renderHeightPx === "number" &&
        Number.isFinite(candidate.renderHeightPx) &&
        candidate.renderHeightPx > 0)) &&
    (candidate.baseline === undefined ||
      (typeof candidate.baseline === "number" &&
        Number.isFinite(candidate.baseline) &&
        candidate.baseline >= 0 &&
        (candidate.renderHeightPx === undefined ||
          candidate.baseline <= candidate.renderHeightPx))) &&
    // Windows writes 5-200 pt, macOS any positive size
    (candidate.fontSizePt === undefined ||
      (typeof candidate.fontSizePt === "number" &&
        Number.isFinite(candidate.fontSizePt) &&
        candidate.fontSizePt > 0)) &&
    (candidate.renderFontSizePt === undefined ||
      (typeof candidate.renderFontSizePt === "number" &&
        Number.isFinite(candidate.renderFontSizePt) &&
        candidate.renderFontSizePt >= 5 &&
        candidate.renderFontSizePt <= 200)) &&
    (candidate.formulaLetterFont === undefined ||
      // formulas saved by earlier versions may name fonts no longer offered
      LEGACY_FORMULA_LETTER_FONTS.includes(candidate.formulaLetterFont)) &&
    (candidate.formulaChineseFont === undefined ||
      FORMULA_CHINESE_FONT_OPTIONS.some(
        (item) => item.id === candidate.formulaChineseFont,
      )) &&
    (candidate.formulaFaces === undefined || validFormulaFacesRecord(candidate.formulaFaces)) &&
    ((candidate.wordInlineOleWidthPt === undefined &&
      candidate.wordInlineOleHeightPt === undefined) ||
      (candidate.displayMode === "inline" &&
        typeof candidate.wordInlineOleWidthPt === "number" &&
        Number.isFinite(candidate.wordInlineOleWidthPt) &&
        candidate.wordInlineOleWidthPt > 0 &&
        typeof candidate.wordInlineOleHeightPt === "number" &&
        Number.isFinite(candidate.wordInlineOleHeightPt) &&
        candidate.wordInlineOleHeightPt > 0)) &&
    (candidate.wordInlineOlePositionPt === undefined ||
      (candidate.displayMode === "inline" &&
        typeof candidate.wordInlineOlePositionPt === "number" &&
        Number.isInteger(candidate.wordInlineOlePositionPt) &&
        candidate.wordInlineOlePositionPt >= -256 &&
        candidate.wordInlineOlePositionPt <= 256)) &&
    (candidate.wordInlineSourceBottomWhitespacePt === undefined ||
      (candidate.displayMode === "inline" &&
        typeof candidate.wordInlineSourceBottomWhitespacePt === "number" &&
        Number.isFinite(candidate.wordInlineSourceBottomWhitespacePt) &&
        candidate.wordInlineSourceBottomWhitespacePt >= 0 &&
        candidate.wordInlineSourceBottomWhitespacePt <= 256)) &&
    (candidate.wordInlinePreviewShiftPx === undefined ||
      (candidate.displayMode === "inline" &&
        typeof candidate.wordInlinePreviewShiftPx === "number" &&
        Number.isFinite(candidate.wordInlinePreviewShiftPx) &&
        candidate.wordInlinePreviewShiftPx >= -2 &&
        candidate.wordInlinePreviewShiftPx <= 2)) &&
    (candidate.nativeOmmlFingerprint === undefined ||
      (typeof candidate.nativeOmmlFingerprint === "string" &&
        /^(?:omml2:)?[0-9a-f]{64}$/i.test(candidate.nativeOmmlFingerprint))) &&
    (candidate.referenceWidthPt === undefined ||
      (typeof candidate.referenceWidthPt === "number" &&
        Number.isFinite(candidate.referenceWidthPt) &&
        candidate.referenceWidthPt > 0)) &&
    (candidate.referenceHeightPt === undefined ||
      (typeof candidate.referenceHeightPt === "number" &&
        Number.isFinite(candidate.referenceHeightPt) &&
        candidate.referenceHeightPt > 0)) &&
    (candidate.referenceBaselinePt === undefined ||
      (typeof candidate.referenceBaselinePt === "number" &&
        Number.isFinite(candidate.referenceBaselinePt) &&
        candidate.referenceBaselinePt <= 0 &&
        candidate.referenceBaselinePt >= -256)) &&
    (candidate.imageInkCenterYRatio === undefined ||
      (typeof candidate.imageInkCenterYRatio === "number" &&
        Number.isFinite(candidate.imageInkCenterYRatio) &&
        candidate.imageInkCenterYRatio >= 0 &&
        candidate.imageInkCenterYRatio <= 1)) &&
    typeof candidate.createdWithVersion === "string" &&
    typeof candidate.updatedWithVersion === "string" &&
    typeof candidate.createdAt === "string" &&
    typeof candidate.updatedAt === "string"
  );
}

function createWindowsFormulaMetadata({
  formulaId,
  title,
  lines,
  codeFormat,
  displayMode = "block",
  numbered = false,
  equationTag,
  renderWidthPx,
  renderHeightPx,
  baseline,
  fontSizePt,
  renderFontSizePt,
  formulaLetterFont,
  formulaChineseFont,
  wordInlineOleWidthPt,
  wordInlineOleHeightPt,
  wordInlineOlePositionPt,
  wordInlineSourceBottomWhitespacePt,
  wordInlinePreviewShiftPx,
  appVersion = CURRENT_VISUALTEX_VERSION,
  original = null,
}: CreateFormulaMetadataInput): VisualTeXFormulaMetadata {
  if (!validFormulaId(formulaId)) {
    throw new Error("VisualTeX formulaId must be a UUID v4.");
  }
  if (!lines.length) {
    throw new Error("VisualTeX formula metadata requires at least one line.");
  }
  const now = new Date().toISOString();
  const resolvedRenderWidth =
    renderWidthPx && Number.isFinite(renderWidthPx) && renderWidthPx > 0
      ? renderWidthPx
      : original?.renderWidthPx;
  const resolvedRenderHeight =
    renderHeightPx && Number.isFinite(renderHeightPx) && renderHeightPx > 0
      ? renderHeightPx
      : original?.renderHeightPx;
  const resolvedBaseline =
    baseline !== undefined &&
    Number.isFinite(baseline) &&
    baseline >= 0 &&
    (resolvedRenderHeight === undefined || baseline <= resolvedRenderHeight)
      ? baseline
      : original?.baseline;
  const normalizeFontSize = (value: number | undefined, fallback: number) => {
    const resolved = Number.isFinite(value) ? (value as number) : fallback;
    return Math.round(Math.min(200, Math.max(5, resolved)) * 100) / 100;
  };
  const resolvedFontSize = normalizeFontSize(
    fontSizePt,
    original?.fontSizePt ?? original?.renderFontSizePt ?? 14,
  );
  const resolvedRenderFontSize = normalizeFontSize(
    renderFontSizePt,
    original?.renderFontSizePt ?? resolvedFontSize,
  );
  const requestedInlineWidth = wordInlineOleWidthPt ?? original?.wordInlineOleWidthPt;
  const requestedInlineHeight = wordInlineOleHeightPt ?? original?.wordInlineOleHeightPt;
  const resolvedInlineOleSize =
    displayMode === "inline" &&
    Number.isFinite(requestedInlineWidth) &&
    (requestedInlineWidth ?? 0) > 0 &&
    Number.isFinite(requestedInlineHeight) &&
    (requestedInlineHeight ?? 0) > 0
      ? {
          wordInlineOleWidthPt: requestedInlineWidth as number,
          wordInlineOleHeightPt: requestedInlineHeight as number,
        }
      : {};
  return {
    schema: VISUALTEX_FORMULA_SCHEMA,
    schemaVersion: VISUALTEX_FORMULA_SCHEMA_VERSION,
    formulaId,
    title,
    latex: lines.map((line) => line.latex).join("\n"),
    lines: lines.map((line) => ({ ...line })),
    codeFormat,
    displayMode,
    numbered,
    ...(displayMode === "block" &&
    (equationTag?.trim() || original?.equationTag?.trim())
      ? { equationTag: equationTag?.trim() || original?.equationTag?.trim() }
      : {}),
    ...(resolvedRenderWidth ? { renderWidthPx: resolvedRenderWidth } : {}),
    ...(resolvedRenderHeight ? { renderHeightPx: resolvedRenderHeight } : {}),
    ...(resolvedBaseline !== undefined ? { baseline: resolvedBaseline } : {}),
    fontSizePt: resolvedFontSize,
    renderFontSizePt: resolvedRenderFontSize,
    ...(formulaLetterFont ?? original?.formulaLetterFont
      ? { formulaLetterFont: formulaLetterFont ?? original?.formulaLetterFont }
      : {}),
    ...(formulaChineseFont ?? original?.formulaChineseFont
      ? { formulaChineseFont: formulaChineseFont ?? original?.formulaChineseFont }
      : {}),
    ...resolvedInlineOleSize,
    ...(displayMode === "inline" &&
    Number.isInteger(wordInlineOlePositionPt ?? original?.wordInlineOlePositionPt)
      ? {
          wordInlineOlePositionPt:
            (wordInlineOlePositionPt ?? original?.wordInlineOlePositionPt) as number,
        }
      : {}),
    ...(displayMode === "inline" &&
    Number.isFinite(
      wordInlineSourceBottomWhitespacePt ??
        original?.wordInlineSourceBottomWhitespacePt,
    )
      ? {
          wordInlineSourceBottomWhitespacePt:
            (wordInlineSourceBottomWhitespacePt ??
              original?.wordInlineSourceBottomWhitespacePt) as number,
        }
      : {}),
    ...(displayMode === "inline" &&
    Number.isFinite(wordInlinePreviewShiftPx ?? original?.wordInlinePreviewShiftPx)
      ? {
          wordInlinePreviewShiftPx:
            (wordInlinePreviewShiftPx ?? original?.wordInlinePreviewShiftPx) as number,
        }
      : {}),
    ...(original?.nativeOmmlFingerprint
      ? { nativeOmmlFingerprint: original.nativeOmmlFingerprint }
      : {}),
    createdWithVersion: original?.createdWithVersion ?? appVersion,
    updatedWithVersion: appVersion,
    createdAt: original?.createdAt ?? now,
    updatedAt: now,
  };
}

function createMacFormulaMetadata({
  formulaId,
  title,
  lines,
  codeFormat,
  sourceLatex,
  displayMode = "block",
  inlineImageMathStyle,
  numbered = false,
  renderWidthPx,
  renderHeightPx,
  fontSizePt,
  formulaLetterFont,
  formulaChineseFont,
  formulaFaces,
  referenceWidthPt,
  referenceHeightPt,
  referenceBaselinePt,
  imageInkCenterYRatio,
  appVersion = CURRENT_VISUALTEX_VERSION,
  original = null,
}: CreateFormulaMetadataInput): VisualTeXFormulaMetadata {
  if (!validFormulaId(formulaId)) {
    throw new Error("VisualTeX formulaId must be a UUID v4.");
  }
  if (!lines.length) {
    throw new Error("VisualTeX formula metadata requires at least one line.");
  }
  const now = new Date().toISOString();
  const resolvedRenderWidth =
    renderWidthPx && Number.isFinite(renderWidthPx) && renderWidthPx > 0
      ? renderWidthPx
      : original?.renderWidthPx;
  const resolvedRenderHeight =
    renderHeightPx && Number.isFinite(renderHeightPx) && renderHeightPx > 0
      ? renderHeightPx
      : original?.renderHeightPx;
  const resolvedFontSize =
    fontSizePt && Number.isFinite(fontSizePt) && fontSizePt > 0
      ? fontSizePt
      : original?.fontSizePt;
  const resolvedFormulaLetterFont = normalizeFormulaLetterFont(
    formulaLetterFont ?? original?.formulaLetterFont,
  );
  const resolvedFormulaChineseFont = normalizeFormulaChineseFont(
    formulaChineseFont ?? original?.formulaChineseFont,
  );
  const resolvedReferenceWidth =
    referenceWidthPt && Number.isFinite(referenceWidthPt) && referenceWidthPt > 0
      ? referenceWidthPt
      : original?.referenceWidthPt;
  const resolvedReferenceHeight =
    referenceHeightPt && Number.isFinite(referenceHeightPt) && referenceHeightPt > 0
      ? referenceHeightPt
      : original?.referenceHeightPt;
  const resolvedReferenceBaseline =
    referenceBaselinePt !== undefined &&
    Number.isFinite(referenceBaselinePt) &&
    referenceBaselinePt <= 0 &&
    referenceBaselinePt >= -256
      ? referenceBaselinePt
      : original?.referenceBaselinePt;
  const resolvedImageInkCenterYRatio =
    imageInkCenterYRatio !== undefined &&
    Number.isFinite(imageInkCenterYRatio) &&
    imageInkCenterYRatio >= 0 &&
    imageInkCenterYRatio <= 1
      ? imageInkCenterYRatio
      : original?.imageInkCenterYRatio;
  const resolvedLatex = sourceLatex?.replace(/\r\n?/g, "\n").trim() ||
    lines.map((line) => line.latex).join("\n");
  return {
    schema: VISUALTEX_FORMULA_SCHEMA,
    schemaVersion: VISUALTEX_FORMULA_SCHEMA_VERSION,
    formulaId,
    title,
    latex: resolvedLatex,
    lines: lines.map((line) => ({ ...line })),
    codeFormat,
    displayMode,
    inlineImageMathStyle:
      inlineImageMathStyle ?? original?.inlineImageMathStyle ?? "text",
    numbered,
    ...(resolvedRenderWidth ? { renderWidthPx: resolvedRenderWidth } : {}),
    ...(resolvedRenderHeight ? { renderHeightPx: resolvedRenderHeight } : {}),
    ...(resolvedFontSize ? { fontSizePt: resolvedFontSize } : {}),
    formulaLetterFont: resolvedFormulaLetterFont,
    formulaChineseFont: resolvedFormulaChineseFont,
    // the layout engine draws every formula in the faces of the Settings, so the formula records those
    formulaFaces: normalizeFormulaFaces(formulaFaces ?? readPersistedFormulaFaces()),
    ...(resolvedReferenceWidth ? { referenceWidthPt: resolvedReferenceWidth } : {}),
    ...(resolvedReferenceHeight ? { referenceHeightPt: resolvedReferenceHeight } : {}),
    ...(resolvedReferenceBaseline !== undefined
      ? { referenceBaselinePt: resolvedReferenceBaseline }
      : {}),
    ...(resolvedImageInkCenterYRatio !== undefined
      ? { imageInkCenterYRatio: resolvedImageInkCenterYRatio }
      : {}),
    createdWithVersion: original?.createdWithVersion ?? appVersion,
    updatedWithVersion: appVersion,
    createdAt: original?.createdAt ?? now,
    updatedAt: now,
  };
}

/** New metadata for a formula, as this app's Office add-ins expect it. */
export function createFormulaMetadata(input: CreateFormulaMetadataInput): VisualTeXFormulaMetadata {
  return APP_PLATFORM === "macos" ? createMacFormulaMetadata(input) : createWindowsFormulaMetadata(input);
}

export function encodeFormulaMetadata(metadata: VisualTeXFormulaMetadata) {
  if (!isVisualTeXFormulaMetadata(metadata)) {
    throw new Error("Cannot encode invalid VisualTeX formula metadata.");
  }
  const compactJson = JSON.stringify(metadata);
  const compressed = deflateSync(strToU8(compactJson), { level: 9 });
  return `${VISUALTEX_METADATA_PREFIX}${bytesToBase64Url(compressed)}`;
}

export function decodeFormulaMetadata(value: string) {
  if (!value.startsWith(VISUALTEX_METADATA_PREFIX)) return null;
  try {
    const compressed = base64UrlToBytes(
      value.slice(VISUALTEX_METADATA_PREFIX.length),
    );
    const parsed: unknown = JSON.parse(strFromU8(inflateSync(compressed)));
    return isVisualTeXFormulaMetadata(parsed) ? parsed : null;
  } catch {
    return null;
  }
}

export function formulaMetadataToXml(metadata: VisualTeXFormulaMetadata) {
  const encoded = encodeFormulaMetadata(metadata);
  return `<?xml version="1.0" encoding="UTF-8"?><visualtexFormula xmlns="${VISUALTEX_FORMULA_XML_NAMESPACE}" formulaId="${escapeXmlAttribute(metadata.formulaId)}"><metadata encoding="deflate-base64url">${encoded}</metadata></visualtexFormula>`;
}

export function formulaMetadataFromXml(xml: string) {
  const root = xml.match(
    /<visualtexFormula\b[^>]*\bformulaId="([^"]+)"[^>]*>/i,
  );
  const payload = xml.match(
    /<metadata\b[^>]*\bencoding="deflate-base64url"[^>]*>([^<]+)<\/metadata>/i,
  );
  if (!root || !payload) return null;
  const metadata = decodeFormulaMetadata(payload[1].trim());
  if (!metadata) return null;
  return metadata.formulaId === unescapeXmlAttribute(root[1]) ? metadata : null;
}
