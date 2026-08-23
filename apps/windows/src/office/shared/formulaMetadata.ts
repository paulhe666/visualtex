import { deflateSync, inflateSync, strFromU8, strToU8 } from "fflate";
import {
  FORMULA_CHINESE_FONT_OPTIONS,
  FORMULA_LETTER_FONT_OPTIONS,
  type FormulaChineseFont,
  type FormulaLetterFont,
} from "../../editor/formulaFontPreferences";

export interface VisualTeXFormulaMetadata {
  schema: "visualtex-formula";
  schemaVersion: 1;
  formulaId: string;
  title: string;
  latex: string;
  lines: Array<{ id: string; latex: string }>;
  codeFormat: string;
  displayMode: "inline" | "block";
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
  /** Semantic Office math size in points. */
  fontSizePt?: number;
  /** Point size used to create the current cached SVG/EMF/PNG preview. */
  renderFontSizePt?: number;
  /** VisualTeX font preferences used for this formula's rendered/native Office form. */
  formulaLetterFont?: FormulaLetterFont;
  formulaChineseFont?: FormulaChineseFont;
  /** Physical Word inline OLE extent retained across OLE/OMML conversions. */
  wordInlineOleWidthPt?: number;
  wordInlineOleHeightPt?: number;
  /** Fingerprint of the exact Word-native OMML source. */
  nativeOmmlFingerprint?: string;
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
  displayMode?: "inline" | "block";
  numbered?: boolean;
  equationTag?: string;
  renderWidthPx?: number;
  renderHeightPx?: number;
  baseline?: number;
  fontSizePt?: number;
  renderFontSizePt?: number;
  formulaLetterFont?: FormulaLetterFont;
  formulaChineseFont?: FormulaChineseFont;
  wordInlineOleWidthPt?: number;
  wordInlineOleHeightPt?: number;
  appVersion?: string;
  original?: VisualTeXFormulaMetadata | null;
}

export const VISUALTEX_FORMULA_SCHEMA = "visualtex-formula" as const;
export const VISUALTEX_FORMULA_SCHEMA_VERSION = 1 as const;
export const VISUALTEX_FORMULA_XML_NAMESPACE = "urn:visualtex:formula:1";
export const VISUALTEX_METADATA_PREFIX = "visualtex:v1:deflate:";
export const CURRENT_VISUALTEX_VERSION = "1.2.5";

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
    (candidate.fontSizePt === undefined ||
      (typeof candidate.fontSizePt === "number" &&
        Number.isFinite(candidate.fontSizePt) &&
        candidate.fontSizePt >= 5 &&
        candidate.fontSizePt <= 200)) &&
    (candidate.renderFontSizePt === undefined ||
      (typeof candidate.renderFontSizePt === "number" &&
        Number.isFinite(candidate.renderFontSizePt) &&
        candidate.renderFontSizePt >= 5 &&
        candidate.renderFontSizePt <= 200)) &&
    (candidate.formulaLetterFont === undefined ||
      FORMULA_LETTER_FONT_OPTIONS.some(
        (item) => item.id === candidate.formulaLetterFont,
      )) &&
    (candidate.formulaChineseFont === undefined ||
      FORMULA_CHINESE_FONT_OPTIONS.some(
        (item) => item.id === candidate.formulaChineseFont,
      )) &&
    ((candidate.wordInlineOleWidthPt === undefined &&
      candidate.wordInlineOleHeightPt === undefined) ||
      (candidate.displayMode === "inline" &&
        typeof candidate.wordInlineOleWidthPt === "number" &&
        Number.isFinite(candidate.wordInlineOleWidthPt) &&
        candidate.wordInlineOleWidthPt > 0 &&
        typeof candidate.wordInlineOleHeightPt === "number" &&
        Number.isFinite(candidate.wordInlineOleHeightPt) &&
        candidate.wordInlineOleHeightPt > 0)) &&
    (candidate.nativeOmmlFingerprint === undefined ||
      (typeof candidate.nativeOmmlFingerprint === "string" &&
        /^[0-9a-f]{64}$/i.test(candidate.nativeOmmlFingerprint))) &&
    typeof candidate.createdWithVersion === "string" &&
    typeof candidate.updatedWithVersion === "string" &&
    typeof candidate.createdAt === "string" &&
    typeof candidate.updatedAt === "string"
  );
}

export function createFormulaMetadata({
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
    ...(original?.nativeOmmlFingerprint
      ? { nativeOmmlFingerprint: original.nativeOmmlFingerprint }
      : {}),
    createdWithVersion: original?.createdWithVersion ?? appVersion,
    updatedWithVersion: appVersion,
    createdAt: original?.createdAt ?? now,
    updatedAt: now,
  };
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
