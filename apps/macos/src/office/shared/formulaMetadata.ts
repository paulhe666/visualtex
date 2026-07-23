import { deflateSync, inflateSync, strFromU8, strToU8 } from "fflate";

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
  /** Natural MathJax export bounds used to preserve PowerPoint's visual scale
   * when a formula is replaced with a longer or taller expression. */
  renderWidthPx?: number;
  renderHeightPx?: number;
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
  renderWidthPx?: number;
  renderHeightPx?: number;
  appVersion?: string;
  original?: VisualTeXFormulaMetadata | null;
}

export const VISUALTEX_FORMULA_SCHEMA = "visualtex-formula" as const;
export const VISUALTEX_FORMULA_SCHEMA_VERSION = 1 as const;
export const VISUALTEX_FORMULA_XML_NAMESPACE = "urn:visualtex:formula:1";
export const VISUALTEX_METADATA_PREFIX = "visualtex:v1:deflate:";
export const CURRENT_VISUALTEX_VERSION = "1.2.2";

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
    (candidate.renderWidthPx === undefined ||
      (typeof candidate.renderWidthPx === "number" &&
        Number.isFinite(candidate.renderWidthPx) &&
        candidate.renderWidthPx > 0)) &&
    (candidate.renderHeightPx === undefined ||
      (typeof candidate.renderHeightPx === "number" &&
        Number.isFinite(candidate.renderHeightPx) &&
        candidate.renderHeightPx > 0)) &&
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
  renderWidthPx,
  renderHeightPx,
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
    ...(resolvedRenderWidth ? { renderWidthPx: resolvedRenderWidth } : {}),
    ...(resolvedRenderHeight ? { renderHeightPx: resolvedRenderHeight } : {}),
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
