import type { NativeSystemMathFontProbe } from "./systemMathGlyphLibrary";

interface NativeSystemMathGlyphOutlinePayload {
  character: string;
  requestedFamily: string;
  resolvedFamily: string;
  fallbackUsed: boolean;
  glyphId: number;
  path: string;
  metrics: {
    widthEm: number;
    ascentEm: number;
    descentEm: number;
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function requiredString(
  source: Record<string, unknown>,
  key: string,
  label: string,
  allowEmpty = false,
) {
  const value = source[key];
  if (typeof value !== "string" || (!allowEmpty && !value.trim())) {
    throw new Error(`VisualTeX ${label}.${key} is invalid.`);
  }
  return value;
}

function requiredBoolean(
  source: Record<string, unknown>,
  key: string,
  label: string,
) {
  const value = source[key];
  if (typeof value !== "boolean") {
    throw new Error(`VisualTeX ${label}.${key} is invalid.`);
  }
  return value;
}

function requiredFiniteNumber(
  source: Record<string, unknown>,
  key: string,
  label: string,
  minimum = Number.NEGATIVE_INFINITY,
) {
  const value = source[key];
  if (typeof value !== "number" || !Number.isFinite(value) || value < minimum) {
    throw new Error(`VisualTeX ${label}.${key} is invalid.`);
  }
  return value;
}

export function decodeNativeSystemMathFontProbes(
  value: unknown,
): NativeSystemMathFontProbe[] {
  if (!Array.isArray(value)) {
    throw new Error("VisualTeX system math font probe returned invalid data.");
  }
  return value.map((entry, index) => {
    if (!isRecord(entry)) {
      throw new Error(`VisualTeX system math font probe[${index}] is invalid.`);
    }
    return {
      requestedFamily: requiredString(entry, "requestedFamily", `system math font probe[${index}]`),
      resolvedFamily: requiredString(entry, "resolvedFamily", `system math font probe[${index}]`),
      available: requiredBoolean(entry, "available", `system math font probe[${index}]`),
    };
  });
}

export function decodeNativeSystemMathGlyphOutline(
  value: unknown,
): NativeSystemMathGlyphOutlinePayload {
  if (!isRecord(value)) {
    throw new Error("VisualTeX system math glyph outline returned invalid data.");
  }
  const metrics = value.metrics;
  if (!isRecord(metrics)) {
    throw new Error("VisualTeX system math glyph outline.metrics is invalid.");
  }
  const glyphId = requiredFiniteNumber(value, "glyphId", "system math glyph outline", 0);
  if (!Number.isInteger(glyphId)) {
    throw new Error("VisualTeX system math glyph outline.glyphId is invalid.");
  }
  return {
    character: requiredString(value, "character", "system math glyph outline"),
    requestedFamily: requiredString(value, "requestedFamily", "system math glyph outline"),
    resolvedFamily: requiredString(value, "resolvedFamily", "system math glyph outline"),
    fallbackUsed: requiredBoolean(value, "fallbackUsed", "system math glyph outline"),
    glyphId,
    path: requiredString(value, "path", "system math glyph outline"),
    metrics: {
      widthEm: requiredFiniteNumber(metrics, "widthEm", "system math glyph outline.metrics", 0),
      ascentEm: requiredFiniteNumber(metrics, "ascentEm", "system math glyph outline.metrics", 0),
      descentEm: requiredFiniteNumber(metrics, "descentEm", "system math glyph outline.metrics", 0),
    },
  };
}
