const POWERPOINT_SHAPE_PREFIX = "VisualTeX_";
const POWERPOINT_OBJECT_PREFIX = "visualtex-ppt:v1:";

export interface NativePowerPointObjectReference {
  slideIndex: number;
  shapeName: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface PowerPointObjectReference {
  slideId: string;
  shapeId: string;
  native?: NativePowerPointObjectReference;
}

function bytesToBase64Url(bytes: Uint8Array) {
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return btoa(binary)
    .replaceAll("+", "-")
    .replaceAll("/", "_")
    .replace(/=+$/g, "");
}

function base64UrlToBytes(value: string) {
  const normalized = value.replaceAll("-", "+").replaceAll("_", "/");
  const binary = atob(
    normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "="),
  );
  return Uint8Array.from(binary, (character) => character.charCodeAt(0));
}

export function powerpointShapeName(formulaId: string) {
  return `${POWERPOINT_SHAPE_PREFIX}${formulaId}`;
}

export function formulaIdFromPowerPointShapeName(name: string) {
  if (!name.startsWith(POWERPOINT_SHAPE_PREFIX)) return null;
  const formulaId = name.slice(POWERPOINT_SHAPE_PREFIX.length);
  return /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
    formulaId,
  )
    ? formulaId
    : null;
}

export function encodePowerPointObjectReference(
  reference: PowerPointObjectReference,
) {
  if (!reference.slideId || !reference.shapeId) {
    throw new Error("PowerPoint object reference requires slideId and shapeId.");
  }
  if (
    reference.native &&
    (!Number.isInteger(reference.native.slideIndex) ||
      reference.native.slideIndex <= 0 ||
      !reference.native.shapeName)
  ) {
    throw new Error("PowerPoint native object reference is invalid.");
  }
  return `${POWERPOINT_OBJECT_PREFIX}${bytesToBase64Url(
    new TextEncoder().encode(JSON.stringify(reference)),
  )}`;
}

export function decodePowerPointObjectReference(value: string | null) {
  if (!value?.startsWith(POWERPOINT_OBJECT_PREFIX)) return null;
  try {
    const parsed = JSON.parse(
      new TextDecoder().decode(
        base64UrlToBytes(value.slice(POWERPOINT_OBJECT_PREFIX.length)),
      ),
    ) as Partial<PowerPointObjectReference>;
    if (
      typeof parsed.slideId !== "string" ||
      parsed.slideId.length === 0 ||
      typeof parsed.shapeId !== "string" ||
      parsed.shapeId.length === 0
    ) {
      return null;
    }
    const native = parsed.native;
    if (native !== undefined) {
      if (
        typeof native !== "object" ||
        native === null ||
        !Number.isInteger(native.slideIndex) ||
        native.slideIndex <= 0 ||
        typeof native.shapeName !== "string" ||
        native.shapeName.length === 0 ||
        ![native.left, native.top, native.width, native.height].every(
          (value) => typeof value === "number" && Number.isFinite(value),
        )
      ) {
        return null;
      }
      return {
        slideId: parsed.slideId,
        shapeId: parsed.shapeId,
        native: {
          slideIndex: native.slideIndex,
          shapeName: native.shapeName,
          left: native.left,
          top: native.top,
          width: native.width,
          height: native.height,
        },
      };
    }
    return { slideId: parsed.slideId, shapeId: parsed.shapeId };
  } catch {
    return null;
  }
}
