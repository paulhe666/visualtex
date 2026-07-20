import {
  getCachedFormulaMetadata,
  getNativePowerPointSelection,
  getNativePowerPointSlideSnapshot,
  markLastNativePowerPointFormula,
  markNativePowerPointSelection,
  putCachedFormulaMetadata,
  replaceLastNativePowerPointFormula,
  revealDesktopApp,
} from "../api/companionClient";
import type {
  NativePowerPointCommitSelection,
  OfficeFormulaSession,
  OfficeSessionMode,
} from "../api/sessionClient";
import { officeErrorMessage } from "../errors";
import {
  createFormulaMetadata,
  decodeFormulaMetadata,
  encodeFormulaMetadata,
  type VisualTeXFormulaMetadata,
} from "../metadata/formulaMetadata";
import {
  decodePowerPointObjectReference,
  encodePowerPointObjectReference,
  formulaIdFromPowerPointShapeName,
  powerpointShapeName,
  type PowerPointObjectReference,
} from "../metadata/powerpointMetadata";
import type {
  OfficeHostAdapter,
  OfficeInteractionTarget,
  OfficeSelectionContext,
} from "./OfficeHostAdapter";

const USE_NATIVE_POWERPOINT_EDIT_TARGET =
  typeof navigator !== "undefined" &&
  /Macintosh|Mac OS X/i.test(navigator.userAgent);

const POWERPOINT_ALT_TEXT_TITLE = "VisualTeX Formula";
const POWERPOINT_FORMULA_ID_TAG = "VISUALTEX_FORMULA_ID";
const POWERPOINT_METADATA_COUNT_TAG = "VISUALTEX_META_COUNT";
const POWERPOINT_METADATA_CHUNK_PREFIX = "VISUALTEX_META_";
const POWERPOINT_METADATA_CHUNK_SIZE = 200;
const MAX_POWERPOINT_WIDTH_PT = 600;
const MAX_POWERPOINT_HEIGHT_PT = 400;
const MIN_POWERPOINT_SIZE_PT = 12;
const POWERPOINT_ASYNC_TIMEOUT_MS = 20_000;

interface PowerPointShapeSnapshot extends PowerPointObjectReference {
  formulaId: string;
  encodedMetadata: string | null;
  nativePresentationIdentity?: string;
  nativeSlideId?: number;
  left: number;
  top: number;
  width: number;
  height: number;
  rotation?: number;
  zOrderPosition?: number;
}

interface PowerPointGeometry {
  left?: number;
  top?: number;
  width: number;
  height: number;
  rotation?: number;
  zOrderPosition?: number;
}

interface PowerPointSlideSnapshot {
  slideId: string;
  shapeIds: Set<string>;
}

interface PowerPointSelectedShapeCore extends PowerPointObjectReference {
  name: string;
  left: number;
  top: number;
  width: number;
  height: number;
  altTextTitle?: string;
  altTextDescription?: string;
  rotation?: number;
  zOrderPosition?: number;
}

function supportsRequirement(name: string, version: string) {
  try {
    return Office.context.requirements.isSetSupported(name, version);
  } catch {
    return false;
  }
}

function supportsPowerPointApi(version: string) {
  return supportsRequirement("PowerPointApi", version);
}

function supportsImageCoercion(version: string) {
  return supportsRequirement("ImageCoercion", version);
}

function requirePowerPointEditingApi() {
  if (!supportsPowerPointApi("1.5")) {
    throw new Error(
      "当前 PowerPoint 版本可以新建 VisualTeX 公式，但不支持重新编辑所选公式。重新编辑需要 PowerPointApi 1.5 或更高版本。",
    );
  }
}

function encodeMetadataTags(metadata: string) {
  const chunks: string[] = [];
  for (
    let offset = 0;
    offset < metadata.length;
    offset += POWERPOINT_METADATA_CHUNK_SIZE
  ) {
    chunks.push(metadata.slice(offset, offset + POWERPOINT_METADATA_CHUNK_SIZE));
  }
  return chunks;
}

function decodeMetadataTags(tags: Map<string, string>) {
  const count = Number(tags.get(POWERPOINT_METADATA_COUNT_TAG) ?? "0");
  if (!Number.isInteger(count) || count <= 0 || count > 4096) return null;
  const chunks: string[] = [];
  for (let index = 0; index < count; index += 1) {
    const value = tags.get(`${POWERPOINT_METADATA_CHUNK_PREFIX}${index}`);
    if (value === undefined) return null;
    chunks.push(value);
  }
  return chunks.join("");
}

function officeAsync<T>(
  invoke: (callback: (result: Office.AsyncResult<T>) => void) => void,
): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    let settled = false;
    const finish = (callback: () => void) => {
      if (settled) return;
      settled = true;
      globalThis.clearTimeout(timeout);
      callback();
    };
    const timeout = window.setTimeout(() => {
      finish(() =>
        reject(
          new Error(
            "PowerPoint 写入公式图片超时。请确认演示文稿仍处于可编辑状态后重试。",
          ),
        ),
      );
    }, POWERPOINT_ASYNC_TIMEOUT_MS);

    try {
      invoke((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          finish(() => resolve(result.value));
        } else {
          finish(() =>
            reject(
              new Error(
                officeErrorMessage(
                  result.error,
                  "PowerPoint 无法把公式图片写入当前幻灯片。",
                ),
              ),
            ),
          );
        }
      });
    } catch (error) {
      finish(() =>
        reject(
          new Error(
            officeErrorMessage(
              error,
              "PowerPoint 无法启动公式图片写入操作。",
            ),
          ),
        ),
      );
    }
  });
}

function calculatePowerPointGeometry(
  widthPx: number,
  heightPx: number,
): PowerPointGeometry {
  const naturalWidth = widthPx * 0.75;
  const naturalHeight = heightPx * 0.75;
  const maximumScale = Math.min(
    MAX_POWERPOINT_WIDTH_PT / naturalWidth,
    MAX_POWERPOINT_HEIGHT_PT / naturalHeight,
  );
  const preferredScale = Math.min(
    1,
    maximumScale,
  );
  // Apply minimum bounds with the same scale on both axes. Independent
  // width/height clamps distort the SVG aspect ratio, making longer edits look
  // progressively squeezed inside the previous formula box.
  const minimumScale = Math.max(
    MIN_POWERPOINT_SIZE_PT / naturalWidth,
    MIN_POWERPOINT_SIZE_PT / naturalHeight,
  );
  const scale = Math.min(maximumScale, Math.max(preferredScale, minimumScale));
  return {
    width: naturalWidth * scale,
    height: naturalHeight * scale,
  };
}

function imageInsertionOptions(
  coercionType: Office.CoercionType,
  geometry: PowerPointGeometry,
) {
  const options: Office.SetSelectedDataOptions = {
    coercionType,
    imageWidth: geometry.width,
    imageHeight: geometry.height,
  };
  if (
    Number.isFinite(geometry.left) &&
    Number.isFinite(geometry.top) &&
    geometry.left !== undefined &&
    geometry.top !== undefined
  ) {
    options.imageLeft = geometry.left;
    options.imageTop = geometry.top;
  }
  return options;
}

async function setSelectedPowerPointImage(
  data: string,
  coercionType: Office.CoercionType,
  geometry: PowerPointGeometry,
) {
  await officeAsync<void>((callback) =>
    Office.context.document.setSelectedDataAsync(
      data,
      imageInsertionOptions(coercionType, geometry),
      callback,
    ),
  );
}

async function insertImageWithFallback(
  session: OfficeFormulaSession,
  geometry: PowerPointGeometry,
) {
  const exportResult = session.exportResult;
  if (!exportResult) {
    throw new Error("VisualTeX Session does not contain an exported formula image.");
  }

  const failures: string[] = [];
  if (exportResult.pngBase64) {
    try {
      await setSelectedPowerPointImage(
        exportResult.pngBase64,
        Office.CoercionType.Image,
        geometry,
      );
      return "png" as const;
    } catch (error) {
      failures.push(
        `PNG: ${officeErrorMessage(error, "PowerPoint 拒绝了 PNG 图片")}`,
      );
    }
  }

  if (
    exportResult.svg &&
    exportResult.svg.trimStart().startsWith("<svg") &&
    supportsImageCoercion("1.2")
  ) {
    try {
      await setSelectedPowerPointImage(
        exportResult.svg,
        Office.CoercionType.XmlSvg,
        geometry,
      );
      return "svg" as const;
    } catch (error) {
      failures.push(
        `SVG: ${officeErrorMessage(error, "PowerPoint 拒绝了 SVG 图片")}`,
      );
    }
  }

  throw new Error(
    failures.length
      ? `PowerPoint 无法插入公式图片。${failures.join("；")}`
      : "VisualTeX 没有可供 PowerPoint 插入的公式图片。",
  );
}

async function getPresentationId() {
  if (!supportsPowerPointApi("1.5")) {
    return Office.context.document.url || null;
  }
  return PowerPoint.run(async (context) => {
    const presentation = context.presentation;
    presentation.load("id");
    await context.sync();
    return presentation.id || Office.context.document.url || null;
  });
}

async function captureSlideSnapshot(preferredSlideId?: string) {
  requirePowerPointEditingApi();
  return PowerPoint.run(async (context) => {
    let slide: PowerPoint.Slide;
    if (preferredSlideId) {
      slide = context.presentation.slides.getItem(preferredSlideId);
      slide.load("id");
    } else {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load("items/id");
      await context.sync();
      if (selectedSlides.items.length < 1) {
        throw new Error("PowerPoint 没有返回当前幻灯片。");
      }
      slide = selectedSlides.items[0];
    }

    slide.shapes.load("items/id");
    await context.sync();
    return {
      slideId: slide.id,
      shapeIds: new Set(slide.shapes.items.map((shape) => shape.id)),
    } satisfies PowerPointSlideSnapshot;
  });
}

async function locateInsertedShape(before: PowerPointSlideSnapshot) {
  requirePowerPointEditingApi();
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItem(before.slideId);
    const selected = context.presentation.getSelectedShapes();
    slide.shapes.load("items/id,items/type");
    selected.load("items/id,items/type");
    await context.sync();

    const insertedCandidates = slide.shapes.items.filter(
      (shape) => !before.shapeIds.has(shape.id),
    );
    const selectedIds = new Set(selected.items.map((shape) => shape.id));
    const selectedInserted = insertedCandidates.find((shape) =>
      selectedIds.has(shape.id),
    );
    const candidate =
      selectedInserted ??
      insertedCandidates[insertedCandidates.length - 1] ??
      (selected.items.length === 1 && selected.items[0].type === "Image"
        ? selected.items[0]
        : null);

    return candidate
      ? ({ slideId: before.slideId, shapeId: candidate.id } satisfies PowerPointObjectReference)
      : null;
  });
}

async function locateInsertedShapeWithRetry(before: PowerPointSlideSnapshot) {
  for (let attempt = 0; attempt < 4; attempt += 1) {
    const reference = await locateInsertedShape(before);
    if (reference) return reference;
    if (attempt < 3) {
      await new Promise((resolve) => window.setTimeout(resolve, 50));
    }
  }
  return null;
}

async function readSelectedPowerPointShapeCore(): Promise<PowerPointSelectedShapeCore | null> {
  requirePowerPointEditingApi();
  return PowerPoint.run(async (context) => {
    const supportsAccessibility = supportsPowerPointApi("1.10");
    const supportsZOrder = supportsPowerPointApi("1.8");
    const selected = context.presentation.getSelectedShapes();
    const properties = [
      "items/id",
      "items/name",
      "items/left",
      "items/top",
      "items/width",
      "items/height",
    ];
    if (supportsAccessibility) {
      properties.push(
        "items/altTextDescription",
        "items/altTextTitle",
        "items/rotation",
      );
    }
    if (supportsZOrder) properties.push("items/zOrderPosition");
    selected.load(properties.join(","));
    await context.sync();
    if (selected.items.length !== 1) return null;

    const shape = selected.items[0];
    const slide = shape.getParentSlide();
    slide.load("id");
    await context.sync();

    return {
      slideId: slide.id,
      shapeId: shape.id,
      name: shape.name,
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
      altTextTitle: supportsAccessibility ? shape.altTextTitle : undefined,
      altTextDescription: supportsAccessibility
        ? shape.altTextDescription
        : undefined,
      rotation: supportsAccessibility ? shape.rotation : undefined,
      zOrderPosition: supportsZOrder ? shape.zOrderPosition : undefined,
    };
  });
}

async function readPowerPointShapeTags(reference: PowerPointObjectReference) {
  return PowerPoint.run(async (context) => {
    const shape = context.presentation.slides
      .getItem(reference.slideId)
      .shapes.getItem(reference.shapeId);
    shape.tags.load("items/key,items/value");
    await context.sync();
    return new Map(
      shape.tags.items.map((tag) => [tag.key.toUpperCase(), tag.value]),
    );
  });
}

async function readSelectedPowerPointShape(): Promise<PowerPointShapeSnapshot | null> {
  const core = await readSelectedPowerPointShapeCore();
  if (!core) return null;

  let formulaId = formulaIdFromPowerPointShapeName(core.name);
  let encodedMetadata =
    core.altTextTitle === POWERPOINT_ALT_TEXT_TITLE && core.altTextDescription
      ? core.altTextDescription
      : null;

  try {
    const tags = await readPowerPointShapeTags(core);
    formulaId ??= tags.get(POWERPOINT_FORMULA_ID_TAG) ?? null;
    encodedMetadata = decodeMetadataTags(tags) ?? encodedMetadata;
  } catch (error) {
    if (!formulaId) {
      throw new Error(
        `PowerPoint 无法读取所选对象的 VisualTeX 标记：${officeErrorMessage(
          error,
          "未知 Office 错误",
        )}`,
      );
    }
    logOptionalDecorationFailure("metadata tag reading", error);
  }

  if (!formulaId) return null;
  return {
    slideId: core.slideId,
    shapeId: core.shapeId,
    formulaId,
    encodedMetadata,
    left: core.left,
    top: core.top,
    width: core.width,
    height: core.height,
    rotation: core.rotation,
    zOrderPosition: core.zOrderPosition,
  };
}

async function readNativeSelectedPowerPointShape(): Promise<PowerPointShapeSnapshot | null> {
  const native = await getNativePowerPointSelection();
  const formulaId = formulaIdFromPowerPointShapeName(native.shapeName);
  if (!formulaId) return null;
  return {
    slideId: `native:${native.slideIndex}`,
    shapeId: native.shapeName,
    native: {
      slideIndex: native.slideIndex,
      shapeName: native.shapeName,
      left: native.left,
      top: native.top,
      width: native.width,
      height: native.height,
    },
    formulaId,
    encodedMetadata: null,
    nativePresentationIdentity: native.presentationIdentity,
    nativeSlideId: native.slideId,
    left: native.left,
    top: native.top,
    width: native.width,
    height: native.height,
  };
}

async function selectOriginalShape(
  reference: PowerPointObjectReference,
): Promise<PowerPointGeometry> {
  if (reference.native) {
    return {
      left: reference.native.left,
      top: reference.native.top,
      width: reference.native.width,
      height: reference.native.height,
    } satisfies PowerPointGeometry;
  }
  requirePowerPointEditingApi();
  return PowerPoint.run(async (context) => {
    const supportsAccessibility = supportsPowerPointApi("1.10");
    const supportsZOrder = supportsPowerPointApi("1.8");
    const slide = context.presentation.slides.getItem(reference.slideId);
    const shape = slide.shapes.getItemOrNullObject(reference.shapeId);
    const properties = [
      "id",
      "left",
      "top",
      "width",
      "height",
      "isNullObject",
    ];
    if (supportsAccessibility) properties.push("rotation");
    if (supportsZOrder) properties.push("zOrderPosition");
    shape.load(properties.join(","));
    await context.sync();
    if (shape.isNullObject) {
      throw new Error(
        "找不到原 VisualTeX Shape。请重新选择该公式后再执行编辑。",
      );
    }
    slide.setSelectedShapes([shape.id]);
    await context.sync();
    return {
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
      rotation: supportsAccessibility ? shape.rotation : undefined,
      zOrderPosition: supportsZOrder ? shape.zOrderPosition : undefined,
    } satisfies PowerPointGeometry;
  });
}

async function applyInsertedShapeCore(
  reference: PowerPointObjectReference,
  metadata: VisualTeXFormulaMetadata,
  geometry: PowerPointGeometry,
) {
  return PowerPoint.run(async (context) => {
    const shape = context.presentation.slides
      .getItem(reference.slideId)
      .shapes.getItemOrNullObject(reference.shapeId);
    shape.load("id,isNullObject");
    await context.sync();
    if (shape.isNullObject) {
      throw new Error("PowerPoint 中找不到刚刚插入的公式图片。");
    }

    shape.name = powerpointShapeName(metadata.formulaId);
    if (geometry.left !== undefined) shape.left = geometry.left;
    if (geometry.top !== undefined) shape.top = geometry.top;
    shape.width = geometry.width;
    shape.height = geometry.height;
    await context.sync();
  });
}

async function writeInsertedShapeTags(
  reference: PowerPointObjectReference,
  metadata: VisualTeXFormulaMetadata,
) {
  const encodedMetadata = encodeFormulaMetadata(metadata);
  const chunks = encodeMetadataTags(encodedMetadata);
  return PowerPoint.run(async (context) => {
    const shape = context.presentation.slides
      .getItem(reference.slideId)
      .shapes.getItem(reference.shapeId);
    shape.tags.add(POWERPOINT_FORMULA_ID_TAG, metadata.formulaId);
    shape.tags.add(POWERPOINT_METADATA_COUNT_TAG, String(chunks.length));
    chunks.forEach((chunk, index) => {
      shape.tags.add(`${POWERPOINT_METADATA_CHUNK_PREFIX}${index}`, chunk);
    });
    await context.sync();
  });
}

async function writeInsertedShapeAccessibility(
  reference: PowerPointObjectReference,
  metadata: VisualTeXFormulaMetadata,
  geometry: PowerPointGeometry,
) {
  if (!supportsPowerPointApi("1.10")) return;
  const encodedMetadata = encodeFormulaMetadata(metadata);
  return PowerPoint.run(async (context) => {
    const shape = context.presentation.slides
      .getItem(reference.slideId)
      .shapes.getItem(reference.shapeId);
    shape.altTextTitle = POWERPOINT_ALT_TEXT_TITLE;
    shape.altTextDescription = encodedMetadata;
    shape.isDecorative = false;
    if (geometry.rotation !== undefined) shape.rotation = geometry.rotation;
    await context.sync();
  });
}

async function deleteOriginalShape(
  originalReference: PowerPointObjectReference,
  insertedReference: PowerPointObjectReference,
) {
  if (
    originalReference.slideId === insertedReference.slideId &&
    originalReference.shapeId === insertedReference.shapeId
  ) {
    return;
  }
  return PowerPoint.run(async (context) => {
    const original = context.presentation.slides
      .getItem(originalReference.slideId)
      .shapes.getItemOrNullObject(originalReference.shapeId);
    original.load("isNullObject");
    await context.sync();
    if (!original.isNullObject) {
      original.delete();
      await context.sync();
    }
  });
}

async function restoreInsertedShapeZOrder(
  reference: PowerPointObjectReference,
  targetZOrder: number | undefined,
) {
  if (!supportsPowerPointApi("1.8") || targetZOrder === undefined) return;
  return PowerPoint.run(async (context) => {
    const inserted = context.presentation.slides
      .getItem(reference.slideId)
      .shapes.getItem(reference.shapeId);
    inserted.load("zOrderPosition");
    await context.sync();
    const difference = inserted.zOrderPosition - targetZOrder;
    const operation = difference > 0 ? "SendBackward" : "BringForward";
    for (
      let index = 0;
      index < Math.min(1024, Math.abs(difference));
      index += 1
    ) {
      inserted.setZOrder(operation);
    }
    if (difference !== 0) await context.sync();
  });
}

async function selectInsertedShape(reference: PowerPointObjectReference) {
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItem(reference.slideId);
    slide.setSelectedShapes([reference.shapeId]);
    await context.sync();
  });
}

function logOptionalDecorationFailure(stage: string, error: unknown) {
  console.warn(
    `VisualTeX PowerPoint ${stage} failed after the formula image was inserted: ${officeErrorMessage(
      error,
      "unknown Office error",
    )}`,
  );
}

function metadataFromPowerPointSession(session: OfficeFormulaSession) {
  return createFormulaMetadata({
    formulaId: session.formulaId,
    title: session.title,
    lines: session.lines,
    codeFormat: session.codeFormat,
    displayMode: session.displayMode,
    renderWidthPx: session.exportResult?.width ?? session.exportWidth,
    renderHeightPx: session.exportResult?.height ?? session.exportHeight,
    baseline: session.exportResult?.baseline,
    original: session.originalMetadata,
  });
}

function createSessionSeed(metadata: VisualTeXFormulaMetadata) {
  return {
    formulaId: metadata.formulaId,
    title: metadata.title,
    lines: metadata.lines.map((line) => ({ ...line })),
    activeLineId: metadata.lines[0]?.id ?? null,
    codeFormat: metadata.codeFormat,
    displayMode: metadata.displayMode,
    originalMetadata: metadata,
    autoCommitOnClose: true,
  };
}

function encodeNativePowerPointEditTarget(slideIndex: number, shapeName: string) {
  const encodedName = Array.from(new TextEncoder().encode(shapeName), (byte) =>
    byte.toString(16).padStart(2, "0"),
  ).join("");
  return `visualtex-ppt-native-edit:${slideIndex}:${encodedName}`;
}

function snapshotFromInteractionTarget(
  target: OfficeInteractionTarget,
): PowerPointShapeSnapshot | null {
  const formulaId = formulaIdFromPowerPointShapeName(target.shapeName);
  if (
    target.host !== "powerpoint" ||
    !formulaId ||
    formulaId !== target.formulaId ||
    !Number.isInteger(target.slideIndex) ||
    (target.slideIndex ?? 0) < 1 ||
    !Number.isFinite(target.left) ||
    !Number.isFinite(target.top) ||
    !Number.isFinite(target.width) ||
    !Number.isFinite(target.height) ||
    (target.width ?? 0) <= 0 ||
    (target.height ?? 0) <= 0
  ) {
    return null;
  }
  const slideIndex = target.slideIndex as number;
  const left = target.left as number;
  const top = target.top as number;
  const width = target.width as number;
  const height = target.height as number;
  return {
    slideId: `native:${slideIndex}`,
    shapeId: target.shapeName,
    native: {
      slideIndex,
      shapeName: target.shapeName,
      left,
      top,
      width,
      height,
    },
    formulaId,
    encodedMetadata: null,
    nativePresentationIdentity: target.presentationIdentity,
    nativeSlideId: target.slideId,
    left,
    top,
    width,
    height,
  };
}

function geometryMatchesNativeCommit(
  shape: PowerPointSelectedShapeCore,
  selection: NativePowerPointCommitSelection,
) {
  const tolerance = 2;
  return (
    Math.abs(shape.left - selection.left) <= tolerance &&
    Math.abs(shape.top - selection.top) <= tolerance &&
    Math.abs(shape.width - selection.width) <= tolerance &&
    Math.abs(shape.height - selection.height) <= tolerance
  );
}

async function locateNativeCommitShapeByIdentity(
  selection: NativePowerPointCommitSelection,
): Promise<PowerPointSelectedShapeCore | null> {
  if (!Number.isInteger(selection.slideIndex) || selection.slideIndex < 1) {
    return null;
  }
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItemAt(selection.slideIndex - 1);
    slide.load("id");
    const shapes = slide.shapes;
    shapes.load("items/id,items/name,items/left,items/top,items/width,items/height");
    await context.sync();

    const exactName = shapes.items.find(
      (shape) => shape.name === selection.shapeName,
    );
    const geometryMatches = shapes.items.filter((shape) =>
      geometryMatchesNativeCommit(
        {
          slideId: slide.id,
          shapeId: shape.id,
          name: shape.name,
          left: shape.left,
          top: shape.top,
          width: shape.width,
          height: shape.height,
        },
        selection,
      ),
    );
    const resolved = exactName ??
      (geometryMatches.length === 1 ? geometryMatches[0] : null);
    if (!resolved) return null;
    return {
      slideId: slide.id,
      shapeId: resolved.id,
      name: resolved.name,
      left: resolved.left,
      top: resolved.top,
      width: resolved.width,
      height: resolved.height,
    };
  });
}

async function selectedNativeCommitShape(
  selection: NativePowerPointCommitSelection,
) {
  for (let attempt = 0; attempt < 6; attempt += 1) {
    // The native transaction already returns the durable slide index, shape
    // name and geometry. Resolve that immutable identity first; PowerPoint can
    // legitimately move the UI selection before Office.js starts finalizing.
    const located = await locateNativeCommitShapeByIdentity(selection).catch(
      () => null,
    );
    if (located) return located;

    // Compatibility fallback for older PowerPoint builds where the shape
    // collection is briefly stale but the newly pasted picture remains selected.
    const selected = await readSelectedPowerPointShapeCore().catch(() => null);
    if (selected && geometryMatchesNativeCommit(selected, selection)) return selected;
    await new Promise((resolve) => window.setTimeout(resolve, 60));
  }
  throw new Error(
    "PowerPoint 已粘贴公式，但无法按返回的 Shape 名称、幻灯片和几何位置重新定位对象。未写入可编辑标记。",
  );
}

async function verifyNativePowerPointDecoration(
  reference: PowerPointObjectReference,
  formulaId: string,
) {
  return PowerPoint.run(async (context) => {
    const shape = context.presentation.slides
      .getItem(reference.slideId)
      .shapes.getItemOrNullObject(reference.shapeId);
    shape.load("isNullObject,name");
    shape.tags.load("items/key,items/value");
    await context.sync();
    if (shape.isNullObject) return false;
    const tags = new Map(
      shape.tags.items.map((tag) => [tag.key.toUpperCase(), tag.value]),
    );
    return (
      shape.name === powerpointShapeName(formulaId) &&
      tags.get(POWERPOINT_FORMULA_ID_TAG) === formulaId &&
      Number(tags.get(POWERPOINT_METADATA_COUNT_TAG) ?? "0") > 0
    );
  });
}

async function finalizeSelectedNativePowerPointShape(
  session: OfficeFormulaSession,
  selection: NativePowerPointCommitSelection,
) {
  const metadata = metadataFromPowerPointSession(session);
  const selected = await selectedNativeCommitShape(selection);
  const reference: PowerPointObjectReference = {
    slideId: selected.slideId,
    shapeId: selected.shapeId,
  };
  const geometry: PowerPointGeometry = {
    left: selection.left,
    top: selection.top,
    width: selection.width,
    height: selection.height,
    rotation: selected.rotation,
    zOrderPosition: selected.zOrderPosition,
  };

  // PowerPoint normalizes a pasted SVG asynchronously and can overwrite the
  // first name assignment. Reapply the complete identity after two host settle
  // points, and require both the shape name and chunked metadata tags to read
  // back before the native Session can be confirmed.
  for (let attempt = 0; attempt < 3; attempt += 1) {
    await applyInsertedShapeCore(reference, metadata, geometry);
    await writeInsertedShapeTags(reference, metadata);
    await writeInsertedShapeAccessibility(reference, metadata, geometry);
    if (await verifyNativePowerPointDecoration(reference, metadata.formulaId)) {
      await putCachedFormulaMetadata(metadata);
      return;
    }
    await new Promise((resolve) => window.setTimeout(resolve, 120));
  }
  throw new Error(
    "PowerPoint 未能持久保存 VisualTeX 公式标记；为避免生成无法再次编辑的图片，本次更新未确认完成。",
  );
}

export class PowerPointAdapter implements OfficeHostAdapter {
  readonly host = "powerpoint" as const;
  private pendingInteractionTarget: OfficeInteractionTarget | null = null;

  prepareInteractionTarget(target: OfficeInteractionTarget) {
    this.pendingInteractionTarget = target;
  }

  async readSelection(mode: OfficeSessionMode): Promise<OfficeSelectionContext> {
    if (mode === "create") {
      const nativeSlide = await getNativePowerPointSlideSnapshot().catch(
        () => null,
      );
      const sourceDocumentId = nativeSlide
        ? `visualtex-ppt-native-presentation:${nativeSlide.presentationIdentity}`
        : await getPresentationId();
      return {
        sourceDocumentId,
        sourceObjectId: nativeSlide
          ? `visualtex-ppt-native-slide:${nativeSlide.slideId}:${nativeSlide.slideIndex}`
          : null,
        sessionSeed: {},
      };
    }

    const interactionTarget = this.pendingInteractionTarget;
    this.pendingInteractionTarget = null;
    let selected = interactionTarget
      ? snapshotFromInteractionTarget(interactionTarget)
      : null;
    let officeSelectionError: unknown = null;
    let nativeSelectionError: unknown = null;
    let nativeSelectionAttempted = false;

    // The macOS native commit path always names the inserted shape and stores
    // its metadata in the companion cache. Read that inexpensive source first
    // instead of waiting for several PowerPoint.run/context.sync round trips.
    // Office.js remains the compatibility fallback for legacy tagged shapes.
    if (!selected && USE_NATIVE_POWERPOINT_EDIT_TARGET) {
      nativeSelectionAttempted = true;
      try {
        selected = await readNativeSelectedPowerPointShape();
      } catch (error) {
        nativeSelectionError = error;
      }
    }

    if (!selected && supportsPowerPointApi("1.5")) {
      try {
        selected = await readSelectedPowerPointShape();
      } catch (error) {
        officeSelectionError = error;
      }
    }

    if (!selected && !nativeSelectionAttempted) {
      try {
        selected = await readNativeSelectedPowerPointShape();
      } catch (error) {
        nativeSelectionError = error;
      }
    }

    if (!selected && nativeSelectionError) {
      if (officeSelectionError) {
        throw new Error(
          `PowerPoint 无法读取所选公式：${officeErrorMessage(
            officeSelectionError,
            "Office.js 选择读取失败",
          )}；${officeErrorMessage(
            nativeSelectionError,
            "macOS 原生选择读取失败",
          )}`,
        );
      }
      throw nativeSelectionError;
    }
    if (!selected) {
      throw new Error(
        "当前选中的对象没有 VisualTeX 标记。请先选择由 VisualTeX 1.0.10 或更高版本插入的公式。",
      );
    }
    const objectMetadata = selected.encodedMetadata
      ? decodeFormulaMetadata(selected.encodedMetadata)
      : null;
    const cachedMetadata = objectMetadata
      ? null
      : await getCachedFormulaMetadata(selected.formulaId).catch(() => null);
    const metadata = objectMetadata ?? cachedMetadata;
    if (!metadata || metadata.formulaId !== selected.formulaId) {
      throw new Error(
        "所选 VisualTeX Shape 的原始 LaTeX metadata 已损坏或缺失，无法安全编辑。",
      );
    }

    const nativeSlide = selected.nativePresentationIdentity
      ? null
      : await getNativePowerPointSlideSnapshot().catch(() => null);
    const sourceDocumentId = selected.nativePresentationIdentity
      ? `visualtex-ppt-native-presentation:${selected.nativePresentationIdentity}`
      : nativeSlide
        ? `visualtex-ppt-native-presentation:${nativeSlide.presentationIdentity}`
        : await getPresentationId();

    const nativeEditTarget = USE_NATIVE_POWERPOINT_EDIT_TARGET
      ? selected.native
        ? {
            slideIndex: selected.native.slideIndex,
            shapeName: selected.native.shapeName,
          }
        : await getNativePowerPointSelection().catch(() => null)
      : null;

    return {
      sourceDocumentId,
      sourceObjectId: nativeEditTarget
        ? encodeNativePowerPointEditTarget(
            nativeEditTarget.slideIndex,
            nativeEditTarget.shapeName,
          )
        : encodePowerPointObjectReference(selected),
      sessionSeed: {
        ...createSessionSeed(metadata),
        exportWidth: selected.width / 0.75,
        exportHeight: selected.height / 0.75,
      },
    };
  }

  async finalizeNativePowerPointCommit(
    session: OfficeFormulaSession,
    selection: NativePowerPointCommitSelection,
  ): Promise<void> {
    await finalizeSelectedNativePowerPointShape(session, selection);
  }

  async applySession(session: OfficeFormulaSession): Promise<void> {
    const metadata = metadataFromPowerPointSession(session);
    await putCachedFormulaMetadata(metadata);

    const supportsEditing = supportsPowerPointApi("1.5");
    const originalReference =
      session.mode === "edit"
        ? decodePowerPointObjectReference(session.sourceObjectId)
        : null;
    if (session.mode === "edit" && !originalReference) {
      throw new Error("PowerPoint Session 缺少原 Shape 标识，无法安全替换。");
    }
    if (
      session.mode === "edit" &&
      !supportsEditing &&
      !originalReference?.native
    ) {
      requirePowerPointEditingApi();
    }

    const geometry = originalReference
      ? await selectOriginalShape(originalReference)
      : calculatePowerPointGeometry(
          session.exportResult?.width ?? session.exportWidth,
          session.exportResult?.height ?? session.exportHeight,
        );

    const nativeBefore = await getNativePowerPointSlideSnapshot().catch(() => null);
    let before: PowerPointSlideSnapshot | null = null;
    if (supportsEditing && !originalReference?.native) {
      try {
        before = await captureSlideSnapshot(
          originalReference ? originalReference.slideId : undefined,
        );
      } catch (error) {
        logOptionalDecorationFailure("pre-insertion snapshot", error);
      }
    }

    await insertImageWithFallback(session, geometry);

    let nativeMarked = false;
    const needsNativeMarking =
      session.mode === "create" || Boolean(originalReference?.native) || !before;
    if (nativeBefore && needsNativeMarking) {
      try {
        if (originalReference?.native) {
          await replaceLastNativePowerPointFormula(
            metadata.formulaId,
            nativeBefore.shapeNames,
            originalReference.native.shapeName,
            {
              left: originalReference.native.left,
              top: originalReference.native.top,
              width: originalReference.native.width,
              height: originalReference.native.height,
            },
          );
        } else {
          await markLastNativePowerPointFormula(
            metadata.formulaId,
            nativeBefore.shapeNames,
          );
        }
        nativeMarked = true;
      } catch (error) {
        if (!originalReference?.native) {
          try {
            await markNativePowerPointSelection(metadata.formulaId);
            nativeMarked = true;
          } catch (fallbackError) {
            logOptionalDecorationFailure(
              "native shape marking",
              new Error(
                `${officeErrorMessage(error, "mark-last failed")}; ${officeErrorMessage(
                  fallbackError,
                  "selected-shape fallback failed",
                )}`,
              ),
            );
          }
        } else {
          throw new Error(
            `PowerPoint 已插入新公式，但无法安全替换原公式：${officeErrorMessage(
              error,
              "macOS 原生替换失败",
            )}`,
          );
        }
      }
    }

    if (originalReference?.native) {
      if (!nativeMarked) {
        throw new Error("PowerPoint 无法标记并替换新公式对象。");
      }
      return;
    }

    if (!before) return;

    let insertedReference: PowerPointObjectReference | null = null;
    try {
      insertedReference = await locateInsertedShapeWithRetry(before);
    } catch (error) {
      logOptionalDecorationFailure("inserted-shape detection", error);
    }
    if (!insertedReference) {
      if (!nativeMarked) {
        logOptionalDecorationFailure(
          "inserted-shape detection",
          new Error("PowerPoint did not expose the inserted image shape"),
        );
      }
      return;
    }

    try {
      await applyInsertedShapeCore(insertedReference, metadata, geometry);
    } catch (error) {
      logOptionalDecorationFailure("shape naming and sizing", error);
      if (!nativeMarked) return;
    }

    try {
      await writeInsertedShapeTags(insertedReference, metadata);
    } catch (error) {
      logOptionalDecorationFailure("metadata tags", error);
    }
    try {
      await writeInsertedShapeAccessibility(insertedReference, metadata, geometry);
    } catch (error) {
      logOptionalDecorationFailure("accessibility metadata", error);
    }

    if (originalReference && !originalReference.native) {
      try {
        await deleteOriginalShape(originalReference, insertedReference);
      } catch (error) {
        logOptionalDecorationFailure("original-shape cleanup", error);
      }
    }
    try {
      await restoreInsertedShapeZOrder(insertedReference, geometry.zOrderPosition);
    } catch (error) {
      logOptionalDecorationFailure("z-order restoration", error);
    }
    try {
      await selectInsertedShape(insertedReference);
    } catch (error) {
      logOptionalDecorationFailure("shape selection", error);
    }
  }

  async openDesktopApp(): Promise<void> {
    await revealDesktopApp();
  }

  showMessage(message: string) {
    const status = document.getElementById("bridge-status");
    if (status) status.textContent = message;
  }
}
