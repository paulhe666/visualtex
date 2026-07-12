import {
  getCachedFormulaMetadata,
  putCachedFormulaMetadata,
} from "../api/companionClient";
import type {
  OfficeFormulaSession,
  OfficeSessionMode,
} from "../api/sessionClient";
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
  OfficeSelectionContext,
} from "./OfficeHostAdapter";

const POWERPOINT_ALT_TEXT_TITLE = "VisualTeX Formula";
const MAX_POWERPOINT_WIDTH_PT = 600;
const MAX_POWERPOINT_HEIGHT_PT = 400;
const MIN_POWERPOINT_SIZE_PT = 12;

interface PowerPointShapeSnapshot extends PowerPointObjectReference {
  formulaId: string;
  encodedMetadata: string | null;
  left: number;
  top: number;
  width: number;
  height: number;
  rotation: number;
  zOrderPosition: number;
}

interface PowerPointGeometry {
  left?: number;
  top?: number;
  width: number;
  height: number;
  rotation: number;
  zOrderPosition?: number;
}

function ensurePowerPointApi() {
  let supported = false;
  try {
    supported = Office.context.requirements.isSetSupported(
      "PowerPointApi",
      "1.10",
    );
  } catch {
    supported = false;
  }
  if (!supported) {
    throw new Error(
      "当前 PowerPoint 版本不支持 VisualTeX 所需的 PowerPointApi 1.10。请更新 Microsoft PowerPoint for Mac。",
    );
  }
}

function officeAsync<T>(
  invoke: (callback: (result: Office.AsyncResult<T>) => void) => void,
): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    invoke((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(
          new Error(result.error?.message ?? "PowerPoint image insertion failed."),
        );
      }
    });
  });
}

function calculatePowerPointGeometry(
  widthPx: number,
  heightPx: number,
): PowerPointGeometry {
  const naturalWidth = Math.max(MIN_POWERPOINT_SIZE_PT, widthPx * 0.75);
  const naturalHeight = Math.max(MIN_POWERPOINT_SIZE_PT, heightPx * 0.75);
  const scale = Math.min(
    1,
    MAX_POWERPOINT_WIDTH_PT / naturalWidth,
    MAX_POWERPOINT_HEIGHT_PT / naturalHeight,
  );
  return {
    width: naturalWidth * scale,
    height: naturalHeight * scale,
    rotation: 0,
  };
}

async function setSelectedPowerPointImage(
  base64: string,
  geometry: PowerPointGeometry,
) {
  await officeAsync<void>((callback) =>
    Office.context.document.setSelectedDataAsync(
      base64,
      {
        coercionType: Office.CoercionType.Image,
        imageLeft: geometry.left,
        imageTop: geometry.top,
        imageWidth: geometry.width,
        imageHeight: geometry.height,
      },
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
  try {
    await setSelectedPowerPointImage(exportResult.svgBase64, geometry);
  } catch (svgError) {
    if (!exportResult.pngBase64) throw svgError;
    await setSelectedPowerPointImage(exportResult.pngBase64, geometry);
  }
}

async function getPresentationId() {
  return PowerPoint.run(async (context) => {
    const presentation = context.presentation;
    presentation.load("id");
    await context.sync();
    return presentation.id || Office.context.document.url || null;
  });
}

async function readSelectedPowerPointShape(): Promise<PowerPointShapeSnapshot | null> {
  return PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load(
      "items/id,items/name,items/altTextDescription,items/altTextTitle,items/left,items/top,items/width,items/height,items/rotation,items/zOrderPosition,items/type",
    );
    await context.sync();
    if (selected.items.length !== 1) return null;
    const shape = selected.items[0];
    const formulaId = formulaIdFromPowerPointShapeName(shape.name);
    if (!formulaId || shape.altTextTitle !== POWERPOINT_ALT_TEXT_TITLE) return null;
    const slide = shape.getParentSlide();
    slide.load("id");
    await context.sync();
    return {
      slideId: slide.id,
      shapeId: shape.id,
      formulaId,
      encodedMetadata: shape.altTextDescription || null,
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
      rotation: shape.rotation,
      zOrderPosition: shape.zOrderPosition,
    };
  });
}

async function selectOriginalShape(reference: PowerPointObjectReference) {
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItem(reference.slideId);
    const shape = slide.shapes.getItemOrNullObject(reference.shapeId);
    shape.load("id,left,top,width,height,rotation,zOrderPosition,name,isNullObject");
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
      rotation: shape.rotation,
      zOrderPosition: shape.zOrderPosition,
    } satisfies PowerPointGeometry;
  });
}

async function decorateInsertedShape(
  metadata: VisualTeXFormulaMetadata,
  geometry: PowerPointGeometry,
  originalReference: PowerPointObjectReference | null,
) {
  return PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load(
      "items/id,items/name,items/left,items/top,items/width,items/height,items/rotation,items/zOrderPosition",
    );
    await context.sync();
    if (selected.items.length !== 1) {
      throw new Error(
        "PowerPoint 未返回唯一的新公式 Shape，无法安全写入 metadata。",
      );
    }
    const inserted = selected.items[0];
    const slide = inserted.getParentSlide();
    slide.load("id");

    inserted.name = powerpointShapeName(metadata.formulaId);
    inserted.altTextTitle = POWERPOINT_ALT_TEXT_TITLE;
    inserted.altTextDescription = encodeFormulaMetadata(metadata);
    inserted.isDecorative = false;
    if (geometry.left !== undefined) inserted.left = geometry.left;
    if (geometry.top !== undefined) inserted.top = geometry.top;
    inserted.width = geometry.width;
    inserted.height = geometry.height;
    inserted.rotation = geometry.rotation;
    await context.sync();

    if (originalReference) {
      const originalSlide = context.presentation.slides.getItem(
        originalReference.slideId,
      );
      const original = originalSlide.shapes.getItemOrNullObject(
        originalReference.shapeId,
      );
      original.load("id,isNullObject");
      await context.sync();
      if (!original.isNullObject && original.id !== inserted.id) {
        original.delete();
        await context.sync();
      }
    }

    inserted.load("zOrderPosition");
    await context.sync();
    const targetZ = geometry.zOrderPosition;
    if (targetZ !== undefined) {
      const difference = inserted.zOrderPosition - targetZ;
      const operation = difference > 0 ? "SendBackward" : "BringForward";
      for (let index = 0; index < Math.min(1024, Math.abs(difference)); index += 1) {
        inserted.setZOrder(operation);
      }
      if (difference !== 0) await context.sync();
    }

    return encodePowerPointObjectReference({
      slideId: slide.id,
      shapeId: inserted.id,
    });
  });
}

function createSessionSeed(metadata: VisualTeXFormulaMetadata) {
  return {
    formulaId: metadata.formulaId,
    title: metadata.title,
    lines: metadata.lines.map((line) => ({ ...line })),
    activeLineId: metadata.lines[0]?.id ?? null,
    codeFormat: metadata.codeFormat,
    originalMetadata: metadata,
    autoCommitOnClose: true,
  };
}

export class PowerPointAdapter implements OfficeHostAdapter {
  readonly host = "powerpoint" as const;

  async readSelection(mode: OfficeSessionMode): Promise<OfficeSelectionContext> {
    ensurePowerPointApi();
    const sourceDocumentId = await getPresentationId();
    if (mode === "create") {
      return {
        sourceDocumentId,
        sourceObjectId: null,
        sessionSeed: {},
      };
    }

    const selected = await readSelectedPowerPointShape();
    if (!selected) {
      throw new Error(
        "当前选中的对象不是 VisualTeX 公式。请先选择由 VisualTeX 插入的公式。",
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

    return {
      sourceDocumentId,
      sourceObjectId: encodePowerPointObjectReference(selected),
      sessionSeed: {
        ...createSessionSeed(metadata),
        exportWidth: selected.width / 0.75,
        exportHeight: selected.height / 0.75,
      },
    };
  }

  async applySession(session: OfficeFormulaSession): Promise<void> {
    ensurePowerPointApi();
    const metadata = createFormulaMetadata({
      formulaId: session.formulaId,
      title: session.title,
      lines: session.lines,
      codeFormat: session.codeFormat,
      displayMode: session.originalMetadata?.displayMode ?? "block",
      original: session.originalMetadata,
    });
    await putCachedFormulaMetadata(metadata);

    const originalReference =
      session.mode === "edit"
        ? decodePowerPointObjectReference(session.sourceObjectId)
        : null;
    if (session.mode === "edit" && !originalReference) {
      throw new Error("PowerPoint Session 缺少原 Shape 标识，无法安全替换。");
    }
    const geometry = originalReference
      ? await selectOriginalShape(originalReference)
      : calculatePowerPointGeometry(
          session.exportResult?.width ?? session.exportWidth,
          session.exportResult?.height ?? session.exportHeight,
        );
    await insertImageWithFallback(session, geometry);
    await decorateInsertedShape(metadata, geometry, originalReference);
  }

  async openDesktopApp(): Promise<void> {
    try {
      Office.context.ui.openBrowserWindow("visualtex://office/start");
    } catch {
      window.location.href = "visualtex://office/start";
    }
  }

  showMessage(message: string) {
    const status = document.getElementById("bridge-status");
    if (status) status.textContent = message;
  }
}
