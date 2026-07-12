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
  readWordDocumentMetadata,
  writeWordDocumentMetadata,
} from "../metadata/wordMetadata";
import type {
  OfficeHostAdapter,
  OfficeSelectionContext,
} from "./OfficeHostAdapter";

const WORD_CONTENT_CONTROL_TITLE = "VisualTeX Formula";
const WORD_CONTENT_CONTROL_TAG_PREFIX = "visualtex:";
const WORD_ALT_TEXT_TITLE_PREFIX = "VisualTeX_";
const MAX_WORD_FORMULA_WIDTH_PT = 500;
const MIN_WORD_FORMULA_SIZE_PT = 12;

interface SelectedWordFormula {
  contentControlId: number;
  formulaId: string;
  encodedMetadata: string | null;
  width: number;
  height: number;
}

function contentControlTag(formulaId: string) {
  return `${WORD_CONTENT_CONTROL_TAG_PREFIX}${formulaId}`;
}

function formulaIdFromTag(tag: string) {
  if (!tag.startsWith(WORD_CONTENT_CONTROL_TAG_PREFIX)) return null;
  const formulaId = tag.slice(WORD_CONTENT_CONTROL_TAG_PREFIX.length);
  return /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
    formulaId,
  )
    ? formulaId
    : null;
}

function wordApiSupported(version: string) {
  try {
    return Office.context.requirements.isSetSupported("WordApi", version);
  } catch {
    return false;
  }
}

function applyPictureSize(
  picture: Word.InlinePicture,
  widthPx: number,
  heightPx: number,
) {
  if (!(widthPx > 0) || !(heightPx > 0)) return;
  const naturalWidthPt = widthPx * 0.75;
  const naturalHeightPt = heightPx * 0.75;
  const scale = Math.min(1, MAX_WORD_FORMULA_WIDTH_PT / naturalWidthPt);
  picture.lockAspectRatio = true;
  picture.width = Math.max(MIN_WORD_FORMULA_SIZE_PT, naturalWidthPt * scale);
  picture.height = Math.max(MIN_WORD_FORMULA_SIZE_PT, naturalHeightPt * scale);
}

async function readSelectedWordFormula(): Promise<SelectedWordFormula | null> {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    const selectedControls = selection.contentControls;
    selectedControls.load("items/id,items/tag,items/title");

    let parent: Word.ContentControl | null = null;
    if (wordApiSupported("1.3")) {
      parent = selection.parentContentControlOrNullObject;
      parent.load("id,tag,title");
    }
    await context.sync();

    const candidates = [
      ...(parent && !parent.isNullObject ? [parent] : []),
      ...selectedControls.items,
    ];
    const control = candidates.find(
      (item, index) =>
        formulaIdFromTag(item.tag) &&
        candidates.findIndex((candidate) => candidate.id === item.id) === index,
    );
    if (!control) return null;

    const formulaId = formulaIdFromTag(control.tag);
    if (!formulaId || control.title !== WORD_CONTENT_CONTROL_TITLE) return null;

    const range = control.getRange("Content");
    const pictures = range.inlinePictures;
    pictures.load(
      "items/altTextDescription,items/altTextTitle,items/width,items/height",
    );
    await context.sync();
    const picture = pictures.items.find(
      (item) => item.altTextTitle === `${WORD_ALT_TEXT_TITLE_PREFIX}${formulaId}`,
    );
    if (!picture) return null;

    return {
      contentControlId: control.id,
      formulaId,
      encodedMetadata: picture.altTextDescription || null,
      width: picture.width,
      height: picture.height,
    };
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

async function insertWordFormula(
  session: OfficeFormulaSession,
  metadata: VisualTeXFormulaMetadata,
  imageBase64: string,
) {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    const picture = selection.insertInlinePictureFromBase64(
      imageBase64,
      "Replace",
    );
    picture.altTextTitle = `${WORD_ALT_TEXT_TITLE_PREFIX}${metadata.formulaId}`;
    picture.altTextDescription = encodeFormulaMetadata(metadata);
    applyPictureSize(
      picture,
      session.exportResult?.width ?? session.exportWidth,
      session.exportResult?.height ?? session.exportHeight,
    );

    const control = picture.insertContentControl();
    control.tag = contentControlTag(metadata.formulaId);
    control.title = WORD_CONTENT_CONTROL_TITLE;
    control.appearance = "BoundingBox";
    control.cannotDelete = false;
    control.cannotEdit = false;
    control.load("id");
    control.getRange("After").select("End");
    await context.sync();
    return String(control.id);
  });
}

async function replaceWordFormula(
  session: OfficeFormulaSession,
  metadata: VisualTeXFormulaMetadata,
  imageBase64: string,
) {
  return Word.run(async (context) => {
    const controls = context.document.contentControls.getByTag(
      contentControlTag(metadata.formulaId),
    );
    controls.load("items/id,items/tag,items/title");
    await context.sync();

    const requestedId = Number(session.sourceObjectId);
    const control =
      controls.items.find(
        (item) => Number.isFinite(requestedId) && item.id === requestedId,
      ) ?? controls.items[0];
    if (!control || control.title !== WORD_CONTENT_CONTROL_TITLE) {
      throw new Error(
        "找不到原 VisualTeX 公式。请重新选择该公式后再执行编辑。",
      );
    }

    const range = control.getRange("Content");
    const picture = range.insertInlinePictureFromBase64(imageBase64, "Replace");
    picture.altTextTitle = `${WORD_ALT_TEXT_TITLE_PREFIX}${metadata.formulaId}`;
    picture.altTextDescription = encodeFormulaMetadata(metadata);
    applyPictureSize(
      picture,
      session.exportResult?.width ?? session.exportWidth,
      session.exportResult?.height ?? session.exportHeight,
    );
    control.tag = contentControlTag(metadata.formulaId);
    control.title = WORD_CONTENT_CONTROL_TITLE;
    control.appearance = "BoundingBox";
    control.getRange("After").select("End");
    await context.sync();
    return String(control.id);
  });
}

async function applyWithImageFallback(
  session: OfficeFormulaSession,
  metadata: VisualTeXFormulaMetadata,
) {
  const exportResult = session.exportResult;
  if (!exportResult) {
    throw new Error("VisualTeX Session does not contain an exported formula image.");
  }
  const apply = session.mode === "edit" ? replaceWordFormula : insertWordFormula;
  try {
    return await apply(session, metadata, exportResult.svgBase64);
  } catch (svgError) {
    if (!exportResult.pngBase64) throw svgError;
    return apply(session, metadata, exportResult.pngBase64);
  }
}

export class WordAdapter implements OfficeHostAdapter {
  readonly host = "word" as const;

  async readSelection(mode: OfficeSessionMode): Promise<OfficeSelectionContext> {
    const sourceDocumentId = Office.context.document.url || null;
    if (mode === "create") {
      return {
        sourceDocumentId,
        sourceObjectId: null,
        sessionSeed: {},
      };
    }

    const selected = await readSelectedWordFormula();
    if (!selected) {
      throw new Error(
        "当前选中的对象不是 VisualTeX 公式。请先选择由 VisualTeX 插入的公式。",
      );
    }
    const objectMetadata = selected.encodedMetadata
      ? decodeFormulaMetadata(selected.encodedMetadata)
      : null;
    const documentMetadata = objectMetadata
      ? null
      : await readWordDocumentMetadata(selected.formulaId);
    const cachedMetadata =
      objectMetadata || documentMetadata
        ? null
        : await getCachedFormulaMetadata(selected.formulaId).catch(() => null);
    const metadata = objectMetadata ?? documentMetadata ?? cachedMetadata;
    if (!metadata || metadata.formulaId !== selected.formulaId) {
      throw new Error(
        "所选 VisualTeX 公式的原始 LaTeX metadata 已损坏或缺失，无法安全编辑。",
      );
    }

    return {
      sourceDocumentId,
      sourceObjectId: String(selected.contentControlId),
      sessionSeed: {
        ...createSessionSeed(metadata),
        exportWidth: selected.width / 0.75,
        exportHeight: selected.height / 0.75,
      },
    };
  }

  async applySession(session: OfficeFormulaSession): Promise<void> {
    const metadata = createFormulaMetadata({
      formulaId: session.formulaId,
      title: session.title,
      lines: session.lines,
      codeFormat: session.codeFormat,
      displayMode: session.originalMetadata?.displayMode ?? "inline",
      original: session.originalMetadata,
    });
    await putCachedFormulaMetadata(metadata);
    await applyWithImageFallback(session, metadata);
    await writeWordDocumentMetadata(metadata);
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
