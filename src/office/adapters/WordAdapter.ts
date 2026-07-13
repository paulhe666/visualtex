import {
  getCachedFormulaMetadata,
  putCachedFormulaMetadata,
  revealDesktopApp,
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

const LEGACY_WORD_CONTENT_CONTROL_TITLE = "VisualTeX Formula";
const LEGACY_WORD_CONTENT_CONTROL_TAG_PREFIX = "visualtex:";
const WORD_ALT_TEXT_TITLE_PREFIX = "VisualTeX_";
const MAX_WORD_FORMULA_WIDTH_PT = 500;
const MIN_WORD_FORMULA_SIZE_PT = 8;

interface SelectedWordFormula {
  formulaId: string;
  encodedMetadata: string | null;
  width: number;
  height: number;
}

interface AppliedPictureSize {
  width: number;
  height: number;
  scale: number;
}

function legacyContentControlTag(formulaId: string) {
  return `${LEGACY_WORD_CONTENT_CONTROL_TAG_PREFIX}${formulaId}`;
}

function formulaIdFromLegacyTag(tag: string) {
  if (!tag.startsWith(LEGACY_WORD_CONTENT_CONTROL_TAG_PREFIX)) return null;
  return validFormulaId(tag.slice(LEGACY_WORD_CONTENT_CONTROL_TAG_PREFIX.length));
}

function formulaIdFromPictureTitle(title: string) {
  if (!title.startsWith(WORD_ALT_TEXT_TITLE_PREFIX)) return null;
  return validFormulaId(title.slice(WORD_ALT_TEXT_TITLE_PREFIX.length));
}

function validFormulaId(value: string) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
    value,
  )
    ? value
    : null;
}

function requirementSupported(name: string, version: string) {
  try {
    return Office.context.requirements.isSetSupported(name, version);
  } catch {
    return false;
  }
}

function wordApiSupported(version: string) {
  return requirementSupported("WordApi", version);
}

function wordDesktopApiSupported(version: string) {
  return requirementSupported("WordApiDesktop", version);
}

function requireWordApi() {
  if (!wordApiSupported("1.1")) {
    throw new Error(
      "当前 Word 版本不支持 VisualTeX 所需的 WordApi 1.1。请更新 Microsoft Word。",
    );
  }
}

function applyPictureSize(
  picture: Word.InlinePicture,
  widthPx: number,
  heightPx: number,
): AppliedPictureSize {
  if (!(widthPx > 0) || !(heightPx > 0)) {
    return { width: 0, height: 0, scale: 1 };
  }
  const naturalWidthPt = widthPx * 0.75;
  const naturalHeightPt = heightPx * 0.75;
  const scale = Math.min(1, MAX_WORD_FORMULA_WIDTH_PT / naturalWidthPt);
  const width = Math.max(MIN_WORD_FORMULA_SIZE_PT, naturalWidthPt * scale);
  const height = Math.max(MIN_WORD_FORMULA_SIZE_PT, naturalHeightPt * scale);
  picture.lockAspectRatio = true;
  picture.width = width;
  picture.height = height;
  return { width, height, scale };
}

function applyPictureMetadata(
  picture: Word.InlinePicture,
  metadata: VisualTeXFormulaMetadata,
) {
  picture.altTextTitle = `${WORD_ALT_TEXT_TITLE_PREFIX}${metadata.formulaId}`;
  picture.altTextDescription = encodeFormulaMetadata(metadata);
}

function applyInlineBaseline(
  picture: Word.InlinePicture,
  _exportResult: OfficeFormulaSession["exportResult"],
  _size: AppliedPictureSize,
) {
  if (!wordDesktopApiSupported("1.3")) return;
  // Word for Mac already aligns an inline picture's bottom edge to the text
  // baseline. Applying a negative font position shifts the visible formula to
  // the top of its selection box on some Office builds. Keep the picture at
  // the native baseline; the export canvas itself is tightly cropped.
  picture.getRange("Whole").font.position = 0;
}

function styleDisplayParagraph(picture: Word.InlinePicture) {
  const paragraph = picture.getRange("Whole").paragraphs.getFirst();
  paragraph.alignment = "Centered";
  paragraph.spaceBefore = 0;
  paragraph.spaceAfter = 0;
}

function configurePicture(
  picture: Word.InlinePicture,
  session: OfficeFormulaSession,
  metadata: VisualTeXFormulaMetadata,
) {
  applyPictureMetadata(picture, metadata);
  const size = applyPictureSize(
    picture,
    session.exportResult?.width ?? session.exportWidth,
    session.exportResult?.height ?? session.exportHeight,
  );
  if (session.displayMode === "inline") {
    applyInlineBaseline(picture, session.exportResult, size);
  } else {
    styleDisplayParagraph(picture);
  }
}

function selectedPictureSnapshot(picture: Word.InlinePicture): SelectedWordFormula | null {
  const formulaId = formulaIdFromPictureTitle(picture.altTextTitle);
  if (!formulaId) return null;
  return {
    formulaId,
    encodedMetadata: picture.altTextDescription || null,
    width: picture.width,
    height: picture.height,
  };
}

async function readSelectedWordFormula(): Promise<SelectedWordFormula | null> {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();

    if (wordApiSupported("1.2")) {
      const selectedPictures = selection.inlinePictures;
      selectedPictures.load(
        "items/altTextDescription,items/altTextTitle,items/width,items/height",
      );
      await context.sync();
      const selected = selectedPictures.items
        .map(selectedPictureSnapshot)
        .find((item): item is SelectedWordFormula => Boolean(item));
      if (selected) return selected;
    }

    // Migration path for formulas created by VisualTeX 1.0.10 and earlier.
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
        formulaIdFromLegacyTag(item.tag) &&
        candidates.findIndex((candidate) => candidate.id === item.id) === index,
    );
    if (!control || control.title !== LEGACY_WORD_CONTENT_CONTROL_TITLE) {
      return null;
    }

    const formulaId = formulaIdFromLegacyTag(control.tag);
    if (!formulaId) return null;
    const pictures = control.getRange("Content").inlinePictures;
    pictures.load(
      "items/altTextDescription,items/altTextTitle,items/width,items/height",
    );
    await context.sync();
    const picture =
      pictures.items.find(
        (item) => item.altTextTitle === `${WORD_ALT_TEXT_TITLE_PREFIX}${formulaId}`,
      ) ?? pictures.items[0];
    if (!picture) return null;
    return {
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
    displayMode: metadata.displayMode,
    originalMetadata: metadata,
    autoCommitOnClose: true,
  };
}

async function removeLegacyContentControls(
  context: Word.RequestContext,
  formulaId: string,
) {
  const controls = context.document.contentControls.getByTag(
    legacyContentControlTag(formulaId),
  );
  controls.load("items/id");
  await context.sync();
  controls.items.forEach((control) => control.delete(true));
  if (controls.items.length) await context.sync();
}

async function findFormulaPicture(
  context: Word.RequestContext,
  formulaId: string,
): Promise<Word.InlinePicture | null> {
  if (wordApiSupported("1.2")) {
    const selected = context.document.getSelection().inlinePictures;
    selected.load("items/altTextTitle");
    await context.sync();
    const selectedPicture = selected.items.find(
      (item) => item.altTextTitle === `${WORD_ALT_TEXT_TITLE_PREFIX}${formulaId}`,
    );
    if (selectedPicture) return selectedPicture;
  }

  const pictures = context.document.body.inlinePictures;
  pictures.load("items/altTextTitle");
  await context.sync();
  return (
    pictures.items.find(
      (item) => item.altTextTitle === `${WORD_ALT_TEXT_TITLE_PREFIX}${formulaId}`,
    ) ?? null
  );
}

async function insertInlineWordFormula(
  context: Word.RequestContext,
  session: OfficeFormulaSession,
  metadata: VisualTeXFormulaMetadata,
  imageBase64: string,
) {
  const selection = context.document.getSelection();
  const picture = selection.insertInlinePictureFromBase64(imageBase64, "Replace");
  configurePicture(picture, session, metadata);
  picture.getRange("End").select("End");
  await context.sync();
}

async function insertDisplayWordFormula(
  context: Word.RequestContext,
  session: OfficeFormulaSession,
  metadata: VisualTeXFormulaMetadata,
  imageBase64: string,
) {
  if (!wordApiSupported("1.2")) {
    throw new Error("当前 Word 版本不支持创建独立的居中行间公式段落。");
  }
  const selection = context.document.getSelection();
  const selectedParagraph = selection.paragraphs.getFirst();
  selectedParagraph.load("text");
  await context.sync();
  const paragraph = selectedParagraph.text.trim()
    ? selection.insertParagraph("", "After")
    : selectedParagraph;
  paragraph.alignment = "Centered";
  paragraph.spaceBefore = 0;
  paragraph.spaceAfter = 0;
  const picture = paragraph.insertInlinePictureFromBase64(imageBase64, "Replace");
  configurePicture(picture, session, metadata);
  paragraph.getRange("After").select("End");
  await context.sync();
}

async function insertWordFormula(
  session: OfficeFormulaSession,
  metadata: VisualTeXFormulaMetadata,
  imageBase64: string,
) {
  return Word.run(async (context) => {
    if (session.displayMode === "block") {
      await insertDisplayWordFormula(context, session, metadata, imageBase64);
    } else {
      await insertInlineWordFormula(context, session, metadata, imageBase64);
    }
    await removeLegacyContentControls(context, metadata.formulaId);
  });
}

async function replaceWordFormula(
  session: OfficeFormulaSession,
  metadata: VisualTeXFormulaMetadata,
  imageBase64: string,
) {
  return Word.run(async (context) => {
    const original = await findFormulaPicture(context, metadata.formulaId);
    if (!original) {
      throw new Error(
        "找不到原 VisualTeX 公式图片。请重新选择该公式后再执行编辑。",
      );
    }

    const range = original.getRange("Whole");
    const picture = range.insertInlinePictureFromBase64(imageBase64, "Replace");
    configurePicture(picture, session, metadata);
    picture.getRange("End").select("End");
    await context.sync();
    await removeLegacyContentControls(context, metadata.formulaId);
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

  // PNG is the most reliable Word representation on macOS. SVG remains a
  // fallback for hosts that reject transparent PNG data.
  if (exportResult.pngBase64) {
    try {
      return await apply(session, metadata, exportResult.pngBase64);
    } catch (pngError) {
      try {
        return await apply(session, metadata, exportResult.svgBase64);
      } catch {
        throw pngError;
      }
    }
  }
  return apply(session, metadata, exportResult.svgBase64);
}

export class WordAdapter implements OfficeHostAdapter {
  readonly host = "word" as const;

  async readSelection(mode: OfficeSessionMode): Promise<OfficeSelectionContext> {
    requireWordApi();
    const sourceDocumentId = Office.context.document.url || null;
    if (mode === "create") {
      return {
        sourceDocumentId,
        sourceObjectId: null,
        sessionSeed: { displayMode: "inline" },
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
      sourceObjectId: selected.formulaId,
      sessionSeed: {
        ...createSessionSeed(metadata),
        exportWidth: selected.width / 0.75,
        exportHeight: selected.height / 0.75,
      },
    };
  }

  async applySession(session: OfficeFormulaSession): Promise<void> {
    requireWordApi();
    const metadata = createFormulaMetadata({
      formulaId: session.formulaId,
      title: session.title,
      lines: session.lines,
      codeFormat: session.codeFormat,
      displayMode: session.displayMode,
      original: session.originalMetadata,
    });
    await putCachedFormulaMetadata(metadata);
    await applyWithImageFallback(session, metadata);
    await writeWordDocumentMetadata(metadata);
  }

  async openDesktopApp(): Promise<void> {
    await revealDesktopApp();
  }

  showMessage(message: string) {
    const status = document.getElementById("bridge-status");
    if (status) status.textContent = message;
  }
}
