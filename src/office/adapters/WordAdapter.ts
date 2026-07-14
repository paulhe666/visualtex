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
  OfficeInteractionTarget,
  OfficeSelectionContext,
} from "./OfficeHostAdapter";

const LEGACY_WORD_CONTENT_CONTROL_TITLE = "VisualTeX Formula";
const LEGACY_WORD_CONTENT_CONTROL_TAG_PREFIX = "visualtex:";
const WORD_ALT_TEXT_TITLE_PREFIX = "VisualTeX_";
const WORD_EQUATION_NUMBER_TAG_PREFIX = "visualtex-equation-number:";
const WORD_EQUATION_NUMBER_TITLE = "VisualTeX Equation Number";
const MAX_WORD_FORMULA_WIDTH_PT = 500;
const WORD_EQUATION_NUMBER_GUTTER_PT = 42;

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

interface EquationPictureRecord {
  picture: Word.InlinePicture;
  metadata: VisualTeXFormulaMetadata;
  parentTable: Word.Table;
}

interface EquationNumberRecord {
  control: Word.ContentControl;
  table: Word.Table;
  pictureCollection: Word.InlinePictureCollection;
  pictures: Word.InlinePicture[];
  used: boolean;
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

function equationNumberTag(formulaId: string) {
  return `${WORD_EQUATION_NUMBER_TAG_PREFIX}${formulaId}`;
}

function formulaIdFromEquationNumberTag(tag: string) {
  if (!tag.startsWith(WORD_EQUATION_NUMBER_TAG_PREFIX)) return null;
  return validFormulaId(tag.slice(WORD_EQUATION_NUMBER_TAG_PREFIX.length));
}

export function equationNumberLabel(index: number) {
  if (!Number.isSafeInteger(index) || index < 1) {
    throw new Error("Equation number must be a positive integer.");
  }
  return `(${index})`;
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
  maximumWidthPt = MAX_WORD_FORMULA_WIDTH_PT,
): AppliedPictureSize {
  if (!(widthPx > 0) || !(heightPx > 0)) {
    return { width: 0, height: 0, scale: 1 };
  }
  const naturalWidthPt = widthPx * 0.75;
  const naturalHeightPt = heightPx * 0.75;
  const scale = Math.min(1, maximumWidthPt / naturalWidthPt);
  // Both dimensions always use exactly the same factor. Never clamp width and
  // height independently: doing so distorts short/tall formulas, and changing
  // a formula's length can otherwise look like its glyphs were compressed.
  const width = naturalWidthPt * scale;
  const height = naturalHeightPt * scale;
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
  exportResult: OfficeFormulaSession["exportResult"],
  size: AppliedPictureSize,
) {
  // VisualTeX targets desktop Word on Windows and macOS. Some Mac builds
  // incorrectly report WordApiDesktop 1.3 as unsupported even though
  // Range.font.position is available, which previously skipped the baseline
  // write entirely. Apply the same formula used by the Windows integration
  // and let invalid/missing export metadata resolve safely to position 0.
  picture.getRange("Whole").font.position = calculateInlineFormulaPosition(
    size.height,
    exportResult?.height,
    exportResult?.baseline,
  );
}

/** Align the mathematical baseline in the exported canvas with Word's text
 * baseline. Word aligns the image's bottom edge by default, so the rendered
 * descender below MathJax's baseline must be moved below Word's baseline. */
export function calculateInlineFormulaPosition(
  actualHeightPoints: number,
  exportedHeight: number | undefined,
  exportedBaseline: number | undefined,
) {
  if (
    !(actualHeightPoints > 0) ||
    !(exportedHeight && exportedHeight > 0) ||
    exportedBaseline === undefined ||
    exportedBaseline < 0 ||
    exportedBaseline > exportedHeight
  ) {
    return 0;
  }
  const descentRatio = (exportedHeight - exportedBaseline) / exportedHeight;
  const downwardShiftPoints = actualHeightPoints * descentRatio;
  if (!(downwardShiftPoints > 0) || !Number.isFinite(downwardShiftPoints)) {
    return 0;
  }
  return -Math.max(0, Math.round(downwardShiftPoints));
}

/** Calculate the exact offset used after Word has scaled the exported image.
 * Kept as a pure helper so the macOS native fallback can write the same value
 * after Office.js has materialized the inline picture. */
export function calculateInlineSessionPosition(session: OfficeFormulaSession) {
  const widthPx = session.exportResult?.width ?? session.exportWidth;
  const heightPx = session.exportResult?.height ?? session.exportHeight;
  if (!(widthPx > 0) || !(heightPx > 0)) return 0;
  const naturalWidthPt = widthPx * 0.75;
  const naturalHeightPt = heightPx * 0.75;
  const scale = Math.min(1, MAX_WORD_FORMULA_WIDTH_PT / naturalWidthPt);
  return calculateInlineFormulaPosition(
    naturalHeightPt * scale,
    session.exportResult?.height,
    session.exportResult?.baseline,
  );
}

function styleDisplayParagraph(picture: Word.InlinePicture) {
  const paragraph = picture.getRange("Whole").paragraphs.getFirst();
  paragraph.alignment = "Centered";
  paragraph.spaceBefore = 0;
  paragraph.spaceAfter = 0;
}

function styleEquationTable(table: Word.Table) {
  table.alignment = "Centered";
  table.verticalAlignment = "Center";
  table.getBorder("All").type = "None";
  table.setCellPadding("Top", 0);
  table.setCellPadding("Bottom", 0);
  table.setCellPadding("Left", 0);
  table.setCellPadding("Right", 0);
  table.autoFitWindow();

  const leftCell = table.getCell(0, 0);
  const formulaCell = table.getCell(0, 1);
  const numberCell = table.getCell(0, 2);
  leftCell.horizontalAlignment = "Left";
  formulaCell.horizontalAlignment = "Centered";
  numberCell.horizontalAlignment = "Right";
  leftCell.verticalAlignment = "Center";
  formulaCell.verticalAlignment = "Center";
  numberCell.verticalAlignment = "Center";
}

function setEquationTableColumnWidths(table: Word.Table, tableWidth: number) {
  if (!(tableWidth > WORD_EQUATION_NUMBER_GUTTER_PT * 2 + 40)) {
    return MAX_WORD_FORMULA_WIDTH_PT;
  }
  const gutter = Math.min(
    WORD_EQUATION_NUMBER_GUTTER_PT,
    tableWidth * 0.12,
  );
  const formulaWidth = tableWidth - gutter * 2;
  table.getCell(0, 0).columnWidth = gutter;
  table.getCell(0, 1).columnWidth = formulaWidth;
  table.getCell(0, 2).columnWidth = gutter;
  return Math.max(40, formulaWidth - 4);
}

function createEquationNumberControl(
  table: Word.Table,
  formulaId: string,
  label: string,
) {
  const numberCell = table.getCell(0, 2);
  numberCell.value = label;
  const numberParagraph = numberCell.body.paragraphs.getFirst();
  numberParagraph.alignment = "Right";
  numberParagraph.spaceBefore = 0;
  numberParagraph.spaceAfter = 0;
  const control = numberCell.body
    .getRange("Content")
    .insertContentControl("PlainText");
  control.title = WORD_EQUATION_NUMBER_TITLE;
  control.tag = equationNumberTag(formulaId);
  control.appearance = "Hidden";
  return control;
}

async function insertNumberedDisplayWordFormula(
  context: Word.RequestContext,
  selectedParagraph: Word.Paragraph,
  paragraphIsEmpty: boolean,
  session: OfficeFormulaSession,
  metadata: VisualTeXFormulaMetadata,
  imageBase64: string,
) {
  const table = selectedParagraph.insertTable(
    1,
    3,
    paragraphIsEmpty ? "Before" : "After",
    [["", "", equationNumberLabel(1)]],
  );
  styleEquationTable(table);
  table.load("width");
  await context.sync();

  const maximumFormulaWidth = setEquationTableColumnWidths(table, table.width);
  const picture = table
    .getCell(0, 1)
    .body.insertInlinePictureFromBase64(imageBase64, "Start");
  configurePicture(picture, session, metadata, maximumFormulaWidth);
  createEquationNumberControl(
    table,
    metadata.formulaId,
    equationNumberLabel(1),
  );
  table.getRange("After").select("End");
  await context.sync();
}

function configurePicture(
  picture: Word.InlinePicture,
  session: OfficeFormulaSession,
  metadata: VisualTeXFormulaMetadata,
  maximumWidthPt = MAX_WORD_FORMULA_WIDTH_PT,
) {
  applyPictureMetadata(picture, metadata);
  const size = applyPictureSize(
    picture,
    session.exportResult?.width ?? session.exportWidth,
    session.exportResult?.height ?? session.exportHeight,
    maximumWidthPt,
  );
  if (session.displayMode !== "inline") {
    styleDisplayParagraph(picture);
  }
  return size;
}

function selectAfterInlinePicture(picture: Word.InlinePicture) {
  const caret = picture.getRange("End");
  // A collapsed range otherwise inherits the picture range's negative baseline
  // position. Resetting the insertion format prevents Return from carrying the
  // formula object/formatting onto the next line in affected Word for Mac builds.
  caret.font.position = 0;
  caret.select("End");
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

function pictureMetadata(picture: Word.InlinePicture) {
  const formulaId = formulaIdFromPictureTitle(picture.altTextTitle);
  if (!formulaId || !picture.altTextDescription) return null;
  const metadata = decodeFormulaMetadata(picture.altTextDescription);
  return metadata?.formulaId === formulaId ? metadata : null;
}

function duplicateMetadataWithFreshFormulaId(
  metadata: VisualTeXFormulaMetadata,
) {
  const bytes = crypto.getRandomValues(new Uint8Array(16));
  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;
  const hex = Array.from(bytes, (value) => value.toString(16).padStart(2, "0"));
  const formulaId = [
    hex.slice(0, 4).join(""),
    hex.slice(4, 6).join(""),
    hex.slice(6, 8).join(""),
    hex.slice(8, 10).join(""),
    hex.slice(10, 16).join(""),
  ].join("-");
  return {
    ...metadata,
    formulaId,
    updatedAt: new Date().toISOString(),
  } satisfies VisualTeXFormulaMetadata;
}

function wordFormulaFromInteractionTarget(
  target: OfficeInteractionTarget,
): SelectedWordFormula | null {
  if (target.host !== "word") return null;
  const metadata = decodeFormulaMetadata(target.shapeName);
  if (
    !metadata ||
    !Number.isFinite(target.width) ||
    !Number.isFinite(target.height) ||
    (target.width ?? 0) <= 0 ||
    (target.height ?? 0) <= 0
  ) {
    return null;
  }
  return {
    formulaId: metadata.formulaId,
    encodedMetadata: target.shapeName,
    width: target.width as number,
    height: target.height as number,
  };
}

async function readWordFormulaByMarker(
  encodedMetadata: string,
): Promise<SelectedWordFormula | null> {
  const metadata = decodeFormulaMetadata(encodedMetadata);
  if (!metadata) return null;

  return Word.run(async (context) => {
    const pictures = context.document.body.inlinePictures;
    pictures.load(
      "items/altTextDescription,items/altTextTitle,items/width,items/height",
    );
    await context.sync();
    const picture = pictures.items.find(
      (candidate) =>
        candidate.altTextDescription === encodedMetadata &&
        candidate.altTextTitle ===
          `${WORD_ALT_TEXT_TITLE_PREFIX}${metadata.formulaId}`,
    );
    return picture ? selectedPictureSnapshot(picture) : null;
  });
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
    const snapshot = {
      formulaId,
      encodedMetadata: picture.altTextDescription || null,
      width: picture.width,
      height: picture.height,
    };
    // Old releases wrapped inline formulas in a rich-text content control.
    // Word expands that wrapper when Return is pressed and can repeat the image
    // on the following line. Unwrap it as soon as it is selected for editing.
    control.delete(true);
    await context.sync();
    return snapshot;
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
    numbered: metadata.displayMode === "block" && Boolean(metadata.numbered),
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
  const size = configurePicture(picture, session, metadata);
  // Word for Mac can discard Font.position when it is assigned in the same
  // request batch that creates the inline picture. Materialize the picture
  // first, then apply the baseline shift to its durable range.
  await context.sync();
  applyInlineBaseline(picture, session.exportResult, size);
  selectAfterInlinePicture(picture);
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
  const paragraphIsEmpty = selectedParagraph.text.trim().length === 0;
  if (metadata.numbered) {
    if (!wordApiSupported("1.3")) {
      throw new Error("当前 Word 版本不支持 VisualTeX 行间公式编号排版。");
    }
    await insertNumberedDisplayWordFormula(
      context,
      selectedParagraph,
      paragraphIsEmpty,
      session,
      metadata,
      imageBase64,
    );
    return;
  }

  const paragraph = !paragraphIsEmpty
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

    let parentTable: Word.Table | null = null;
    let maximumFormulaWidth = MAX_WORD_FORMULA_WIDTH_PT;
    if (
      metadata.displayMode === "block" &&
      (metadata.numbered || session.originalMetadata?.numbered)
    ) {
      const candidate = original.parentTableOrNullObject;
      candidate.load("isNullObject");
      await context.sync();
      if (!candidate.isNullObject) {
        const formulaCell = candidate.getCell(0, 1);
        formulaCell.load("width");
        await context.sync();
        parentTable = candidate;
        maximumFormulaWidth = Math.max(40, formulaCell.width - 4);
      }
    }

    const range = original.getRange("Whole");
    const picture = range.insertInlinePictureFromBase64(imageBase64, "Replace");
    const size = configurePicture(
      picture,
      session,
      metadata,
      maximumFormulaWidth,
    );
    if (session.displayMode === "inline") {
      // See insertInlineWordFormula: the Mac host only persists the run-level
      // baseline shift after the replacement picture has been synchronized.
      await context.sync();
      applyInlineBaseline(picture, session.exportResult, size);
    }
    if (session.displayMode === "inline") {
      selectAfterInlinePicture(picture);
    } else if (parentTable) {
      parentTable.getRange("After").select("End");
    } else {
      picture.getRange("End").select("End");
    }
    await context.sync();
    await removeLegacyContentControls(context, metadata.formulaId);
  });
}

function canDeleteMigratedEquationParagraph(text: string) {
  const visible = text.replace(/[\r\a\t\s]/g, "");
  return visible.length === 0 || /^\(\d+\)$/.test(visible);
}

async function migrateUnscaffoldedNumberedPictures(
  context: Word.RequestContext,
  records: EquationPictureRecord[],
) {
  if (!records.length) return false;

  const snapshots = records.map((record) => {
    const paragraph = record.picture.paragraph;
    const image = record.picture.getBase64ImageSrc();
    const font = record.picture.getRange("Whole").font;
    paragraph.load("text");
    if (wordDesktopApiSupported("1.3")) font.load("position");
    return { ...record, paragraph, image, font, table: null as Word.Table | null };
  });
  await context.sync();

  for (const snapshot of snapshots) {
    const table = snapshot.paragraph.insertTable(1, 3, "After", [
      ["", "", equationNumberLabel(1)],
    ]);
    styleEquationTable(table);
    table.load("width");
    snapshot.table = table;
  }
  await context.sync();

  for (const snapshot of snapshots) {
    const table = snapshot.table;
    if (!table) continue;
    const maximumWidth = setEquationTableColumnWidths(table, table.width);
    const scale = Math.min(1, maximumWidth / snapshot.picture.width);
    const replacement = table
      .getCell(0, 1)
      .body.insertInlinePictureFromBase64(snapshot.image.value, "Start");
    applyPictureMetadata(replacement, snapshot.metadata);
    replacement.lockAspectRatio = true;
    replacement.width = snapshot.picture.width * scale;
    replacement.height = snapshot.picture.height * scale;
    if (wordDesktopApiSupported("1.3")) {
      replacement.getRange("Whole").font.position = Math.round(
        snapshot.font.position * scale,
      );
    }
    styleDisplayParagraph(replacement);
    createEquationNumberControl(
      table,
      snapshot.metadata.formulaId,
      equationNumberLabel(1),
    );
  }
  // Make every replacement durable before removing the copied/original image.
  await context.sync();

  for (const snapshot of snapshots) {
    snapshot.picture.delete();
    if (canDeleteMigratedEquationParagraph(snapshot.paragraph.text)) {
      snapshot.paragraph.delete();
    }
  }
  await context.sync();
  return true;
}

async function reconcileWordEquationNumbers(
  context: Word.RequestContext,
  allowMigration: boolean,
) {
  const pictures = context.document.body.inlinePictures;
  const controls = context.document.contentControls;
  pictures.load(
    "items/altTextDescription,items/altTextTitle,items/width,items/height",
  );
  controls.load("items/tag,items/title");
  await context.sync();

  for (const control of controls.items) {
    if (
      control.title === LEGACY_WORD_CONTENT_CONTROL_TITLE &&
      formulaIdFromLegacyTag(control.tag)
    ) {
      control.delete(true);
    }
  }

  const formulaRecords = pictures.items
    .map((picture) => {
      const metadata = pictureMetadata(picture);
      if (!metadata) return null;
      const parentTable = picture.parentTableOrNullObject;
      parentTable.load("isNullObject");
      return { picture, metadata, parentTable };
    })
    .filter((record): record is EquationPictureRecord => Boolean(record));
  const deduplicatedMetadata: VisualTeXFormulaMetadata[] = [];
  const seenFormulaIds = new Set<string>();
  for (const record of formulaRecords) {
    if (seenFormulaIds.has(record.metadata.formulaId)) {
      record.metadata = duplicateMetadataWithFreshFormulaId(record.metadata);
      applyPictureMetadata(record.picture, record.metadata);
      deduplicatedMetadata.push(record.metadata);
    }
    seenFormulaIds.add(record.metadata.formulaId);
  }
  const numberControls = controls.items.filter(
    (control) =>
      control.title === WORD_EQUATION_NUMBER_TITLE &&
      Boolean(formulaIdFromEquationNumberTag(control.tag)),
  );
  const controlTables = numberControls.map((control) => {
    const table = control.parentTableOrNullObject;
    table.load("isNullObject");
    return { control, table };
  });
  await context.sync();

  const numberRecords: EquationNumberRecord[] = [];
  for (const record of controlTables) {
    if (record.table.isNullObject) {
      record.control.delete(false);
      continue;
    }
    const tablePictures = record.table.getRange("Whole").inlinePictures;
    tablePictures.load("items/altTextDescription,items/altTextTitle");
    numberRecords.push({
      control: record.control,
      table: record.table,
      pictureCollection: tablePictures,
      // Office.js proxy collections throw PropertyNotLoaded when `.items` is
      // touched before the request context has synchronized the load above.
      pictures: [],
      used: false,
    });
  }
  await context.sync();

  // Collection items become available only after the preceding sync.
  for (const record of numberRecords) {
    record.pictures = record.pictureCollection.items;
  }

  const unscaffolded: EquationPictureRecord[] = [];
  const assignedNumberRecords: Array<{
    record: EquationNumberRecord;
    formulaId: string;
    label: string;
  }> = [];
  let assigned = 0;
  const scaffoldedFormulaIds = new Set<string>();

  // Word for Mac doesn't consistently include pictures nested inside table
  // cells in `document.body.inlinePictures`. Numbered equations already have
  // an explicit table/control scaffold, so enumerate those table collections
  // directly and keep their document order from the content-control collection.
  for (const numberRecord of numberRecords) {
    let metadata = numberRecord.pictures
      .map(pictureMetadata)
      .find((candidate) => candidate?.displayMode === "block" && candidate.numbered);
    if (!metadata) continue;

    const existingFormulaId = metadata.formulaId;
    const picture = numberRecord.pictures.find(
      (candidate) => pictureMetadata(candidate)?.formulaId === existingFormulaId,
    );
    if (scaffoldedFormulaIds.has(metadata.formulaId) && picture) {
      metadata = duplicateMetadataWithFreshFormulaId(metadata);
      applyPictureMetadata(picture, metadata);
      deduplicatedMetadata.push(metadata);
    }

    scaffoldedFormulaIds.add(metadata.formulaId);
    numberRecord.used = true;
    assigned += 1;
    assignedNumberRecords.push({
      record: numberRecord,
      formulaId: metadata.formulaId,
      label: equationNumberLabel(assigned),
    });
  }

  for (const record of formulaRecords) {
    if (record.metadata.displayMode !== "block" || !record.metadata.numbered) {
      continue;
    }
    if (scaffoldedFormulaIds.has(record.metadata.formulaId)) continue;
    if (allowMigration && record.parentTable.isNullObject) {
      unscaffolded.push(record);
    }
  }

  for (const record of numberRecords.filter((candidate) => !candidate.used)) {
    const tableMetadata = record.pictures
      .map(pictureMetadata)
      .filter((metadata): metadata is VisualTeXFormulaMetadata => Boolean(metadata));
    if (!tableMetadata.length) {
      record.table.delete();
      continue;
    }
    if (!tableMetadata.some((metadata) => metadata.numbered)) {
      record.control.delete(false);
      record.table.getCell(0, 2).value = "";
    }
  }

  // Rebuild assigned controls instead of replacing their content in place.
  // Word for Mac can leave a cell-wide PlainText control visually unchanged
  // when its Content range is replaced, and copied equations can retain a
  // stale formula id in the control tag. Recreating the tiny hidden control
  // updates the visible label and repairs the tag in one deterministic pass.
  for (const { record } of assignedNumberRecords) {
    record.control.delete(false);
  }
  await context.sync();

  for (const { record, formulaId, label } of assignedNumberRecords) {
    createEquationNumberControl(record.table, formulaId, label);
  }
  await context.sync();

  const migrated = await migrateUnscaffoldedNumberedPictures(
    context,
    unscaffolded,
  );
  return { assigned, migrated, deduplicatedMetadata };
}

/** Refresh all VisualTeX display-equation numbers in document order. Copied
 * numbered pictures that lost their layout scaffold are migrated into the
 * same centered, borderless three-column structure before a second pass. */
export async function refreshWordEquationNumbers() {
  requireWordApi();
  if (!wordApiSupported("1.3")) {
    throw new Error("当前 Word 版本不支持 VisualTeX 公式编号重排。");
  }
  const first = await Word.run((context) =>
    reconcileWordEquationNumbers(context, true),
  );
  await Promise.all(
    first.deduplicatedMetadata.map((metadata) =>
      putCachedFormulaMetadata(metadata).catch(() => undefined),
    ),
  );
  if (!first.migrated) return first.assigned;
  const second = await Word.run((context) =>
    reconcileWordEquationNumbers(context, false),
  );
  await Promise.all(
    second.deduplicatedMetadata.map((metadata) =>
      putCachedFormulaMetadata(metadata).catch(() => undefined),
    ),
  );
  return second.assigned;
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
  /**
   * A Mac Word commit has a second, native baseline step after Office.js has
   * made the picture durable. If that native step fails, retrying the same
   * Session must not insert/replace the picture a second time. The bridge and
   * adapter live for the lifetime of the add-in command, so the Session id is
   * the correct idempotency key here.
   */
  private readonly imageAppliedSessionIds = new Set<string>();
  private readonly nativeFormulaMarkers = new Map<string, string>();
  private pendingInteractionTarget: OfficeInteractionTarget | null = null;

  prepareInteractionTarget(target: OfficeInteractionTarget) {
    if (target.host !== "word") return;
    this.pendingInteractionTarget = target;
  }

  getNativeWordFormulaMarker(sessionId: string) {
    return this.nativeFormulaMarkers.get(sessionId) ?? null;
  }

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

    const interactionTarget = this.pendingInteractionTarget;
    this.pendingInteractionTarget = null;
    const interactionFormula = interactionTarget
      ? wordFormulaFromInteractionTarget(interactionTarget)
      : null;
    const selected = interactionFormula
      ? interactionFormula
      : interactionTarget?.shapeName
        ? await readWordFormulaByMarker(interactionTarget.shapeName)
        : await readSelectedWordFormula();
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
      numbered: session.displayMode === "block" && Boolean(session.numbered),
      renderWidthPx: session.exportResult?.width ?? session.exportWidth,
      renderHeightPx: session.exportResult?.height ?? session.exportHeight,
      original: session.originalMetadata,
    });
    await putCachedFormulaMetadata(metadata);
    if (!this.imageAppliedSessionIds.has(session.id)) {
      await applyWithImageFallback(session, metadata);
      this.imageAppliedSessionIds.add(session.id);
      // This is byte-for-byte the payload written to altTextDescription by
      // applyPictureMetadata. The macOS native verifier can therefore locate
      // the durable picture without depending on Word's mutable insertion point.
      this.nativeFormulaMarkers.set(session.id, encodeFormulaMetadata(metadata));
    }
    await writeWordDocumentMetadata(metadata);
    if (metadata.numbered || session.originalMetadata?.numbered) {
      await refreshWordEquationNumbers();
    }
  }

  async updateEquationNumbers(): Promise<number> {
    return refreshWordEquationNumbers();
  }

  async openDesktopApp(): Promise<void> {
    await revealDesktopApp();
  }

  showMessage(message: string) {
    const status = document.getElementById("bridge-status");
    if (status) status.textContent = message;
  }
}
