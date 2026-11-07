// VisualTeX document editor (docs/document-editor/DESIGN.md).
export { documentSchema, isFormulaNodeName, FONT_SIZES } from "./schema.ts";
export { idsPlugin, newNodeId, assignIds } from "./ids.ts";
export {
  closeFormula,
  FORMULA_PARTS,
  FORMULAS_REFRESH_EVENT,
  findNodeById,
  formulaKey,
  formulaParts,
  formulaViewOf,
  openFormula,
  openFormulaId,
  type FormulaPart,
  type FormulaEntry,
  type FormulaExit,
  type FormulaHost,
  type FormulaKey,
  type FormulaSelection,
  type FormulaSession,
  type FormulaSessionCallbacks,
} from "./formulas.ts";
export * from "./commands.ts";
export { createDocumentState, footnoteEditorPlugins, type FootnoteEditorOptions, insertEquationReference, numberedEquations, slicePlainText, documentPlugins, equationNumbers, isDocumentEmpty, mathTextToSlice, numberingKey } from "./plugins.ts";
export { hasMathDelimiters, parseMathText, splitInlineMath, type MathTextBlock } from "./mathText.ts";
export { documentToLatex, escapeLatexText, sliceToLatex, type LatexExportOptions } from "./latex.ts";
export { parseDocument, serializeDocument, DOCUMENT_FILE_FORMAT } from "./fileFormat.ts";
export * from "./blocks.ts";
export {
  commandMenuEntries,
  commandMenuKey,
  commandToken,
  dismissCommandMenu,
  DOCUMENT_COMMANDS,
  MARKDOWN_COMMANDS,
  highlightCommandEntry,
  matchCommands,
  runCommandEntry,
  startCommand,
  type CommandEntry,
  type CommandPreview,
  type CommandRun,
  type CommandToken,
  type MathCommandProvider,
  type MathCommandSuggestion,
} from "./suggest.ts";
export {
  CODE_LANGUAGES,
  codeBlockKeys,
  insertCodeBlock,
  type CodeBlockOptions,
  type CodeHighlighter,
  type CodeToken,
} from "./codeBlocks.ts";
export {
  documentToMarkdown,
  looksLikeMarkdown,
  looksLikeMarkdownDocument,
  markdownToDoc,
  markdownToSlice,
  sliceToMarkdown,
  MARKDOWN_HIGHLIGHT,
  markdownSourceKey,
  markdownSourceOf,
  readMarkdown,
  type MarkdownExportOptions,
  type MarkdownSource,
} from "./markdown.ts";
export { documentDragKey, documentDragPlugin } from "./documentDrag.ts";
export { latexToMarkdown, looksLikeLatexSource } from "./latexImport.ts";
export { decodeText, findMarkdownFile, readLatexProject, type LatexProject, type ProjectFile } from "./latexProject.ts";
export { LINK_REQUEST_EVENT, insertLink, linkAt, normalizeHref, removeLink, requestLink, setLink, type LinkRange } from "./links.ts";
export { markdownEnter, markdownHeadingEnter, markdownInputKey, markdownInputPlugin, markdownKeymap, type InputMode, type MarkdownRuleSwitches } from "./markdownInput.ts";
export {
  clearFormat,
  collapseTransaction,
  markdownLiveKey,
  markdownLivePlugin,
  markdownRedo,
  markdownUndo,
  shownMarkdown,
  sliceWithoutMarkdownSyntax,
  typoraFormat,
  typoraLink,
  withMarkdownHidden,
  withoutMarkdownSyntax,
} from "./markdownLive.ts";
export { parseInline, writeInline, type InlineKind, type InlineSpan } from "./inlineMarkdown.ts";
export { deleteSelectionWithParts, mixedSelectionPlugin, selectionHasFormulaParts, selectionSliceWithParts, toggleDocumentBold } from "./mixedSelection.ts";
export { SelectGesture } from "./formulas.ts";
export {
  CHINESE_FONT_SIZES,
  DEFAULT_DOCUMENT_SETTINGS,
  LINE_SPACINGS,
  PAPER_SIZES,
  MARGIN_PRESETS,
  MIN_TEXT_MM,
  POINT_FONT_SIZES,
  SINGLE_LINE_HEIGHT,
  clampMargin,
  type MarginName,
  bodyFontFamily,
  bracketLabel,
  equationLabels,
  type EquationBrackets,
  type EquationNumbering,
  type HeadingNumbering,
  chineseSizeName,
  chosenSettings,
  documentSettings,
  setDocumentSettings,
  settingsPreamble,
  textArea,
  type DocumentSettings,
  type PaperName,
} from "./settings.ts";
export {
  FOOTNOTE_EDIT_EVENT,
  CITATION_REQUEST_EVENT,
  addReference,
  addReferenceEntry,
  requestCitation,
  bibliographyEntries,
  citationKey,
  citeLabel,
  footnotes,
  footnoteEditorDoc,
  footnoteEditorText,
  insertCitation,
  insertFootnote,
  notesKey,
  requestFootnoteEdit,
  setFootnoteText,
  type BibEntryInfo,
  type FootnoteEditRequest,
} from "./notes.ts";
export { paginationKey, paginationPlugin, refreshPagination, type PageGeometry, type PaginationOptions } from "./pagination.ts";
export {
  DEFAULT_IMAGE_WIDTH,
  IMAGE_GAP,
  documentImageKeys,
  imageEntry,
  imageDataUrl,
  imageFiles,
  imageUrl,
  insertImagesAtSelection,
  isImageLine,
  registerImageBytes,
  registerImageFile,
  resolveImageSource,
  setImageAlign,
  setImageFileResolver,
  setImageWrap,
  type ImageLabels,
} from "./images.ts";
export { graphicsPaths } from "./latexImport.ts";
export { IMAGE_REQUEST_EVENT, requestImage } from "./images.ts";
export { formatPageNumber, hasPageNumbering, pageCommandLatex, pageLabels, parsePageCommand, PAGE_NUMBER_STYLES } from "./pageCommands.ts";
export { richBlockViews, richBlocksPlugin, safeHtml } from "./richBlocks.ts";
export { emojiOf } from "./emoji.ts";
export { boardCommand, boardFixes, boardHeight, boardItems, boardRows, canonicalRow, isBoard, isBoardImage, lineAsBoard, DEFAULT_BOARD_LABELS, type BoardLabels, type BoardCommand } from "./imageBoards.ts";
