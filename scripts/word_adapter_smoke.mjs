import assert from "node:assert/strict";
import { createFormulaMetadata } from "../src/office/metadata/formulaMetadata.ts";

let cachedMetadata = null;
let rejectPngOnce = false;
const allPictures = [];
const paragraphs = [];
const tables = [];
const legacyControls = [];
const equationNumberControls = [];
let selectedPictures = [];
let lastCaretFontPosition = null;
let wordRunCount = 0;

class FakeCollection {
  constructor(itemsProvider) {
    this.itemsProvider = itemsProvider;
  }
  get items() {
    return this.itemsProvider();
  }
  load() {
    return this;
  }
}

class FakeParagraph {
  constructor(text = "") {
    this.text = text;
    this.alignment = "Left";
    this.spaceBefore = 0;
    this.spaceAfter = 0;
    this.pictures = [];
  }
  load() {
    return this;
  }
  getRange() {
    return new FakeRange({ paragraph: this });
  }
  insertInlinePictureFromBase64(base64, location) {
    if (location === "Replace") {
      for (const picture of [...this.pictures]) picture.remove();
    }
    const picture = new FakePicture(base64, this);
    this.pictures.push(picture);
    allPictures.push(picture);
    selectedPictures = [picture];
    return picture;
  }
  insertTable(_rowCount, _columnCount, _location, values) {
    const table = new FakeTable(values);
    tables.push(table);
    return table;
  }
  delete() {
    const index = paragraphs.indexOf(this);
    if (index >= 0) paragraphs.splice(index, 1);
  }
}

class FakePicture {
  constructor(base64, paragraph, table = null) {
    this.base64 = base64;
    this.paragraph = paragraph;
    this.altTextTitle = "";
    this.altTextDescription = "";
    this.width = 0;
    this.height = 0;
    this.lockAspectRatio = false;
    this.fontPosition = 0;
    this.table = table;
  }
  get parentTableOrNullObject() {
    return this.table ?? new FakeNullTable();
  }
  getRange(location = "Whole") {
    if (location === "End") {
      return new FakeRange({ paragraph: this.paragraph, caret: true });
    }
    return new FakeRange({ picture: this, paragraph: this.paragraph });
  }
  remove() {
    const globalIndex = allPictures.indexOf(this);
    if (globalIndex >= 0) allPictures.splice(globalIndex, 1);
    const paragraphIndex = this.paragraph.pictures.indexOf(this);
    if (paragraphIndex >= 0) this.paragraph.pictures.splice(paragraphIndex, 1);
    selectedPictures = selectedPictures.filter((picture) => picture !== this);
  }
  delete() {
    this.remove();
  }
  getBase64ImageSrc() {
    return { value: this.base64 };
  }
}

class FakeRange {
  constructor({
    picture = null,
    paragraph = null,
    table = null,
    control = null,
    cell = null,
    caret = false,
  } = {}) {
    this.picture = picture;
    this.paragraph = paragraph ?? paragraphs[0];
    this.table = table;
    this.control = control;
    this.cell = cell;
    this.caret = caret;
    this.inlinePictures = new FakeCollection(() =>
      picture
        ? [picture]
        : table
          ? table.pictures
          : selectedPictures,
    );
    this.contentControls = new FakeCollection(() => []);
    this.parentContentControlOrNullObject = {
      isNullObject: true,
      load() {
        return this;
      },
    };
    this.paragraphs = {
      getFirst: () => this.paragraph,
    };
    this.font = {};
    Object.defineProperty(this.font, "position", {
      get: () => this.picture?.fontPosition ?? (this.caret ? lastCaretFontPosition ?? 0 : 0),
      set: (value) => {
        if (this.picture) this.picture.fontPosition = value;
        if (this.caret) lastCaretFontPosition = value;
      },
    });
    this.font.load = () => this.font;
  }
  insertInlinePictureFromBase64(base64, location) {
    if (rejectPngOnce && base64 === "png-base64") {
      rejectPngOnce = false;
      throw new Error("PNG unsupported by simulated Word host");
    }
    if (location === "Replace" && this.picture) {
      const paragraph = this.picture.paragraph;
      this.picture.remove();
      const replacement = new FakePicture(base64, paragraph, this.picture.table);
      paragraph.pictures.push(replacement);
      allPictures.push(replacement);
      selectedPictures = [replacement];
      return replacement;
    }
    if (location === "Replace") {
      for (const picture of [...selectedPictures]) picture.remove();
    }
    const picture = new FakePicture(base64, this.paragraph);
    this.paragraph.pictures.push(picture);
    allPictures.push(picture);
    selectedPictures = [picture];
    return picture;
  }
  insertParagraph(text, location) {
    assert.equal(location, "After");
    const paragraph = new FakeParagraph(text);
    paragraphs.push(paragraph);
    return paragraph;
  }
  insertContentControl() {
    assert.ok(this.cell, "number controls must be created inside a table cell");
    const control = new FakeEquationNumberControl(this.cell.table, this.cell);
    equationNumberControls.push(control);
    return control;
  }
  insertText(text, location) {
    assert.equal(location, "Replace");
    if (this.control) {
      this.control.text = text;
      this.control.cell.value = text;
    }
    return this;
  }
  select() {}
}

class FakeNullTable {
  constructor() {
    this.isNullObject = true;
  }
  load() {
    return this;
  }
}

class FakeBody {
  constructor(cell) {
    this.cell = cell;
    this.paragraph = new FakeParagraph("");
    this.paragraphs = { getFirst: () => this.paragraph };
    this.contentControls = new FakeCollection(() =>
      equationNumberControls.filter(
        (control) => !control.deleted && control.cell === this.cell,
      ),
    );
  }
  insertInlinePictureFromBase64(base64) {
    const picture = new FakePicture(base64, this.paragraph, this.cell.table);
    this.paragraph.pictures.push(picture);
    allPictures.push(picture);
    selectedPictures = [picture];
    return picture;
  }
  getRange() {
    return new FakeRange({
      paragraph: this.paragraph,
      table: this.cell.table,
      cell: this.cell,
    });
  }
}

class FakeCell {
  constructor(table, index, value = "") {
    this.table = table;
    this.index = index;
    this._value = value;
    this.body = new FakeBody(this);
    this.horizontalAlignment = "Left";
    this.verticalAlignment = "Top";
    this.columnWidth = 0;
  }
  get value() {
    return this._value;
  }
  set value(value) {
    this._value = value;
    this.body.paragraph.text = value;
  }
}

class FakeTable {
  constructor(values = [["", "", ""]]) {
    this.isNullObject = false;
    this.width = 468;
    this.cells = [0, 1, 2].map(
      (index) => new FakeCell(this, index, values?.[0]?.[index] ?? ""),
    );
    this.deleted = false;
  }
  get pictures() {
    return this.cells.flatMap((cell) => cell.body.paragraph.pictures);
  }
  load() {
    return this;
  }
  getCell(_row, column) {
    return this.cells[column];
  }
  getBorder() {
    return { type: "Single" };
  }
  setCellPadding() {}
  autoFitWindow() {}
  getRange(location = "Whole") {
    return new FakeRange({ table: this, caret: location === "After" });
  }
  delete() {
    this.deleted = true;
    for (const picture of [...this.pictures]) picture.remove();
    for (const control of equationNumberControls.filter(
      (item) => item.table === this,
    )) {
      control.deleted = true;
    }
  }
}

class FakeLegacyControl {
  constructor(formulaId) {
    this.id = 1;
    this.tag = `visualtex:${formulaId}`;
    this.title = "VisualTeX Formula";
    this.deleted = false;
  }
  load() {
    return this;
  }
  delete(keepContent) {
    assert.equal(keepContent, true);
    this.deleted = true;
  }
  get parentTableOrNullObject() {
    return new FakeNullTable();
  }
}

class FakeEquationNumberControl {
  constructor(table, cell) {
    this.table = table;
    this.cell = cell;
    this.title = "";
    this.tag = "";
    this.text = cell.value;
    this.deleted = false;
  }
  get parentTableOrNullObject() {
    return this.table;
  }
  getRange() {
    return new FakeRange({
      table: this.table,
      control: this,
      cell: this.cell,
    });
  }
  delete(keepContent) {
    this.deleted = true;
    if (!keepContent) this.cell.value = "";
  }
}

class FakeContentControlCollection extends FakeCollection {
  constructor() {
    super(() => [
      ...legacyControls.filter((control) => !control.deleted),
      ...equationNumberControls.filter((control) => !control.deleted),
    ]);
  }
  getByTag(tag) {
    return new FakeCollection(() =>
      this.items.filter((control) => control.tag === tag),
    );
  }
}

const initialParagraph = new FakeParagraph("");
paragraphs.push(initialParagraph);
const selection = new FakeRange({ paragraph: initialParagraph });

const fakeDocument = {
  getSelection() {
    selection.paragraph = selectedPictures[0]?.paragraph ?? selection.paragraph;
    return selection;
  },
  body: {
    inlinePictures: new FakeCollection(() => allPictures),
  },
  contentControls: new FakeContentControlCollection(),
};

const context = {
  document: fakeDocument,
  async sync() {},
};

globalThis.window = {
  __VISUALTEX_INSTALL_TOKEN__: "test-install-token",
  location: { href: "" },
};
globalThis.document = {
  querySelector() {
    return null;
  },
  getElementById() {
    return null;
  },
};

globalThis.Office = {
  HostType: { Word: "Word", PowerPoint: "PowerPoint" },
  AsyncResultStatus: { Succeeded: "succeeded" },
  context: {
    document: {
      url: "file:///tmp/word-adapter-smoke.docx",
      customXmlParts: null,
    },
    requirements: {
      isSetSupported(name, version) {
        if (name === "CustomXmlParts") return false;
        if (name === "WordApi") return Number(version) <= 1.6;
        // Reproduce Mac Word builds that under-report WordApiDesktop 1.3 even
        // though Range.font.position is available and writable.
        if (name === "WordApiDesktop") return false;
        return false;
      },
    },
  },
};

globalThis.Word = {
  async run(callback) {
    wordRunCount += 1;
    return callback(context);
  },
};

globalThis.fetch = async (input, init = {}) => {
  const url = String(input);
  if (url.endsWith("/api/v1/app/reveal")) {
    return new Response(null, { status: 204 });
  }
  if (!url.includes("/api/v1/formulas/")) {
    throw new Error(`Unexpected fetch: ${url}`);
  }
  if (init.method === "PUT") {
    cachedMetadata = JSON.parse(init.body);
    return new Response(JSON.stringify(cachedMetadata), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  }
  if (cachedMetadata) {
    return new Response(JSON.stringify(cachedMetadata), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  }
  return new Response("", { status: 404 });
};

const {
  WordAdapter,
  calculateInlineFormulaPosition,
  calculateInlineSessionPosition,
  equationNumberLabel,
} = await import("../src/office/adapters/WordAdapter.ts");
const adapter = new WordAdapter();

assert.equal(calculateInlineFormulaPosition(15, 20, 15), -4);
assert.equal(calculateInlineFormulaPosition(30, 40, 30), -8);
assert.equal(calculateInlineFormulaPosition(20, 20, 20), 0);
assert.equal(calculateInlineFormulaPosition(20, 20, undefined), 0);
assert.equal(equationNumberLabel(12), "(12)");
assert.throws(() => equationNumberLabel(0));

function createSession(formulaId, lineId, displayMode = "inline") {
  return {
    id: crypto.randomUUID(),
    mode: "create",
    host: "word",
    formulaId,
    sourceDocumentId: null,
    sourceObjectId: null,
    title: "Word Formula",
    lines: [{ id: lineId, latex: String.raw`\frac{a}{b}+\text{测试}` }],
    activeLineId: lineId,
    codeFormat: "raw",
    displayMode,
    numbered: false,
    exportWidth: 120,
    exportHeight: 48,
    exportResult: {
      svg: "<svg></svg>",
      svgBase64: "svg-base64",
      pngBase64: "png-base64",
      width: 120,
      height: 48,
      baseline: 36,
    },
    originalMetadata: null,
    dirty: true,
    status: "committing",
    autoCommitOnClose: true,
    explicitCancel: false,
    error: null,
    createdAt: Date.now(),
    updatedAt: Date.now(),
    expiresAt: Date.now() + 1000,
  };
}

const formulaId = crypto.randomUUID();
const lineId = crypto.randomUUID();
const inlineSession = createSession(formulaId, lineId, "inline");
assert.equal(
  calculateInlineSessionPosition(inlineSession),
  -9,
  "native Word fallback must reuse the exact scaled Office.js offset",
);
const runsBeforeInlineCreate = wordRunCount;
await adapter.applySession(inlineSession);
assert.equal(
  wordRunCount - runsBeforeInlineCreate,
  1,
  "ordinary inline commits must not trigger a full-document numbering scan",
);
assert.equal(allPictures.length, 1, "inline create must insert one picture");
assert.equal(legacyControls.length, 0, "new formulas must not create content controls");
let picture = allPictures[0];
assert.equal(picture.base64, "png-base64");
assert.equal(picture.altTextTitle, `VisualTeX_${formulaId}`);
assert.match(picture.altTextDescription, /^visualtex:v1:deflate:/);
assert.equal(
  picture.fontPosition,
  -9,
  "inline formula must align its exported mathematical baseline with Word text",
);
assert.equal(
  lastCaretFontPosition,
  0,
  "the caret after an inline formula must reset the formula's baseline shift",
);
assert.equal(cachedMetadata.displayMode, "inline");

const runsAfterInlineCreate = wordRunCount;
await adapter.applySession(inlineSession);
assert.equal(
  wordRunCount,
  runsAfterInlineCreate,
  "retrying the same Session after a later native failure must not write the picture twice",
);
assert.equal(allPictures.length, 1, "a retry of the same create Session must remain idempotent");

selectedPictures = [picture];
const selectionContext = await adapter.readSelection("edit");
assert.equal(selectionContext.sourceObjectId, formulaId);
assert.equal(selectionContext.sessionSeed.formulaId, formulaId);
assert.equal(selectionContext.sessionSeed.displayMode, "inline");

const originalMetadata = createFormulaMetadata({
  formulaId,
  title: inlineSession.title,
  lines: inlineSession.lines,
  codeFormat: inlineSession.codeFormat,
  displayMode: "inline",
});
const legacyControl = new FakeLegacyControl(formulaId);
legacyControls.push(legacyControl);

const editSession = {
  ...inlineSession,
  id: crypto.randomUUID(),
  mode: "edit",
  sourceObjectId: formulaId,
  originalMetadata,
  lines: [{ id: lineId, latex: "x=y" }],
  exportResult: {
    ...inlineSession.exportResult,
    pngBase64: "updated-png-base64",
  },
};
await adapter.applySession(editSession);
assert.equal(allPictures.length, 1, "editing must replace instead of duplicating");
picture = allPictures[0];
assert.equal(picture.base64, "updated-png-base64");
assert.equal(legacyControl.deleted, true, "editing must remove the old bounding box");

selectedPictures = [picture];
const secondEditContext = await adapter.readSelection("edit");
assert.equal(secondEditContext.sessionSeed.lines[0].latex, "x=y");
const secondEditSession = {
  ...editSession,
  id: crypto.randomUUID(),
  sourceObjectId: secondEditContext.sourceObjectId,
  originalMetadata: secondEditContext.sessionSeed.originalMetadata,
  lines: [{ id: lineId, latex: "x=z" }],
  exportResult: {
    ...editSession.exportResult,
    pngBase64: "second-update-png-base64",
  },
};
await adapter.applySession(secondEditSession);
assert.equal(allPictures.length, 1);
assert.equal(allPictures[0].base64, "second-update-png-base64");

const blockFormulaId = crypto.randomUUID();
const blockLineId = crypto.randomUUID();
selectedPictures = [];
selection.paragraph = new FakeParagraph("");
paragraphs.push(selection.paragraph);
const blockSession = createSession(blockFormulaId, blockLineId, "block");
const runsBeforeUnnumberedBlock = wordRunCount;
await adapter.applySession(blockSession);
assert.equal(
  wordRunCount - runsBeforeUnnumberedBlock,
  1,
  "unnumbered display commits must not trigger a full-document numbering scan",
);
const blockPicture = allPictures.find(
  (item) => item.altTextTitle === `VisualTeX_${blockFormulaId}`,
);
assert.ok(blockPicture, "display formula must insert a picture");
assert.equal(blockPicture.paragraph.alignment, "Centered");
assert.equal(blockPicture.fontPosition, 0, "display formula must not use inline baseline shift");

const numberedFormulaId = crypto.randomUUID();
const numberedLineId = crypto.randomUUID();
selectedPictures = [];
selection.paragraph = new FakeParagraph("");
paragraphs.push(selection.paragraph);
const numberedSession = {
  ...createSession(numberedFormulaId, numberedLineId, "block"),
  numbered: true,
};
await adapter.applySession(numberedSession);
const numberedPicture = allPictures.find(
  (item) => item.altTextTitle === `VisualTeX_${numberedFormulaId}`,
);
assert.ok(numberedPicture?.table, "numbered display formula must use a table scaffold");
assert.equal(numberedPicture.table.cells[1].horizontalAlignment, "Centered");
assert.equal(numberedPicture.table.cells[2].horizontalAlignment, "Right");
assert.equal(numberedPicture.table.cells[2].verticalAlignment, "Center");
assert.equal(numberedPicture.table.cells[2].value, "(1)");
assert.equal(
  numberedPicture.width / numberedPicture.height,
  numberedSession.exportResult.width / numberedSession.exportResult.height,
  "numbered formula sizing must preserve the exported aspect ratio",
);

// Copying a complete numbered equation duplicates its metadata/formulaId.
// Refresh must allocate a fresh identity for the copy and number both in order.
selectedPictures = [];
selection.paragraph = new FakeParagraph("");
paragraphs.push(selection.paragraph);
await adapter.applySession({ ...numberedSession, id: crypto.randomUUID() });
const numberedPictures = allPictures.filter(
  (item) => item.table && item.table.cells[2].value.startsWith("("),
);
const numberedTitles = numberedPictures
  .filter((item) => item.table.cells[2].value.startsWith("("))
  .map((item) => item.altTextTitle);
assert.equal(new Set(numberedTitles).size, numberedTitles.length);
const activeNumberLabels = equationNumberControls
  .filter((control) => !control.deleted)
  .map((control) => control.cell.value);
assert.deepEqual(activeNumberLabels, ["(1)", "(2)"]);

// Deleting a numbered formula and refreshing removes its orphan scaffold and
// closes the sequence without leaving a stale label.
const firstNumberedTable = numberedPicture.table;
numberedPicture.delete();
const updatedCount = await adapter.updateEquationNumbers();
assert.equal(updatedCount, 1);
assert.equal(firstNumberedTable.deleted, true);
assert.deepEqual(
  equationNumberControls
    .filter((control) => !control.deleted)
    .map((control) => control.cell.value),
  ["(1)"],
);

rejectPngOnce = true;
const fallbackFormulaId = crypto.randomUUID();
const fallbackLineId = crypto.randomUUID();
selectedPictures = [];
selection.paragraph = new FakeParagraph("");
paragraphs.push(selection.paragraph);
await adapter.applySession(createSession(fallbackFormulaId, fallbackLineId, "inline"));
const fallbackPicture = allPictures.find(
  (item) => item.altTextTitle === `VisualTeX_${fallbackFormulaId}`,
);
assert.equal(fallbackPicture.base64, "svg-base64", "SVG must be used after PNG rejection");

console.log(
  "Word adapter baseline, caret, numbering, deduplication, migration and repeated-edit tests passed",
);
