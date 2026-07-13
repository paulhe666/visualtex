import assert from "node:assert/strict";
import { createFormulaMetadata } from "../src/office/metadata/formulaMetadata.ts";

let cachedMetadata = null;
let rejectPngOnce = false;
const allPictures = [];
const paragraphs = [];
const legacyControls = [];
let selectedPictures = [];

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
}

class FakePicture {
  constructor(base64, paragraph) {
    this.base64 = base64;
    this.paragraph = paragraph;
    this.altTextTitle = "";
    this.altTextDescription = "";
    this.width = 0;
    this.height = 0;
    this.lockAspectRatio = false;
    this.fontPosition = 0;
  }
  getRange() {
    return new FakeRange({ picture: this, paragraph: this.paragraph });
  }
  remove() {
    const globalIndex = allPictures.indexOf(this);
    if (globalIndex >= 0) allPictures.splice(globalIndex, 1);
    const paragraphIndex = this.paragraph.pictures.indexOf(this);
    if (paragraphIndex >= 0) this.paragraph.pictures.splice(paragraphIndex, 1);
    selectedPictures = selectedPictures.filter((picture) => picture !== this);
  }
}

class FakeRange {
  constructor({ picture = null, paragraph = null } = {}) {
    this.picture = picture;
    this.paragraph = paragraph ?? paragraphs[0];
    this.inlinePictures = new FakeCollection(() =>
      picture ? [picture] : selectedPictures,
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
      get: () => this.picture?.fontPosition ?? 0,
      set: (value) => {
        if (this.picture) this.picture.fontPosition = value;
      },
    });
  }
  insertInlinePictureFromBase64(base64, location) {
    if (rejectPngOnce && base64 === "png-base64") {
      rejectPngOnce = false;
      throw new Error("PNG unsupported by simulated Word host");
    }
    if (location === "Replace" && this.picture) {
      const paragraph = this.picture.paragraph;
      this.picture.remove();
      const replacement = new FakePicture(base64, paragraph);
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
  select() {}
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
  contentControls: {
    getByTag(tag) {
      return new FakeCollection(() =>
        legacyControls.filter((control) => control.tag === tag && !control.deleted),
      );
    },
  },
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
        if (name === "WordApiDesktop") return Number(version) <= 1.4;
        return false;
      },
    },
  },
};

globalThis.Word = {
  async run(callback) {
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

const { WordAdapter } = await import("../src/office/adapters/WordAdapter.ts");
const adapter = new WordAdapter();

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
await adapter.applySession(inlineSession);
assert.equal(allPictures.length, 1, "inline create must insert one picture");
assert.equal(legacyControls.length, 0, "new formulas must not create content controls");
let picture = allPictures[0];
assert.equal(picture.base64, "png-base64");
assert.equal(picture.altTextTitle, `VisualTeX_${formulaId}`);
assert.match(picture.altTextDescription, /^visualtex:v1:deflate:/);
assert.equal(
  picture.fontPosition,
  0,
  "inline formula must remain centered in its image selection box",
);
assert.equal(cachedMetadata.displayMode, "inline");

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
await adapter.applySession(blockSession);
const blockPicture = allPictures.find(
  (item) => item.altTextTitle === `VisualTeX_${blockFormulaId}`,
);
assert.ok(blockPicture, "display formula must insert a picture");
assert.equal(blockPicture.paragraph.alignment, "Centered");
assert.equal(blockPicture.fontPosition, 0, "display formula must not use inline baseline shift");

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

console.log("Word adapter inline, display, migration and repeated-edit tests passed");
