import assert from "node:assert/strict";
import { createFormulaMetadata } from "../src/office/metadata/formulaMetadata.ts";

const formulaId = crypto.randomUUID();
const lineId = crypto.randomUUID();
let cachedMetadata = null;
let rejectSvgOnce = true;
let nextControlId = 1;
const controls = [];

class FakePicture {
  constructor(base64) {
    this.base64 = base64;
    this.altTextTitle = "";
    this.altTextDescription = "";
    this.width = 0;
    this.height = 0;
    this.lockAspectRatio = false;
    this.control = null;
  }

  insertContentControl() {
    const control = new FakeControl(this);
    this.control = control;
    controls.push(control);
    selection.contentControls.items = [control];
    return control;
  }
}

class FakeRange {
  constructor(control = null) {
    this.control = control;
    this.inlinePictures = {
      items: control ? [control.picture] : [],
      load() {
        return this;
      },
    };
  }

  insertInlinePictureFromBase64(base64, location) {
    if (rejectSvgOnce && base64 === "svg-base64") {
      rejectSvgOnce = false;
      throw new Error("SVG unsupported by simulated Word host");
    }
    const picture = new FakePicture(base64);
    if (location === "Replace" && this.control) {
      this.control.picture = picture;
      picture.control = this.control;
      this.inlinePictures.items = [picture];
    }
    return picture;
  }

  insertContentControl() {
    throw new Error("unused");
  }

  select(mode) {
    this.selectedMode = mode;
  }
}

class FakeControl {
  constructor(picture) {
    this.id = nextControlId++;
    this.tag = "";
    this.title = "";
    this.appearance = "";
    this.cannotDelete = false;
    this.cannotEdit = false;
    this.isNullObject = false;
    this.picture = picture;
  }

  load() {
    return this;
  }

  getRange(location) {
    if (location === "Content") return new FakeRange(this);
    return new FakeRange();
  }
}

const selection = new FakeRange();
selection.contentControls = {
  items: [],
  load() {
    return this;
  },
};
selection.parentContentControlOrNullObject = {
  isNullObject: true,
  load() {
    return this;
  },
};

const fakeDocument = {
  getSelection() {
    return selection;
  },
  contentControls: {
    getByTag(tag) {
      return {
        items: controls.filter((control) => control.tag === tag),
        load() {
          return this;
        },
      };
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
        if (name === "WordApi" && version === "1.3") return false;
        return true;
      },
    },
    ui: {
      openBrowserWindow() {},
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

const createSession = {
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

await adapter.applySession(createSession);
assert.equal(controls.length, 1);
assert.equal(controls[0].picture.base64, "png-base64");
assert.equal(controls[0].tag, `visualtex:${formulaId}`);
assert.equal(controls[0].title, "VisualTeX Formula");
assert.equal(controls[0].picture.altTextTitle, `VisualTeX_${formulaId}`);
assert.match(controls[0].picture.altTextDescription, /^visualtex:v1:deflate:/);
assert.equal(cachedMetadata.formulaId, formulaId);

const selectionContext = await adapter.readSelection("edit");
assert.equal(selectionContext.sourceObjectId, String(controls[0].id));
assert.equal(selectionContext.sessionSeed.formulaId, formulaId);
assert.equal(
  selectionContext.sessionSeed.lines[0].latex,
  createSession.lines[0].latex,
);

const originalMetadata = createFormulaMetadata({
  formulaId,
  title: createSession.title,
  lines: createSession.lines,
  codeFormat: createSession.codeFormat,
  displayMode: "inline",
});
const editSession = {
  ...createSession,
  mode: "edit",
  sourceObjectId: String(controls[0].id),
  originalMetadata,
  lines: [{ id: lineId, latex: "x=y" }],
  exportResult: {
    ...createSession.exportResult,
    svgBase64: "updated-svg-base64",
  },
};

await adapter.applySession(editSession);
assert.equal(controls.length, 1, "editing must not create another content control");
assert.equal(
  controls[0].picture.base64,
  "updated-svg-base64",
  "editing must replace the original picture",
);
assert.equal(controls[0].tag, `visualtex:${formulaId}`);

console.log("Word adapter smoke test passed");
