import assert from "node:assert/strict";
import {
  decodePowerPointObjectReference,
  formulaIdFromPowerPointShapeName,
} from "../src/office/metadata/powerpointMetadata.ts";

const formulaId = crypto.randomUUID();
const lineId = crypto.randomUUID();
let cachedMetadata = null;
let rejectSvgOnce = true;
let nextShapeId = 1;
let currentSelection = [];

class FakeShape {
  constructor(slide, base64, geometry = {}) {
    this.slide = slide;
    this.id = `shape-${nextShapeId++}`;
    this.base64 = base64;
    this.name = `Picture ${this.id}`;
    this.altTextDescription = "";
    this.altTextTitle = "";
    this.isDecorative = false;
    this.left = geometry.imageLeft ?? 0;
    this.top = geometry.imageTop ?? 0;
    this.width = geometry.imageWidth ?? 100;
    this.height = geometry.imageHeight ?? 40;
    this.rotation = 0;
    this.type = "Image";
    this.isNullObject = false;
    this.deleted = false;
  }

  get zOrderPosition() {
    return this.slide.shapes.items.indexOf(this);
  }

  load() {
    return this;
  }

  getParentSlide() {
    return this.slide;
  }

  setZOrder(operation) {
    const items = this.slide.shapes.items;
    const index = items.indexOf(this);
    if (index < 0) return;
    if (operation === "SendBackward" && index > 0) {
      [items[index - 1], items[index]] = [items[index], items[index - 1]];
    }
    if (operation === "BringForward" && index < items.length - 1) {
      [items[index], items[index + 1]] = [items[index + 1], items[index]];
    }
    if (operation === "SendToBack") {
      items.splice(index, 1);
      items.unshift(this);
    }
    if (operation === "BringToFront") {
      items.splice(index, 1);
      items.push(this);
    }
  }

  delete() {
    const index = this.slide.shapes.items.indexOf(this);
    if (index >= 0) this.slide.shapes.items.splice(index, 1);
    this.deleted = true;
    currentSelection = currentSelection.filter((shape) => shape !== this);
  }
}

class FakeShapeCollection {
  constructor(slide) {
    this.slide = slide;
    this.items = [];
  }

  getItemOrNullObject(id) {
    return (
      this.items.find((shape) => shape.id === id) ?? {
        id,
        isNullObject: true,
        load() {
          return this;
        },
      }
    );
  }
}

class FakeSlide {
  constructor(id) {
    this.id = id;
    this.shapes = new FakeShapeCollection(this);
  }

  load() {
    return this;
  }

  setSelectedShapes(shapeIds) {
    currentSelection = shapeIds
      .map((id) => this.shapes.items.find((shape) => shape.id === id))
      .filter(Boolean);
  }
}

const activeSlide = new FakeSlide("slide-1");
const background = new FakeShape(activeSlide, "background");
background.name = "Background";
const foreground = new FakeShape(activeSlide, "foreground");
foreground.name = "Foreground";
activeSlide.shapes.items.push(background, foreground);

const slides = {
  getItem(id) {
    assert.equal(id, activeSlide.id);
    return activeSlide;
  },
};

const presentation = {
  id: "presentation-1",
  slides,
  load() {
    return this;
  },
  getSelectedShapes() {
    return {
      get items() {
        return currentSelection;
      },
      load() {
        return this;
      },
    };
  },
};

const context = {
  presentation,
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
  AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
  CoercionType: { Image: "image" },
  context: {
    document: {
      url: "file:///tmp/powerpoint-adapter-smoke.pptx",
      setSelectedDataAsync(base64, options, callback) {
        assert.equal(options.coercionType, "image");
        if (rejectSvgOnce && base64 === "svg-base64") {
          rejectSvgOnce = false;
          callback({
            status: "failed",
            error: { message: "SVG unsupported by simulated PowerPoint host" },
          });
          return;
        }
        const shape = new FakeShape(activeSlide, base64, options);
        activeSlide.shapes.items.push(shape);
        currentSelection = [shape];
        callback({ status: "succeeded", value: undefined });
      },
    },
    requirements: {
      isSetSupported(name, version) {
        return name === "PowerPointApi" && version === "1.10";
      },
    },
    ui: {
      openBrowserWindow() {},
    },
  },
};

globalThis.PowerPoint = {
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

const { PowerPointAdapter } = await import(
  "../src/office/adapters/PowerPointAdapter.ts"
);
const adapter = new PowerPointAdapter();

const createSession = {
  id: crypto.randomUUID(),
  mode: "create",
  host: "powerpoint",
  formulaId,
  sourceDocumentId: null,
  sourceObjectId: null,
  title: "PowerPoint Formula",
  lines: [{ id: lineId, latex: String.raw`\sum_{i=1}^{n}i^2` }],
  activeLineId: lineId,
  codeFormat: "raw",
  exportWidth: 160,
  exportHeight: 64,
  exportResult: {
    svg: "<svg></svg>",
    svgBase64: "svg-base64",
    pngBase64: "png-base64",
    width: 160,
    height: 64,
    baseline: 48,
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
const createdShape = currentSelection[0];
assert.equal(createdShape.base64, "png-base64");
assert.equal(createdShape.name, `VisualTeX_${formulaId}`);
assert.equal(createdShape.altTextTitle, "VisualTeX Formula");
assert.match(createdShape.altTextDescription, /^visualtex:v1:deflate:/);
assert.equal(formulaIdFromPowerPointShapeName(createdShape.name), formulaId);
assert.equal(cachedMetadata.formulaId, formulaId);

const createGeometry = {
  width: createdShape.width,
  height: createdShape.height,
};
assert.ok(createGeometry.width > 0 && createGeometry.height > 0);

const editContext = await adapter.readSelection("edit");
const reference = decodePowerPointObjectReference(editContext.sourceObjectId);
assert.deepEqual(reference, {
  slideId: activeSlide.id,
  shapeId: createdShape.id,
});
assert.equal(editContext.sessionSeed.formulaId, formulaId);
assert.equal(
  editContext.sessionSeed.lines[0].latex,
  createSession.lines[0].latex,
);

createdShape.left = 100;
createdShape.top = 120;
createdShape.width = 220;
createdShape.height = 70;
createdShape.rotation = 15;
activeSlide.shapes.items.splice(
  activeSlide.shapes.items.indexOf(createdShape),
  1,
);
activeSlide.shapes.items.splice(1, 0, createdShape);
currentSelection = [createdShape];
const originalZ = createdShape.zOrderPosition;
const originalGeometry = {
  left: createdShape.left,
  top: createdShape.top,
  width: createdShape.width,
  height: createdShape.height,
  rotation: createdShape.rotation,
};

const editSession = {
  ...createSession,
  mode: "edit",
  sourceObjectId: editContext.sourceObjectId,
  originalMetadata: editContext.sessionSeed.originalMetadata,
  lines: [{ id: lineId, latex: "x=y" }],
  exportResult: {
    ...createSession.exportResult,
    svgBase64: "updated-svg-base64",
  },
};

await adapter.applySession(editSession);
const visualShapes = activeSlide.shapes.items.filter(
  (shape) => formulaIdFromPowerPointShapeName(shape.name) === formulaId,
);
assert.equal(visualShapes.length, 1, "editing must leave exactly one VisualTeX shape");
const updatedShape = visualShapes[0];
assert.equal(updatedShape.base64, "updated-svg-base64");
assert.deepEqual(
  {
    left: updatedShape.left,
    top: updatedShape.top,
    width: updatedShape.width,
    height: updatedShape.height,
    rotation: updatedShape.rotation,
  },
  originalGeometry,
);
assert.equal(updatedShape.zOrderPosition, originalZ);
assert.equal(createdShape.deleted, true, "original shape must be deleted after success");

const ordinary = new FakeShape(activeSlide, "ordinary");
ordinary.name = "Ordinary Picture";
ordinary.altTextTitle = "";
activeSlide.shapes.items.push(ordinary);
currentSelection = [ordinary];
await assert.rejects(
  () => adapter.readSelection("edit"),
  /不是 VisualTeX 公式/,
);

console.log("PowerPoint adapter smoke test passed");
