import assert from "node:assert/strict";
import {
  decodePowerPointObjectReference,
  formulaIdFromPowerPointShapeName,
} from "../src/office/metadata/powerpointMetadata.ts";
import { officeErrorMessage } from "../src/office/errors.ts";

const apiLevel = process.env.POWERPOINT_API_LEVEL ?? "1.10";
const failTagWrites = process.env.POWERPOINT_FAIL_TAGS === "1";
const inPlaceEdit = process.env.POWERPOINT_IN_PLACE_EDIT === "1";
const macosNativeFirst = process.env.POWERPOINT_MACOS_NATIVE_FIRST === "1";

if (macosNativeFirst) {
  Object.defineProperty(globalThis, "navigator", {
    configurable: true,
    value: { userAgent: "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_5)" },
  });
}

function apiAtLeast(required) {
  const [major, minor] = apiLevel.split(".").map(Number);
  const [requiredMajor, requiredMinor] = required.split(".").map(Number);
  return major > requiredMajor || (major === requiredMajor && minor >= requiredMinor);
}

const formulaId = crypto.randomUUID();
const lineId = crypto.randomUUID();
let cachedMetadata = null;
const rejectPngOnce = process.env.POWERPOINT_REJECT_PNG === "1";
let pngRejected = false;
let nextShapeId = 1;
let currentSelection = [];
const insertionCalls = [];
let revealRequestCount = 0;
let powerPointRunCount = 0;
let nativeSelectionRequestCount = 0;
let nativeSlideSnapshotRequestCount = 0;

class FakeTagCollection {
  constructor() {
    this.items = [];
  }

  add(key, value) {
    if (failTagWrites) {
      throw {
        name: "RichApi.Error",
        code: "GeneralException",
        message: "Simulated PowerPoint tag failure",
        debugInfo: { errorLocation: "Shape.tags.add" },
      };
    }
    const normalized = key.toUpperCase();
    const existing = this.items.find((item) => item.key === normalized);
    if (existing) {
      existing.value = value;
    } else {
      this.items.push({ key: normalized, value });
    }
  }

  load() {
    return this;
  }
}

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
    this.tags = new FakeTagCollection();
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

  load() {
    return this;
  }

  getItem(id) {
    const shape = this.items.find((item) => item.id === id);
    if (!shape) throw new Error(`Shape not found: ${id}`);
    return shape;
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
  getItemAt(index) {
    assert.equal(index, 0);
    return activeSlide;
  },
};

function scopedCollection(itemsProvider) {
  return {
    get items() {
      return itemsProvider();
    },
    load() {
      return this;
    },
  };
}

const presentation = {
  id: "presentation-1",
  slides,
  load() {
    return this;
  },
  getSelectedShapes() {
    return scopedCollection(() => currentSelection);
  },
  getSelectedSlides() {
    return scopedCollection(() => [activeSlide]);
  },
};

const context = {
  presentation,
  async sync() {},
};

globalThis.window = {
  __VISUALTEX_INSTALL_TOKEN__: "test-install-token",
  location: { href: "" },
  setTimeout,
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
  CoercionType: { Image: "image", XmlSvg: "xmlSvg" },
  context: {
    document: {
      url: "file:///tmp/powerpoint-adapter-smoke.pptx",
      setSelectedDataAsync(base64, options, callback) {
        insertionCalls.push({ base64, options: { ...options } });
        if (
          rejectPngOnce &&
          !pngRejected &&
          options.coercionType === "image"
        ) {
          pngRejected = true;
          callback({
            status: "failed",
            error: {
              code: "UnsupportedDataObject",
              message: "PNG unsupported by simulated PowerPoint host",
            },
          });
          return;
        }
        if (
          inPlaceEdit &&
          currentSelection.length === 1 &&
          formulaIdFromPowerPointShapeName(currentSelection[0].name)
        ) {
          const shape = currentSelection[0];
          shape.base64 = base64;
          shape.left = options.imageLeft ?? shape.left;
          shape.top = options.imageTop ?? shape.top;
          shape.width = options.imageWidth ?? shape.width;
          shape.height = options.imageHeight ?? shape.height;
        } else {
          const shape = new FakeShape(activeSlide, base64, options);
          activeSlide.shapes.items.push(shape);
          // PowerPoint for Mac does not reliably leave the inserted image selected.
          currentSelection = [];
        }
        callback({ status: "succeeded", value: undefined });
      },
    },
    requirements: {
      isSetSupported(name, version) {
        if (name === "PowerPointApi") return apiAtLeast(version);
        if (name === "ImageCoercion") return Number(version) <= 1.2;
        return false;
      },
    },
    ui: {
      openBrowserWindow() {},
    },
  },
};

globalThis.PowerPoint = {
  async run(callback) {
    powerPointRunCount += 1;
    return callback(context);
  },
};

globalThis.fetch = async (input, init = {}) => {
  const url = String(input);
  if (url.endsWith("/api/v1/app/reveal")) {
    assert.equal(init.method, "POST");
    assert.equal(init.headers["X-VisualTeX-Install-Token"], "test-install-token");
    revealRequestCount += 1;
    return new Response(null, { status: 204 });
  }
  if (url.endsWith("/api/v1/powerpoint/slide/snapshot")) {
    nativeSlideSnapshotRequestCount += 1;
    return new Response(
      JSON.stringify({
        presentationIdentity: "/tmp/powerpoint-adapter-smoke.pptx",
        slideIndex: 1,
        slideId: 256,
        shapeCount: activeSlide.shapes.items.length,
        shapeNames: activeSlide.shapes.items.map((shape) => shape.name),
      }),
      { status: 200, headers: { "Content-Type": "application/json" } },
    );
  }
  if (url.endsWith("/api/v1/powerpoint/selection")) {
    nativeSelectionRequestCount += 1;
    if (currentSelection.length !== 1) {
      return new Response(JSON.stringify({ error: "Select exactly one shape" }), {
        status: 400,
        headers: { "Content-Type": "application/json" },
      });
    }
    const shape = currentSelection[0];
    return new Response(
      JSON.stringify({
        shapeName: shape.name,
        slideIndex: 1,
        slideId: 256,
        presentationIdentity: "/tmp/powerpoint-adapter-smoke.pptx",
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
      }),
      { status: 200, headers: { "Content-Type": "application/json" } },
    );
  }
  if (url.endsWith("/api/v1/powerpoint/selection/mark")) {
    const { formulaId } = JSON.parse(init.body);
    if (currentSelection.length !== 1) {
      return new Response(JSON.stringify({ error: "Select exactly one shape" }), {
        status: 400,
        headers: { "Content-Type": "application/json" },
      });
    }
    currentSelection[0].name = `VisualTeX_${formulaId}`;
    const shape = currentSelection[0];
    return new Response(
      JSON.stringify({
        shapeName: shape.name,
        slideIndex: 1,
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
      }),
      { status: 200, headers: { "Content-Type": "application/json" } },
    );
  }
  if (url.endsWith("/api/v1/powerpoint/shape/mark-last")) {
    const { formulaId, previousShapeNames } = JSON.parse(init.body);
    const inserted = activeSlide.shapes.items.filter(
      (shape) => !previousShapeNames.includes(shape.name),
    );
    const shape = inserted.length === 1 ? inserted[0] : currentSelection[0];
    if (!shape) {
      return new Response(JSON.stringify({ error: "No inserted shape" }), {
        status: 400,
        headers: { "Content-Type": "application/json" },
      });
    }
    shape.name = `VisualTeX_${formulaId}`;
    return new Response(
      JSON.stringify({
        shapeName: shape.name,
        slideIndex: 1,
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
      }),
      { status: 200, headers: { "Content-Type": "application/json" } },
    );
  }
  if (url.endsWith("/api/v1/powerpoint/shape/replace-last")) {
    const {
      formulaId,
      previousShapeNames,
      originalShapeName,
      left,
      top,
      width,
      height,
    } = JSON.parse(init.body);
    const original = activeSlide.shapes.items.find(
      (shape) => shape.name === originalShapeName,
    );
    if (!original) {
      return new Response(JSON.stringify({ error: "Original shape missing" }), {
        status: 400,
        headers: { "Content-Type": "application/json" },
      });
    }
    const inserted = activeSlide.shapes.items.filter(
      (shape) => !previousShapeNames.includes(shape.name),
    );
    const shape = inserted.length === 1 ? inserted[0] : original;
    if (shape !== original) {
      original.name = `VisualTeXOld_${formulaId}`;
      shape.name = `VisualTeX_${formulaId}`;
      original.delete();
    } else {
      shape.name = `VisualTeX_${formulaId}`;
    }
    shape.left = left;
    shape.top = top;
    shape.width = width;
    shape.height = height;
    return new Response(
      JSON.stringify({
        shapeName: shape.name,
        slideIndex: 1,
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
      }),
      { status: 200, headers: { "Content-Type": "application/json" } },
    );
  }
  if (url.endsWith("/api/v1/powerpoint/events?cursor=0")) {
    return new Response("[]", {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
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

assert.equal(
  officeErrorMessage(
    {
      name: "RichApi.Error",
      code: "GeneralException",
      message: "PowerPoint operation failed",
      debugInfo: { errorLocation: "Shape.tags.add" },
    },
    "fallback",
  ),
  "PowerPoint operation failed (code=GeneralException, location=Shape.tags.add)",
);

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
  displayMode: "block",
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

const shapeCountBeforeCreate = activeSlide.shapes.items.length;
await adapter.applySession(createSession);
const expectedCreateImage = rejectPngOnce ? "<svg></svg>" : "png-base64";
const createdShape = activeSlide.shapes.items.find(
  (shape) => shape.base64 === expectedCreateImage,
);
assert.ok(createdShape, "formula insertion must leave a visible image on the slide");
assert.equal(
  activeSlide.shapes.items.length,
  shapeCountBeforeCreate + 1,
  "create must add exactly one image",
);
assert.equal(cachedMetadata.formulaId, formulaId);
assert.equal(insertionCalls[0].options.coercionType, "image");
if (rejectPngOnce) {
  assert.equal(insertionCalls[1].options.coercionType, "xmlSvg");
  assert.equal(
    insertionCalls[1].base64,
    "<svg></svg>",
    "XmlSvg must receive raw SVG XML instead of Base64",
  );
}
for (const call of insertionCalls.slice(0, rejectPngOnce ? 2 : 1)) {
  assert.equal(
    Object.hasOwn(call.options, "imageLeft"),
    false,
    "new formula insertion must omit undefined imageLeft",
  );
  assert.equal(
    Object.hasOwn(call.options, "imageTop"),
    false,
    "new formula insertion must omit undefined imageTop",
  );
}

assert.equal(createdShape.name, `VisualTeX_${formulaId}`);
assert.equal(formulaIdFromPowerPointShapeName(createdShape.name), formulaId);
if (apiAtLeast("1.5") && !failTagWrites) {
  assert.ok(
    createdShape.tags.items.some(
      (tag) => tag.key === "VISUALTEX_FORMULA_ID" && tag.value === formulaId,
    ),
  );
} else {
  assert.equal(
    createdShape.tags.items.length,
    0,
    apiAtLeast("1.5")
      ? "optional tag failure must not roll back the inserted image"
      : "native fallback does not require Office.js shape tags",
  );
}
if (apiAtLeast("1.10")) {
  assert.equal(createdShape.altTextTitle, "VisualTeX Formula");
  assert.match(createdShape.altTextDescription, /^visualtex:v1:deflate:/);
} else {
  assert.equal(createdShape.altTextTitle, "");
  assert.equal(createdShape.altTextDescription, "");
}

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

const powerPointRunsBeforeEditRead = powerPointRunCount;
const nativeReadsBeforeEditRead = nativeSelectionRequestCount;
const nativeSnapshotsBeforeEditRead = nativeSlideSnapshotRequestCount;
if (macosNativeFirst) {
  adapter.prepareInteractionTarget({
    host: "powerpoint",
    formulaId,
    shapeName: createdShape.name,
    slideIndex: 1,
    slideId: 256,
    presentationIdentity: "/tmp/powerpoint-adapter-smoke.pptx",
    left: createdShape.left,
    top: createdShape.top,
    width: createdShape.width,
    height: createdShape.height,
  });
  currentSelection = [];
  const targetedEditContext = await adapter.readSelection("edit");
  assert.match(
    targetedEditContext.sourceObjectId,
    /^visualtex-ppt-native-edit:1:/,
    "double-click editing must preserve the captured native target",
  );
  assert.equal(
    nativeSelectionRequestCount,
    nativeReadsBeforeEditRead,
    "double-click editing must not re-read PowerPoint's mutable selection",
  );
  assert.equal(targetedEditContext.sessionSeed.formulaId, formulaId);
  currentSelection = [createdShape];
}
const editContext = await adapter.readSelection("edit");
if (macosNativeFirst) {
  assert.match(
    editContext.sourceObjectId,
    /^visualtex-ppt-native-edit:1:/,
    "macOS edits must preserve an exact native shape target",
  );
  assert.equal(
    nativeSelectionRequestCount,
    nativeReadsBeforeEditRead + 1,
    "macOS edit selection should require one native shape read",
  );
  assert.equal(
    powerPointRunCount,
    powerPointRunsBeforeEditRead,
    "macOS named formulas must bypass the slower Office.js selection path",
  );
  assert.equal(
    nativeSlideSnapshotRequestCount,
    nativeSnapshotsBeforeEditRead,
    "the native selection payload must avoid a second serialized slide query",
  );
  assert.equal(
    editContext.sourceDocumentId,
    "visualtex-ppt-native-presentation:/tmp/powerpoint-adapter-smoke.pptx",
  );

  // Native PowerPoint can move the UI selection immediately after paste. The
  // finalizer must locate the durable shape by the immutable slide/name/geometry
  // returned by the native transaction, not by whatever is currently selected.
  const nativeFinalizeFormulaId = crypto.randomUUID();
  const nativeFinalizeLineId = crypto.randomUUID();
  const nativeFinalizeShape = new FakeShape(activeSlide, "native-svg", {
    imageLeft: 240,
    imageTop: 96,
    imageWidth: 180,
    imageHeight: 54,
  });
  nativeFinalizeShape.name = `VisualTeX_${nativeFinalizeFormulaId}`;
  activeSlide.shapes.items.push(nativeFinalizeShape);
  currentSelection = [];
  await adapter.finalizeNativePowerPointCommit(
    {
      ...createSession,
      id: crypto.randomUUID(),
      formulaId: nativeFinalizeFormulaId,
      lines: [{ id: nativeFinalizeLineId, latex: String.raw`\\int_0^1 x^2\\,dx` }],
      activeLineId: nativeFinalizeLineId,
    },
    {
      shapeName: nativeFinalizeShape.name,
      slideIndex: 1,
      slideId: 256,
      presentationIdentity: "/tmp/powerpoint-adapter-smoke.pptx",
      left: nativeFinalizeShape.left,
      top: nativeFinalizeShape.top,
      width: nativeFinalizeShape.width,
      height: nativeFinalizeShape.height,
    },
  );
  assert.equal(currentSelection.length, 0, "native finalization must not require UI selection");
  assert.ok(
    nativeFinalizeShape.tags.items.some(
      (tag) =>
        tag.key === "VISUALTEX_FORMULA_ID" &&
        tag.value === nativeFinalizeFormulaId,
    ),
    "native finalization must persist editable metadata on the exact returned shape",
  );
  assert.equal(cachedMetadata.formulaId, nativeFinalizeFormulaId);
  console.log("PowerPoint macOS native-first edit and unselected-finalize checks passed");
  process.exit(0);
}
const reference = decodePowerPointObjectReference(editContext.sourceObjectId);
if (apiAtLeast("1.5")) {
  assert.deepEqual(reference, {
    slideId: activeSlide.id,
    shapeId: createdShape.id,
  });
} else {
  assert.equal(reference.slideId, "native:1");
  assert.equal(reference.shapeId, createdShape.name);
  assert.equal(reference.native.slideIndex, 1);
  assert.equal(reference.native.shapeName, createdShape.name);
}
assert.equal(editContext.sessionSeed.formulaId, formulaId);
assert.equal(
  editContext.sessionSeed.lines[0].latex,
  createSession.lines[0].latex,
);

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
    pngBase64: "updated-png-base64",
    svg: "<svg id=\"updated\"></svg>",
    svgBase64: "updated-svg-base64",
  },
};

const editCallStart = insertionCalls.length;
await adapter.applySession(editSession);
const visualShapes = activeSlide.shapes.items.filter(
  (shape) => formulaIdFromPowerPointShapeName(shape.name) === formulaId,
);
assert.equal(visualShapes.length, 1, "editing must leave exactly one VisualTeX shape");
const updatedShape = visualShapes[0];
assert.equal(updatedShape.base64, "updated-png-base64");
assert.equal(insertionCalls[editCallStart].options.coercionType, "image");
assert.equal(insertionCalls[editCallStart].options.imageLeft, originalGeometry.left);
assert.equal(insertionCalls[editCallStart].options.imageTop, originalGeometry.top);
assert.deepEqual(
  {
    left: updatedShape.left,
    top: updatedShape.top,
    width: updatedShape.width,
    height: updatedShape.height,
  },
  {
    left: originalGeometry.left,
    top: originalGeometry.top,
    width: originalGeometry.width,
    height: originalGeometry.height,
  },
);
if (apiAtLeast("1.10") || inPlaceEdit) {
  assert.equal(updatedShape.rotation, originalGeometry.rotation);
} else {
  assert.equal(updatedShape.rotation, 0);
}
if (apiAtLeast("1.8")) {
  assert.equal(updatedShape.zOrderPosition, originalZ);
}
assert.equal(
  createdShape.deleted,
  !inPlaceEdit,
  inPlaceEdit
    ? "in-place replacement must keep the original shape object"
    : "new-shape replacement must delete the original shape",
);
if (apiAtLeast("1.5") && !inPlaceEdit) {
  assert.equal(currentSelection[0], updatedShape, "updated formula should be selected");
}

currentSelection = [updatedShape];
const secondEditContext = await adapter.readSelection("edit");
assert.equal(secondEditContext.sessionSeed.lines[0].latex, "x=y");
const secondEditSession = {
  ...editSession,
  sourceObjectId: secondEditContext.sourceObjectId,
  originalMetadata: secondEditContext.sessionSeed.originalMetadata,
  lines: [{ id: lineId, latex: "x=z" }],
  exportResult: {
    ...editSession.exportResult,
    pngBase64: "second-updated-png-base64",
  },
};
await adapter.applySession(secondEditSession);
const secondVisualShapes = activeSlide.shapes.items.filter(
  (shape) => formulaIdFromPowerPointShapeName(shape.name) === formulaId,
);
assert.equal(secondVisualShapes.length, 1, "a second edit must remain editable");
assert.equal(secondVisualShapes[0].base64, "second-updated-png-base64");

const ordinary = new FakeShape(activeSlide, "ordinary");
ordinary.name = "Ordinary Picture";
ordinary.altTextTitle = "";
activeSlide.shapes.items.push(ordinary);
currentSelection = [ordinary];
await assert.rejects(
  () => adapter.readSelection("edit"),
  /没有 VisualTeX 标记|不是 VisualTeX 公式/,
);

await adapter.openDesktopApp();
assert.equal(revealRequestCount, 1, "Open VisualTeX must call the local reveal API");

console.log(
  `PowerPoint ${apiLevel} adapter compatibility passed${
    failTagWrites ? " with optional tag failure" : ""
  }`,
);
