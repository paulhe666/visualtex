import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 700;
const vitePort = 22300 + offset;
const debugPort = 29300 + offset;
const baseUrl = `http://127.0.0.1:${vitePort}`;
const chromeProfile = `/tmp/visualtex-custom-symbol-glyph-compiler-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      if ((await fetch(url)).ok) return;
    } catch {
      // Retry during process startup.
    }
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

class CdpClient {
  constructor(url) {
    this.url = url;
    this.nextId = 1;
    this.pending = new Map();
  }
  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((resolve, reject) => {
      this.socket.addEventListener("open", resolve, { once: true });
      this.socket.addEventListener("error", reject, { once: true });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      if (!message.id) return;
      const pending = this.pending.get(message.id);
      if (!pending) return;
      this.pending.delete(message.id);
      if (message.error) pending.reject(new Error(message.error.message));
      else pending.resolve(message.result);
    });
  }
  send(method, params = {}) {
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`CDP ${method} timed out`));
      }, 15000);
      this.pending.set(id, {
        resolve: (value) => {
          clearTimeout(timeout);
          resolve(value);
        },
        reject: (error) => {
          clearTimeout(timeout);
          reject(error);
        },
      });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }
  async evaluate(expression) {
    const result = await this.send("Runtime.evaluate", {
      expression,
      awaitPromise: true,
      returnByValue: true,
    });
    if (result.exceptionDetails) {
      throw new Error(
        result.exceptionDetails.exception?.description ||
          result.exceptionDetails.text ||
          "Runtime.evaluate failed",
      );
    }
    return result.result.value;
  }
  close() {
    this.socket?.close();
  }
}

async function main() {
  const vite = spawn(
    process.execPath,
    [
      "node_modules/vite/bin/vite.js",
      "--host",
      "127.0.0.1",
      "--port",
      String(vitePort),
      "--strictPort",
    ],
    { cwd: process.cwd(), stdio: "ignore" },
  );
  let chrome;
  let client;
  try {
    await waitFor(baseUrl);
    chrome = spawn(
      chromePath,
      [
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`,
        `--user-data-dir=${chromeProfile}`,
        "--window-size=1400,1000",
        baseUrl,
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const page = targets.find(
      (target) => target.type === "page" && target.url.startsWith(baseUrl),
    );
    assert.ok(page, "VisualTeX browser target must exist");
    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(450);

    const probe = await client.evaluate(`(async () => {
      const compiler = await import("/src/math/customSymbolGlyphCompiler.ts");
      const designer = await import("/src/math/customSymbolDesignerCompiler.ts");
      const designerTypes = await import("/src/math/customSymbolDesignerTypes.ts");
      const designerGeometry = await import("/src/math/customSymbolDesignerGeometry.ts");
      const designerArchive = await import("/src/math/customSymbolDesignerArchive.ts");
      const systemGlyphs = await import("/src/math/customSymbolSystemGlyphs.ts");
      const registry = await import("/src/math/customSymbolRegistry.ts");
      const registration = await import("/src/math/customSymbolRegistration.ts");
      const runtime = await import("/src/export/runtime.ts");
      const copyService = await import("/src/clipboard/LatexCopyService.ts");
      const officeRender = await import("/src/office/shared/formulaRenderArtifacts.ts");
      localStorage.removeItem("visualtex.custom-symbols.v1");
      registry.refreshCustomSymbolLibraryFromStorage();
      const samples = [
        ["alpha", "\\\\alpha"],
        ["integral", "\\\\int"],
        ["sum", "\\\\sum"],
        ["arrow", "\\\\rightarrow"],
        ["blackboard", "\\\\mathbb{R}"],
        ["partial", "\\\\partial"],
        ["fragment", "\\\\alpha+\\\\beta"],
      ];
      const output = [];
      const BS = String.fromCharCode(92);
      for (let index = 0; index < samples.length; index += 1) {
        const [name, encoded] = samples[index];
        const source = encoded.replaceAll("\\\\", BS);
        const asset = compiler.compileLatexGlyphAsset(source);
        const command = "vtxcompiler" + String.fromCharCode(97 + index);
        const id = "glyph-compiler-" + name;
        const now = Date.now();
        registration.registerCustomSymbolSafely({
          id,
          command,
          name,
          role: "ordinary",
          limitsBehavior: "auto",
          metrics: asset.metrics,
          artwork: { shapes: asset.shapes },
          ommlFallback: null,
          createdAt: now,
          updatedAt: now,
        });
        const original = runtime.latexToSvg(source, {
          displayMode: false,
          fontSizePt: 12,
          paddingPx: 0,
          background: "transparent",
        });
        const roundTrip = runtime.latexToSvg(BS + command, {
          displayMode: false,
          fontSizePt: 12,
          paddingPx: 0,
          background: "transparent",
        });
        output.push({
          name,
          metrics: asset.metrics,
          shapeCount: asset.shapes.length,
          kinds: asset.shapes.map((shape) => shape.kind),
          matrices: asset.shapes.map((shape) => shape.transform?.matrix || null),
          original: {
            width: original.width,
            height: original.height,
            baseline: original.baseline,
          },
          roundTrip: {
            width: roundTrip.width,
            height: roundTrip.height,
            baseline: roundTrip.baseline,
            hasArtwork: roundTrip.svg.includes('data-visualtex-custom-symbol="' + id + '"'),
            containsText: /<text\\b|<foreignObject\\b|<image\\b/i.test(roundTrip.svg),
          },
        });
      }
      const partialAsset = compiler.compileLatexGlyphAsset(BS + "partial");
      const firstSlice = designer.glyphLayerFromAsset(partialAsset, {
        id: "partial-top",
        name: "partial top",
      });
      firstSlice.clipRect = {
        x: 0,
        y: 0,
        width: partialAsset.metrics.widthEm * 1000,
        height: partialAsset.metrics.ascentEm * 1000 * 0.52,
      };
      firstSlice.transform = {
        ...firstSlice.transform,
        translateX: 35,
        translateY: -18,
        scaleX: 1.08,
        scaleY: 0.94,
        rotateDeg: 7,
      };
      const secondSlice = designer.glyphLayerFromAsset(partialAsset, {
        id: "partial-bottom",
        name: "partial bottom",
      });
      secondSlice.clipRect = {
        x: 0,
        y: partialAsset.metrics.ascentEm * 1000 * 0.44,
        width: partialAsset.metrics.widthEm * 1000,
        height:
          (partialAsset.metrics.ascentEm + partialAsset.metrics.descentEm) * 1000 * 0.56,
      };
      const hiddenSlice = designer.glyphLayerFromAsset(partialAsset, {
        id: "partial-hidden",
        name: "hidden",
      });
      hiddenSlice.visible = false;
      const designDocument = designerTypes.createEmptyCustomSymbolDesignerDocument();
      designDocument.name = "Sliced partial";
      designDocument.command = "vtxdesignerslice";
      designDocument.metrics = { ...partialAsset.metrics };
      designDocument.layers = [firstSlice, secondSlice, hiddenSlice];
      const definition = designer.customSymbolDefinitionFromDesignerDocument(
        designDocument,
        { id: "designer-slice-roundtrip" },
      );
      registration.registerCustomSymbolSafely(definition);
      const designerSvg = runtime.latexToSvg(BS + "vtxdesignerslice", {
        displayMode: false,
        fontSizePt: 12,
        paddingPx: 0,
        background: "transparent",
      }).svg;
      const restoredDesigner = designerArchive.restoreCustomSymbolDesignerDocument(definition);
      const restoredArtwork = designer.compileCustomSymbolDesignerArtwork(
        restoredDesigner.document,
      );
      const designerProbe = {
        sourcePathPreserved:
          firstSlice.asset.shapes[0]?.kind === "path" &&
          secondSlice.asset.shapes[0]?.kind === "path" &&
          firstSlice.asset.shapes[0]?.d === secondSlice.asset.shapes[0]?.d,
        shapeCount: definition.artwork.shapes.length,
        clipCount: definition.artwork.shapes.filter((shape) => shape.clipRect).length,
        matrixCount: definition.artwork.shapes.filter(
          (shape) => Array.isArray(shape.transform?.matrix),
        ).length,
        firstTransform: definition.artwork.shapes[0]?.transform || null,
        hasTwoCropViewports:
          (designerSvg.match(/overflow="hidden"/g) || []).length === 2,
        cropInsideUserTransform:
          /<g transform="[^"]*translate\([^"]+\)"><svg[^>]*overflow="hidden"/.test(designerSvg),
        hasArtwork: designerSvg.includes(
          'data-visualtex-custom-symbol="designer-slice-roundtrip"',
        ),
        archiveAssetCount: definition.designerSource?.assets?.length ?? -1,
        archiveLayerCount: definition.designerSource?.layers?.length ?? -1,
        restoredSourceMode: restoredDesigner.sourceMode,
        restoredLayerCount: restoredDesigner.document.layers.length,
        restoredHiddenCount: restoredDesigner.document.layers.filter(
          (layer) => !layer.visible,
        ).length,
        restoredClipCount: restoredDesigner.document.layers.filter(
          (layer) => layer.clipRect,
        ).length,
        restoredGlyphSources: restoredDesigner.document.layers
          .filter((layer) => layer.kind === "glyph")
          .map((layer) => layer.asset.sourceLatex),
        restoredArtworkShapeCount: restoredArtwork.length,
        restoredFirstTransform:
          restoredDesigner.document.layers[0]?.transform || null,
        restoredCanvasMetrics: restoredDesigner.document.metrics,
      };

      const wideDocument = designerTypes.createEmptyCustomSymbolDesignerDocument();
      wideDocument.name = "Wide canvas tight symbol";
      wideDocument.command = "vtxtightcrop";
      wideDocument.metrics = { widthEm: 6, ascentEm: 4, descentEm: 2 };
      const tightLayer = designer.glyphLayerFromAsset(partialAsset, {
        id: "tight-partial",
        name: "tight partial",
      });
      tightLayer.transform = {
        ...tightLayer.transform,
        translateX: 220,
        translateY: 4000 - partialAsset.metrics.ascentEm * 1000,
      };
      wideDocument.layers = [tightLayer];
      const tightDefinition = designer.customSymbolDefinitionFromDesignerDocument(
        wideDocument,
        { id: "designer-tight-crop-roundtrip" },
      );
      const restoredWide = designerArchive.restoreCustomSymbolDesignerDocument(
        tightDefinition,
      );
      const tightCropProbe = {
        designerWidth: wideDocument.metrics.widthEm,
        runtimeWidth: tightDefinition.metrics.widthEm,
        runtimeAscent: tightDefinition.metrics.ascentEm,
        runtimeDescent: tightDefinition.metrics.descentEm,
        archivedWidth: tightDefinition.designerSource?.metrics?.widthEm ?? -1,
        restoredWidth: restoredWide.document.metrics.widthEm,
        runtimeTranslateX:
          tightDefinition.artwork.shapes[0]?.transform?.translateX ?? null,
        restoredTranslateX:
          restoredWide.document.layers[0]?.transform.translateX ?? null,
      };

      const geometryDocument = designerTypes.createEmptyCustomSymbolDesignerDocument();
      geometryDocument.name = "Mixed geometry";
      geometryDocument.command = "vtxgeometrymix";
      geometryDocument.role = "relation";
      geometryDocument.ommlFallback = BS + "approx";
      geometryDocument.metrics = { ...partialAsset.metrics };
      const circle = designerGeometry.createCustomSymbolGeometryLayer(
        "circle",
        geometryDocument.metrics,
        { id: "geometry-circle" },
      );
      const line = designerGeometry.createCustomSymbolGeometryLayer(
        "line",
        geometryDocument.metrics,
        { id: "geometry-line" },
      );
      const arrow = designerGeometry.createCustomSymbolGeometryLayer(
        "arrow",
        geometryDocument.metrics,
        { id: "geometry-arrow" },
      );
      geometryDocument.layers = [firstSlice, circle, line, arrow];
      const geometryDefinition = designer.customSymbolDefinitionFromDesignerDocument(
        geometryDocument,
        { id: "designer-geometry-roundtrip" },
      );
      registration.registerCustomSymbolSafely(geometryDefinition);
      const geometrySvgArtifact = runtime.latexToSvg(BS + "vtxgeometrymix", {
        displayMode: false,
        fontSizePt: 12,
        paddingPx: 0,
        background: "transparent",
      });
      const geometrySvg = geometrySvgArtifact.svg;
      const geometryPng = await runtime.svgToPng(geometrySvgArtifact, {
        scale: 2,
        background: "transparent",
      });
      const formattedLatex = copyService.formatLatex(
        BS + "vtxgeometrymix",
        "display-bracket",
      );
      const wordArtifact = officeRender.renderOfficeFormulaArtifacts({
        lines: [{ id: "designer-geometry-word", latex: BS + "vtxgeometrymix" }],
        codeFormat: "raw",
        displayMode: "inline",
        host: "word",
        includeWordOmml: true,
      });
      const powerpointArtifact = officeRender.renderOfficeFormulaArtifacts({
        lines: [{ id: "designer-geometry-powerpoint", latex: BS + "vtxgeometrymix" }],
        codeFormat: "raw",
        displayMode: "inline",
        host: "powerpoint",
        includeWordOmml: false,
      });
      const restoredGeometry = designerArchive.restoreCustomSymbolDesignerDocument(
        geometryDefinition,
      );
      const legacyDefinition = {
        ...geometryDefinition,
        id: "designer-legacy-roundtrip",
        command: "vtxlegacyroundtrip",
        designerSource: null,
      };
      const restoredLegacy = designerArchive.restoreCustomSymbolDesignerDocument(
        legacyDefinition,
      );
      const geometryProbe = {
        shapeCount: geometryDefinition.artwork.shapes.length,
        hasArtwork: geometrySvg.includes(
          'data-visualtex-custom-symbol="designer-geometry-roundtrip"',
        ),
        hasCircle: geometrySvg.includes("<circle "),
        hasLine: geometrySvg.includes("<line "),
        hasArrowPath: geometrySvg.includes("M0 90H300M220 10L300 90L220 170"),
        hasCrop: geometrySvg.includes('overflow="hidden"'),
        formattedLatex,
        pngType: geometryPng.blob.type,
        pngSize: geometryPng.blob.size,
        pngWidth: geometryPng.width,
        pngHeight: geometryPng.height,
        wordCanonicalLatex: wordArtifact.canonicalLatex,
        wordSvgHasArtwork: wordArtifact.svg.svg.includes(
          'data-visualtex-custom-symbol="designer-geometry-roundtrip"',
        ),
        wordOmml: wordArtifact.omml?.omml || "",
        powerpointCanonicalLatex: powerpointArtifact.canonicalLatex,
        powerpointSvgHasArtwork: powerpointArtifact.svg.svg.includes(
          'data-visualtex-custom-symbol="designer-geometry-roundtrip"',
        ),
        powerpointOmmlIsNull: powerpointArtifact.omml === null,
        archiveAssetCount: geometryDefinition.designerSource?.assets?.length ?? -1,
        archiveLayerCount: geometryDefinition.designerSource?.layers?.length ?? -1,
        restoredSourceMode: restoredGeometry.sourceMode,
        restoredLayerKinds: restoredGeometry.document.layers.map(
          (layer) => layer.kind,
        ),
        legacySourceMode: restoredLegacy.sourceMode,
        legacyLayerCount: restoredLegacy.document.layers.length,
        legacyLayerKinds: restoredLegacy.document.layers.map(
          (layer) => layer.kind,
        ),
      };

      const cambriaPreset = systemGlyphs.systemFontPresetById("cambria-math");
      const styledSystemGlyph = await systemGlyphs.createSystemFontGlyphAssetAsync({
        character: "𝑥",
        font: cambriaPreset,
        italic: true,
      });
      const ordinarySystemGlyph = await systemGlyphs.createSystemFontGlyphAssetAsync({
        character: "A",
        font: cambriaPreset,
        italic: true,
      });
      const systemDocument = designerTypes.createEmptyCustomSymbolDesignerDocument();
      systemDocument.name = "System mathematical glyph";
      systemDocument.command = "vtxsystemglyph";
      systemDocument.ommlFallback = BS + "mathit{x}";
      systemDocument.metrics = { ...styledSystemGlyph.asset.metrics };
      systemDocument.layers = [
        designer.glyphLayerFromAsset(styledSystemGlyph.asset, {
          id: "system-math-italic-x",
          name: "Cambria mathematical italic x",
        }),
      ];
      const systemDefinition = designer.customSymbolDefinitionFromDesignerDocument(
        systemDocument,
        { id: "designer-system-glyph-roundtrip" },
      );
      registration.registerCustomSymbolSafely(systemDefinition);
      const systemArtifact = runtime.latexToSvg(BS + "vtxsystemglyph", {
        displayMode: false,
        fontSizePt: 12,
        paddingPx: 0,
        background: "transparent",
      });
      const systemPng = await runtime.svgToPng(systemArtifact, {
        scale: 2,
        background: "transparent",
      });
      const restoredSystem = designerArchive.restoreCustomSymbolDesignerDocument(
        systemDefinition,
      );
      const styledShape = styledSystemGlyph.asset.shapes[0] || null;
      const ordinaryShape = ordinarySystemGlyph.asset.shapes[0] || null;
      const searchResults = systemGlyphs.searchSystemGlyphs(
        "math-alphanumeric",
        "U+1D465",
        20,
      );
      const completeMathematicalAlphabet = systemGlyphs.searchSystemGlyphs(
        "math-alphanumeric",
        "",
      );
      const fontAvailability = await systemGlyphs.detectSystemMathFontAvailability();
      const systemGlyphProbe = {
        searchHasItalicX: searchResults.some(
          (glyph) => glyph.codePoint === 0x1d465,
        ),
        completeAlphabetCount: completeMathematicalAlphabet.length,
        completeAlphabetHasLastCodePoint: completeMathematicalAlphabet.some(
          (glyph) => glyph.codePoint === 0x1d7ff,
        ),
        styledVectorOutline: styledSystemGlyph.vectorOutline,
        styledShapeKind: styledShape?.kind || "",
        styledFontStyle:
          styledShape?.kind === "text" ? styledShape.fontStyle || "" : "path",
        ordinaryShapeKind: ordinaryShape?.kind || "",
        ordinaryFontStyle:
          ordinaryShape?.kind === "text" ? ordinaryShape.fontStyle || "" : "path",
        svgHasArtwork: systemArtifact.svg.includes(
          'data-visualtex-custom-symbol="designer-system-glyph-roundtrip"',
        ),
        svgHasSystemText: systemArtifact.svg.includes("<text "),
        pngType: systemPng.blob.type,
        pngSize: systemPng.blob.size,
        archiveLayerCount: systemDefinition.designerSource?.layers?.length ?? -1,
        restoredSourceMode: restoredSystem.sourceMode,
        restoredSourceLatex:
          restoredSystem.document.layers[0]?.kind === "glyph"
            ? restoredSystem.document.layers[0].asset.sourceLatex
            : "",
        availabilityKeys: Object.keys(fontAvailability).sort(),
      };

      const effectsDocument = designerTypes.createEmptyCustomSymbolDesignerDocument();
      effectsDocument.name = "Transformed outlined perspective glyph";
      effectsDocument.command = "vtxeffectglyph";
      const effectsLayer = designer.glyphLayerFromAsset(partialAsset, {
        id: "effect-partial",
        name: "effect partial",
      });
      const effectsCenterX = partialAsset.metrics.widthEm * 500;
      const effectsCenterY =
        (partialAsset.metrics.ascentEm + partialAsset.metrics.descentEm) * 500;
      effectsLayer.transform = {
        ...effectsLayer.transform,
        translateX: 120,
        translateY: -80,
        scaleX: -1.2,
        scaleY: 1.1,
        skewXDeg: -8,
        rotateDeg: 27,
        originX: effectsCenterX,
        originY: effectsCenterY,
      };
      effectsLayer.effects = {
        outline: { enabled: true, width: 84 },
        perspective: { enabled: true, depth: 180, angleDeg: 32, steps: 4 },
      };
      effectsDocument.metrics = { widthEm: 5, ascentEm: 4, descentEm: 2 };
      effectsDocument.layers = [effectsLayer];
      const localEffectsShapes = designer.customSymbolDesignerLayerLocalShapes(
        effectsLayer,
      );
      const effectsDefinition = designer.customSymbolDefinitionFromDesignerDocument(
        effectsDocument,
        { id: "designer-effects-roundtrip" },
      );
      const restoredEffects = designerArchive.restoreCustomSymbolDesignerDocument(
        effectsDefinition,
      );
      registration.registerCustomSymbolSafely(effectsDefinition);
      const effectsSvg = runtime.latexToSvg(BS + "vtxeffectglyph", {
        displayMode: false,
        fontSizePt: 12,
        paddingPx: 0,
        background: "transparent",
      }).svg;
      const frontEffectShape = localEffectsShapes.at(-1) || null;
      const restoredEffectsLayer = restoredEffects.document.layers[0] || null;
      const effectsProbe = {
        sourceShapeCount: partialAsset.shapes.length,
        localShapeCount: localEffectsShapes.length,
        expectedShapeCount: partialAsset.shapes.length * 5,
        frontFill: frontEffectShape?.fill,
        frontStrokeWidth: frontEffectShape?.strokeWidth ?? 0,
        compiledScaleX: effectsDefinition.artwork.shapes[0]?.transform?.scaleX ?? 0,
        compiledSkewX:
          effectsDefinition.artwork.shapes[0]?.transform?.skewXDeg ?? 0,
        compiledRotation:
          effectsDefinition.artwork.shapes[0]?.transform?.rotateDeg ?? 0,
        compiledOriginX:
          effectsDefinition.artwork.shapes[0]?.transform?.originX ?? null,
        restoredEffects:
          restoredEffectsLayer && "effects" in restoredEffectsLayer
            ? restoredEffectsLayer.effects || null
            : null,
        svgTransforms: effectsSvg.match(/transform="[^"]+"/g) || [],
        svgHasScaleFlip: effectsSvg.includes("scale(-1.2 1.1)"),
        svgHasSkew: effectsSvg.includes("skewX(-8)"),
        svgHasRotation: effectsSvg.includes("rotate(27)"),
      };

      const fitDocument = designerTypes.createEmptyCustomSymbolDesignerDocument();
      fitDocument.metrics = { widthEm: 10, ascentEm: 7, descentEm: 3 };
      const fitLayer = designer.glyphLayerFromAsset(partialAsset, {
        id: "fit-partial",
        name: "fit partial",
      });
      fitLayer.transform = {
        ...fitLayer.transform,
        translateX: 7200,
        translateY: 5200,
        rotateDeg: 18,
        originX: partialAsset.metrics.widthEm * 500,
        originY:
          (partialAsset.metrics.ascentEm + partialAsset.metrics.descentEm) * 500,
      };
      fitDocument.layers = [fitLayer];
      const fittedDocument = designer.fitCustomSymbolDesignerDocumentToArtwork(
        fitDocument,
      );
      const fitProbe = {
        sourceWidth: fitDocument.metrics.widthEm,
        fittedWidth: fittedDocument.metrics.widthEm,
        fittedHeight:
          fittedDocument.metrics.ascentEm + fittedDocument.metrics.descentEm,
        sourceTranslateX: fitLayer.transform.translateX,
        fittedTranslateX:
          fittedDocument.layers[0]?.transform.translateX ?? null,
        fittedTranslateY:
          fittedDocument.layers[0]?.transform.translateY ?? null,
      };

      const originalGetBBox = SVGGraphicsElement.prototype.getBBox;
      let fallbackStrokeProbe;
      SVGGraphicsElement.prototype.getBBox = function (...args) {
        if (args.length > 0) {
          throw new Error("Simulated WebKit getBBox options fallback");
        }
        return originalGetBBox.apply(this, args);
      };
      try {
        const createFallbackDocument = (outlined) => {
          const documentState =
            designerTypes.createEmptyCustomSymbolDesignerDocument();
          documentState.metrics = { widthEm: 8, ascentEm: 6, descentEm: 2 };
          const layer = designer.glyphLayerFromAsset(partialAsset, {
            id: outlined ? "fallback-outline" : "fallback-plain",
            name: outlined ? "fallback outline" : "fallback plain",
          });
          layer.transform = {
            ...layer.transform,
            translateX: 1900,
            translateY: 1200,
            scaleX: 2.1,
            scaleY: 1.35,
            rotateDeg: 13,
          };
          if (outlined) {
            layer.effects = {
              outline: { enabled: true, width: 320 },
            };
          }
          documentState.layers = [layer];
          return documentState;
        };
        const plainDocument = createFallbackDocument(false);
        const outlinedDocument = createFallbackDocument(true);
        const fittedPlain = designer.fitCustomSymbolDesignerDocumentToArtwork(
          plainDocument,
        );
        const fittedOutlined = designer.fitCustomSymbolDesignerDocumentToArtwork(
          outlinedDocument,
        );
        const plainDefinition = designer.customSymbolDefinitionFromDesignerDocument(
          plainDocument,
          { id: "fallback-plain-definition" },
        );
        const outlinedDefinition =
          designer.customSymbolDefinitionFromDesignerDocument(outlinedDocument, {
            id: "fallback-outline-definition",
          });
        fallbackStrokeProbe = {
          fittedPlainWidth: fittedPlain.metrics.widthEm,
          fittedOutlinedWidth: fittedOutlined.metrics.widthEm,
          fittedPlainHeight:
            fittedPlain.metrics.ascentEm + fittedPlain.metrics.descentEm,
          fittedOutlinedHeight:
            fittedOutlined.metrics.ascentEm + fittedOutlined.metrics.descentEm,
          runtimePlainWidth: plainDefinition.metrics.widthEm,
          runtimeOutlinedWidth: outlinedDefinition.metrics.widthEm,
          runtimePlainHeight:
            plainDefinition.metrics.ascentEm + plainDefinition.metrics.descentEm,
          runtimeOutlinedHeight:
            outlinedDefinition.metrics.ascentEm +
            outlinedDefinition.metrics.descentEm,
        };
      } finally {
        SVGGraphicsElement.prototype.getBBox = originalGetBBox;
      }

      for (const symbol of registry.readCustomSymbolLibrary().symbols) {
        registry.deleteCustomSymbol(symbol.id);
      }
      return {
        results: output,
        designerProbe,
        geometryProbe,
        tightCropProbe,
        systemGlyphProbe,
        effectsProbe,
        fitProbe,
        fallbackStrokeProbe,
      };
    })()`);

    const {
      results,
      designerProbe,
      geometryProbe,
      tightCropProbe,
      systemGlyphProbe,
      effectsProbe,
      fitProbe,
      fallbackStrokeProbe,
    } = probe;
    assert.equal(results.length, 7);
    for (const result of results) {
      assert.ok(result.shapeCount > 0, `${result.name} should compile vector shapes`);
      assert.ok(result.kinds.every((kind) => ["path", "circle", "ellipse", "line", "rect", "polygon"].includes(kind)));
      assert.equal(result.roundTrip.hasArtwork, true, `${result.name} custom artwork marker`);
      assert.equal(result.roundTrip.containsText, false, `${result.name} must remain pure vector`);
      assert.ok(result.metrics.widthEm > 0);
      assert.ok(result.metrics.ascentEm > 0);
      for (const matrix of result.matrices.filter(Boolean)) {
        assert.equal(matrix.length, 6);
        assert.ok(matrix.every(Number.isFinite));
        const determinant = matrix[0] * matrix[3] - matrix[1] * matrix[2];
        assert.ok(Math.abs(determinant) > 0.0000001);
      }
      assert.ok(
        Math.abs(result.original.width - result.roundTrip.width) < 0.05,
        `${result.name} width changed: ${JSON.stringify(result)}`,
      );
      assert.ok(
        Math.abs(result.original.height - result.roundTrip.height) < 0.05,
        `${result.name} height changed: ${JSON.stringify(result)}`,
      );
      assert.ok(
        Math.abs(result.original.baseline - result.roundTrip.baseline) < 0.05,
        `${result.name} baseline changed: ${JSON.stringify(result)}`,
      );
    }
    assert.ok(
      results.find((result) => result.name === "fragment")?.shapeCount >= 2,
      "A multi-glyph LaTeX fragment must flatten into multiple editable shapes",
    );
    assert.equal(designerProbe.sourcePathPreserved, true);
    assert.equal(designerProbe.shapeCount, 2, "Hidden designer layers must not compile");
    assert.equal(designerProbe.clipCount, 2);
    assert.equal(designerProbe.matrixCount, 2);
    assert.equal(designerProbe.hasTwoCropViewports, true);
    assert.equal(designerProbe.cropInsideUserTransform, true);
    assert.equal(designerProbe.hasArtwork, true);
    assert.equal(designerProbe.archiveAssetCount, 1);
    assert.equal(designerProbe.archiveLayerCount, 3);
    assert.equal(designerProbe.restoredSourceMode, "editable");
    assert.equal(designerProbe.restoredLayerCount, 3);
    assert.equal(designerProbe.restoredHiddenCount, 1);
    assert.equal(designerProbe.restoredClipCount, 2);
    assert.deepEqual(designerProbe.restoredGlyphSources, ["\\partial", "\\partial", "\\partial"]);
    assert.equal(designerProbe.restoredArtworkShapeCount, designerProbe.shapeCount);
    assert.notEqual(
      designerProbe.firstTransform.translateX,
      35,
      "Runtime artwork should be normalized into its tight registered box",
    );
    assert.equal(designerProbe.restoredFirstTransform.translateX, 35);
    assert.equal(designerProbe.restoredFirstTransform.translateY, -18);
    assert.equal(designerProbe.restoredFirstTransform.scaleX, 1.08);
    assert.equal(designerProbe.restoredFirstTransform.scaleY, 0.94);
    assert.equal(designerProbe.restoredFirstTransform.rotateDeg, 7);
    assert.ok(Array.isArray(designerProbe.firstTransform.matrix));
    assert.ok(tightCropProbe.runtimeWidth < 2, `Wide designer canvas leaked into runtime width: ${JSON.stringify(tightCropProbe)}`);
    assert.ok(tightCropProbe.runtimeWidth < tightCropProbe.designerWidth * 0.34);
    assert.equal(tightCropProbe.archivedWidth, 6);
    assert.equal(tightCropProbe.restoredWidth, 6);
    assert.equal(tightCropProbe.restoredTranslateX, 220);
    assert.notEqual(tightCropProbe.runtimeTranslateX, 220);
    assert.equal(geometryProbe.shapeCount, 4);
    assert.equal(geometryProbe.hasArtwork, true);
    assert.equal(geometryProbe.hasCircle, true);
    assert.equal(geometryProbe.hasLine, true);
    assert.equal(geometryProbe.hasArrowPath, true);
    assert.equal(geometryProbe.hasCrop, true);
    assert.match(geometryProbe.formattedLatex, /\\vtxgeometrymix/);
    assert.match(geometryProbe.formattedLatex, /^\\\[/);
    assert.match(geometryProbe.formattedLatex, /\\\]$/);
    assert.equal(geometryProbe.pngType, "image/png");
    assert.ok(geometryProbe.pngSize > 100);
    assert.ok(geometryProbe.pngWidth > 0 && geometryProbe.pngHeight > 0);
    assert.equal(geometryProbe.wordCanonicalLatex, "\\vtxgeometrymix");
    assert.equal(geometryProbe.wordSvgHasArtwork, true);
    assert.match(geometryProbe.wordOmml, /≈/);
    assert.doesNotMatch(geometryProbe.wordOmml, /vtxgeometrymix/);
    assert.equal(geometryProbe.powerpointCanonicalLatex, "\\vtxgeometrymix");
    assert.equal(geometryProbe.powerpointSvgHasArtwork, true);
    assert.equal(geometryProbe.powerpointOmmlIsNull, true);
    assert.equal(geometryProbe.archiveAssetCount, 1);
    assert.equal(geometryProbe.archiveLayerCount, 4);
    assert.equal(geometryProbe.restoredSourceMode, "editable");
    assert.deepEqual(geometryProbe.restoredLayerKinds, ["glyph", "geometry", "geometry", "geometry"]);
    assert.equal(geometryProbe.legacySourceMode, "flattened-legacy");
    assert.equal(geometryProbe.legacyLayerCount, geometryProbe.shapeCount);
    assert.ok(geometryProbe.legacyLayerKinds.every((kind) => kind === "geometry"));

    assert.equal(systemGlyphProbe.searchHasItalicX, true);
    assert.ok(systemGlyphProbe.completeAlphabetCount > 900);
    assert.equal(systemGlyphProbe.completeAlphabetHasLastCodePoint, true);
    assert.equal(systemGlyphProbe.styledVectorOutline, false);
    assert.equal(systemGlyphProbe.styledShapeKind, "text");
    assert.equal(
      systemGlyphProbe.styledFontStyle,
      "normal",
      "A mathematical-alphanumeric italic glyph must not receive a second synthetic italic slant",
    );
    assert.equal(systemGlyphProbe.ordinaryShapeKind, "text");
    assert.equal(
      systemGlyphProbe.ordinaryFontStyle,
      "italic",
      "An ordinary system-font letter must honor the requested mathematical italic style",
    );
    assert.equal(systemGlyphProbe.svgHasArtwork, true);
    assert.equal(systemGlyphProbe.svgHasSystemText, true);
    assert.equal(systemGlyphProbe.pngType, "image/png");
    assert.ok(systemGlyphProbe.pngSize > 100);
    assert.equal(systemGlyphProbe.archiveLayerCount, 1);
    assert.equal(systemGlyphProbe.restoredSourceMode, "editable");
    assert.equal(systemGlyphProbe.restoredSourceLatex, "𝑥");
    assert.deepEqual(systemGlyphProbe.availabilityKeys, [
      "cambria-math",
      "latin-modern-math",
      "stix-two-math",
      "system-serif",
      "xits-math",
    ]);

    assert.ok(effectsProbe.sourceShapeCount > 0);
    assert.equal(effectsProbe.localShapeCount, effectsProbe.expectedShapeCount);
    assert.equal(effectsProbe.frontFill, false);
    assert.ok(effectsProbe.frontStrokeWidth > 0);
    assert.equal(effectsProbe.compiledScaleX, -1.2);
    assert.equal(effectsProbe.compiledSkewX, -8);
    assert.equal(effectsProbe.compiledRotation, 27);
    assert.ok(Number.isFinite(effectsProbe.compiledOriginX));
    assert.deepEqual(effectsProbe.restoredEffects, {
      outline: { enabled: true, width: 84 },
      perspective: { enabled: true, depth: 180, angleDeg: 32, steps: 4 },
    });
    assert.equal(
      effectsProbe.svgHasScaleFlip,
      true,
      `Runtime SVG lost the horizontal flip: ${JSON.stringify(effectsProbe.svgTransforms)}`,
    );
    assert.equal(effectsProbe.svgHasSkew, true);
    assert.equal(effectsProbe.svgHasRotation, true);

    assert.ok(fitProbe.fittedWidth > 0);
    assert.ok(fitProbe.fittedHeight > 0);
    assert.ok(
      fitProbe.fittedWidth < fitProbe.sourceWidth * 0.4,
      `Fit-to-content did not tighten the designer output: ${JSON.stringify(fitProbe)}`,
    );
    assert.equal(fitProbe.sourceTranslateX, 7200);
    assert.notEqual(fitProbe.fittedTranslateX, fitProbe.sourceTranslateX);
    assert.ok(Number.isFinite(fitProbe.fittedTranslateX));
    assert.ok(Number.isFinite(fitProbe.fittedTranslateY));

    assert.ok(fallbackStrokeProbe);
    assert.ok(
      fallbackStrokeProbe.fittedOutlinedWidth >
        fallbackStrokeProbe.fittedPlainWidth + 0.2,
      `The WebKit getBBox fallback must reserve horizontal outline width: ${JSON.stringify(fallbackStrokeProbe)}`,
    );
    assert.ok(
      fallbackStrokeProbe.fittedOutlinedHeight >
        fallbackStrokeProbe.fittedPlainHeight + 0.2,
      `The WebKit getBBox fallback must reserve vertical outline width: ${JSON.stringify(fallbackStrokeProbe)}`,
    );
    assert.ok(
      fallbackStrokeProbe.runtimeOutlinedWidth >
        fallbackStrokeProbe.runtimePlainWidth + 0.2,
    );
    assert.ok(
      fallbackStrokeProbe.runtimeOutlinedHeight >
        fallbackStrokeProbe.runtimePlainHeight + 0.2,
    );

    console.log(
      "Custom symbol LaTeX/system glyph compilation, fit-to-content, transforms, outline/perspective, geometry composition, and editable round-trip regression passed",
    );
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    vite.kill("SIGTERM");
    await sleep(220);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => undefined);
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
