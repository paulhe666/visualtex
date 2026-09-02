import { spawnSync } from "node:child_process";
import {
  existsSync,
  mkdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { homedir } from "node:os";
import { join, resolve } from "node:path";
import { inflateRawSync } from "node:zlib";

const repositoryRoot = resolve(new URL("..", import.meta.url).pathname);
const expectedDocumentName =
  process.argv.find((value) => value.startsWith("--document="))?.slice("--document=".length) ??
  "50_inline_50_display_meaningful.docx";
const createDocumentSnapshot = !process.argv.includes("--no-snapshot");
const scratch = join(
  homedir(),
  "Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch",
);
const runtime = join(
  homedir(),
  "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime",
);
const snapshotPath = join(scratch, "word-image-baseline-active-snapshot.docx");
const pdfPath = join(scratch, "word-image-baseline-active.pdf");
const reportPath = join(scratch, "word-image-baseline-active.json");
const requestPath = join(runtime, "document-import-regression-pdf-path.txt");
const statusPath = join(runtime, "document-import-regression-pdf-status.txt");
mkdirSync(scratch, { recursive: true });
mkdirSync(runtime, { recursive: true });

function appleScript(lines, timeout = 120_000) {
  const result = spawnSync(
    "/usr/bin/osascript",
    lines.flatMap((line) => ["-e", line]),
    {
      encoding: "utf8",
      timeout,
      maxBuffer: 16 * 1024 * 1024,
    },
  );
  if (result.status !== 0) {
    throw new Error(
      result.stderr.trim() || result.stdout.trim() || "AppleScript failed",
    );
  }
  return result.stdout.trimEnd();
}

function parseScalar(value) {
  if (value === "missing value" || value === "") return null;
  if (value === "true") return true;
  if (value === "false") return false;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : value;
}

function decodeMetadata(encoded) {
  const prefix = "visualtex:v1:deflate:";
  if (!encoded.startsWith(prefix)) return null;
  const compressed = Buffer.from(encoded.slice(prefix.length), "base64url");
  return JSON.parse(inflateRawSync(compressed).toString("utf8"));
}

function collectWordLayout() {
  const raw = appleScript([
    'tell application "Microsoft Word"',
    'if not (exists active document) then error "Microsoft Word has no active document"',
    "set sourceDocument to active document",
    'set outputText to "DOC" & tab & (name of sourceDocument) & tab & (saved of sourceDocument) & tab & (count of inline shapes of sourceDocument) & tab & (count of math objects of sourceDocument) & linefeed',
    "repeat with shapeIndex from 1 to (count of inline shapes of sourceDocument)",
    "set formulaShape to inline shape shapeIndex of sourceDocument",
    "set formulaRange to text object of formulaShape",
    "set formulaFont to font object of formulaRange",
    "set paragraphRange to text object of paragraph 1 of formulaRange",
    'set outputText to outputText & "IMAGE" & tab & shapeIndex & tab & (start of content of formulaRange) & tab & (end of content of formulaRange) & tab & (start of content of paragraphRange) & tab & (end of content of paragraphRange) & tab & (size of formulaFont) & tab & (font position of formulaFont) & tab & (width of formulaShape) & tab & (height of formulaShape) & tab & (get range information formulaRange information type horizontal position relative to page) & tab & (get range information formulaRange information type vertical position relative to page) & tab & (title of formulaShape) & tab & (alternative text of formulaShape) & linefeed',
    "end repeat",
    "repeat with mathIndex from 1 to (count of math objects of sourceDocument)",
    "set formulaMath to math object mathIndex of sourceDocument",
    "set formulaRange to text range of formulaMath",
    "set formulaFont to font object of formulaRange",
    "set paragraphRange to text object of paragraph 1 of formulaRange",
    "set formulaEndPosition to end of content of formulaRange",
    "set formulaEndRange to create range sourceDocument start formulaEndPosition end formulaEndPosition",
    'set outputText to outputText & "OMML" & tab & mathIndex & tab & (start of content of formulaRange) & tab & (end of content of formulaRange) & tab & (start of content of paragraphRange) & tab & (end of content of paragraphRange) & tab & (size of formulaFont) & tab & (font position of formulaFont) & tab & (display type of formulaMath) & tab & (get range information formulaRange information type horizontal position relative to page) & tab & (get range information formulaEndRange information type horizontal position relative to page) & tab & (get range information formulaRange information type vertical position relative to page) & linefeed',
    "end repeat",
    "return outputText",
    "end tell",
  ]);

  const lines = raw.split(/\r?\n/).filter(Boolean);
  const [documentLine, ...objectLines] = lines;
  const documentFields = documentLine.split("\t");
  if (documentFields[0] !== "DOC") {
    throw new Error(`Unexpected Word probe header: ${documentLine}`);
  }
  const document = {
    name: documentFields[1],
    saved: parseScalar(documentFields[2]),
    inlineShapeCount: parseScalar(documentFields[3]),
    ommlCount: parseScalar(documentFields[4]),
  };
  const images = [];
  const omml = [];
  for (const line of objectLines) {
    const fields = line.split("\t");
    if (fields[0] === "IMAGE") {
      const metadata = decodeMetadata(fields[13] ?? "");
      images.push({
        index: parseScalar(fields[1]),
        rangeStart: parseScalar(fields[2]),
        rangeEnd: parseScalar(fields[3]),
        paragraphStart: parseScalar(fields[4]),
        paragraphEnd: parseScalar(fields[5]),
        fontSizePt: parseScalar(fields[6]),
        fontPositionPt: parseScalar(fields[7]),
        widthPt: parseScalar(fields[8]),
        heightPt: parseScalar(fields[9]),
        pageXPt: parseScalar(fields[10]),
        pageYPt: parseScalar(fields[11]),
        title: fields[12] ?? "",
        alternativeText: fields[13] ?? "",
        metadata,
      });
    } else if (fields[0] === "OMML") {
      omml.push({
        index: parseScalar(fields[1]),
        rangeStart: parseScalar(fields[2]),
        rangeEnd: parseScalar(fields[3]),
        paragraphStart: parseScalar(fields[4]),
        paragraphEnd: parseScalar(fields[5]),
        fontSizePt: parseScalar(fields[6]),
        fontPositionPt: parseScalar(fields[7]),
        displayType: fields[8] ?? "",
        pageXPt: parseScalar(fields[9]),
        pageEndXPt: parseScalar(fields[10]),
        pageYPt: parseScalar(fields[11]),
      });
    }
  }
  return { document, images, omml };
}

function createSnapshot() {
  rmSync(snapshotPath, { force: true });
  appleScript([
    'tell application "Microsoft Word"',
    "set sourceDocument to active document",
    `if (name of sourceDocument) is not ${JSON.stringify(expectedDocumentName)} then error "Unexpected active document: " & (name of sourceDocument)`,
    "set sourceRange to text object of sourceDocument",
    "copy object sourceRange",
    "set probeDocument to make new document",
    "paste object text object of probeDocument",
    `save as probeDocument file name ${JSON.stringify(snapshotPath)} file format format document default add to recent files false`,
    "close probeDocument saving no",
    "activate object sourceDocument",
    "end tell",
  ]);
  if (!existsSync(snapshotPath)) {
    throw new Error(`Word did not create the baseline snapshot: ${snapshotPath}`);
  }
}

function exportActivePdf() {
  rmSync(pdfPath, { force: true });
  rmSync(statusPath, { force: true });
  writeFileSync(requestPath, pdfPath, { mode: 0o600 });
  appleScript([
    'tell application "Microsoft Word"',
    `if (name of active document) is not ${JSON.stringify(expectedDocumentName)} then error "Unexpected active document before PDF export"`,
    'run VB macro macro name "VisualTeX_ExportActiveDocumentPdfForRegression"',
    "end tell",
  ]);
  const status = existsSync(statusPath)
    ? readFileSync(statusPath, "utf8").trim()
    : "missing";
  if (!status.startsWith("ok|") || !existsSync(pdfPath)) {
    throw new Error(`PDF export failed: ${status}`);
  }
  return status;
}

function collectRasterGeometry() {
  const result = spawnSync(
    "/usr/bin/swift",
    [join(repositoryRoot, "scripts/pdf_formula_geometry.swift"), pdfPath, "--raster-only"],
    {
      encoding: "utf8",
      timeout: 120_000,
      maxBuffer: 16 * 1024 * 1024,
    },
  );
  if (result.status !== 0) {
    throw new Error(
      result.stderr.trim() || result.stdout.trim() || "PDF raster geometry probe failed",
    );
  }
  return JSON.parse(result.stdout);
}

function rasterInkBounds(
  rasterGeometry,
  pageXPt,
  pageEndXPt,
  targetCenterYPt,
  verticalRadiusPt,
) {
  if (
    ![pageXPt, pageEndXPt, targetCenterYPt, verticalRadiusPt].every(Number.isFinite) ||
    pageEndXPt <= pageXPt ||
    verticalRadiusPt <= 0
  ) {
    return null;
  }
  const components = (
    rasterGeometry.rasterComponents ??
    (rasterGeometry.rasterBands ?? []).flatMap((band) => band.components ?? [])
  ).filter(
    (component) =>
      component.centerX >= pageXPt - 0.75 &&
      component.centerX <= pageEndXPt + 0.75 &&
      component.centerY >= targetCenterYPt - verticalRadiusPt &&
      component.centerY <= targetCenterYPt + verticalRadiusPt,
  );
  if (components.length === 0) return null;
  const minX = Math.min(...components.map((component) => component.minX));
  const minY = Math.min(...components.map((component) => component.minY));
  const maxX = Math.max(...components.map((component) => component.maxX));
  const maxY = Math.max(...components.map((component) => component.maxY));
  return {
    minX,
    minY,
    maxX,
    maxY,
    width: maxX - minX,
    height: maxY - minY,
    centerY: (minY + maxY) / 2,
    targetCenterYPt,
    verticalRadiusPt,
    componentCount: components.length,
  };
}

function sameLinePairs(images, omml, rasterGeometry) {
  const visualImages = images.filter((image) => image.metadata?.schema === "visualtex-formula");
  return omml.map((nativeEquation) => {
    const candidates = visualImages
      .filter(
        (image) =>
          typeof image.pageYPt === "number" &&
          typeof nativeEquation.pageYPt === "number" &&
          Math.abs(image.pageYPt - nativeEquation.pageYPt) <= 0.1,
      )
      .map((image) => ({
        image,
        distancePt:
          typeof image.pageXPt === "number" && typeof nativeEquation.pageXPt === "number"
            ? Math.abs(image.pageXPt - nativeEquation.pageXPt)
            : Number.POSITIVE_INFINITY,
      }))
      .sort((left, right) => left.distancePt - right.distancePt);
    const nearestImage = candidates[0]?.image ?? null;
    const imageInk = nearestImage && typeof nearestImage.heightPt === "number"
      ? rasterInkBounds(
          rasterGeometry,
          nearestImage.pageXPt,
          typeof nearestImage.widthPt === "number"
            ? nearestImage.pageXPt + nearestImage.widthPt
            : null,
          nearestImage.pageYPt + nearestImage.heightPt / 2,
          Math.max(3, nearestImage.heightPt / 2 + 2),
        )
      : null;
    const ommlInk = rasterInkBounds(
      rasterGeometry,
      nativeEquation.pageXPt,
      nativeEquation.pageEndXPt,
      imageInk?.centerY ?? nativeEquation.pageYPt,
      Math.max(4, (nearestImage?.heightPt ?? 0) / 2 + 3),
    );
    const inkDelta = ommlInk && imageInk
      ? {
          topPt: imageInk.minY - ommlInk.minY,
          bottomPt: imageInk.maxY - ommlInk.maxY,
          centerPt: imageInk.centerY - ommlInk.centerY,
          heightPt: imageInk.height - ommlInk.height,
        }
      : null;
    return {
      omml: nativeEquation,
      nearestImage,
      horizontalDistancePt: Number.isFinite(candidates[0]?.distancePt)
        ? candidates[0].distancePt
        : null,
      ommlInk,
      imageInk,
      inkDelta,
    };
  });
}

const layout = collectWordLayout();
if (layout.document.name !== expectedDocumentName) {
  throw new Error(
    `Expected ${expectedDocumentName}, found ${layout.document.name}. Keep the regression document active.`,
  );
}
if (createDocumentSnapshot) createSnapshot();
const pdfStatus = exportActivePdf();
const rasterGeometry = collectRasterGeometry();
const pairs = sameLinePairs(layout.images, layout.omml, rasterGeometry);
const report = {
  generatedAt: new Date().toISOString(),
  document: layout.document,
  paths: {
    snapshotPath: createDocumentSnapshot ? snapshotPath : null,
    pdfPath,
    reportPath,
  },
  pdfStatus,
  images: layout.images,
  omml: layout.omml,
  rasterGeometry,
  sameLinePairs: pairs,
};
writeFileSync(reportPath, `${JSON.stringify(report, null, 2)}\n`, "utf8");

console.log(
  JSON.stringify(
    {
      document: layout.document,
      visualTexImageCount: layout.images.filter(
        (image) => image.metadata?.schema === "visualtex-formula",
      ).length,
      ommlCount: layout.omml.length,
      sameLinePairs: pairs.map((pair) => ({
        omml: pair.omml,
        image: pair.nearestImage
          ? {
              index: pair.nearestImage.index,
              rangeStart: pair.nearestImage.rangeStart,
              rangeEnd: pair.nearestImage.rangeEnd,
              paragraphStart: pair.nearestImage.paragraphStart,
              paragraphEnd: pair.nearestImage.paragraphEnd,
              fontSizePt: pair.nearestImage.fontSizePt,
              fontPositionPt: pair.nearestImage.fontPositionPt,
              widthPt: pair.nearestImage.widthPt,
              heightPt: pair.nearestImage.heightPt,
              pageXPt: pair.nearestImage.pageXPt,
              pageYPt: pair.nearestImage.pageYPt,
              metadata: pair.nearestImage.metadata
                ? {
                    fontSizePt: pair.nearestImage.metadata.fontSizePt,
                    referenceWidthPt: pair.nearestImage.metadata.referenceWidthPt,
                    referenceHeightPt: pair.nearestImage.metadata.referenceHeightPt,
                    referenceBaselinePt: pair.nearestImage.metadata.referenceBaselinePt,
                    renderWidthPx: pair.nearestImage.metadata.renderWidthPx,
                    renderHeightPx: pair.nearestImage.metadata.renderHeightPx,
                    latex: pair.nearestImage.metadata.lines
                      ?.map((line) => line.latex)
                      .join("\\\\"),
                  }
                : null,
            }
          : null,
        horizontalDistancePt: pair.horizontalDistancePt,
        ommlInk: pair.ommlInk,
        imageInk: pair.imageInk,
        inkDelta: pair.inkDelta,
      })),
      snapshotPath: createDocumentSnapshot ? snapshotPath : null,
      pdfPath,
      reportPath,
    },
    null,
    2,
  ),
);
