import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, "..");
const adapterPath = path.join(root, "office/macos-offline/word/VTWordAdapter.bas");
const wordAdapter = fs.readFileSync(adapterPath, "utf8");

const failures = [];
const expect = (condition, message) => {
  if (!condition) failures.push(message);
};

function normalizeHeadingNumberText(value) {
  let text = String(value ?? "").trim().replaceAll("\t", "").replaceAll("\u00a0", "");
  while (text.length > 0 && ".-、:：".includes(text.at(-1))) {
    text = text.slice(0, -1);
  }
  return text.trim();
}

function resolveScopes(paragraphs, sequenceAnchors, mode) {
  const targetLevel = mode === "chapter" ? 1 : mode === "section" ? 2 : 0;
  if (!targetLevel) throw new Error(`unsupported mode: ${mode}`);

  const prefixes = new Array(sequenceAnchors.length).fill("");
  const localOrdinals = new Array(sequenceAnchors.length).fill(0);
  let chapterCount = 0;
  let sectionCount = 0;
  let currentPrefix = "";
  let currentSectionPrefix = "";
  let currentLocalOrdinal = 0;
  let itemIndex = 0;

  const assignFormula = () => {
    currentLocalOrdinal += 1;
    localOrdinals[itemIndex] = currentLocalOrdinal;
    if (targetLevel === 1) {
      prefixes[itemIndex] = currentPrefix || "0";
    } else if (currentSectionPrefix) {
      prefixes[itemIndex] = currentSectionPrefix;
    } else if (currentPrefix) {
      prefixes[itemIndex] = `${currentPrefix}.0`;
    } else {
      prefixes[itemIndex] = "0.0";
    }
    itemIndex += 1;
  };

  for (const paragraph of paragraphs) {
    while (itemIndex < sequenceAnchors.length && sequenceAnchors[itemIndex] <= paragraph.start) {
      assignFormula();
    }
    if (itemIndex >= sequenceAnchors.length) break;

    const level = paragraph.outlineLevel;
    if (level === 1) {
      chapterCount += 1;
      sectionCount = 0;
      const listText = normalizeHeadingNumberText(paragraph.listString);
      currentPrefix = listText || String(chapterCount);
      currentSectionPrefix = "";
      currentLocalOrdinal = 0;
    } else if (level === 2) {
      sectionCount += 1;
      const listText = normalizeHeadingNumberText(paragraph.listString);
      currentSectionPrefix =
        listText || `${currentPrefix || String(chapterCount > 0 ? chapterCount : 0)}.${sectionCount}`;
      if (targetLevel === 2) currentLocalOrdinal = 0;
    }
  }

  while (itemIndex < sequenceAnchors.length) assignFormula();
  return { prefixes, localOrdinals };
}

function expectResolution(name, paragraphs, anchors, mode, expectedPrefixes, expectedOrdinals) {
  const actual = resolveScopes(paragraphs, anchors, mode);
  expect(
    JSON.stringify(actual.prefixes) === JSON.stringify(expectedPrefixes),
    `${name}: prefix mismatch ${JSON.stringify(actual.prefixes)} != ${JSON.stringify(expectedPrefixes)}`,
  );
  expect(
    JSON.stringify(actual.localOrdinals) === JSON.stringify(expectedOrdinals),
    `${name}: local ordinal mismatch ${JSON.stringify(actual.localOrdinals)} != ${JSON.stringify(expectedOrdinals)}`,
  );
  return actual;
}

// Before any heading, formulas live in explicit zero scope rather than being
// treated as if chapter/section 1 already existed.
expectResolution(
  "before-first-chapter",
  [],
  [20, 40],
  "chapter",
  ["0", "0"],
  [1, 2],
);
expectResolution(
  "before-first-section",
  [],
  [20, 40],
  "section",
  ["0.0", "0.0"],
  [1, 2],
);

// Standard Heading 1 paragraphs with no automatic list numbering: VisualTeX's
// fallback heading count must produce 1-1,1-2,2-1,2-2.
const standardHeadings = [
  { start: 10, outlineLevel: 1, listString: "" },
  { start: 100, outlineLevel: 1, listString: "" },
];
expectResolution(
  "standard-heading-1",
  standardHeadings,
  [20, 40, 120, 140],
  "chapter",
  ["1", "1", "2", "2"],
  [1, 2, 1, 2],
);

// Multi-level/TOC numbering must use ListString for the visible prefix while
// still resetting the local Equation ordinal from the same paragraph boundary.
const listedHeadings = [
  { start: 10, outlineLevel: 1, listString: "1." },
  { start: 100, outlineLevel: 1, listString: "2、" },
];
expectResolution(
  "multilevel-list-chapter",
  listedHeadings,
  [20, 40, 120, 140],
  "chapter",
  ["1", "1", "2", "2"],
  [1, 2, 1, 2],
);

const sectionHeadings = [
  { start: 10, outlineLevel: 1, listString: "1." },
  { start: 30, outlineLevel: 2, listString: "1.1." },
  { start: 80, outlineLevel: 2, listString: "1.2." },
  { start: 120, outlineLevel: 1, listString: "2." },
  { start: 140, outlineLevel: 2, listString: "2.1." },
];
expectResolution(
  "multilevel-list-section",
  sectionHeadings,
  [40, 90, 150, 170],
  "section",
  ["1.1", "1.2", "2.1", "2.1"],
  [1, 1, 1, 2],
);

// A chapter without Heading 2 is section 0, not an implicit section 1. This is
// what prevents formulas before the first real section from duplicating 2.1.1.
expectResolution(
  "zero-section-per-chapter",
  standardHeadings,
  [20, 40, 60, 120],
  "section",
  ["1.0", "1.0", "1.0", "2.0"],
  [1, 2, 3, 1],
);

const zeroSectionHeadings = [
  { start: 10, outlineLevel: 1, listString: "1." },
  { start: 100, outlineLevel: 1, listString: "2." },
  { start: 160, outlineLevel: 2, listString: "2.1." },
];
expectResolution(
  "chapter-2-before-and-after-first-section",
  zeroSectionHeadings,
  [120, 140, 180, 200],
  "section",
  ["2.0", "2.0", "2.1", "2.1"],
  [1, 2, 1, 2],
);

// Issue #11 case: the paragraph style is intentionally irrelevant here; only
// OutlineLevel=1 is required, matching Word users who set an outline/TOC level
// without using the built-in Heading 1 style.
const customOutlineHeadings = [
  { start: 10, outlineLevel: 1, listString: "" },
  { start: 100, outlineLevel: 1, listString: "" },
];
expectResolution(
  "custom-outline-level",
  customOutlineHeadings,
  [20, 40, 120, 140],
  "chapter",
  ["1", "1", "2", "2"],
  [1, 2, 1, 2],
);

function formulaDisplayMap(formulas) {
  const result = new Map();
  for (const type of ["image", "omml"]) {
    const sorted = formulas
      .filter((item) => item.type === type)
      .sort((a, b) => a.anchor - b.anchor);
    const resolved = resolveScopes(
      customOutlineHeadings,
      sorted.map((item) => item.anchor),
      "chapter",
    );
    sorted.forEach((item, index) => {
      result.set(item.id, `${resolved.prefixes[index]}-${resolved.localOrdinals[index]}`);
    });
  }
  return result;
}

let formulas = [
  { id: "image-a", type: "image", anchor: 20 },
  { id: "omml-b", type: "omml", anchor: 40 },
  { id: "image-c", type: "image", anchor: 120 },
  { id: "omml-d", type: "omml", anchor: 140 },
];
let display = formulaDisplayMap(formulas);
expect(display.get("image-a") === "1-1", "mixed image/OMML: image-a must be 1-1");
expect(display.get("omml-b") === "1-1", "mixed image/OMML: omml-b must independently be 1-1");
expect(display.get("image-c") === "2-1", "mixed image/OMML: image-c must be 2-1");
expect(display.get("omml-d") === "2-1", "mixed image/OMML: omml-d must independently be 2-1");
const referenceTarget = "omml-d";
expect(display.get(referenceTarget) === "2-1", "body REF target must initially resolve to 2-1");

formulas.push({ id: "image-inserted", type: "image", anchor: 130 });
display = formulaDisplayMap(formulas);
expect(display.get("image-c") === "2-1", "second-chapter insertion must keep the first formula at 2-1");
expect(display.get("image-inserted") === "2-2", "second-chapter insertion must become 2-2");
expect(display.get("omml-d") === "2-1", "image insertion must not advance the OMML sequence");
expect(display.get(referenceTarget) === "2-1", "body REF identity must remain on the independent OMML sequence");

formulas = formulas.filter((item) => item.id !== "image-a" && item.id !== "image-c");
display = formulaDisplayMap(formulas);
expect(display.get("omml-b") === "1-1", "first-chapter deletion/update must compact to 1-1");
expect(display.get("image-inserted") === "2-1", "second-chapter deletion/update must compact to 2-1");
expect(display.get("omml-d") === "2-1", "image deletion must not renumber the OMML target");
expect(display.get(referenceTarget) === "2-1", "body REF identity must still point to the same OMML formula after deletion/update");

// Source-level invariants: the production VBA must implement exactly the same
// local-ordinal architecture rather than reverting to Word's Heading-style-only
// SEQ \\s behavior.
expect(wordAdapter.includes("ByRef localOrdinals() As Long"), "production scan must output local ordinals");
expect(wordAdapter.includes('headingPrefixes(itemIndex) = "0"'), "chapter mode must use chapter 0 before the first Heading 1");
expect(wordAdapter.includes('headingPrefixes(itemIndex) = currentPrefix & ".0"'), "section mode must use section 0 before the first Heading 2 of a chapter");
expect(wordAdapter.includes('headingPrefixes(itemIndex) = "0.0"'), "section mode must use 0.0 before any chapter or section");
expect(wordAdapter.includes("VTImageSequenceInstruction"), "image numbering must have a dedicated sequence-instruction builder");
expect(wordAdapter.includes('VTEquationSequenceFieldCodeForOrdinal & " \\r 1"'), "first formula in a scope must use native SEQ \\r 1");
expect(!wordAdapter.includes('" \\s " & CStr(restartLevel)'), "production numbering must not rely on native SEQ \\s heading detection");
expect(wordAdapter.includes("VTEquationSequenceResultText(sequenceField) = CStr(sequenceOrdinal)"), "SEQ validation must compare the exact expected local ordinal");
expect(wordAdapter.includes('regressionStage = "verify-two-real-chapters"'), "real-host regression must create and verify two real chapter boundaries");
expect(wordAdapter.includes('"initial=1-1,1-2,2-1,2-2"'), "real-host regression must record the Issue #11 expected sequence");
const mixedRegressionStart = wordAdapter.indexOf("Public Sub VisualTeX_RunWordMixedNativeImageChapterAppendRegression()");
const mixedRegressionEnd = wordAdapter.indexOf("Public Sub ", mixedRegressionStart + 16);
const mixedRegressionSource = wordAdapter.slice(
  mixedRegressionStart,
  mixedRegressionEnd > mixedRegressionStart ? mixedRegressionEnd : undefined,
);
expect(mixedRegressionStart >= 0, "packaged VBA must retain the dedicated mixed image/OMML numbering regression");
expect(mixedRegressionSource.includes("VTRegressionCreateNumberedNative"), "dedicated mixed regression must create native OMML formulas");
expect(mixedRegressionSource.includes("VTRegressionCreateNumberedImage"), "dedicated mixed regression must create image formulas");
expect(mixedRegressionSource.includes('"independentSequenceStreams=PASS"'), "dedicated mixed regression must verify independent image and OMML sequence streams");
expect(wordAdapter.includes('VT_WORD_IMAGE_SEQUENCE_NAME As String = "VisualTeXEquation"'), "production image numbering must use its private VisualTeXEquation stream");
expect(wordAdapter.includes('"standardHeading1=PASS"'), "real-host regression must cover built-in Heading 1");
expect(wordAdapter.includes('"customOutlineLevel1=PASS"'), "real-host regression must cover custom outline-level headings");
expect(wordAdapter.includes('"secondChapterInsertion=PASS"'), "real-host regression must cover insertion inside chapter 2");
expect(wordAdapter.includes('"deleteAndUpdate=PASS"'), "real-host regression must cover delete + Update Numbers in both chapters");
expect(wordAdapter.includes('"bodyCrossReference=PASS"'), "real-host regression must cover a surviving VisualTeX body REF");
expect(wordAdapter.includes('"wordBuiltInPrivateRef=PASS"'), "real-host regression must cover a Word-built-in private _Ref across a local SEQ reset");
expect(wordAdapter.includes("Private Function VTReferenceResultMatchesEquationNumber"), "private _Ref fallback must match the complete formatted Equation number");
expect(wordAdapter.includes("matchingFormulaCount = 1"), "private _Ref fallback must reject ambiguous formatted-number matches");
expect(wordAdapter.includes("If VTEquationNumberingMode(documentObject) = _\n       VT_WORD_NUMBERING_MODE_SEQUENCE Then"), "only plain sequence mode may use the legacy integer ordinal _Ref fallback");

const reconcileStart = wordAdapter.indexOf("Private Sub VTReconcileEquationNumbers");
const reconcileEnd = wordAdapter.indexOf("Private Function VTCountManagedEquationSequences", reconcileStart);
const reconcileSource = wordAdapter.slice(reconcileStart, reconcileEnd > reconcileStart ? reconcileEnd : undefined);
expect(!reconcileSource.includes("VTCaptureBodyEquationReferenceBindings"), "reconcile must not inspect Word-private _Ref bindings");
expect(!reconcileSource.includes("VTRestoreBodyEquationReferenceBindings"), "reconcile must not rebuild Word-private _Ref bindings");
expect(
  wordAdapter.includes("numberingMode = VT_WORD_NUMBERING_MODE_SEQUENCE And _\n        VTCanUseEquationTailFastPath"),
  "chapter/section tail appends must not bypass the shared scope scan",
);

if (failures.length > 0) {
  console.error("VisualTeX Word chapter scope numbering regression: FAIL");
  for (const failure of failures) console.error(`FAIL ${failure}`);
  process.exitCode = 1;
} else {
  console.log("VisualTeX Word chapter scope numbering regression: PASS");
  console.log("PASS standard Heading 1 => 1-1,1-2,2-1,2-2");
  console.log("PASS multi-level ListString chapter/section scopes");
  console.log("PASS custom OutlineLevel chapter reset (Issue #11)");
  console.log("PASS mixed image/OMML insertion, deletion/update, and REF identity model");
}
