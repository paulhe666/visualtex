import assert from "node:assert/strict";
import {
  decodeOfficeFormulaSession,
  decodePreparedPowerPointCommit,
} from "../src/office/shared/sessionPayloadValidation.ts";

const session = {
  id: "session_0123456789abcdef",
  mode: "create",
  host: "word",
  formulaId: "4bf2217c-f29e-4f77-98b8-7258be8f63ae",
  sourceDocumentId: null,
  sourceObjectId: null,
  title: "",
  lines: [{ id: "line-1", latex: "x^2" }],
  activeLineId: "line-1",
  codeFormat: "latex",
  displayMode: "block",
  numbered: false,
  fontSizePt: 14,
  formulaLetterFont: "katex",
  formulaChineseFont: "system",
  exportWidth: 640,
  exportHeight: 180,
  exportResult: null,
  originalMetadata: null,
  dirty: false,
  status: "editing",
  autoCommitOnClose: true,
  explicitCancel: false,
  error: null,
  createdAt: 1,
  updatedAt: 2,
  expiresAt: 3,
};

const clone = <T>(value: T): T => JSON.parse(JSON.stringify(value)) as T;

assert.equal(decodeOfficeFormulaSession(session), session);
assert.equal(
  decodeOfficeFormulaSession({
    ...clone(session),
    operation: "nativeToImage",
    exportResult: {
      svg: "<svg></svg>",
      svgBase64: "PHN2Zz48L3N2Zz4=",
      pngBase64: null,
      ommlBase64: null,
      ommlDocxBase64: "UEsDBA==",
      width: 120.5,
      height: 36,
      baseline: 24,
      inkTopRatio: 0.1,
      inkBottomRatio: 0.9,
      inkCenterYRatio: 0.5,
    },
  }).operation,
  "nativeToImage",
);

for (const [label, mutate] of [
  ["null root", () => null],
  ["invalid mode", () => ({ ...clone(session), mode: "open" })],
  ["missing host", () => {
    const value = clone(session) as Record<string, unknown>;
    delete value.host;
    return value;
  }],
  ["malformed lines", () => ({ ...clone(session), lines: [null] })],
  ["invalid operation", () => ({ ...clone(session), operation: "convert" })],
  ["invalid font size", () => ({ ...clone(session), fontSizePt: 0 })],
  ["invalid letter font", () => ({ ...clone(session), formulaLetterFont: "unknown" })],
  ["non-finite geometry", () => ({ ...clone(session), exportWidth: Infinity })],
  ["invalid status", () => ({ ...clone(session), status: "done" })],
  ["invalid metadata", () => ({ ...clone(session), originalMetadata: [] })],
  ["invalid export", () => ({
    ...clone(session),
    exportResult: {
      svg: "<svg></svg>",
      svgBase64: "",
      width: -1,
      height: 20,
    },
  })],
] as const) {
  assert.throws(
    () => decodeOfficeFormulaSession(mutate()),
    /returned invalid data/,
    label,
  );
}

const prepared = {
  session,
  selection: {
    shapeName: "VisualTeX_1",
    slideIndex: 1,
    slideId: 42,
    presentationIdentity: "presentation-1",
    left: 10,
    top: 20,
    width: 100,
    height: 30,
  },
};
assert.equal(decodePreparedPowerPointCommit(prepared), prepared);
assert.throws(
  () =>
    decodePreparedPowerPointCommit({
      ...prepared,
      selection: { ...prepared.selection, width: "100" },
    }),
  /selection\.width/,
);

console.log("VisualTeX macOS Office Session payload safety regression passed");
