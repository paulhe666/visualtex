import assert from "node:assert/strict";
import {
  decodeOfficeBatchConversion,
  decodeOfficeFormulaSession,
  decodeOfficePreferences,
  decodeOfficeThemeStatus,
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
  objectMode: "nativeOle",
  numbered: false,
  mathTypeNumberPosition: "right",
  fontSizePt: 14,
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
{
  const legacy = clone(session) as Record<string, unknown>;
  delete legacy.objectMode;
  delete legacy.numbered;
  delete legacy.mathTypeNumberPosition;
  delete legacy.fontSizePt;
  const normalized = decodeOfficeFormulaSession(legacy);
  assert.equal(normalized.objectMode, "nativeOle");
  assert.equal(normalized.numbered, false);
  assert.equal(normalized.mathTypeNumberPosition, "right");
  assert.equal(normalized.fontSizePt, 14);
}
assert.equal(
  decodeOfficeFormulaSession({
    ...clone(session),
    exportResult: {
      svg: "<svg></svg>",
      svgBase64: "PHN2Zz48L3N2Zz4=",
      mathMl: "<math></math>",
      pngBase64: null,
      width: 120.5,
      height: 36,
      baseline: 24,
      formulaLetterFont: "katex",
      formulaChineseFont: "system",
    },
  }).status,
  "editing",
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
  ["invalid object mode", () => ({ ...clone(session), objectMode: "picture" })],
  ["invalid number position", () => ({ ...clone(session), mathTypeNumberPosition: "center" })],
  ["invalid font size", () => ({ ...clone(session), fontSizePt: 0 })],
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

assert.deepEqual(
  decodeOfficeThemeStatus({ theme: "dark", editorLayout: "standard" }),
  { theme: "dark", editorLayout: "standard" },
);
assert.throws(
  () => decodeOfficeThemeStatus({ theme: "dark", editorLayout: "wide" }),
  /officeTheme\.editorLayout/,
);

assert.deepEqual(
  decodeOfficePreferences({
    powerpointDefaultFontSizePt: 18,
    editorPreferences: {
      settings: { language: "cn" },
      customTheme: { id: "custom" },
    },
  }),
  {
    powerpointDefaultFontSizePt: 18,
    editorPreferences: {
      settings: { language: "cn" },
      customTheme: { id: "custom" },
    },
  },
);
assert.throws(
  () =>
    decodeOfficePreferences({
      powerpointDefaultFontSizePt: 18,
      editorPreferences: { settings: [] },
    }),
  /editorPreferences\.settings/,
);

assert.deepEqual(decodeOfficeBatchConversion({ sessionIds: ["a", "b"] }), {
  sessionIds: ["a", "b"],
});
assert.throws(
  () => decodeOfficeBatchConversion({ sessionIds: ["a", null] }),
  /sessionIds\[1\]/,
);

console.log("VisualTeX Windows Office Session payload safety regression passed");
