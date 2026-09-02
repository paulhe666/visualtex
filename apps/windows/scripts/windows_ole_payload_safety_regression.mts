import assert from "node:assert/strict";
import {
  decodeOfficeBridgeEvents,
  decodeOfficeBridgeResponseEnvelope,
  decodeOfficeSelectionResult,
  decodeUpdatedEquationNumberResult,
} from "../src/office/windows-ole/windowsOlePayloadValidation.ts";

const requestId = "80feee7e-476a-4f83-91bf-51d02545bb45";
const response = {
  protocolVersion: 1,
  id: requestId,
  ok: true,
  result: { updated: 4 },
};
assert.equal(decodeOfficeBridgeResponseEnvelope(response, requestId), response);
assert.throws(
  () =>
    decodeOfficeBridgeResponseEnvelope(
      { ...response, protocolVersion: 2 },
      requestId,
    ),
  /response\.protocolVersion/,
);
assert.throws(
  () =>
    decodeOfficeBridgeResponseEnvelope(
      { ...response, id: "stale-request" },
      requestId,
    ),
  /response\.id/,
);
assert.throws(
  () =>
    decodeOfficeBridgeResponseEnvelope(
      {
        protocolVersion: 1,
        id: requestId,
        ok: false,
        error: { code: 7, message: "failed" },
      },
      requestId,
    ),
  /response\.error\.code/,
);

const metadata = {
  schema: "visualtex-formula",
  schemaVersion: 1,
  formulaId: "4bf2217c-f29e-4f77-98b8-7258be8f63ae",
  title: "",
  latex: "x^2",
  lines: [{ id: "line-1", latex: "x^2" }],
  codeFormat: "raw",
  displayMode: "inline",
  numbered: false,
  createdWithVersion: "1.2.5",
  updatedWithVersion: "1.2.5",
  createdAt: "2026-09-02T00:00:00.000Z",
  updatedAt: "2026-09-02T00:00:00.000Z",
} as const;
const selection = {
  host: "word",
  documentId: "document-1",
  objectId: "object-1",
  readOnly: false,
  formulaId: metadata.formulaId,
  metadata,
};
assert.equal(decodeOfficeSelectionResult(selection), selection);
assert.throws(
  () => decodeOfficeSelectionResult({ ...selection, readOnly: "false" }),
  /selection\.readOnly/,
);
assert.throws(
  () =>
    decodeOfficeSelectionResult({
      ...selection,
      formulaId: "f4ccaf97-b8bc-44fd-82ec-0446c67e0ae7",
    }),
  /selection\.metadata\.formulaId/,
);

assert.deepEqual(decodeUpdatedEquationNumberResult({ updated: 0 }), {
  updated: 0,
});
assert.throws(
  () => decodeUpdatedEquationNumberResult({ updated: -1 }),
  /updateEquationNumbers\.updated/,
);
assert.throws(
  () => decodeUpdatedEquationNumberResult({ updated: 1.5 }),
  /updateEquationNumbers\.updated/,
);

const events = [
  {
    protocolVersion: 1,
    event: "formula-edited",
    payload: null,
    cursor: 10,
  },
];
assert.deepEqual(decodeOfficeBridgeEvents(events), events);
assert.throws(
  () => decodeOfficeBridgeEvents({ events }),
  /events/,
);
assert.throws(
  () => decodeOfficeBridgeEvents([{ ...events[0], protocolVersion: 2 }]),
  /events\[0\]\.protocolVersion/,
);
assert.throws(
  () => {
    const { payload: _payload, ...withoutPayload } = events[0];
    decodeOfficeBridgeEvents([withoutPayload]);
  },
  /events\[0\]\.payload/,
);

console.log("VisualTeX Windows OLE payload safety regression passed");
