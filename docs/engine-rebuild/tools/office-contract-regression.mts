import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { decodeOfficeFormulaSession } from '../../../apps/windows/src/office/shared/sessionPayloadValidation.ts';
import { createFormulaMetadata, encodeFormulaMetadata, decodeFormulaMetadata,
  formulaMetadataToXml, formulaMetadataFromXml, isVisualTeXFormulaMetadata,
  VISUALTEX_METADATA_PREFIX } from '../../../apps/windows/src/office/shared/formulaMetadata.ts';
import { getOfficeSession, createOfficeSession, updateOfficeSession, commitWindowsOfficeSession }
  from '../../../apps/windows/src/office/shared/sessionClient.ts';

// Fixed times and identities. Geometry, source, fingerprint and numbering are NOT normalized away.
const metadata = {
  schema: 'visualtex-formula' as const, schemaVersion: 1 as const,
  formulaId: '4bf2217c-f29e-4f77-98b8-7258be8f63ae', title: 'engine-boundary-fixture',
  latex: String.raw`\frac{x_1+\text{速度}}{y^2}`, lines: [{ id: 'line-1', latex: String.raw`\frac{x_1+\text{速度}}{y^2}` }],
  codeFormat: 'latex', displayMode: 'inline' as const, numbered: false,
  renderWidthPx: 96, renderHeightPx: 48, baseline: 31.25,
  fontSizePt: 13.25, renderFontSizePt: 13.25, wordInlineOleWidthPt: 72, wordInlineOleHeightPt: 36,
  nativeOmmlFingerprint: createHash('sha256').update('engine-contract-owned-test-source').digest('hex'), createdWithVersion: '1.2.7', updatedWithVersion: '1.2.7',
  createdAt: '2026-09-13T00:00:00.000Z', updatedAt: '2026-09-13T00:00:00.000Z',
};
const exportResult = { svg: '<svg xmlns="http://www.w3.org/2000/svg" width="96" height="48" viewBox="0 0 96 48"><path d="M0 31.25H96"/></svg>',
  svgBase64: '', mathMl: '<math><mi>x</mi></math>', width: 96, height: 48, baseline: 31.25 };
exportResult.svgBase64 = Buffer.from(exportResult.svg).toString('base64');
const fixture = {
  id: 'session_engine_boundary_1', mode: 'edit', host: 'word', formulaId: metadata.formulaId,
  sourceDocumentId: 'owned-test-document', sourceObjectId: 'existing-object-42', title: metadata.title,
  lines: metadata.lines, activeLineId: 'line-1', codeFormat: 'latex', displayMode: 'inline', objectMode: 'nativeOle',
  numbered: false, mathTypeNumberPosition: 'right', fontSizePt: 13.25, exportWidth: 96, exportHeight: 48,
  exportResult, originalMetadata: metadata, dirty: false, status: 'editing', autoCommitOnClose: false,
  explicitCancel: false, error: null, createdAt: 1, updatedAt: 2, expiresAt: 3,
};
let assertions = 0;
function check(name: string, f: () => void) { f(); assertions++; console.log(`PASS ${name}`); }
check('real payload decoder retains all semantic data', () => assert.deepEqual(decodeOfficeFormulaSession(structuredClone(fixture)), fixture));
check('metadata compressed source/identity/geometry roundtrip', () => {
  const encoded = encodeFormulaMetadata(metadata);
  assert.ok(encoded.startsWith(VISUALTEX_METADATA_PREFIX));
  assert.deepEqual(decodeFormulaMetadata(encoded), metadata);
  assert.deepEqual(formulaMetadataFromXml(formulaMetadataToXml(metadata)), metadata);
});
check('unchanged native source and dimensions survive metadata refresh', () => {
  const next = createFormulaMetadata({ formulaId: metadata.formulaId, title: metadata.title, lines: metadata.lines,
    codeFormat: 'latex', displayMode: 'inline', original: metadata });
  const { updatedAt, ...rest } = next;
  const { updatedAt: ignored, ...expected } = metadata;
  assert.deepEqual(rest, expected);
  assert.ok(updatedAt.length > 0);
});
check('optional/null export semantics and MathML-only path', () => {
  const sparse = structuredClone(fixture);
  sparse.objectMode = 'wordOmml'; sparse.exportResult.svg = ''; sparse.exportResult.svgBase64 = '';
  assert.equal(decodeOfficeFormulaSession(sparse).exportResult?.mathMl, '<math><mi>x</mi></math>');
  assert.equal(decodeOfficeFormulaSession({ ...fixture, exportResult: null }).exportResult, null);
});
check('legacy defaults retain current implementation semantics', () => {
  const legacy: any = structuredClone(fixture);
  for (const key of ['objectMode', 'numbered', 'mathTypeNumberPosition', 'fontSizePt']) delete legacy[key];
  const value = decodeOfficeFormulaSession(legacy);
  assert.equal(value.objectMode, 'nativeOle'); assert.equal(value.numbered, false);
  assert.equal(value.mathTypeNumberPosition, 'right'); assert.equal(value.fontSizePt, 14);
});
check('invalid baseline rejected by real metadata validator', () => assert.equal(isVisualTeXFormulaMetadata({ ...metadata, baseline: 49 }), false));
check('field rename mutation rejected by real decoder', () => {
  const bad: any = structuredClone(fixture);
  bad.exportResult.svg_base64 = bad.exportResult.svgBase64; delete bad.exportResult.svgBase64;
  assert.throws(() => decodeOfficeFormulaSession(bad));
});
check('unit drift mutation detected without normalizing dimensions', () => {
  const wrong = decodeOfficeFormulaSession({ ...structuredClone(fixture), exportResult: { ...exportResult, width: 72 } });
  assert.throws(() => assert.deepEqual(wrong, fixture));
});

const originalFetch = globalThis.fetch;
const requests: { path: string; init: RequestInit }[] = [];
(globalThis as any).window = { __VISUALTEX_INSTALL_TOKEN__: 'synthetic-contract-test-token' };
globalThis.fetch = (async (path: any, init: RequestInit = {}) => {
  requests.push({ path: String(path), init });
  return new Response(JSON.stringify(fixture), { status: 200, headers: { 'content-type': 'application/json' } });
}) as typeof fetch;
try {
  const id = 'session/with space';
  await getOfficeSession(id);
  await createOfficeSession({ mode: 'edit', host: 'word', sourceDocumentId: 'owned-test-document', sourceObjectId: 'existing-object-42' });
  await updateOfficeSession(id, { exportResult, dirty: true });
  await commitWindowsOfficeSession(id);
  check('actual client routes, auth, methods and identity encoding', () => {
    assert.deepEqual(requests.map(r => [r.path, r.init.method ?? 'GET']), [
      ['/api/v1/sessions/session%2Fwith%20space', 'GET'], ['/api/v1/sessions', 'POST'],
      ['/api/v1/sessions/session%2Fwith%20space', 'PATCH'], ['/api/v1/windows/sessions/session%2Fwith%20space/commit', 'POST'],
    ]);
    for (const r of requests) {
      assert.equal(r.init.credentials, 'same-origin'); assert.equal(r.init.cache, 'no-store');
      assert.equal(new Headers(r.init.headers).get('X-VisualTeX-Install-Token'), 'synthetic-contract-test-token');
    }
    assert.deepEqual(JSON.parse(requests[2].init.body as string), { exportResult, dirty: true });
  });
} finally { globalThis.fetch = originalFetch; delete (globalThis as any).window; }
console.log(JSON.stringify({ status: 'PASS', checks: assertions, realOfficeActions: 0, fixtureDimensions: [96, 48, 31.25] }));
