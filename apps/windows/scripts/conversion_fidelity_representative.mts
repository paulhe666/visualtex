import fs from 'node:fs';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { latexToMathMl } from '../src/export/runtime';

type DisplayMode = 'inline' | 'block';
type Formula = {
  index: number; id: string; latex: string; displayMode: DisplayMode; category: string;
  expectedKnownLimitations: string[]; mathMl?: string; parseStatus?: string;
  exportError?: string; mathMlSha256?: string; nativeExportStatus?: string;
  visualFidelityStatus?: string; sourceOnlyGateStatus?: string;
};
const output = path.resolve(process.argv[2] ?? '../../artifacts/conversion-symbol-audit-20261002/representative20');
const outputFiles = ['manifest.json', 'summary.json', 'case-list.json', 'excluded-diagnostics.json'];
for (const name of outputFiles) {
  if (fs.existsSync(path.join(output, name))) throw new Error('Refusing to replace evidence: ' + path.join(output, name));
}
const limitations = {
  alternate: 'Known prior codec loss of MathJax alternate-glyph presentation metadata. Scalar or signature equality does not establish short/ordinary glyph fidelity.',
  style: 'Explicit operator font/style requires both codec and native visual review; source mathvariant preservation alone is not a pixel-fidelity pass.',
  delimiter: 'These delimiter types had prior native tofu/missing-glyph evidence. Managed serialization success does not resolve native font authority or glyph availability.',
  scripts: 'Script structure and exact slot attachment require structural comparison and native review; text/scalar equality alone is insufficient.',
  prescript: 'Empty-base script forms must be gated individually. Prior source/OMML validator/native preparation restrictions can block a case without making the entire corpus untestable.',
};
const formulas: Formula[] = [];
function add(id: string, latex: string, displayMode: DisplayMode, category: string, expectedKnownLimitations: string[] = []) {
  formulas.push({ index: formulas.length + 1, id, latex, displayMode, category, expectedKnownLimitations });
}
add('F01', String.raw`\mathrm{a\parallel b\quad a\shortparallel b\quad a\nparallel b\quad a\nshortparallel b}`, 'inline', 'parallel-normal', [limitations.alternate, limitations.style]);
add('F02', String.raw`\mathit{a\parallel b\quad a\shortparallel b\quad a\nparallel b\quad a\nshortparallel b}`, 'inline', 'parallel-italic', [limitations.alternate, limitations.style]);
add('F03', String.raw`\mathbfit{a\parallel b\quad a\shortparallel b\quad a\nparallel b\quad a\nshortparallel b}`, 'block', 'parallel-bolditalic', [limitations.alternate, limitations.style]);
add('F04', String.raw`\mathrm{\angle ABC+\measuredangle DEF+\sphericalangle GHI}`, 'inline', 'angles-normal', [limitations.style]);
add('F05', String.raw`\mathit{\angle ABC+\measuredangle DEF+\sphericalangle GHI}`, 'inline', 'angles-italic', [limitations.style]);
add('F06', String.raw`\mathbfit{\angle ABC+\measuredangle DEF+\sphericalangle GHI}`, 'block', 'angles-bolditalic', [limitations.style]);
add('F07', String.raw`\left\ulcorner x\right\urcorner+\left\llcorner y\right\lrcorner`, 'block', 'missing-delimiters-corners', [limitations.delimiter]);
add('F08', String.raw`\left\lmoustache x\right\rmoustache+\left\bracevert y\right.`, 'block', 'missing-delimiters-moustaches-bracevert', [limitations.delimiter]);
add('F09', String.raw`\left\lgroup x\right\rgroup+\left\arrowvert y\right.`, 'block', 'missing-delimiters-groups-arrowvert', [limitations.delimiter]);
add('F10', String.raw`\emptyset\ne\varnothing\quad x^{\emptyset}_{\varnothing}`, 'inline', 'emptyset-varnothing-contrast', [limitations.alternate, limitations.scripts]);
add('F11', String.raw`\alpha+\upalpha+\mathrm{\alpha}+\Gamma+\mathit{\Gamma}`, 'inline', 'greek-default-upright-italic', ['Prior explicit-upright Greek style loss requires source and recovered style comparison, then native review.']);
add('F12', String.raw`x^{a_{b^c}}_{i^{j_k}}+y_{a^{b_c}}^{m_n}`, 'block', 'nested-super-subscript', [limitations.scripts]);
add('F13', String.raw`\left(x_i+y_j\right)^{n+1}`, 'block', 'whole-fence-superscript', [limitations.scripts]);
add('F14', String.raw`{\frac{a_i^2}{b_j}}^{n+1}_{k}`, 'block', 'fraction-base-super-subscript', [limitations.scripts]);
add('F15', String.raw`\sqrt[3]{x^{a_b}}+\sqrt{x_i^2+y_i^2}`, 'block', 'root-index-and-nested-scripts', [limitations.scripts]);
add('F16', String.raw`\sum_{i=1}^{n}x_i^2+\int_0^1x^2\,\mathrm{d}x`, 'block', 'sum-integral-limit-placement', [limitations.scripts]);
add('F17', String.raw`\left(\begin{matrix}a_i&b^2\\c_j&d_k^3\end{matrix}\right)^2`, 'block', 'matrix-fence-superscript', [limitations.scripts]);
add('F18', String.raw`\begin{cases}x^2&x>0\\x_2&x\le0\end{cases}`, 'block', 'cases-and-scripts', [limitations.scripts]);
add('F19', String.raw`{}^{14}_{6}\mathrm{C}`, 'inline', 'empty-base-isotope-prescript', [limitations.prescript, limitations.scripts]);
add('F20', String.raw`a+{}^{2}x`, 'inline', 'empty-base-leading-superscript', [limitations.prescript, limitations.scripts]);
if (formulas.length !== 20 || new Set(formulas.map(x => x.id)).size !== 20) throw new Error('Representative corpus must contain exactly 20 unique cases');
const excludedDiagnostics: Formula[] = [{
  index: 1, id: 'D01-half-ceiling-native-crash', latex: String.raw`\left\lceil x\right.`,
  displayMode: 'block', category: 'excluded-known-native-crash',
  expectedKnownLimitations: ['Known native half-ceiling renderer crash. Excluded from representative batch and from managed batch manifest; run only as a separately authorized isolated diagnostic, never infer safety from successful MathML export.'],
}];
for (const formula of [...formulas, ...excludedDiagnostics]) {
  try {
    formula.mathMl = await latexToMathMl(formula.latex, formula.displayMode === 'block');
    formula.parseStatus = 'passed';
    formula.mathMlSha256 = createHash('sha256').update(formula.mathMl).digest('hex');
  } catch (error) {
    formula.mathMl = '';
    formula.parseStatus = 'rejected';
    formula.exportError = String(error);
  }
  formula.nativeExportStatus = 'not-tested';
  formula.visualFidelityStatus = 'not-tested';
  formula.sourceOnlyGateStatus = 'awaiting-managed-codec-classification';
}
const counts = {
  cases: formulas.length, parsePassed: formulas.filter(f => f.parseStatus === 'passed').length,
  parseRejected: formulas.filter(f => f.parseStatus === 'rejected').length,
  excludedDiagnosticCases: excludedDiagnostics.length, nativeExportTested: 0, visualFidelityTested: 0,
};
const sourceModes = ['wordOmml', 'nativeOle', 'mathTypeOle'];
const directions = sourceModes.flatMap(source => sourceModes.filter(target => target !== source).map(target => ({ source, target })));
const scope = 'Twenty representative formula inputs exported through the production latexToMathMl function. Source/codec evidence only; not exhaustive symbol coverage and not native, Word, or pixel fidelity acceptance. The known half-ceiling native crash is excluded from this batch.';
const manifest = {
  schemaVersion: 1, acceptanceCorpus: true, createdUtc: new Date().toISOString(), scope,
  sourceExport: 'apps/windows/src/export/runtime::latexToMathMl',
  acceptedCandidate: 'artifacts/omml-edit-audit-20261001/build-cp194-owned-ole-lifetime-r2/VisualTeX.WordVsto.dll',
  sourceModes, directions, counts,
  requiredDelimiterCommands: ['ulcorner','urcorner','llcorner','lrcorner','lmoustache','rmoustache','bracevert','lgroup','rgroup','arrowvert'],
  gatePolicy: 'Classify each formula/route independently. A rejected preparation or serializer exception is a blocked case, not a corpus-wide failure or visual pass. Preserve original manifest; store observed codec status in a separate report.',
  formulas,
};
fs.mkdirSync(output, { recursive: true });
fs.writeFileSync(path.join(output, 'manifest.json'), JSON.stringify(manifest, null, 2));
fs.writeFileSync(path.join(output, 'case-list.json'), JSON.stringify(formulas.map(({id,latex,displayMode,category,expectedKnownLimitations}) => ({id,latex,displayMode,category,expectedKnownLimitations})), null, 2));
fs.writeFileSync(path.join(output, 'excluded-diagnostics.json'), JSON.stringify({ schemaVersion: 1, includeInRepresentativeBatch: false, nativeExecutionAuthorized: false, scope: 'Source export only; native diagnostic was not executed.', formulas: excludedDiagnostics }, null, 2));
fs.writeFileSync(path.join(output, 'summary.json'), JSON.stringify({ schemaVersion: 1, scope, counts, failures: formulas.filter(f => f.parseStatus === 'rejected'), files: outputFiles.map(name => path.join(output,name)) }, null, 2));
console.log(JSON.stringify({ counts, manifest: path.join(output, 'manifest.json'), cases: formulas.map(({id,category,parseStatus})=>({id,category,parseStatus})) }, null, 2));
if (counts.parseRejected) process.exitCode = 1;
