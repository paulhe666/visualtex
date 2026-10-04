import { readFileSync, writeFileSync } from 'node:fs';
import { latexToMathMl } from '../src/export/runtime.ts';
const base = '../../artifacts/conversion-symbol-audit-20261002/cp249-handoff/';
const sources = [String.raw`\sin x`, String.raw`\operatorname{sin}x`, String.raw`\mathit{\operatorname{sin}x}`, String.raw`\cos x`, String.raw`\operatorname{cos}x`, String.raw`\operatorname{foo}x`];
if (process.argv[2] === 'prepare') {
 const rows = [];
 for (let i = 0; i < sources.length; i++) rows.push({ id: 'P' + (i + 1), latex: sources[i], mathMl: await latexToMathMl(sources[i], false) });
 writeFileSync(base + 'function-input.json', JSON.stringify(rows, null, 2));
} else {
 const report = JSON.parse(readFileSync(base + 'function-product.json', 'utf8').replace(/^\uFEFF/, ''));
 for (const row of report.rows) {
  row.sourceReparsed = await latexToMathMl(row.sourceRecovered, false);
  row.nativeReparsed = await latexToMathMl(row.nativeRecovered, false);
 }
 report.nativeCollisions = report.rows.flatMap((a: any, i: number) => report.rows.slice(i + 1).filter((b: any) => a.nativeBase64 === b.nativeBase64).map((b: any) => [a.id, b.id]));
 writeFileSync(base + 'function-result.json', JSON.stringify(report, null, 2));
 console.log(JSON.stringify({rows: report.rows.map((r: any) => ({id:r.id,latex:r.latex,sourceRecovered:r.sourceRecovered,nativeRecovered:r.nativeRecovered})),nativeCollisions:report.nativeCollisions},null,2));
}
