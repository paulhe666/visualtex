// How many of VisualTeX's commands the Math Characters output writes
// exactly (apps/shared/visualtex-app/src/anywhere/unicodeMath.ts): every built-in command's preview
// formula, empty slots filled with x, through the export path's MathML.
// "exact": no conversion notes; "approx": linearized (notes); "error": the
// MathML path cannot read it (the output falls back to LaTeX).
//   tsx scripts/anywhere_unicode_survey.mts [out.json]
import { DOMParser } from "@xmldom/xmldom";
import { latexToMathMl } from "../src/export/runtime.ts";
import { mathMlToUnicode } from "../../shared/visualtex-app/src/anywhere/unicodeMath.ts";
import { commandRegistry } from "../src/autocomplete/commandRegistry.ts";
import { writeFileSync } from "node:fs";
const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml") as never;
const seen = new Set<string>();
const commands = commandRegistry.filter((c: any) => c.previewLatex && !seen.has(c.command) && seen.add(c.command));
const rows: any[] = [];
for (const c of commands as any[]) {
  // an empty slot holds a letter, as a typed formula would
  const latex = c.previewLatex.replace(/\\placeholder\{\}/g, "x");
  const row: any = { command: c.command, latex, category: c.category };
  try {
    const out = mathMlToUnicode(latexToMathMl(latex, false), parse);
    row.text = out.text; row.notes = out.notes;
    row.status = !out.text ? "empty" : out.notes.length ? "approx" : "exact";
  } catch (e) { row.status = "error"; row.error = String((e as Error)?.message ?? e).slice(0, 120); }
  rows.push(row);
}
writeFileSync(process.argv[2] ?? "anywhere-unicode-survey.json", JSON.stringify(rows, null, 1));
const count = (s: string) => rows.filter((r) => r.status === s).length;
console.log(`commands ${rows.length}: exact ${count("exact")}, approx ${count("approx")}, empty ${count("empty")}, error ${count("error")}`);
const notes = new Map<string, number>();
for (const r of rows) for (const n of r.notes ?? []) notes.set(n, (notes.get(n) ?? 0) + 1);
for (const [n, k] of [...notes].sort((a, b) => b[1] - a[1])) console.log(String(k).padStart(5), n);
