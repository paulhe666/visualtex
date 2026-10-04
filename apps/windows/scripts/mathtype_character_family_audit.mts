import { readFile, writeFile, mkdir } from "node:fs/promises";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { latexToMathMl } from "../src/export/runtime.ts";

// Use the production frontend, not handwritten mathvariant combinations, to
// define the input contract. Save exact inputs even when conversion fails.
const root = resolve(dirname(fileURLToPath(import.meta.url)), "../../..");
const output = resolve(root, "artifacts/mathtype-root-probes-20260930/mathjax-character-family-input.json");
const cs = await readFile(resolve(root, "apps/windows/src-windows/VisualTeX.WindowsOffice.VstoShared/MathMlToLatexConverter.cs"), "utf8");
const head = cs.slice(cs.indexOf("TokenMap ="), cs.indexOf("public static string Convert"));
const commands = new Set<string>();
for (const m of head.matchAll(/\["([^"]+)"\]\s*=\s*(@?)"([^"]*)"/gu)) {
  const latex = (m[2] ? m[3] : m[3].replaceAll("\\\\", "\\")).trim();
  if (latex) commands.add(latex);
}
for (const c of "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789") commands.add(c);
const shapes: [string,(s:string)=>string][] = [
  ["bare",s=>s],
  ["script",s=>String.raw`100\%{${s}}_{2}^{3}+7`],
  ...["dot","ddot","hat","tilde","bar","vec","overleftarrow","overleftrightarrow"].map(
    mark=>[mark,(s:string)=>`\\${mark}{${s}}`] as [string,(s:string)=>string]),
  ["bold-ddot",s=>String.raw`\boldsymbol{\ddot{${s}}}`],
  ["nested-accent",s=>String.raw`\hat{\ddot{${s}}}`],
  ...["mathrm","mathbf","mathit","mathbfit","mathcal","mathfrak","mathsf","mathtt","mathbb"].map(
    font=>[font,(s:string)=>`\\${font}{${s}}`] as [string,(s:string)=>string]),
];
const rows: {Id:string,Latex:string,MathMl:string}[] = [];
const rejected: {Id:string,Latex:string,Error:string}[] = [];
for (const command of [...commands].sort()) for (const [kind,shape] of shapes) {
  const latex = shape(command), id = `${kind}/${command}`;
  try {
    const mathMl = await latexToMathMl(latex, true);
    if (!/<math\b/u.test(mathMl) || /<merror\b/u.test(mathMl))
      throw new Error(`Frontend produced no valid math: ${mathMl}`);
    rows.push({Id:id,Latex:latex,MathMl:mathMl});
  } catch (error) { rejected.push({Id:id,Latex:latex,Error:String(error)}); }
}
await mkdir(dirname(output), {recursive:true});
await writeFile(output,JSON.stringify(rows,null,2),{flag:"wx"});
await writeFile(output.replace(".json","-rejected.json"),JSON.stringify(rejected,null,2),{flag:"wx"});
console.log(JSON.stringify({Commands:commands.size,Shapes:shapes.length,Accepted:rows.length,Rejected:rejected.length,Output:output}));
