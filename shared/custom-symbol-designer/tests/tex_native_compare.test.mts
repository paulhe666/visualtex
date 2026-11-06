// The exported command against TeX itself (needs pdflatex; skipped without it):
//   npx tsx apps/shared/custom-symbol-designer/tests/tex_native_compare.test.mts
//   (Node 22: node --experimental-strip-types <file>)
// A symbol built on \displaystyle\int must set as \int does: same box, scripts
// and limits in the same place, \limits, \nolimits and automatic, display and
// text style, in any context. Two forms:
// - kept as the real operator (a dot over \int): checked in Computer Modern
//   and, when installed, in newtxmath (the form does not depend on the font);
// - drawn as a box (the base given an effect): checked in Computer Modern,
//   whose metrics the box carries.
// TeX measures each formula (width, height, depth); pixels are compared
// (ignoring the overlay) when mutool and Python's Pillow are installed.
import assert from "node:assert/strict";
import { execFileSync, spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { customSymbolTexExport, nativeOperatorBase } from "../src/tex.ts";
import { computeSymbolBox } from "../src/symbolBox.ts";
import type { CustomSymbolDesignerDocument, CustomSymbolDesignerLayer } from "../src/designerTypes.ts";
import type { CustomSymbolLimitsBehavior } from "../src/symbolTypes.ts";

const has = (command: string, args: string[]) => spawnSync(command, args, { stdio: "ignore" }).status === 0;
if (!has("pdflatex", ["--version"])) {
  console.log("custom symbol tex native compare: skipped (no pdflatex)");
  process.exit(0);
}
const kpse = (file: string) => spawnSync("kpsewhich", [file], { encoding: "utf8" }).stdout?.trim() ?? "";

// the engine's measurements of Computer Modern's integrals (as TeX's fonts give them)
const DISPLAY = { widthEm: 1, ascentEm: 1.36112, descentEm: 0.86112 };
const TEXT = { widthEm: 0.66667, ascentEm: 0.80556, descentEm: 0.30556 };
const ITALIC: Record<string, number> = { "\\displaystyle\\int": 0.44444, "\\textstyle\\int": 0.19444 };
const marker = [{ kind: "path" as const, d: "M0 0L1 0", fill: true, transform: { matrix: [1, 0, 0, -1, 0, 0] as [number, number, number, number, number, number] } }];

function integralCommand(command: string, limits: CustomSymbolLimitsBehavior, form: "native" | "box") {
  const layers: CustomSymbolDesignerLayer[] = [
    {
      id: "i",
      name: "int",
      kind: "glyph",
      visible: true,
      locked: false,
      transform: { translateX: 0, translateY: 3000 - DISPLAY.ascentEm * 1000 },
      clipRect: null,
      // an effect keeps the base from being the real operator: the box form
      ...(form === "box" ? { effects: { perspective: { enabled: true, depth: 0.001, angleDeg: 0, steps: 1 } } } : {}),
      asset: { sourceLatex: "\\displaystyle\\int", metrics: DISPLAY, shapes: marker },
    },
  ];
  if (form === "native") {
    // a dot over the integral, inside its box
    layers.push({
      id: "dot",
      name: "cdot",
      kind: "glyph",
      visible: true,
      locked: false,
      transform: { translateX: 300, translateY: 2600 },
      clipRect: null,
      asset: { sourceLatex: "\\cdot", metrics: { widthEm: 0.27778, ascentEm: 0.31, descentEm: 0 }, shapes: marker },
    });
  }
  const document: CustomSymbolDesignerDocument = {
    version: 1,
    symbolId: null,
    name: command,
    command,
    role: "operator",
    limitsBehavior: limits,
    metrics: { widthEm: 3.2, ascentEm: 3, descentEm: 1.5 },
    ommlFallback: null,
    layers,
  };
  const box = computeSymbolBox(document, { left: 50, top: 1700, right: 950, bottom: 3800 }, (s) => ITALIC[s]);
  assert.equal(Boolean(nativeOperatorBase(document, box)), form === "native", `${command}: ${form} form`);
  return customSymbolTexExport(document, box, {
    textStyle: (source) => (source === "\\displaystyle\\int" ? { metrics: TEXT, italicEm: ITALIC["\\textstyle\\int"] } : undefined),
  }).code;
}

const formulas = [
  (op: string) => `\\displaystyle a${op} b`,
  (op: string) => `\\displaystyle ${op}^{1}x`,
  (op: string) => `\\displaystyle ${op}_{\\Omega}x`,
  (op: string) => `\\displaystyle ${op}_{0}^{1}x`,
  (op: string) => `\\displaystyle ${op}_{x=-\\infty}^{\\infty}f`,
  (op: string) => `\\textstyle ${op}_0^\\infty x`,
  (op: string) => `\\textstyle a${op} b`,
  (op: string) => `\\displaystyle\\sum${op}_D f+${op}`,
  // (in script styles LaTeX sets \int in cmex7, a design of its own: the box
  // form scales the text size, so a script-size symbol is only close there)
];
// [form, behaviour, command, native \int, what follows the command where it is used]
const cases: Array<["native" | "box", CustomSymbolLimitsBehavior, string, string, string]> = [];
for (const form of ["native", "box"] as const) {
  const p = form === "native" ? "n" : "b";
  cases.push(
    [form, "nolimits", `${p}intn`, "\\intop\\nolimits", ""],
    [form, "limits", `${p}intl`, "\\intop\\limits", ""],
    [form, "auto", `${p}inta`, "\\intop\\displaylimits", ""],
    // \limits and \nolimits typed after the command, as after \int
    [form, "auto", `${p}intx`, "\\intop\\limits", "\\limits"],
    [form, "auto", `${p}inty`, "\\intop\\nolimits", "\\nolimits"],
  );
}
const definitions = cases.map(([form, limits, command]) => integralCommand(command, limits, form)).join("\n");

const dir = mkdtempSync(join(tmpdir(), "vtx-tex-native-"));
const fonts: Array<[string, string]> = [["cm", ""]];
if (kpse("newtxmath.sty")) {
  const map = kpse("newtx.map") ? "\\pdfmapfile{+newtx.map}" : "";
  fonts.push(["ntx", `${map}\\usepackage{newtxtext,newtxmath}`]);
}
const preamble = (font: string) => ["\\usepackage{amsmath,graphicx,pict2e,trimclip}", font, definitions].join("\n");
let checked = 0;
for (const [fontName, fontPackages] of fonts) {
  const body: string[] = [];
  const pairs: string[] = [];
  for (const [form, , command, native, after] of cases) {
    if (fontName !== "cm" && form === "box") continue; // the box carries Computer Modern's metrics
    for (const [index, formula] of formulas.entries()) {
      const id = `${command}${index}`;
      pairs.push(id);
      for (const [tag, op] of [["n", native], ["c", `\\${command}${after}`]] as const) {
        body.push(
          `\\setbox0\\hbox{$${formula(op)}$\\vrule width 1pt height 1pt}`,
          `\\typeout{VTX ${id} ${tag} \\the\\wd0/\\the\\ht0/\\the\\dp0}`,
        );
      }
    }
  }
  const name = `compare-${fontName}`;
  writeFileSync(join(dir, `${name}.tex`), ["\\documentclass{article}", preamble(fontPackages), "\\begin{document}", ...body, "\\end{document}"].join("\n"));
  const run = spawnSync("pdflatex", ["-interaction=nonstopmode", `${name}.tex`], { cwd: dir, encoding: "utf8" });
  const log = readFileSync(join(dir, `${name}.log`), "utf8");
  assert.ok(!/^! /m.test(log), `LaTeX errors (${fontName}):\n${log.split("\n").filter((line) => line.startsWith("!")).join("\n")}\n${run.stdout.slice(-600)}`);
  const measured = new Map<string, number[]>();
  for (const match of log.matchAll(/^VTX (\S+) ([nc]) ([\d.-]+)pt\/([\d.-]+)pt\/([\d.-]+)pt$/gm)) {
    measured.set(`${match[1]} ${match[2]}`, [Number(match[3]), Number(match[4]), Number(match[5])]);
  }
  const failures: string[] = [];
  for (const id of pairs) {
    const native = measured.get(`${id} n`);
    const custom = measured.get(`${id} c`);
    assert.ok(native && custom, `missing measurement for ${id} (${fontName})`);
    const worst = Math.max(...native.map((value, index) => Math.abs(value - custom[index])));
    if (worst > 0.01) failures.push(`${id}: native ${native.join("/")} custom ${custom.join("/")}`);
  }
  assert.deepEqual(failures, [], `the custom integral sets as \\int does (${fontName})`);
  checked += pairs.length;
}

// pixels (Computer Modern, the box form; the native form draws a dot more)
if (has("mutool", ["-v"]) || has("mutool", [])) {
  const pngs: string[] = [];
  for (const [form, , command, native, after] of cases) {
    if (form !== "box") continue;
    for (const [index, formula] of formulas.entries()) {
      for (const [tag, op] of [["n", native], ["c", `\\${command}${after}`]] as const) {
        const name = `px-${command}${index}-${tag}`;
        writeFileSync(join(dir, `${name}.tex`), ["\\documentclass[border=1pt]{standalone}", preamble(""), "\\begin{document}", `$${formula(op)}$`, "\\end{document}"].join("\n"));
        spawnSync("pdflatex", ["-interaction=nonstopmode", `${name}.tex`], { cwd: dir, stdio: "ignore" });
        spawnSync("mutool", ["draw", "-q", "-r", "1200", "-o", `${name}.png`, `${name}.pdf`], { cwd: dir, stdio: "ignore" });
        pngs.push(name);
      }
    }
  }
  const script = `
import sys
from PIL import Image, ImageChops
d=sys.argv[1]; names=sys.argv[2:]
worst=0
for i in range(0,len(names),2):
    a=Image.open(f"{d}/{names[i]}.png").convert("L"); b=Image.open(f"{d}/{names[i+1]}.png").convert("L")
    if a.size!=b.size: print("SIZE",names[i],a.size,b.size); worst=max(worst,1.0); continue
    diff=ImageChops.difference(a,b).point(lambda v:255 if v>128 else 0)
    mis=sum(1 for v in diff.getdata() if v); ink=sum(1 for v in a.getdata() if v<128)
    worst=max(worst,mis/max(1,ink))
print("WORST", worst)
`;
  try {
    const output = execFileSync("python3", ["-W", "ignore", "-c", script, dir, ...pngs], { encoding: "utf8" });
    const worst = Number(/WORST ([\d.]+)/.exec(output)?.[1] ?? "1");
    assert.ok(worst < 0.02, `pixels differ by up to ${(worst * 100).toFixed(2)}% of the ink`);
  } catch (error) {
    if ((error as { code?: string }).code === "ENOENT") console.log("pixel comparison skipped (no python3)");
    else throw error;
  }
}
console.log(`custom symbol tex native compare: ok (${checked} formulas, fonts ${fonts.map(([name]) => name).join("+")}, ${dir})`);
