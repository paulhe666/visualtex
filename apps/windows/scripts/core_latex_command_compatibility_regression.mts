import assert from "node:assert/strict";
import { convertLatexToMarkup } from "mathlive";
import { commandRegistry } from "../src/autocomplete/commandRegistry";
import { VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS } from "../src/math/mathLiveCompatibilityMacros";
import { VISUALTEX_CORE_LATEX_ALIASES } from "../src/math/coreLatexAliases";
import { VISUALTEX_PACKAGE_COMPATIBILITY_MACROS } from "../src/math/packageMacroCompatibility";
import { VISUALTEX_MATHML_MACROS, VISUALTEX_SVG_MACROS, normalizePackageLatexCommands } from "../src/math/latexCompatibility";
import { formatLatex, parseLatexSourceDraft } from "../src/clipboard/LatexCopyService";
import { latexToMathMl, latexToSvg } from "../src/export/runtime";
import { mathjax } from "mathjax-full/js/mathjax.js";
import { TeX } from "mathjax-full/js/input/tex.js";
import { SVG } from "mathjax-full/js/output/svg.js";
import { AllPackages } from "mathjax-full/js/input/tex/AllPackages.js";
import { STATE } from "mathjax-full/js/core/MathItem.js";
import { TEXCLASS, type MmlNode } from "mathjax-full/js/core/MmlTree/MmlNode.js";
import { SerializedMmlVisitor } from "mathjax-full/js/core/MmlTree/SerializedMmlVisitor.js";

const registryCommands = new Set(commandRegistry.map((entry) => entry.command));

for (const alias of VISUALTEX_CORE_LATEX_ALIASES) {
  const command = `\\${alias.name}`;
  assert.equal(
    registryCommands.has(command),
    true,
    `${command} is renderable but missing from semantic completion`,
  );
  const markup = convertLatexToMarkup(command, {
    macros: VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS,
  });
  assert.doesNotMatch(
    markup,
    /ML__error/,
    `${command} still renders as a MathLive error`,
  );
}

const parameterizedSamples = new Map([
  ["\\pod", String.raw`x\equiv y\pod{n}`],
  ["\\substack", String.raw`\sum_{\substack{i<j\\i+j=n}} a_{ij}`],
  ["\\prescript", String.raw`\prescript{14}{6}{C}`],
  ["\\sideset", String.raw`\sideset{_a^b}{_c^d}{\sum}`],
]);

for (const [command, latex] of parameterizedSamples) {
  assert.equal(
    registryCommands.has(command),
    true,
    `${command} is missing from semantic completion`,
  );
  const markup = convertLatexToMarkup(latex, {
    macros: VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS,
  });
  assert.doesNotMatch(
    markup,
    /ML__error/,
    `${command} still renders as a MathLive error`,
  );
}

for (const command of ["\\c", "\\quad", "\\qquad"]) {
  const entry = commandRegistry.find((candidate) => candidate.command === command);
  assert.ok(entry, `${command} is missing from semantic completion`);
  assert.notEqual(
    entry.previewLatex,
    command,
    `${command} still uses an invisible bare-command preview`,
  );
  const markup = convertLatexToMarkup(entry.previewLatex, {
    macros: VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS,
  });
  assert.doesNotMatch(markup, /ML__error/, `${command} preview has a parse error`);
}

// One scalar can have distinct named atom classes. Check the shared registry,
// both actual MathJax configurations, the kernel's expansion as the editor draws
// it, and source copy.
const slantedSymbols = [
  ["varparallel", "\\mathrel{⫽}", TEXCLASS.REL],
  ["parallelslant", "\\mathrel{⫽}", TEXCLASS.REL],
  ["sslash", "\\mathbin{⫽}", TEXCLASS.BIN],
] as const;
for (const [name, replacement] of slantedSymbols) {
  assert.deepEqual(VISUALTEX_PACKAGE_COMPATIBILITY_MACROS[name], [replacement, 0]);
  const macro = VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS[name];
  assert.ok(typeof macro === "object" && macro !== null && "def" in macro);
  assert.equal(macro.def, replacement);
  assert.equal(macro.expand, false, "Editing must retain the authored package spelling");
  for (const source of ["a\\" + name + " b", "x^{a\\" + name + " b}_{i}"]) {
    const markup = convertLatexToMarkup(source, { macros: VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS });
    assert.doesNotMatch(markup, /ML__error/);
    assert.match(markup, /⫽|&#x2afd;/i, "The editor must draw the double solidus, not upright parallel");
    assert.equal(formatLatex(source, "raw"), source);
    assert.equal(parseLatexSourceDraft(source, "raw").valid, true);
    assert.ok(normalizePackageLatexCommands(source).includes(replacement));
    for (const display of [false, true]) {
      const mathMl = await latexToMathMl(source, display);
      assert.match(mathMl, /⫽|&#x2AFD;/i);
      assert.doesNotMatch(mathMl, /∥|&#x2225;|merror|varparallel|parallelslant|sslash/i);
      const { svg } = await latexToSvg(source, { displayMode: display, fontSizePt: 12, paddingPx: 8, background: "transparent" });
      assert.match(svg, /⫽|data-c="2AFD"/i);
      assert.match(svg, name === "sslash" ? /data-mjx-texclass="BIN"/ : /data-mjx-texclass="REL"/);
      assert.doesNotMatch(svg, /data-mml-node="merror"/);
      // A U+2AFD <text> fallback is currently permitted by the SVG exporter;
      // this asserts scalar/class fidelity, not self-contained font geometry.
    }
  }
}

for (const macros of [VISUALTEX_MATHML_MACROS, VISUALTEX_SVG_MACROS]) {
  const input = new TeX({
    packages: AllPackages.filter((name) => name !== "noundefined" && name !== "noerrors"),
    macros,
    formatError: (_jax: unknown, error: unknown) => { throw error; },
  });
  const document = mathjax.document("", { InputJax: input, OutputJax: new SVG() });
  const visitor = new SerializedMmlVisitor(document.mmlFactory);
  for (const [name, , expectedClass] of slantedSymbols) {
    for (const source of ["a\\" + name + " b", "x^{a\\" + name + " b}_{i}"]) {
      const root = document.convert(source, { end: STATE.COMPILED }) as unknown as MmlNode;
      let typedGlyph = false;
      root.walkTree((node) => {
        const atom = node as MmlNode;
        if (atom.kind === "TeXAtom" && atom.texClass === expectedClass
            && /⫽|&#x2AFD;/i.test(visitor.visitTree(atom))) typedGlyph = true;
      });
      assert.ok(typedGlyph, name + " lost its explicit atom class in the actual MathJax tree");
    }
  }
  const raw = document.convert("a⫽b", { end: STATE.COMPILED }) as unknown as MmlNode;
  let rawBinary = false;
  raw.walkTree((node) => {
    const atom = node as MmlNode;
    if (atom.kind === "mo" && atom.texClass === TEXCLASS.BIN
        && /⫽|&#x2AFD;/i.test(visitor.visitTree(atom))) rawBinary = true;
  });
  assert.ok(rawBinary, "Raw U+2AFD must remain a binary operator");
  assert.throws(() => document.convert("\\varparalle", { end: STATE.COMPILED }));
}
for (const source of ["\\parallel", "\\mathit{\\parallel}", "\\mathbfit{\\parallel}"]) {
  const mathMl = await latexToMathMl(source);
  assert.match(mathMl, /∥|&#x2225;/i);
  assert.doesNotMatch(mathMl, /⫽|&#x2AFD;/i);
}
for (const name of ["varparalle", "varparallelinv", "nvarparallel", "nparallelslant"]) {
  assert.equal(Object.hasOwn(VISUALTEX_PACKAGE_COMPATIBILITY_MACROS, name), false);
  assert.equal(registryCommands.has("\\" + name), false);
}

for (const [command, body] of [["mathrel", "⫽"], ["mathbin", "="], ["mathop", "x+y"]]) {
  const source = "\\" + command + "{" + body + "}";
  assert.ok(registryCommands.has("\\" + command));
  assert.equal(parseLatexSourceDraft(source, "raw").valid, true);
  assert.equal(formatLatex(source, "raw"), source);
  assert.doesNotMatch(convertLatexToMarkup(source, { macros: VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS }), /ML__error/);
  assert.doesNotMatch(await latexToMathMl(source), /merror/);
}
for (const command of ["mathrel", "mathbin"]) {
  assert.equal(parseLatexSourceDraft("\\" + command, "raw").valid, false);
}

// AMS alternates are native commands, not compatibility macro approximations.
for (const [name, scalar] of [
  ["shortparallel", "2225"], ["nshortparallel", "2226"],
]) {
  const command = "\\" + name;
  assert.ok(registryCommands.has(command));
  assert.equal(Object.hasOwn(VISUALTEX_PACKAGE_COMPATIBILITY_MACROS, name), false);
  for (const source of [command, "a" + command + " b",
    "x^{" + command + "}_{i}", "\\mathit{" + command + "}",
    "\\mathbf{" + command + "}", "\\mathbfit{" + command + "}"]) {
    assert.equal(parseLatexSourceDraft(source, "raw").valid, true, source);
    assert.equal(formatLatex(source, "raw"), source);
    const markup = convertLatexToMarkup(source, { macros: VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS });
    assert.doesNotMatch(markup, /ML__error/);
    // the editor draws AMS alternates from the AMS symbol font
    assert.match(markup, /'VTX AMS Symbols B \d+'/, "AMS alternate glyph lost");
    for (const displayMode of [false, true]) {
      const mathMl = await latexToMathMl(source, displayMode);
      assert.match(mathMl, /data-mjx-alternate="1"/);
      assert.ok(mathMl.includes("&#x" + scalar + ";"), mathMl);
      assert.doesNotMatch(mathMl, /\uE007|&#xE007;/i, "Font-local PUA cannot become semantic MathML");
      const { svg } = await latexToSvg(source, {displayMode, fontSizePt: 12, paddingPx: 8, background: "white"});
      assert.ok(svg.includes("-TEX-V-" + scalar), "Real alternate path lost: " + source);
      assert.doesNotMatch(svg, /<text\b|merror/);
    }
  }
}

console.log(
  `Core LaTeX compatibility regression passed (${VISUALTEX_CORE_LATEX_ALIASES.length} atomic aliases + ${parameterizedSamples.size} structured commands + visible native spacing/accent previews).`,
);
