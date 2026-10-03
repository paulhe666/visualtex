import assert from "node:assert/strict";
import { convertLatexToMarkup } from "mathlive";
import { commandRegistry } from "../src/autocomplete/commandRegistry";
import { VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS } from "../src/math/mathLiveCompatibilityMacros";
import { VISUALTEX_CORE_LATEX_ALIASES } from "../src/math/coreLatexAliases";

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

console.log(
  `Core LaTeX compatibility regression passed (${VISUALTEX_CORE_LATEX_ALIASES.length} atomic aliases + ${parameterizedSamples.size} structured commands + visible native spacing/accent previews).`,
);
