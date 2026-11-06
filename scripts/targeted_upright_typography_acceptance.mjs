import assert from "node:assert/strict";
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import {
  normalizeChineseLatex,
  normalizeContextualUprightSymbols,
} from "../src/editor/normalizeChineseLatex.ts";
import { latexToMathMl } from "../src/export/runtime.ts";

const outputPath = process.argv[2]
  ? path.resolve(process.argv[2])
  : path.resolve("src-windows/artifacts/targeted-upright-crossref/upright-cases.json");

const forbiddenInternalCommands = [
  "differentialD",
  "capitalDifferentialD",
  "exponentialE",
  "imaginaryI",
  "imaginaryJ",
];

const cases = [
  {
    name: "first-order differential typed as dy/dx",
    input: String.raw`\frac{dy}{dx}`,
    expectedFragments: [String.raw`\mathrm{d}y`, String.raw`\mathrm{d}x`],
    normalMathMlTokens: ["d"],
  },
  {
    name: "arbitrary Latin differential operand",
    input: String.raw`\frac{dQ}{dt}`,
    expectedFragments: [String.raw`\mathrm{d}Q`, String.raw`\mathrm{d}t`],
    normalMathMlTokens: ["d"],
  },
  {
    name: "Greek differential operand and variable",
    input: String.raw`\frac{d\Phi}{d\theta}`,
    expectedFragments: [String.raw`\mathrm{d}\Phi`, String.raw`\mathrm{d}\theta`],
    normalMathMlTokens: ["d"],
  },
  {
    name: "Greek differential in an asymmetric fraction",
    input: String.raw`\frac{\partial f}{d\alpha}`,
    expectedFragments: [String.raw`\partial f`, String.raw`\mathrm{d}\alpha`],
    normalMathMlTokens: ["d"],
  },
  {
    name: "spaced Greek differential in a reciprocal fraction",
    input: String.raw`\frac{1}{d \theta}`,
    expectedFragments: [String.raw`\mathrm{d} \theta`],
    normalMathMlTokens: ["d"],
  },
  {
    name: "styled vector differential operand",
    input: String.raw`\frac{d\mathbf{r}}{dt}`,
    expectedFragments: [String.raw`\mathrm{d}\mathbf{r}`, String.raw`\mathrm{d}t`],
    normalMathMlTokens: ["d"],
  },
  {
    name: "parenthesized differential expression",
    input: String.raw`\frac{d(xy+\theta)}{dx}`,
    expectedFragments: [String.raw`\mathrm{d}(xy+\theta)`, String.raw`\mathrm{d}x`],
    normalMathMlTokens: ["d"],
  },
  {
    name: "second-order differential",
    input: String.raw`\frac{d^2y}{dx^2}`,
    expectedFragments: [String.raw`\mathrm{d}^2y`, String.raw`\mathrm{d}x^2`],
    normalMathMlTokens: ["d"],
  },
  {
    name: "mixed higher-order differential",
    input: String.raw`\frac{d^2(\alpha+\beta)}{dx\,dy}`,
    expectedFragments: [
      String.raw`\mathrm{d}^2(\alpha+\beta)`,
      String.raw`\mathrm{d}x`,
      String.raw`\mathrm{d}y`,
    ],
    normalMathMlTokens: ["d"],
  },
  {
    name: "subscripted differential variable",
    input: String.raw`\frac{du_i}{dt_j}`,
    expectedFragments: [String.raw`\mathrm{d}u_i`, String.raw`\mathrm{d}t_j`],
    normalMathMlTokens: ["d"],
  },
  {
    name: "nth-order differential",
    input: String.raw`\frac{d^{n}y}{dx^{n}}`,
    expectedFragments: [String.raw`\mathrm{d}^{n}y`, String.raw`\mathrm{d}x^{n}`],
    normalMathMlTokens: ["d"],
  },
  {
    name: "integral differential element",
    input: String.raw`\int_0^1 f(x)dx`,
    expectedFragments: [String.raw`\mathrm{d}x`],
    normalMathMlTokens: ["d"],
  },
  {
    name: "ordinary d identifiers remain italic",
    input: String.raw`d+df+dx`,
    expectedNormalized: String.raw`d+df+dx`,
    expectedFragments: [String.raw`d+df+dx`],
    normalMathMlTokens: [],
  },
  {
    name: "integrand d identifiers are not trailing differentials",
    input: String.raw`\int_0^1 f(dx)+d(x)`,
    expectedNormalized: String.raw`\int_0^1 f(dx)+d(x)`,
    expectedFragments: [String.raw`f(dx)+d(x)`],
    normalMathMlTokens: [],
  },
  {
    name: "MathLive canonical upright commands",
    input: String.raw`\frac{\differentialD y}{\differentialD x}+\exponentialE^{\imaginaryI\theta}+\imaginaryJ`,
    expectedFragments: [
      String.raw`\mathrm{d}y`,
      String.raw`\mathrm{d}x`,
      String.raw`\mathrm{e}^{\mathrm{i}\theta}`,
      String.raw`\mathrm{j}`,
    ],
    normalMathMlTokens: ["d", "e", "i", "j"],
  },
  {
    name: "capital differential operator",
    input: String.raw`\capitalDifferentialD F`,
    expectedFragments: [String.raw`\mathrm{D}F`],
    normalMathMlTokens: ["D"],
  },
  {
    name: "elementary and hyperbolic functions",
    input: "sin(x)+cos(x)+tan(x)+cot(x)+sec(x)+csc(x)+arcsin(x)+arccos(x)+arctan(x)+sinh(x)+cosh(x)+tanh(x)+coth(x)",
    expectedFragments: [
      String.raw`\sin`, String.raw`\cos`, String.raw`\tan`, String.raw`\cot`,
      String.raw`\sec`, String.raw`\csc`, String.raw`\arcsin`, String.raw`\arccos`,
      String.raw`\arctan`, String.raw`\sinh`, String.raw`\cosh`, String.raw`\tanh`,
      String.raw`\coth`,
    ],
    normalMathMlTokens: [
      "sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos",
      "arctan", "sinh", "cosh", "tanh", "coth",
    ],
  },
  {
    name: "exponential logarithmic and limit operators",
    input: "exp(x)+ln(x)+log(x)+lg(x)+lim(x)+limsup(x)+liminf(x)+max(x)+min(x)+sup(x)+inf(x)",
    expectedFragments: [
      String.raw`\exp`, String.raw`\ln`, String.raw`\log`, String.raw`\lg`,
      String.raw`\lim`, String.raw`\limsup`, String.raw`\liminf`, String.raw`\max`,
      String.raw`\min`, String.raw`\sup`, String.raw`\inf`,
    ],
    normalMathMlTokens: [
      "exp", "ln", "log", "lg", "lim", "lim sup", "lim inf", "max", "min", "sup", "inf",
    ],
  },
  {
    name: "linear algebra and named upright operators",
    input: "det(A)+dim(V)+ker(A)+gcd(a,b)+lcm(a,b)+rank(A)+tr(A)+diag(A)+sgn(x)+erf(x)+erfc(x)",
    expectedFragments: [
      String.raw`\det`, String.raw`\dim`, String.raw`\ker`, String.raw`\gcd`,
      String.raw`\operatorname{lcm}`, String.raw`\operatorname{rank}`,
      String.raw`\operatorname{tr}`, String.raw`\operatorname{diag}`,
      String.raw`\operatorname{sgn}`, String.raw`\operatorname{erf}`,
      String.raw`\operatorname{erfc}`,
    ],
    normalMathMlTokens: [
      "det", "dim", "ker", "gcd", "lcm", "rank", "tr", "diag", "sgn", "erf", "erfc",
    ],
  },
  {
    name: "real and imaginary part operators",
    input: "Re(z)+Im(z)",
    expectedFragments: [String.raw`\operatorname{Re}`, String.raw`\operatorname{Im}`],
    normalMathMlTokens: ["Re", "Im"],
  },
  {
    name: "Euler constant and imaginary unit",
    input: String.raw`e^{i\theta}+\exp(i\phi)`,
    expectedFragments: [
      String.raw`\mathrm{e}^{\mathrm{i}\theta}`,
      String.raw`\exp(\mathrm{i}\phi)`,
    ],
    normalMathMlTokens: ["e", "i"],
  },
  {
    name: "already-upright input remains idempotent",
    input: String.raw`\frac{\mathrm{d}y}{\mathrm{d}x}+\mathrm{e}^{\mathrm{i}\theta}`,
    expectedFragments: [
      String.raw`\mathrm{d}y`,
      String.raw`\mathrm{d}x`,
      String.raw`\mathrm{e}^{\mathrm{i}\theta}`,
    ],
    normalMathMlTokens: ["d", "e", "i"],
  },
  {
    name: "text commands remain untouched",
    input: String.raw`\text{sin dx e i}+x`,
    expectedFragments: [String.raw`\text{sin dx e i}`],
    normalMathMlTokens: [],
  },
  // The one upright rule (docs/upright-typography-plan-20261015.md).
  ...[
    // A unit group is upright as a whole: a certain member (km, kg), a
    // number and a space before it, or the whole formula (decision A1).
    ["unit group with a certain member", "km/h", String.raw`\mathrm{km}/\mathrm{h}`],
    ["unit group as the whole formula", "m/s", String.raw`\mathrm{m}/\mathrm{s}`],
    ["single-letter unit joined to a certain unit", "J/kg", String.raw`\mathrm{J}/\mathrm{kg}`],
    ["unit exponent stays with its unit", "kg/m^3", String.raw`\mathrm{kg}/\mathrm{m}^3`],
    ["units after a number and a space", String.raw`a=9.8\ m/s^2`, String.raw`a=9.8\ \mathrm{m}/\mathrm{s}^2`],
    ["units after scientific notation", String.raw`c=3\times10^8\ m/s`, String.raw`c=3\times10^8\ \mathrm{m}/\mathrm{s}`],
    ["product of units", String.raw`N\cdot m`, String.raw`\mathrm{N}\cdot \mathrm{m}`],
    ["minute after a number", String.raw`5\ min`, String.raw`5\ \mathrm{min}`],
    ["old minimum shortcut after a number", String.raw`5\ \min`, String.raw`5\ \mathrm{min}`],
    ["old operator units become upright symbols", String.raw`\operatorname{\mathrm{kg}}/m^3`, String.raw`\mathrm{kg}/\mathrm{m}^3`],
    ["pH is a fixed name", "pH=7", String.raw`\mathrm{pH}=7`],
    ["descriptive subscript word", "v_{max}", String.raw`v_{\mathrm{max}}`],
    ["ambiguous function with an argument", String.raw`lg\ x`, String.raw`\lg\ x`],
    ["exponential with a variable exponent", String.raw`e^{-t/\tau}`, String.raw`\mathrm{e}^{-t/\tau}`],
    ["imaginary unit in a complex number", "z=3+4i", String.raw`z=3+4\mathrm{i}`],
    // Values stay italic.
    ["products of variables stay italic", "F=mg+E-3t+2h", "F=mg+E-3t+2h"],
    ["variable ratio stays italic", "v=s/t", "v=s/t"],
    ["lg without an argument is l times g", "l=lg", "l=lg"],
    ["explicit italic is kept", String.raw`\mathit{lg}`, String.raw`\mathit{lg}`],
    ["elementary charge squared stays italic", "q=e^2", "q=e^2"],
    ["Reynolds number stays italic", String.raw`Re=\rho vL/\mu`, String.raw`Re=\rho vL/\mu`],
    ["single-letter subscripts stay italic", "T_{c}+x_{i}+a_{2i}", "T_{c}+x_{i}+a_{2i}"],
    ["a lone unit letter stays italic", "m^2", "m^2"],
    ["deci-prefixed spelling is still a differential", String.raw`\frac{dB}{dt}`, String.raw`\frac{\mathrm{d}B}{\mathrm{d}t}`],
    ["dimensions are not math", String.raw`x\hspace{1cm}y`, String.raw`x\hspace{1cm}y`],
  ].map(([name, input, expected]) => ({
    name,
    input,
    expectedNormalized: expected,
    expectedFragments: [expected],
    normalMathMlTokens: [],
  })),
];

function assertNoInternalCommands(value, stage, caseName) {
  for (const command of forbiddenInternalCommands) {
    assert.ok(
      !value.includes(command),
      `${caseName}: ${stage} still contains MathLive internal command ${command}`,
    );
  }
}

function assertNormalMathMlToken(mathMl, token, caseName) {
  const escaped = token.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const explicitNormalPattern = new RegExp(
    `<mi\\b(?=[^>]*mathvariant=["']normal["'])[^>]*>\\s*${escaped}\\s*</mi>`,
    "i",
  );
  if (token.length === 1) {
    assert.match(
      mathMl,
      explicitNormalPattern,
      `${caseName}: single-character MathML token ${token} is not explicitly upright/normal`,
    );
    return;
  }

  const uprightNodeTexts = Array.from(
    mathMl.matchAll(
      /<m(?:i|o)\b(?![^>]*mathvariant=["']italic["'])[^>]*>([\s\S]*?)<\/m(?:i|o)>/gi,
    ),
    (match) => match[1].replace(/<[^>]+>/g, "").replace(/\s+/g, ""),
  );
  const compactToken = token.replace(/\s+/g, "");
  const wholeTokenFound = uprightNodeTexts.some(
    (value) => value.toLocaleLowerCase() === compactToken.toLocaleLowerCase(),
  );
  const tokenParts = token.split(/\s+/).filter(Boolean);
  const splitTokenFound = tokenParts.length > 1 && tokenParts.every((part) =>
    uprightNodeTexts.some(
      (value) => value.toLocaleLowerCase() === part.toLocaleLowerCase(),
    ),
  );
  assert.ok(
    wholeTokenFound || splitTokenFound,
    `${caseName}: MathML token ${token} is not represented by upright identifier/operator nodes`,
  );
}

const reportCases = [];
console.log("[Upright 1/3] Normalizing typed and MathLive-internal LaTeX...");
for (const item of cases) {
  const normalized = normalizeChineseLatex(item.input);
  const normalizedAgain = normalizeChineseLatex(normalized);
  assert.equal(
    normalizedAgain,
    normalized,
    `${item.name}: normalization is not idempotent`,
  );
  if (item.expectedNormalized !== undefined) {
    assert.equal(
      normalized,
      item.expectedNormalized,
      `${item.name}: normalization changed an ordinary italic identifier`,
    );
  }
  for (const fragment of item.expectedFragments) {
    assert.ok(
      normalized.includes(fragment),
      `${item.name}: missing normalized fragment ${fragment}; got ${normalized}`,
    );
  }
  assertNoInternalCommands(normalized, "normalized LaTeX", item.name);

  const mathMl = await latexToMathMl(normalized, true);
  assert.match(mathMl, /^<math\b/);
  assertNoInternalCommands(mathMl, "MathML", item.name);
  for (const token of item.normalMathMlTokens) {
    assertNormalMathMlToken(mathMl, token, item.name);
  }

  reportCases.push({
    name: item.name,
    input: item.input,
    normalized,
    mathMl,
    expectedFragments: item.expectedFragments,
    normalMathMlTokens: item.normalMathMlTokens,
  });
  console.log(`  PASS ${item.name}`);
  console.log(`       input      = ${item.input}`);
  console.log(`       normalized = ${normalized}`);
}

// While the caret follows a letter, the run may still grow (cm -> cmax).
assert.equal(normalizeContextualUprightSymbols("cm", { deferLetterRuns: true }), "cm");
assert.equal(normalizeContextualUprightSymbols("cm"), String.raw`\mathrm{cm}`);
console.log("  PASS an unfinished letter run waits");

console.log("[Upright 2/3] Verifying all configured common upright operators...");
const operatorCaseNames = reportCases
  .filter((item) => item.normalMathMlTokens.length > 0)
  .map((item) => item.name);
assert.ok(operatorCaseNames.length >= 10, "Upright operator coverage is unexpectedly small");
console.log(`  Covered ${operatorCaseNames.length} formula classes and ${reportCases.reduce((sum, item) => sum + item.normalMathMlTokens.length, 0)} upright MathML token checks.`);

console.log("[Upright 3/3] Writing the MathML fixture consumed by the Word OMML acceptance...");
await mkdir(path.dirname(outputPath), { recursive: true });
await writeFile(
  outputPath,
  JSON.stringify(
    {
      generatedAt: new Date().toISOString(),
      forbiddenInternalCommands,
      cases: reportCases,
    },
    null,
    2,
  ),
  "utf8",
);
console.log(`Targeted upright typography acceptance passed (${reportCases.length} cases).`);
console.log(`Fixture: ${outputPath}`);
