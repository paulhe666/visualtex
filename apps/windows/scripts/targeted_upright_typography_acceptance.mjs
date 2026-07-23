import assert from "node:assert/strict";
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import {
  normalizeChineseLatex,
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

  const mathMl = latexToMathMl(normalized, true);
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
