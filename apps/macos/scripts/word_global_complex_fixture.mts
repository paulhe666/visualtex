import { writeFileSync, mkdirSync } from "node:fs";
import { createRequire } from "node:module";
import { join, resolve } from "node:path";
import { latexToSvg } from "../src/export/latexToSvg.ts";
import { createFormulaMetadata, encodeFormulaMetadata } from "../src/office/shared/formulaMetadata.ts";

// The Word driver opens these real rendered images, then performs both
// conversions through the production commands. No placeholder artwork.
const destination = resolve(process.argv[2]);
const dependencyRoot = process.argv[3];
const requireDependency = createRequire(join(dependencyRoot, "package.json"));
const sharp = requireDependency("sharp");
mkdirSync(destination, { recursive: true });
const sources = [
  String.raw`(a+b)^n=\sum_{k=0}^{n}\binom{n}{k}a^{n-k}b^k`,
  String.raw`\alpha>b\quad\text{中文说明}`,
  String.raw`x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}+\frac{\int_0^1e^{-t^2}\,dt}{1+\frac{1}{1+x^2}}`,
  String.raw`i\hbar\frac{\partial\Psi}{\partial t}=\hat H\Psi,\quad\oint_C\mathbf{F}\cdot d\mathbf{r}=\iint_S(\nabla\times\mathbf{F})\cdot d\mathbf{S}`,
  String.raw`A=\begin{pmatrix}a_{11}&a_{12}\\a_{21}&a_{22}\end{pmatrix},\quad\det(A-\lambda I)=0`,
  String.raw`f(x)=\begin{cases}\frac{\sin x}{x},&x\ne0\\1,&x=0\end{cases}`,
];
const results = [];
for (const [index, latex] of sources.entries()) {
  const id = `81818181-8181-4818-8818-${String(index + 1).padStart(12, "0")}`;
  const displayMode = index === 1 ? "inline" : "block";
  const fontSizePt = [14, 11, 14, 11, 18, 14][index];
  const rendered = latexToSvg(latex, { displayMode: displayMode === "block", fontSizePt: 14,
    paddingPx: displayMode === "inline" ? 1 : 2, background: "transparent", forceExplicitBlack: true,
    formulaLetterFont: "times", formulaChineseFont: "songti" });
  const fit = Math.min(1, 500 / (rendered.width * 0.75 * 1.067));
  const referenceWidthPt = rendered.width * 0.75 * 1.067 * fit;
  const referenceHeightPt = rendered.height * 0.75 * fit;
  const referenceBaselinePt = -(rendered.height - rendered.baseline) * 0.75 * fit;
  const metadata = createFormulaMetadata({ formulaId: id, title: `Complex Word regression ${index + 1}`,
    lines: [{ id: `91919191-9191-4919-8919-${String(index + 1).padStart(12, "0")}`, latex }],
    codeFormat: "raw", sourceLatex: latex, displayMode, numbered: index !== 1,
    fontSizePt, renderWidthPx: rendered.width, renderHeightPx: rendered.height,
    referenceWidthPt, referenceHeightPt, referenceBaselinePt,
    formulaLetterFont: "times", formulaChineseFont: "songti", appVersion: "1.2.6" });
  writeFileSync(join(destination, `${index}.svg`), rendered.svg);
  await sharp(Buffer.from(rendered.svg), { density: 144 }).png().toFile(join(destination, `${index}.png`));
  results.push({ metadata, encodedMetadata: encodeFormulaMetadata(metadata),
    width: referenceWidthPt * fontSizePt / 14, height: referenceHeightPt * fontSizePt / 14 });
}
writeFileSync(join(destination, "fixtures.json"), JSON.stringify(results, null, 2));
console.log(`Rendered ${results.length} real complex formula images in ${destination}`);
