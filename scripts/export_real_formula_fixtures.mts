import { mkdir, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { latexToSvg } from "../src/export/runtime";

const formulas = [
  String.raw`\int ac`,
  String.raw`\int x\,dy`,
  String.raw`\alpha\beta d f d f aaaaabbbbb`,
  String.raw`\frac{a+b}{c+d}+\sqrt{x^2+y^2}`,
  String.raw`\int_0^1 x^2\,dx+\sum_{n=1}^{\infty}\frac{1}{n^2}`,
];

const outputRoot = resolve(
  process.argv[2] ?? "src-windows/artifacts/real-formula-fixtures",
);
await mkdir(outputRoot, { recursive: true });

const manifest = formulas.map((latex, index) => {
  const result = latexToSvg(latex, {
    displayMode: true,
    fontSizePt: 12,
    paddingPx: 8,
    background: "transparent",
  });
  return {
    id: `formula-${index + 1}`,
    latex,
    fileName: `formula-${index + 1}.svg`,
    width: result.width,
    height: result.height,
    baseline: result.baseline,
    svg: result.svg,
  };
});

for (const item of manifest) {
  await writeFile(resolve(outputRoot, item.fileName), item.svg, "utf8");
}
await writeFile(
  resolve(outputRoot, "manifest.json"),
  JSON.stringify(
    manifest.map(({ svg: _svg, ...item }) => item),
    null,
    2,
  ),
  "utf8",
);

for (const item of manifest) {
  console.log(
    `${item.id}\t${item.width.toFixed(4)}x${item.height.toFixed(4)}\t${item.latex}`,
  );
}
