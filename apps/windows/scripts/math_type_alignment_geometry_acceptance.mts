import { sharedSourceUrl } from "../../shared/visualtex-app/testing/sharedSourceUrl.mjs";
import assert from "node:assert/strict";
import { withAppBrowser } from "../../shared/visualtex-app/testing/browser_test_runtime.mjs";

const exactUserAligned = String.raw`\begin{aligned}
\langle p_1,p_0\rangle &\leftarrow \operatorname{umul}(a,b)=ab
&&\text{Double word product}\\
p_0 &\leftarrow \operatorname{umullo}(a,b)=(ab)\bmod\beta
&&\text{Low word}\\
p_1 &\leftarrow \operatorname{umulhi}(a,b)=\left\lfloor\frac{ab}{\beta}\right\rfloor
&&\text{High word.}
\end{aligned}`;

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

async function main() {
  await withAppBrowser(async ({ baseUrl, client }) => {
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(1000);
    const result = await client.send("Runtime.evaluate", {
      expression: `(async () => {
        const runtime = await import('${baseUrl}${sharedSourceUrl("export/runtime.ts")}');
        const geometry = await import('${baseUrl}/src/office/shared/mathTypeAlignmentGeometry.ts');
        const source = ${JSON.stringify(exactUserAligned)};
        const svg = await runtime.latexToSvg(source, {
          displayMode: true,
          fontSizePt: 12,
          paddingPx: 0,
          paddingXPx: 0,
          paddingYPx: 6,
          background: 'transparent',
          formulaLetterFont: 'times',
          renderer: 'mathjax',
        });
        const rawMathMl = await runtime.latexToMathMl(source, true);
        const annotatedMathMl = geometry.annotateMathTypeAlignmentGeometry(
          rawMathMl,
          svg.svg,
          svg.width,
        );
        const doc = new DOMParser().parseFromString(annotatedMathMl, 'application/xml');
        const table = [...doc.getElementsByTagName('*')].find(
          (element) => element.localName === 'mtable'
        );
        return {
          stops: table?.getAttribute('data-visualtex-mtef-ruler-stops') ?? '',
          columnalign: table?.getAttribute('columnalign') ?? '',
          rowCount: table ? [...table.children].filter((e) => e.localName === 'mtr').length : 0,
          svgWidth: svg.width,
          annotatedMathMl,
        };
      })()`,
      awaitPromise: true,
      returnByValue: true,
    });
    if (result.exceptionDetails) {
      throw new Error(
        result.exceptionDetails.exception?.description ||
          result.exceptionDetails.text ||
          "Runtime.evaluate failed",
      );
    }
    const value = result.result.value as {
      stops: string;
      columnalign: string;
      rowCount: number;
      svgWidth: number;
      annotatedMathMl: string;
    };
    const stops = value.stops
      .split(",")
      .filter(Boolean)
      .map(Number);
    assert.deepEqual(value.columnalign.trim().split(/\s+/), ["right", "left", "right", "left"]);
    assert.equal(value.rowCount, 3);
    assert.equal(stops.length, 2, `Expected two MathType ruler stops, got '${value.stops}'`);
    assert.ok(stops[0] > 0 && stops[1] > stops[0]);
    assert.ok(stops.every((stop) => Number.isInteger(stop) && stop <= 0xffff));
    assert.ok(value.annotatedMathMl.includes('data-visualtex-mtef-ruler-stops='));
    console.log(
      `MathType aligned SVG geometry acceptance passed: rulerStops=${stops.join('/')}, ` +
        `rows=${value.rowCount}, svgWidth=${value.svgWidth.toFixed(3)}px.`,
    );
  }, { server: "dev", windowSize: null });
}

await main();
