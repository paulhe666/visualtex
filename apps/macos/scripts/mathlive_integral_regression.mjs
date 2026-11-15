import assert from "node:assert/strict";
import process from "node:process";
import { withAppBrowser } from "../../shared/visualtex-app/testing/browser_test_runtime.mjs";

const rareIntegralOnly = process.argv.includes("--rare-integrals");
const sleep = (milliseconds) =>
  new Promise((resolve) => setTimeout(resolve, milliseconds));

const rareIntegralCommands = [
  "iiiint",
  "idotsint",
  "dotsint",
  "sqint",
  "sqiint",
  "ointclockwise",
  "varointctrclockwise",
  "varoiint",
  "landupint",
  "landdownint",
  "sumint",
  "intbar",
  "intBar",
  "fint",
  "cirfnint",
  "awint",
  "rppolint",
  "scpolint",
  "npolint",
  "pointint",
  "quatint",
  "intlarhk",
  "intx",
  "intcap",
  "intcup",
  "upint",
  "lowint",
  "intclockwise",
  "varointclockwise",
  "ointctrclockwise",
  "intctrclockwise",
];

function approximatelyEqual(actual, expected, tolerance, message) {
  assert.ok(
    Math.abs(actual - expected) <= tolerance,
    `${message}: expected ${expected} ± ${tolerance}, got ${actual}`,
  );
}

async function main() {
  await withAppBrowser(async ({ baseUrl, client }) => {

    const evaluate = async (expression) => {
      const result = await client.send("Runtime.evaluate", {
        expression,
        awaitPromise: true,
        returnByValue: true,
      });
      if (result.exceptionDetails) {
        throw new Error(
          result.exceptionDetails.exception?.description ??
            result.exceptionDetails.text ??
            "Runtime.evaluate failed",
        );
      }
      return result.result.value;
    };

    const started = Date.now();
    while (Date.now() - started < 12_000) {
      if (
        await evaluate(
          `Boolean(document.querySelector("math-field.visual-mathfield")?.shadowRoot)`,
        )
      ) {
        break;
      }
      await sleep(50);
    }

    const formulas = [
      String.raw`\iint_{a}^{b}`,
      String.raw`\oiint_{a}^{b}`,
      String.raw`\iint\limits_{a}^{b}`,
      String.raw`\oiint\limits_{a}^{b}`,
      String.raw`\iint\nolimits_{a}^{b}`,
      String.raw`\oiint\nolimits_{a}^{b}`,
      String.raw`\frac{\iint_{a}^{b}}{x}`,
      String.raw`\frac{\oiint_{a}^{b}}{x}`,
      String.raw`\iiint_{a}^{b}`,
      String.raw`\oiiint_{a}^{b}`,
      ...rareIntegralCommands.map((command) => `\\${command}_{a}^{b}`),
      ...rareIntegralCommands.map(
        (command) => `\\frac{\\${command}_{a}^{b}}{x}`,
      ),
      String.raw`\iiiint\limits_{a}^{b}`,
      String.raw`\iiiint\nolimits_{a}^{b}`,
      String.raw`\awint\limits_{a}^{b}`,
      String.raw`\intctrclockwise\nolimits_{a}^{b}`,
    ];
    const measurements = await evaluate(`(async () => {
      const field = document.querySelector("math-field.visual-mathfield");
      if (!field?.shadowRoot) throw new Error("Mathfield did not mount.");
      await document.fonts.ready;
      const formulas = ${JSON.stringify(formulas)};
      const results = {};
      const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      for (const latex of formulas) {
        field.setValue(latex, {
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        await frame();
        const root = field.shadowRoot;
        const operator = root.querySelector(".ML__op-symbol");
        if (!operator) throw new Error("Missing operator for " + latex);
        const operatorBounds = operator.getBoundingClientRect();
        const operatorFontSize = Number.parseFloat(getComputedStyle(operator).fontSize);
        const overlay = getComputedStyle(operator, "::after");
        const svgPath = operator.querySelector("svg path");
        const svgPathBounds = svgPath?.getBoundingClientRect();
        const scriptBounds = Array.from(root.querySelectorAll(".ML__mathit"))
          .filter((node) => node.textContent === "a" || node.textContent === "b")
          .map((node) => {
            const bounds = node.getBoundingClientRect();
            return {
              text: node.textContent,
              x: (bounds.left + bounds.width / 2 - operatorBounds.left) / operatorFontSize,
              y: (bounds.top + bounds.height / 2 - (operatorBounds.top + operatorBounds.height / 2)) / operatorFontSize,
            };
          })
          .sort((left, right) => left.text.localeCompare(right.text));
        results[latex] = {
          value: field.getValue("latex"),
          text: operator.textContent,
          classes: Array.from(operator.classList),
          hasError: Boolean(root.querySelector(".ML__error")),
          hasAdjacentScripts: Boolean(root.querySelector(".ML__msubsup")),
          svgPathCount: operator.querySelectorAll("svg path").length,
          svgPath: svgPath?.getAttribute("d") ?? "",
          svgPathBounds: svgPathBounds
            ? {
                width: svgPathBounds.width / operatorFontSize,
                height: svgPathBounds.height / operatorFontSize,
              }
            : null,
          svgViewBox: operator.querySelector("svg")?.getAttribute("viewBox") ?? "",
          width: operatorBounds.width / operatorFontSize,
          height: operatorBounds.height / operatorFontSize,
          overlay: {
            content: overlay.content,
            width: Number.parseFloat(overlay.width) / operatorFontSize,
            height: Number.parseFloat(overlay.height) / operatorFontSize,
            mask: overlay.webkitMaskImage || overlay.maskImage,
          },
          scripts: scriptBounds,
        };
      }
      document.querySelector('[data-toolbar-view="tools"]')?.click();
      await frame();
      document.querySelector('[data-category="calculus"]')?.click();
      await frame();
      return {
        formulas: results,
        globalStyle: Boolean(document.getElementById("visualtex-mathlive-contour-integral-style")),
        shadowStyle: Boolean(field.shadowRoot.getElementById("visualtex-mathlive-contour-integral-shadow-style")),
        previewContours: document.querySelectorAll(".math-preview .visualtex-oiint, .math-preview .visualtex-oiiint").length,
      };
    })()`);

    assert.equal(measurements.globalStyle, true, "global contour style");
    assert.equal(measurements.shadowStyle, true, "shadow contour style");
    assert.ok(measurements.previewContours > 0, "static MathPreview contour class");

    const result = measurements.formulas;
    const comparePair = (baseLatex, contourLatex, expectedClass) => {
      const base = result[baseLatex];
      const contour = result[contourLatex];
      assert.ok(base && contour, `${contourLatex} measurements`);
      assert.ok(contour.value.includes(contourLatex.includes("oiiint") ? "\\oiiint" : "\\oiint"));
      assert.ok(contour.classes.includes(expectedClass), `${expectedClass} class`);
      assert.equal(contour.text, base.text, `${contourLatex} full-size base glyph`);
      approximatelyEqual(contour.width, base.width, 0.02, `${contourLatex} width`);
      approximatelyEqual(contour.height, base.height, 0.02, `${contourLatex} height`);
      assert.equal(contour.scripts.length, 2, `${contourLatex} scripts`);
      assert.equal(base.scripts.length, 2, `${baseLatex} scripts`);
      for (let index = 0; index < 2; index += 1) {
        assert.equal(contour.scripts[index].text, base.scripts[index].text);
        approximatelyEqual(
          contour.scripts[index].x,
          base.scripts[index].x,
          0.03,
          `${contourLatex} ${contour.scripts[index].text} x placement`,
        );
        approximatelyEqual(
          contour.scripts[index].y,
          base.scripts[index].y,
          0.03,
          `${contourLatex} ${contour.scripts[index].text} y placement`,
        );
      }
      assert.notEqual(contour.overlay.content, "none", `${contourLatex} oval`);
      assert.match(contour.overlay.mask, /data:image\/svg\+xml/, `${contourLatex} oval mask`);
    };

    if (!rareIntegralOnly) {
      comparePair(formulas[0], formulas[1], "visualtex-oiint");
      comparePair(formulas[2], formulas[3], "visualtex-oiint");
      comparePair(formulas[4], formulas[5], "visualtex-oiint");
      comparePair(formulas[6], formulas[7], "visualtex-oiint");
      comparePair(formulas[8], formulas[9], "visualtex-oiiint");

      approximatelyEqual(result[formulas[1]].overlay.width, 1.472, 0.02, "oiint Size2 oval width");
      approximatelyEqual(result[formulas[1]].overlay.height, 0.659, 0.02, "oiint Size2 oval height");
      approximatelyEqual(result[formulas[7]].overlay.width, 0.957, 0.02, "oiint Size1 oval width");
      approximatelyEqual(result[formulas[7]].overlay.height, 0.499, 0.02, "oiint Size1 oval height");
      approximatelyEqual(result[formulas[9]].overlay.width, 1.98, 0.02, "oiiint Size2 oval width");
      approximatelyEqual(result[formulas[9]].overlay.height, 0.659, 0.02, "oiiint Size2 oval height");
    }

    for (const command of rareIntegralCommands) {
      const latex = `\\${command}_{a}^{b}`;
      const measurement = result[latex];
      assert.ok(measurement, `${latex} measurements`);
      assert.equal(measurement.hasError, false, `${latex} parses without an error atom`);
      assert.ok(
        measurement.value.includes(`\\${command}`),
        `${latex} preserves its command during serialization`,
      );
      assert.ok(
        measurement.classes.includes("visualtex-integral-svg"),
        `${latex} uses the native SVG integral box`,
      );
      assert.ok(measurement.classes.includes("ML__large-op"), `${latex} uses its large glyph`);
      assert.equal(measurement.svgPathCount, 1, `${latex} has one vector path`);
      assert.ok(measurement.svgPath.length > 20, `${latex} has non-empty path data`);
      assert.ok(measurement.svgPathBounds?.width > 0.25, `${latex} paints a visible path width`);
      assert.ok(measurement.svgPathBounds?.height > 1.5, `${latex} paints a visible large path`);
      assert.match(
        measurement.svgViewBox,
        /^-?[\d.]+ -?[\d.]+ [\d.]+ [\d.]+$/,
        `${latex} has a numeric SVG viewBox`,
      );
      assert.ok(measurement.width > 0.25, `${latex} has a visible width`);
      assert.ok(measurement.height > 1.5, `${latex} has large-operator height`);
      assert.equal(measurement.hasAdjacentScripts, true, `${latex} defaults to adjacent scripts`);

      const smallLatex = `\\frac{\\${command}_{a}^{b}}{x}`;
      const smallMeasurement = result[smallLatex];
      assert.ok(smallMeasurement, `${smallLatex} measurements`);
      assert.equal(smallMeasurement.hasError, false, `${smallLatex} parses without an error atom`);
      assert.ok(
        smallMeasurement.classes.includes("visualtex-integral-svg"),
        `${smallLatex} uses the native SVG integral box`,
      );
      assert.ok(
        smallMeasurement.classes.includes("ML__small-op"),
        `${smallLatex} uses its small glyph`,
      );
      assert.equal(smallMeasurement.svgPathCount, 1, `${smallLatex} has one vector path`);
      assert.ok(smallMeasurement.svgPath.length > 20, `${smallLatex} has non-empty path data`);
      assert.ok(
        smallMeasurement.svgPathBounds?.width > 0.2,
        `${smallLatex} paints a visible path width`,
      );
      assert.ok(
        smallMeasurement.svgPathBounds?.height > 0.7,
        `${smallLatex} paints a visible small path`,
      );
      assert.ok(smallMeasurement.height > 0.7, `${smallLatex} has visible operator height`);
      assert.ok(
        smallMeasurement.height < measurement.height,
        `${smallLatex} is smaller than its display-style glyph`,
      );
    }

    const iiiintLimits = result[String.raw`\iiiint\limits_{a}^{b}`];
    assert.equal(iiiintLimits.hasError, false, "iiiint limits parses");
    assert.equal(iiiintLimits.hasAdjacentScripts, false, "iiiint limits uses over-under scripts");
    assert.ok(iiiintLimits.value.includes(String.raw`\iiiint\limits`), "iiiint limits serializes");

    const iiiintNolimits = result[String.raw`\iiiint\nolimits_{a}^{b}`];
    assert.equal(iiiintNolimits.hasError, false, "iiiint nolimits parses");
    assert.equal(iiiintNolimits.hasAdjacentScripts, true, "iiiint nolimits uses adjacent scripts");
    assert.ok(iiiintNolimits.value.includes(String.raw`\iiiint\nolimits`), "iiiint nolimits serializes");

    const awintLimits = result[String.raw`\awint\limits_{a}^{b}`];
    assert.equal(awintLimits.hasError, false, "awint limits parses");
    assert.equal(awintLimits.hasAdjacentScripts, false, "awint limits uses over-under scripts");
    assert.ok(awintLimits.value.includes(String.raw`\awint\limits`), "awint limits serializes");

    const intctrclockwiseNolimits =
      result[String.raw`\intctrclockwise\nolimits_{a}^{b}`];
    assert.equal(intctrclockwiseNolimits.hasError, false, "intctrclockwise nolimits parses");
    assert.equal(
      intctrclockwiseNolimits.hasAdjacentScripts,
      true,
      "intctrclockwise nolimits uses adjacent scripts",
    );
    assert.ok(
      intctrclockwiseNolimits.value.includes(String.raw`\intctrclockwise\nolimits`),
      "intctrclockwise nolimits serializes",
    );
    assert.equal(
      result[String.raw`\dotsint_{a}^{b}`].svgPath,
      result[String.raw`\idotsint_{a}^{b}`].svgPath,
      "dotsint alias uses the idotsint esint10 outline",
    );
    assert.equal(
      result[String.raw`\ointclockwise_{a}^{b}`].svgPath,
      result[String.raw`\intclockwise_{a}^{b}`].svgPath,
      "MathLive intclockwise spelling uses the esint ointclockwise outline",
    );
    assert.equal(
      result[String.raw`\awint_{a}^{b}`].svgPath,
      result[String.raw`\intctrclockwise_{a}^{b}`].svgPath,
      "awint and intctrclockwise share U+2A11 geometry",
    );

    console.log("MathLive integral regression passed.");
  });
}

await main();
