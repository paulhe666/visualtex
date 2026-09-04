import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const rareIntegralOnly = process.argv.includes("--rare-integrals");
const offset = process.pid % 1000;
const previewPort = 17600 + offset;
const debugPort = 22600 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = `/tmp/visualtex-mathlive-integrals-${process.pid}`;
const chromePath =
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
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

async function waitFor(url, timeoutMs = 15_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while the local preview or browser starts.
    }
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

class CdpClient {
  constructor(url) {
    this.url = url;
    this.nextId = 1;
    this.pending = new Map();
  }

  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((resolve, reject) => {
      this.socket.addEventListener("open", resolve, { once: true });
      this.socket.addEventListener("error", reject, { once: true });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      if (!message.id) return;
      const pending = this.pending.get(message.id);
      if (!pending) return;
      this.pending.delete(message.id);
      if (message.error) pending.reject(new Error(message.error.message));
      else pending.resolve(message.result);
    });
  }

  send(method, params = {}) {
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }

  close() {
    this.socket?.close();
  }
}

function approximatelyEqual(actual, expected, tolerance, message) {
  assert.ok(
    Math.abs(actual - expected) <= tolerance,
    `${message}: expected ${expected} ± ${tolerance}, got ${actual}`,
  );
}

async function main() {
  const preview = spawn(
    process.execPath,
    [
      "node_modules/vite/bin/vite.js",
      "preview",
      "--host",
      "127.0.0.1",
      "--port",
      String(previewPort),
      "--strictPort",
    ],
    { cwd: process.cwd(), stdio: "ignore" },
  );
  let chrome;
  let client;

  try {
    await waitFor(baseUrl);
    chrome = spawn(
      chromePath,
      [
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`,
        `--user-data-dir=${chromeProfile}`,
        "--window-size=1400,1000",
        baseUrl,
      ],
      { stdio: "ignore" },
    );

    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const page =
      targets.find(
        (target) => target.type === "page" && target.url.startsWith(baseUrl),
      ) ?? targets.find((target) => target.type === "page");
    if (!page) throw new Error("No Chrome page target found.");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");

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
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(120);
    await rm(chromeProfile, {
      recursive: true,
      force: true,
      maxRetries: 5,
      retryDelay: 100,
    });
  }
}

await main();
