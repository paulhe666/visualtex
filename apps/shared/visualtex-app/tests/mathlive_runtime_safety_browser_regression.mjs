import assert from "node:assert/strict";
import process from "node:process";
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (milliseconds) =>
  new Promise((resolve) => setTimeout(resolve, milliseconds));

async function main() {
  await withAppBrowser(async ({ baseUrl, client }) => {

    const evaluate = async (expression) => {
      const result = await client.send("Runtime.evaluate", {
        expression,
        awaitPromise: true,
        returnByValue: true,
      });
      if (result.exceptionDetails) {
        const description = result.exceptionDetails.exception?.description;
        throw new Error(
          description || result.exceptionDetails.text || "Runtime.evaluate failed",
        );
      }
      return result.result.value;
    };

    await client.send("Page.navigate", { url: baseUrl });
    await sleep(500);
    await evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(650);
    await evaluate(`new Promise((resolve, reject) => {
      const started = performance.now();
      const check = () => {
        const field = document.querySelector("math-field");
        if (field?.shadowRoot) return resolve(true);
        if (performance.now() - started > 8000) {
          return reject(new Error("Mathfield did not mount"));
        }
        setTimeout(check, 30);
      };
      check();
    })`);

    client.events.length = 0;
    const probe = await evaluate(`(async () => {
      globalThis.__visualtexRuntimeSafetyErrors = [];
      globalThis.__visualtexXssProbe = 0;
      const recordError = (kind, value) => {
        globalThis.__visualtexRuntimeSafetyErrors.push(
          kind + ":" + (value?.message ?? String(value ?? "unknown")),
        );
      };
      window.addEventListener("error", (event) => recordError("error", event.error ?? event.message));
      window.addEventListener("unhandledrejection", (event) => recordError("rejection", event.reason));

      const appField = document.querySelector("math-field");
      const probeHost = document.createElement("div");
      probeHost.dataset.visualtexRuntimeSafetyProbe = "true";
      document.body.append(probeHost);
      const field = new window.MathfieldElement();
      field.value = "";
      probeHost.append(field);
      // Force the transient root shape seen in the reported lifecycle. This
      // invokes the real MathLive setter and would throw at its old mode
      // write, while leaving the rest of the model intact for reparse.
      const modelRoot = field._mathfield.model.root;
      Object.defineProperty(modelRoot, "firstChild", {
        configurable: true,
        get: () => undefined,
      });
      const emptyRootBeforeMacroUpdate = !modelRoot.firstChild;
      const originalMacros = field.macros;
      field.macros = {
        ...originalMacros,
        visualtexruntimeprobe: "\\\\mathrm{VT}",
      };
      const macroCommitted = Boolean(field.macros?.visualtexruntimeprobe);
      const emptyRootAfterMacroUpdate = !modelRoot.firstChild;
      delete modelRoot.firstChild;

      field.setValue("\\\\visualtexruntimeprobe");
      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      const macroRendered = !field.shadowRoot?.querySelector(".ML__error");
      const macroValue = field.getValue();

      const attack = String.raw\`\\text{<img src=x onerror="globalThis.__visualtexXssProbe=1"> & 'quoted'}\`;
      field.setValue(attack);
      appField.setValue(attack);
      appField.dispatchEvent(new InputEvent("input", {
        bubbles: true,
        composed: true,
        inputType: "insertText",
      }));
      await new Promise((resolve) => setTimeout(resolve, 250));

      const shadowImages =
        (field.shadowRoot?.querySelectorAll("img").length ?? 0) +
        (appField.shadowRoot?.querySelectorAll("img").length ?? 0);
      const documentImages = document.querySelectorAll('img[src="x"]').length;
      const mathMl = field.getValue("math-ml");
      const escapedMathMl =
        !mathMl.includes("<img") &&
        mathMl.includes("&lt;img") &&
        mathMl.includes("&amp;");

      return {
        emptyRootBeforeMacroUpdate,
        emptyRootAfterMacroUpdate,
        macroCommitted,
        macroRendered,
        macroValue,
        shadowImages,
        documentImages,
        xssProbe: globalThis.__visualtexXssProbe,
        escapedMathMl,
        mathMl,
        runtimeErrors: globalThis.__visualtexRuntimeSafetyErrors,
      };
    })()`);
    await sleep(150);
    console.log(JSON.stringify(probe, null, 2));

    assert.equal(
      probe.emptyRootBeforeMacroUpdate,
      true,
      "the regression must exercise an empty mounted MathLive model",
    );
    assert.equal(probe.emptyRootAfterMacroUpdate, true);
    assert.equal(probe.macroCommitted, true);
    assert.equal(probe.macroRendered, true);
    assert.match(probe.macroValue, /visualtexruntimeprobe/);
    assert.equal(probe.shadowImages, 0, "text markup must not create an image node");
    assert.equal(probe.documentImages, 0, "preview markup must not create an image node");
    assert.equal(probe.xssProbe, 0, "text markup must not execute an event handler");
    assert.equal(probe.escapedMathMl, true, `unsafe MathML: ${probe.mathMl}`);
    assert.deepEqual(probe.runtimeErrors, []);

    const exceptions = client.events.filter(
      (event) => event.method === "Runtime.exceptionThrown",
    );
    assert.equal(
      exceptions.length,
      0,
      `unexpected browser exceptions: ${JSON.stringify(exceptions)}`,
    );

    // A persisted pathological formula used to crash the entire application
    // during FormulaField's auxiliary row-height markup pass. Reproduce a real
    // recursive MathLive box overflow and verify that startup remains usable and
    // the original source stays persisted for manual recovery.
    const persistedProbe = await evaluate(`(() => {
      let deepLatex = "x";
      for (let index = 0; index < 900; index += 1) {
        deepLatex = "\\\\begin{array}{c}" + deepLatex + "\\\\end{array}";
      }
      localStorage.setItem("visualtex-editor", JSON.stringify({
        state: {
          title: "Pathological persisted formula",
          lines: [{ id: "deep-array", latex: deepLatex, mode: "display" }],
          activeLineId: "deep-array",
          history: [],
        },
        version: 0,
      }));
      return { length: deepLatex.length };
    })()`);
    client.events.length = 0;
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(1200);
    const startupSafety = await evaluate(`(() => {
      const root = document.getElementById("root");
      const stored = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      const persistedLatex = stored?.state?.lines?.[0]?.latex ?? "";
      return {
        shell: Boolean(document.querySelector(".app-shell")),
        wholeAppFailure: (root?.innerText ?? "").includes("VisualTeX 界面加载失败"),
        isolatedFormulaFailure: Boolean(document.querySelector(".formula-field-render-fallback")),
        mathfieldMounted: Boolean(document.querySelector("math-field")),
        persistedLength: persistedLatex.length,
      };
    })()`);
    console.log(JSON.stringify({ persistedProbe, startupSafety }, null, 2));
    assert.equal(startupSafety.shell, true, "pathological persisted formula must not destroy the app shell");
    assert.equal(startupSafety.wholeAppFailure, false, "pathological persisted formula must not reach the root error boundary");
    assert.equal(
      startupSafety.isolatedFormulaFailure,
      true,
      "structurally unsafe persisted LaTeX must be blocked before entering MathLive",
    );
    assert.equal(
      startupSafety.persistedLength,
      persistedProbe.length,
      "formula isolation must not discard the persisted LaTeX source",
    );

    const customTileProbe = await evaluate(`(() => {
      let deepLatex = "x";
      for (let index = 0; index < 900; index += 1) {
        deepLatex = "\\\\begin{array}{c}" + deepLatex + "\\\\end{array}";
      }
      localStorage.setItem("visualtex-editor", JSON.stringify({
        state: {
          title: "Safe document",
          lines: [{ id: "safe-line", latex: "x+1", mode: "display" }],
          activeLineId: "safe-line",
          history: [],
        },
        version: 0,
      }));
      localStorage.setItem("visualtex-custom-formula-tiles", JSON.stringify({
        version: 3,
        sections: [{ id: "custom", name: "Custom", createdAt: 0 }],
        tiles: [{
          id: "pathological-preview",
          latex: deepLatex,
          sectionId: "custom",
          color: null,
          createdAt: 0,
        }],
      }));
      return { length: deepLatex.length };
    })()`);
    client.events.length = 0;
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(1200);
    const customTileSafety = await evaluate(`(() => {
      const root = document.getElementById("root");
      return {
        shell: Boolean(document.querySelector(".app-shell")),
        wholeAppFailure: (root?.innerText ?? "").includes("VisualTeX 界面加载失败"),
        mathfieldMounted: Boolean(document.querySelector("math-field")),
        customTileStillStored: (localStorage.getItem("visualtex-custom-formula-tiles") ?? "").includes("pathological-preview"),
      };
    })()`);
    console.log(JSON.stringify({ customTileProbe, customTileSafety }, null, 2));
    assert.equal(customTileSafety.shell, true, "a pathological custom-tile preview must not destroy the app shell");
    assert.equal(customTileSafety.wholeAppFailure, false);
    assert.equal(customTileSafety.mathfieldMounted, true);
    assert.equal(customTileSafety.customTileStillStored, true, "preview isolation must not silently delete user toolbar data");

    console.log("VisualTeX MathLive browser runtime safety regression passed");
  });
}

main().catch((error) => {
  console.error(error instanceof Error ? error.stack : error);
  process.exitCode = 1;
});
