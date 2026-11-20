// the OCR model choice persists and warms up the chosen model
import assert from "node:assert/strict";

export const names = ["ocr-model-selection"];

// before the page loads
export async function beforeNavigation({ client }) {
  await client.send("Page.addScriptToEvaluateOnNewDocument", {
    source: `(() => {
          const callbacks = new Map();
          let callbackId = 1;
          window.__visualtexOcrModelSelectionProbe = { prewarmed: [] };
          globalThis.isTauri = true;
          window.__TAURI_INTERNALS__ = {
            metadata: {
              currentWindow: { label: "main" },
              currentWebview: { label: "main" },
            },
            transformCallback(callback, once = false) {
              const id = callbackId++;
              callbacks.set(id, { callback, once });
              return id;
            },
            unregisterCallback(id) {
              callbacks.delete(id);
            },
            async invoke(command, args) {
              if (command === "get_ocr_provider_configuration") {
                return {
                  activeProvider: "local",
                  openAiCompatible: {
                    protocol: "responses",
                    baseUrl: "https://api.openai.com/v1",
                    model: "",
                    prompt: "",
                    hasApiKey: false,
                  },
                  ollama: {
                    baseUrl: "http://127.0.0.1:11434",
                    model: "",
                    prompt: "",
                  },
                  mathpix: {
                    baseUrl: "https://api.mathpix.com",
                    appId: "",
                    hasAppKey: false,
                  },
                  paddleOcr: {
                    model: "PaddleOCR-VL-1.6",
                    hasAccessToken: false,
                  },
                  simpleTex: {
                    model: "standard",
                    hasAccessToken: false,
                  },
                };
              }
              if (command === "get_ocr_runtime_status") {
                return {
                  installed: true,
                  pythonPath: "/fake/python",
                  pythonVersion: "3.10.0",
                  paddleVersion: "3.3.1",
                  paddleocrVersion: "3.7.0",
                  runtimePath: "/fake/runtime",
                  offlineBundleAvailable: true,
                  installedModels: ["PP-FormulaNet_plus-M"],
                  defaultModel: "PP-FormulaNet_plus-M",
                  message: "Fake M-only OCR runtime",
                };
              }
              if (command === "prewarm_ocr_model" || command === "warmup_ocr_model") { // macOS, Windows
                window.__visualtexOcrModelSelectionProbe.prewarmed.push(args?.model ?? null);
                return null;
              }
              if (command === "configure_silent_ocr") return null;
              if (command === "plugin:event|listen") return 1;
              if (command === "plugin:event|unlisten") return null;
              return null;
            },
          };
        })();`,
  });
}

export async function run({ evaluate, sleep, waitForEvaluation }) {
  await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector(".canvas-ocr-model select")),
      }))()`, "OCR model selector");

  await sleep(1500);
  const selectModel = async (model) => {
    await evaluate(`(() => {
          const select = document.querySelector(".canvas-ocr-model select");
          if (!select) return;
          const setter = Object.getOwnPropertyDescriptor(
            HTMLSelectElement.prototype,
            "value",
          )?.set;
          setter?.call(select, ${JSON.stringify(model)});
          select.dispatchEvent(new Event("change", { bubbles: true }));
        })()`);
    await sleep(650);
    return evaluate(`(() => {
          const select = document.querySelector(".canvas-ocr-model select");
          return {
            selected: select?.value ?? "",
            stored: localStorage.getItem("visualtex.ocr.model"),
            prewarmed: window.__visualtexOcrModelSelectionProbe?.prewarmed ?? [],
          };
        })()`);
  };

  const sState = await selectModel("local:PP-FormulaNet_plus-S");
  assert.equal(sState.selected, "local:PP-FormulaNet_plus-S", JSON.stringify(sState));
  assert.equal(sState.stored, "PP-FormulaNet_plus-S", JSON.stringify(sState));
  assert.ok(
    sState.prewarmed.includes("PP-FormulaNet_plus-M"),
    JSON.stringify(sState),
  );

  await evaluate(`document.querySelector('button[aria-label="图片公式识别"]')?.click()`);
  const sDialogState = await waitForEvaluation(`(() => {
        const dialog = document.querySelector(".ocr-dialog");
        const select = dialog?.querySelector(".ocr-model-field select");
        const warning = dialog?.querySelector(".ocr-model-warning strong")?.textContent ?? "";
        const importButton = [...(dialog?.querySelectorAll(".ocr-install-actions button") ?? [])]
          .find((button) => button.textContent?.includes("导入模型包"));
        return {
          ready:
            Boolean(dialog) &&
            select?.value === "PP-FormulaNet_plus-S" &&
            warning.includes("尚未安装") &&
            Boolean(importButton),
          selected: select?.value ?? "",
          warning,
          importButton: importButton?.textContent?.trim() ?? "",
        };
      })()`, "S model remains selected in OCR model manager");

  await evaluate(`(() => {
        const select = document.querySelector(".ocr-dialog .ocr-model-field select");
        if (!select) return;
        const setter = Object.getOwnPropertyDescriptor(
          HTMLSelectElement.prototype,
          "value",
        )?.set;
        setter?.call(select, "PP-FormulaNet_plus-L");
        select.dispatchEvent(new Event("change", { bubbles: true }));
      })()`);
  await sleep(650);
  const lState = await evaluate(`(() => ({
        selected: document.querySelector(".ocr-dialog .ocr-model-field select")?.value ?? "",
        toolbarSelected: document.querySelector(".canvas-ocr-model select")?.value ?? "",
        stored: localStorage.getItem("visualtex.ocr.model"),
        warning: document.querySelector(".ocr-dialog .ocr-model-warning strong")?.textContent ?? "",
        prewarmed: window.__visualtexOcrModelSelectionProbe?.prewarmed ?? [],
      }))()`);
  assert.equal(lState.selected, "PP-FormulaNet_plus-L", JSON.stringify(lState));
  assert.equal(lState.toolbarSelected, "local:PP-FormulaNet_plus-L", JSON.stringify(lState));
  assert.equal(lState.stored, "PP-FormulaNet_plus-L", JSON.stringify(lState));
  assert.match(lState.warning, /尚未安装/);

  console.log(JSON.stringify({ sState, sDialogState, lState }, null, 2));
  console.log("Targeted OCR model selection persistence regression passed");
  return;
}
