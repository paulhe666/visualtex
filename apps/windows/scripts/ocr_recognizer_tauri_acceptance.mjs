import assert from "node:assert/strict";

const debugPort = Number.parseInt(process.env.VISUALTEX_CDP_PORT ?? "19333", 10);
const endpoint = `http://127.0.0.1:${debugPort}`;
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

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

async function main() {
  const targets = await (await fetch(`${endpoint}/json/list`)).json();
  const page = targets.find(
    (target) => target.type === "page" && target.url.startsWith("http://tauri.localhost"),
  );
  if (!page) throw new Error(`No VisualTeX Tauri WebView2 page found at ${endpoint}`);

  const client = new CdpClient(page.webSocketDebuggerUrl);
  await client.connect();
  await client.send("Runtime.enable");

  const evaluate = async (expression) => {
    const response = await client.send("Runtime.evaluate", {
      expression,
      awaitPromise: true,
      returnByValue: true,
    });
    if (response.exceptionDetails) {
      throw new Error(
        response.exceptionDetails.exception?.description ||
          response.exceptionDetails.text ||
          "Runtime.evaluate failed",
      );
    }
    return response.result.value;
  };

  const waitFor = async (expression, description, timeoutMs = 8000) => {
    const started = Date.now();
    let lastValue;
    while (Date.now() - started < timeoutMs) {
      lastValue = await evaluate(expression);
      if (lastValue?.ready) return lastValue;
      await sleep(80);
    }
    throw new Error(`Timed out waiting for ${description}: ${JSON.stringify(lastValue)}`);
  };

  const selectExpression = `(() => {
    const select = document.querySelector(
      'select[aria-label="OCR 识别器"], select[aria-label="OCR recognizer"]',
    );
    if (!select) {
      return {
        ready: false,
        selects: [...document.querySelectorAll("select")].map((item) => ({
          ariaLabel: item.getAttribute("aria-label"),
          value: item.value,
        })),
      };
    }
    return {
      ready: true,
      ariaLabel: select.getAttribute("aria-label"),
      value: select.value,
      groups: [...select.querySelectorAll("optgroup")].map((group) => ({
        label: group.label,
        options: [...group.querySelectorAll("option")].map((option) => ({
          value: option.value,
          text: option.textContent ?? "",
          selected: option.selected,
        })),
      })),
      options: [...select.options].map((option) => ({
        value: option.value,
        text: option.textContent ?? "",
        selected: option.selected,
      })),
    };
  })()`;

  const initialUi = await waitFor(selectExpression, "OCR recognizer select");
  const expectedLocal = [
    "local:PP-FormulaNet_plus-S",
    "local:PP-FormulaNet_plus-M",
    "local:PP-FormulaNet_plus-L",
  ];
  const expectedProviders = [
    "provider:openai-compatible",
    "provider:ollama",
    "provider:mathpix",
    "provider:paddleocr",
    "provider:simpletex",
  ];
  const optionValues = initialUi.options.map((item) => item.value);
  assert.deepEqual(optionValues.slice(0, 3), expectedLocal);
  for (const provider of expectedProviders) {
    assert.ok(optionValues.includes(provider), `Missing OCR provider option: ${provider}`);
  }
  assert.equal(initialUi.options.length, 8, "OCR recognizer select must expose 3 local models + 5 APIs");
  assert.equal(initialUi.groups.length, 2, "OCR recognizer select must group local models and APIs");

  const nativeInitial = await evaluate(`window.__TAURI_INTERNALS__.invoke("get_ocr_provider_configuration")`);
  assert.ok(nativeInitial && typeof nativeInitial === "object");

  // Test preparation only: keep Local active, but make the isolated test app's
  // Ollama provider valid. The operation under test below is the real React select
  // change, which must persist the provider through the production native command.
  const prepared = await evaluate(`(async () => {
    const current = await window.__TAURI_INTERNALS__.invoke("get_ocr_provider_configuration");
    return window.__TAURI_INTERNALS__.invoke("save_ocr_provider_configuration", {
      configuration: {
        activeProvider: "local",
        openAiCompatible: {
          protocol: current.openAiCompatible.protocol,
          baseUrl: current.openAiCompatible.baseUrl,
          model: current.openAiCompatible.model,
          prompt: current.openAiCompatible.prompt,
        },
        ollama: {
          baseUrl: current.ollama.baseUrl || "http://127.0.0.1:11434",
          model: current.ollama.model || "visualtex-ocr-ui-acceptance",
          prompt: current.ollama.prompt,
        },
        mathpix: {
          baseUrl: current.mathpix.baseUrl,
          appId: current.mathpix.appId,
        },
        paddleOcr: { model: current.paddleOcr.model },
        simpleTex: { model: current.simpleTex.model },
      },
    });
  })()`);
  assert.equal(prepared.activeProvider, "local");

  const switchResult = await evaluate(`(() => {
    const select = document.querySelector(
      'select[aria-label="OCR 识别器"], select[aria-label="OCR recognizer"]',
    );
    if (!select) return { ready: false };
    select.value = "provider:ollama";
    select.dispatchEvent(new Event("change", { bubbles: true }));
    return { ready: true, valueAfterDispatch: select.value };
  })()`);
  assert.ok(switchResult.ready);

  const switched = await waitFor(`(async () => {
    const select = document.querySelector(
      'select[aria-label="OCR 识别器"], select[aria-label="OCR recognizer"]',
    );
    const native = await window.__TAURI_INTERNALS__.invoke("get_ocr_provider_configuration");
    return {
      ready: select?.value === "provider:ollama" && native.activeProvider === "ollama",
      selectValue: select?.value ?? "",
      activeProvider: native.activeProvider,
    };
  })()`, "UI switch to Ollama provider");

  await evaluate(`(() => {
    const button = document.querySelector(
      'button[aria-label="图片公式识别"], button[aria-label="Recognize formula image"]',
    );
    button?.click();
    return Boolean(button);
  })()`);
  const apiDialogState = await waitFor(`(() => {
    const dialog = document.querySelector(".ocr-dialog");
    const providerSelect = dialog?.querySelector(".ocr-provider-card select");
    return {
      ready: Boolean(dialog && providerSelect?.value === "ollama"),
      provider: providerSelect?.value ?? "",
      providerText: dialog?.querySelector(".ocr-provider-heading span:last-child")?.textContent ?? "",
      hasApiRuntimeCard: Boolean(dialog?.querySelector(".ocr-api-runtime-card")),
      hasLocalRuntimeCard: Boolean(
        dialog?.querySelector(".ocr-runtime-card:not(.ocr-api-runtime-card)"),
      ),
    };
  })()`, "OCR settings dialog in API mode");
  assert.equal(apiDialogState.provider, "ollama");
  assert.equal(
    apiDialogState.hasApiRuntimeCard,
    true,
    "API-mode OCR settings must show the API status card",
  );
  assert.equal(
    apiDialogState.hasLocalRuntimeCard,
    false,
    "API-mode OCR settings must not mount the local runtime checker",
  );
  await evaluate(`(() => {
    const button = document.querySelector(
      'button[aria-label="关闭 OCR"], button[aria-label="Close OCR"]',
    );
    button?.click();
    return Boolean(button);
  })()`);
  await waitFor(`(() => ({ ready: !document.querySelector(".ocr-dialog") }))()`, "OCR settings dialog closed");

  // Restore the isolated test app to Local through the same production UI path.
  await evaluate(`(() => {
    const select = document.querySelector(
      'select[aria-label="OCR 识别器"], select[aria-label="OCR recognizer"]',
    );
    if (!select) return false;
    select.value = "local:PP-FormulaNet_plus-M";
    select.dispatchEvent(new Event("change", { bubbles: true }));
    return true;
  })()`);
  const restored = await waitFor(`(async () => {
    const select = document.querySelector(
      'select[aria-label="OCR 识别器"], select[aria-label="OCR recognizer"]',
    );
    const native = await window.__TAURI_INTERNALS__.invoke("get_ocr_provider_configuration");
    return {
      ready: select?.value === "local:PP-FormulaNet_plus-M" && native.activeProvider === "local",
      selectValue: select?.value ?? "",
      activeProvider: native.activeProvider,
    };
  })()`, "UI switch back to local OCR");

  console.log(JSON.stringify({ initialUi, nativeInitial, switched, restored }, null, 2));
  client.close();
  console.log("Real Tauri/WebView2 OCR recognizer acceptance passed.");
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
