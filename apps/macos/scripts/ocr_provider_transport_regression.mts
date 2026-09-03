import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import {
  configureOcrTransport,
  getOcrProviderConfiguration,
  listenOcrProviderConfigurationChanged,
  listenOcrRecognitionProgress,
  recognizeFormulaImage,
  saveOcrProviderConfiguration,
  type OcrProviderConfigurationUpdate,
  type OcrTransport,
} from "../src/ocr/ocrService.ts";
import {
  activeOcrQuickSelectionId,
  buildOcrQuickSelectionOptions,
  createOcrProviderQuickUpdate,
  parseOcrQuickSelection,
  providerOcrSelectionId,
} from "../src/ocr/ocrQuickSelection.ts";

const calls: Array<{ command: string; args?: Record<string, unknown> }> = [];
const eventHandlers = new Map<string, (event: {
  event: string;
  id: number;
  payload: unknown;
}) => void>();
const providerView = {
  activeProvider: "openai-compatible" as const,
  openAiCompatible: {
    protocol: "responses" as const,
    baseUrl: "https://api.example.test/v1",
    model: "vision-model",
    prompt: "Return formula JSON",
    hasApiKey: true,
  },
  ollama: {
    baseUrl: "http://127.0.0.1:11434",
    model: "vision-model",
    prompt: "Return formula JSON",
  },
  mathpix: {
    baseUrl: "https://api.mathpix.com",
    appId: "app-id",
    hasAppKey: false,
  },
  paddleOcr: {
    model: "PaddleOCR-VL-1.6" as const,
    hasAccessToken: true,
  },
  simpleTex: {
    model: "standard" as const,
    hasAccessToken: true,
  },
};

const transport: OcrTransport = {
  environment: "desktop",
  async invoke<T>(command: string, args?: Record<string, unknown>): Promise<T> {
    calls.push({ command, args });
    if (command === "get_ocr_provider_configuration") return providerView as T;
    if (command === "save_ocr_provider_configuration") return providerView as T;
    if (command === "recognize_formula_image") {
      return {
        provider: "openai-compatible",
        model: "vision-model",
        elapsedMs: 37,
        processedWidth: 800,
        processedHeight: 240,
        backgroundInverted: false,
        backgroundLuminance: 0,
        formulas: [{ latex: "x=1" }, { latex: "y=2" }],
      } as T;
    }
    throw new Error(`Unexpected command: ${command}`);
  },
  async listen(eventName, handler) {
    eventHandlers.set(eventName, handler);
    return () => eventHandlers.delete(eventName);
  },
};

configureOcrTransport(transport);
assert.deepEqual(await getOcrProviderConfiguration(), providerView);

const update: OcrProviderConfigurationUpdate = {
  activeProvider: "openai-compatible",
  openAiCompatible: {
    protocol: "responses",
    baseUrl: "https://api.example.test/v1",
    model: "vision-model",
    prompt: "Return formula JSON",
    apiKey: "secret",
  },
  ollama: {
    baseUrl: "http://127.0.0.1:11434",
    model: "vision-model",
    prompt: "Return formula JSON",
  },
  mathpix: {
    baseUrl: "https://api.mathpix.com",
    appId: "app-id",
  },
  paddleOcr: {
    model: "PaddleOCR-VL-1.6",
    accessToken: "paddle-secret",
  },
  simpleTex: {
    model: "standard",
    accessToken: "simpletex-secret",
  },
};
assert.deepEqual(await saveOcrProviderConfiguration(update), providerView);
assert.deepEqual(calls[1], {
  command: "save_ocr_provider_configuration",
  args: { configuration: update },
});

const recognition = await recognizeFormulaImage({
  bytes: [137, 80, 78, 71],
  extension: "png",
  model: "PP-FormulaNet_plus-M",
});
assert.equal(recognition.provider, "openai-compatible");
assert.deepEqual(recognition.formulas.map((item) => item.latex), ["x=1", "y=2"]);
assert.equal(calls[2].command, "recognize_formula_image");
assert.deepEqual(calls[2].args, {
  request: {
    bytes: [137, 80, 78, 71],
    extension: "png",
    model: "PP-FormulaNet_plus-M",
  },
});

let synchronizedProvider = "";
await listenOcrProviderConfigurationChanged((configuration) => {
  synchronizedProvider = configuration.activeProvider;
});
eventHandlers.get("ocr-provider-configuration-changed")?.({
  event: "ocr-provider-configuration-changed",
  id: 2,
  payload: { ...providerView, activeProvider: "paddleocr" },
});
assert.equal(synchronizedProvider, "paddleocr");

const quickOptions = buildOcrQuickSelectionOptions(providerView);
assert.equal(
  activeOcrQuickSelectionId(providerView, "PP-FormulaNet_plus-M"),
  providerOcrSelectionId("openai-compatible"),
);
assert.ok(
  quickOptions.some(
    (option) =>
      option.id === providerOcrSelectionId("paddleocr") &&
      option.labelZh.includes("PaddleOCR"),
  ),
);
assert.deepEqual(parseOcrQuickSelection("provider:simpletex:turbo"), {
  kind: "provider",
  provider: "simpletex",
  simpleTexModel: "turbo",
});
const quickUpdate = createOcrProviderQuickUpdate(
  providerView,
  parseOcrQuickSelection("provider:simpletex:turbo") as Extract<
    NonNullable<ReturnType<typeof parseOcrQuickSelection>>,
    { kind: "provider" }
  >,
);
assert.equal(quickUpdate.activeProvider, "simpletex");
assert.equal(quickUpdate.simpleTex.model, "turbo");
assert.equal(quickUpdate.openAiCompatible.model, "vision-model");
assert.ok(!("apiKey" in quickUpdate.openAiCompatible));

const receivedStages: string[] = [];
await listenOcrRecognitionProgress((progress) => {
  receivedStages.push(progress.stage);
});
for (const stage of ["api-submit", "api-queued", "api-inference", "api-result"]) {
  eventHandlers.get("ocr-recognition-progress")?.({
    event: "ocr-recognition-progress",
    id: 1,
    payload: {
      event: "progress",
      id: "remote-1",
      stage,
      message: `PaddleOCR ${stage}`,
      model: "PP-FormulaNet_plus-M",
    },
  });
}
assert.deepEqual(receivedStages, [
  "api-submit",
  "api-queued",
  "api-inference",
  "api-result",
]);

for (const sourcePath of [
  "../src/App.tsx",
  "../src/components/OcrDialog.tsx",
  "../src/office/dialog/OfficeDialogApp.tsx",
]) {
  const source = readFileSync(new URL(sourcePath, import.meta.url), "utf8");
  const subscriptionLine = source
    .split("\n")
    .find((line) => line.includes("unlisten = await listenOcrRecognitionProgress"));
  assert.ok(subscriptionLine, `${sourcePath} must subscribe to OCR progress`);
  assert.ok(
    subscriptionLine.trimStart().startsWith("unlisten ="),
    `${sourcePath} must subscribe for remote providers as well as local OCR`,
  );
}

console.log("macOS OCR provider transport regression passed.");
