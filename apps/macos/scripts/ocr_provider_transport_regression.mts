import assert from "node:assert/strict";
import {
  configureOcrTransport,
  getOcrProviderConfiguration,
  recognizeFormulaImage,
  saveOcrProviderConfiguration,
  type OcrProviderConfigurationUpdate,
  type OcrTransport,
} from "../src/ocr/ocrService.ts";

const calls: Array<{ command: string; args?: Record<string, unknown> }> = [];
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
  async listen() {
    return () => undefined;
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
  }
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

console.log("macOS OCR provider transport regression passed.");
