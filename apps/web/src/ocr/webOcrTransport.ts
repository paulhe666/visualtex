// The shared OCR service's backend in a browser (configureOcrTransport): the online services of webOcrService.ts,
// reached from the page and, for the services that refuse browser requests, through the site's relay worker. It
// answers the commands the editor uses outside the OCR window: the provider configuration (the recognizer menu)
// and recognizing a pasted image. Local models and their runtime belong to the desktop app.
import type {
  OcrProviderConfiguration,
  OcrProviderConfigurationUpdate,
  OcrTransport,
} from "../../../shared/visualtex-app/src/ocr/ocrService";
import {
  loadWebOcrConfiguration,
  recognizeFormulaWithWebApi,
  saveWebOcrConfiguration,
  type WebOcrConfiguration,
  type WebOcrProvider,
} from "./webOcrService";

const WEB_PROVIDERS = new Set<string>(["openai-compatible", "mathpix", "paddleocr", "simpletex"]);
const MIME_BY_EXTENSION: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  webp: "image/webp",
  bmp: "image/bmp",
  tif: "image/tiff",
  tiff: "image/tiff",
};

/** The browser's configuration in the shape the desktop backend answers (secrets only as "set or not"). */
export function webOcrProviderConfiguration(configuration: WebOcrConfiguration = loadWebOcrConfiguration()): OcrProviderConfiguration {
  return {
    activeProvider: configuration.activeProvider,
    openAiCompatible: {
      protocol: configuration.openAiCompatible.protocol,
      baseUrl: configuration.openAiCompatible.baseUrl,
      model: configuration.openAiCompatible.model,
      prompt: configuration.openAiCompatible.prompt,
      hasApiKey: Boolean(configuration.openAiCompatible.apiKey),
    },
    ollama: { baseUrl: "", model: "", prompt: "" },
    mathpix: {
      baseUrl: configuration.mathpix.baseUrl,
      appId: configuration.mathpix.appId,
      hasAppKey: Boolean(configuration.mathpix.appKey),
    },
    paddleOcr: { model: configuration.paddleOcr.model, hasAccessToken: Boolean(configuration.paddleOcr.accessToken) },
    simpleTex: { model: configuration.simpleTex.model, hasAccessToken: Boolean(configuration.simpleTex.accessToken) },
  };
}

/** A desktop-shaped update applied to the browser's configuration (a secret left out stays as it is). */
function saveUpdate(update: OcrProviderConfigurationUpdate): OcrProviderConfiguration {
  const current = loadWebOcrConfiguration();
  if (!WEB_PROVIDERS.has(update.activeProvider)) {
    throw new Error("网页版只支持在线 OCR 服务（OpenAI 兼容 API、Mathpix、PaddleOCR、SimpleTex）");
  }
  const next: WebOcrConfiguration = {
    activeProvider: update.activeProvider as WebOcrProvider,
    openAiCompatible: {
      protocol: update.openAiCompatible.protocol,
      baseUrl: update.openAiCompatible.baseUrl,
      model: update.openAiCompatible.model,
      prompt: update.openAiCompatible.prompt,
      apiKey: update.openAiCompatible.clearApiKey ? "" : update.openAiCompatible.apiKey ?? current.openAiCompatible.apiKey,
    },
    mathpix: {
      baseUrl: update.mathpix.baseUrl,
      appId: update.mathpix.appId,
      appKey: update.mathpix.clearAppKey ? "" : update.mathpix.appKey ?? current.mathpix.appKey,
    },
    paddleOcr: {
      model: current.paddleOcr.model,
      accessToken: update.paddleOcr.clearAccessToken ? "" : update.paddleOcr.accessToken ?? current.paddleOcr.accessToken,
    },
    simpleTex: {
      model: update.simpleTex.model,
      accessToken: update.simpleTex.clearAccessToken ? "" : update.simpleTex.accessToken ?? current.simpleTex.accessToken,
    },
  };
  saveWebOcrConfiguration(next);
  return webOcrProviderConfiguration(loadWebOcrConfiguration());
}

let cancelled = 0;

async function recognize(request: { bytes: number[]; extension: string }) {
  const run = cancelled;
  const file = new File([Uint8Array.from(request.bytes)], `image.${request.extension}`, {
    type: MIME_BY_EXTENSION[request.extension] ?? "image/png",
  });
  const result = await recognizeFormulaWithWebApi(file, loadWebOcrConfiguration());
  if (run !== cancelled) throw new Error("OCR_CANCELLED");
  return {
    provider: result.provider,
    model: result.model,
    elapsedMs: result.elapsedMs,
    processedWidth: 0,
    processedHeight: 0,
    backgroundInverted: false,
    backgroundLuminance: 1,
    formulas: result.formulas.map((latex) => ({ latex })),
  };
}

export const webOcrTransport: OcrTransport = {
  environment: "web",
  async invoke<T>(command: string, args?: Record<string, unknown>): Promise<T> {
    switch (command) {
      case "get_ocr_provider_configuration":
        return webOcrProviderConfiguration() as T;
      case "save_ocr_provider_configuration":
        return saveUpdate(args?.configuration as OcrProviderConfigurationUpdate) as T;
      case "recognize_formula_image":
        return (await recognize(args?.request as { bytes: number[]; extension: string })) as T;
      case "cancel_ocr_recognition":
        // the request in flight cannot be withdrawn; its answer is dropped
        cancelled += 1;
        return undefined as T;
      default:
        throw new Error(`"${command}" is only available in the VisualTeX desktop app`);
    }
  },
  async listen() {
    // no progress events: the services answer once
    return () => undefined;
  },
};
