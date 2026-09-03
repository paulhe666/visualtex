import {
  OCR_MODELS,
  type OcrModelName,
  type OcrProviderConfiguration,
  type OcrProviderConfigurationUpdate,
  type OcrProviderId,
  type SimpleTexApiModel,
} from "./ocrService";

export type OcrQuickSelectionGroup = "local" | "api";

export interface OcrQuickSelectionOption {
  id: string;
  group: OcrQuickSelectionGroup;
  labelZh: string;
  labelEn: string;
}

export type ParsedOcrQuickSelection =
  | { kind: "local"; model: OcrModelName }
  | {
      kind: "provider";
      provider: Exclude<OcrProviderId, "local">;
      simpleTexModel?: SimpleTexApiModel;
    };

const API_PROVIDERS = new Set<Exclude<OcrProviderId, "local">>([
  "openai-compatible",
  "ollama",
  "mathpix",
  "paddleocr",
  "simpletex",
]);

export const LOCAL_OCR_QUICK_CONFIGURATION: OcrProviderConfiguration = {
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

export function localOcrSelectionId(model: OcrModelName) {
  return `local:${model}`;
}

export function providerOcrSelectionId(
  provider: Exclude<OcrProviderId, "local">,
  simpleTexModel?: SimpleTexApiModel,
) {
  return simpleTexModel
    ? `provider:${provider}:${simpleTexModel}`
    : `provider:${provider}`;
}

export function parseOcrQuickSelection(
  value: string,
): ParsedOcrQuickSelection | null {
  const [kind, identifier, detail, ...extra] = value.split(":");
  if (extra.length > 0) return null;
  if (kind === "local") {
    const model = OCR_MODELS.find((item) => item.id === identifier)?.id;
    return model ? { kind: "local", model } : null;
  }
  if (
    kind !== "provider" ||
    !API_PROVIDERS.has(identifier as Exclude<OcrProviderId, "local">)
  ) {
    return null;
  }
  const provider = identifier as Exclude<OcrProviderId, "local">;
  if (provider === "simpletex") {
    if (detail !== "standard" && detail !== "turbo") return null;
    return { kind: "provider", provider, simpleTexModel: detail };
  }
  return detail === undefined ? { kind: "provider", provider } : null;
}

export function activeOcrQuickSelectionId(
  configuration: OcrProviderConfiguration,
  localModel: OcrModelName,
) {
  if (configuration.activeProvider === "local") {
    return localOcrSelectionId(localModel);
  }
  if (configuration.activeProvider === "simpletex") {
    return providerOcrSelectionId(
      "simpletex",
      configuration.simpleTex.model,
    );
  }
  return providerOcrSelectionId(configuration.activeProvider);
}

function configuredModelLabel(model: string, isEn: boolean) {
  const normalized = model.trim();
  return normalized || (isEn ? "not configured" : "未配置");
}

export function buildOcrQuickSelectionOptions(
  configuration: OcrProviderConfiguration,
): OcrQuickSelectionOption[] {
  const localOptions = OCR_MODELS.map((item) => ({
    id: localOcrSelectionId(item.id),
    group: "local" as const,
    labelZh: item.labelZh,
    labelEn: item.labelEn,
  }));
  const openAiModelZh = configuredModelLabel(
    configuration.openAiCompatible.model,
    false,
  );
  const openAiModelEn = configuredModelLabel(
    configuration.openAiCompatible.model,
    true,
  );
  const ollamaModelZh = configuredModelLabel(configuration.ollama.model, false);
  const ollamaModelEn = configuredModelLabel(configuration.ollama.model, true);
  const mathpixReady = Boolean(
    configuration.mathpix.appId.trim() && configuration.mathpix.hasAppKey,
  );
  const paddleReady = configuration.paddleOcr.hasAccessToken;
  const simpleTexReady = configuration.simpleTex.hasAccessToken;

  return [
    ...localOptions,
    {
      id: providerOcrSelectionId("openai-compatible"),
      group: "api",
      labelZh: `OpenAI 兼容 · ${openAiModelZh}`,
      labelEn: `OpenAI compatible · ${openAiModelEn}`,
    },
    {
      id: providerOcrSelectionId("ollama"),
      group: "api",
      labelZh: `Ollama · ${ollamaModelZh}`,
      labelEn: `Ollama · ${ollamaModelEn}`,
    },
    {
      id: providerOcrSelectionId("mathpix"),
      group: "api",
      labelZh: mathpixReady ? "Mathpix" : "Mathpix · 未配置",
      labelEn: mathpixReady ? "Mathpix" : "Mathpix · not configured",
    },
    {
      id: providerOcrSelectionId("paddleocr"),
      group: "api",
      labelZh: paddleReady
        ? `PaddleOCR · ${configuration.paddleOcr.model}`
        : "PaddleOCR · 未配置",
      labelEn: paddleReady
        ? `PaddleOCR · ${configuration.paddleOcr.model}`
        : "PaddleOCR · not configured",
    },
    {
      id: providerOcrSelectionId("simpletex", "standard"),
      group: "api",
      labelZh: simpleTexReady
        ? "SimpleTex · 标准模型"
        : "SimpleTex 标准 · 未配置",
      labelEn: simpleTexReady
        ? "SimpleTex · Standard"
        : "SimpleTex Standard · not configured",
    },
    {
      id: providerOcrSelectionId("simpletex", "turbo"),
      group: "api",
      labelZh: simpleTexReady
        ? "SimpleTex · 极速模型"
        : "SimpleTex 极速 · 未配置",
      labelEn: simpleTexReady
        ? "SimpleTex · Turbo"
        : "SimpleTex Turbo · not configured",
    },
  ];
}

export function createOcrProviderQuickUpdate(
  configuration: OcrProviderConfiguration,
  selection: Extract<ParsedOcrQuickSelection, { kind: "provider" }> | null,
): OcrProviderConfigurationUpdate {
  return {
    activeProvider: selection?.provider ?? "local",
    openAiCompatible: {
      protocol: configuration.openAiCompatible.protocol,
      baseUrl: configuration.openAiCompatible.baseUrl,
      model: configuration.openAiCompatible.model,
      prompt: configuration.openAiCompatible.prompt,
    },
    ollama: {
      baseUrl: configuration.ollama.baseUrl,
      model: configuration.ollama.model,
      prompt: configuration.ollama.prompt,
    },
    mathpix: {
      baseUrl: configuration.mathpix.baseUrl,
      appId: configuration.mathpix.appId,
    },
    paddleOcr: {
      model: configuration.paddleOcr.model,
    },
    simpleTex: {
      model: selection?.simpleTexModel ?? configuration.simpleTex.model,
    },
  };
}
