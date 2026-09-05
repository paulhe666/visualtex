export type WebOcrProvider =
  | "openai-compatible"
  | "mathpix"
  | "paddleocr"
  | "simpletex";

export type OpenAiCompatibleProtocol = "responses" | "chat-completions";
export type SimpleTexModel = "standard" | "turbo";

export interface WebOcrConfiguration {
  activeProvider: WebOcrProvider;
  openAiCompatible: {
    protocol: OpenAiCompatibleProtocol;
    baseUrl: string;
    model: string;
    prompt: string;
    apiKey: string;
  };
  mathpix: {
    baseUrl: string;
    appId: string;
    appKey: string;
  };
  paddleOcr: {
    model: "PaddleOCR-VL-1.6";
    accessToken: string;
  };
  simpleTex: {
    model: SimpleTexModel;
    accessToken: string;
  };
}

export interface WebOcrProgress {
  stage: "upload" | "queued" | "recognizing" | "result";
  messageZh: string;
  messageEn: string;
}

export interface WebOcrResult {
  provider: WebOcrProvider;
  model: string;
  elapsedMs: number;
  formulas: string[];
}

const PUBLIC_CONFIGURATION_KEY = "visualtex.web.ocr.configuration.v1";
const SECRET_CONFIGURATION_KEY = "visualtex.web.ocr.secrets.v1";
const PADDLE_JOBS_URL = "/api/ocr/paddle/jobs";
const SIMPLETEX_ENDPOINTS: Record<SimpleTexModel, string> = {
  standard: "/api/ocr/simpletex?model=standard",
  turbo: "/api/ocr/simpletex?model=turbo",
};
const DEFAULT_PROMPT =
  "Recognize every mathematical formula in this image. Return only JSON in the form {\"formulas\":[{\"latex\":\"...\"}]}. Preserve LaTeX structure and keep separate formula rows separate.";

export const DEFAULT_WEB_OCR_CONFIGURATION: WebOcrConfiguration = {
  activeProvider: "simpletex",
  openAiCompatible: {
    protocol: "responses",
    baseUrl: "https://api.openai.com/v1",
    model: "gpt-5-mini",
    prompt: DEFAULT_PROMPT,
    apiKey: "",
  },
  mathpix: {
    baseUrl: "https://api.mathpix.com",
    appId: "",
    appKey: "",
  },
  paddleOcr: {
    model: "PaddleOCR-VL-1.6",
    accessToken: "",
  },
  simpleTex: {
    model: "standard",
    accessToken: "",
  },
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function readJson(storage: Storage, key: string): unknown {
  try {
    const raw = storage.getItem(key);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function cloneDefaults(): WebOcrConfiguration {
  return {
    activeProvider: DEFAULT_WEB_OCR_CONFIGURATION.activeProvider,
    openAiCompatible: { ...DEFAULT_WEB_OCR_CONFIGURATION.openAiCompatible },
    mathpix: { ...DEFAULT_WEB_OCR_CONFIGURATION.mathpix },
    paddleOcr: { ...DEFAULT_WEB_OCR_CONFIGURATION.paddleOcr },
    simpleTex: { ...DEFAULT_WEB_OCR_CONFIGURATION.simpleTex },
  };
}

export function loadWebOcrConfiguration(): WebOcrConfiguration {
  const next = cloneDefaults();
  if (typeof window === "undefined") return next;

  const publicValue = readJson(window.localStorage, PUBLIC_CONFIGURATION_KEY);
  if (isRecord(publicValue)) {
    if (
      publicValue.activeProvider === "openai-compatible" ||
      publicValue.activeProvider === "mathpix" ||
      publicValue.activeProvider === "paddleocr" ||
      publicValue.activeProvider === "simpletex"
    ) {
      next.activeProvider = publicValue.activeProvider;
    }
    const openAi = isRecord(publicValue.openAiCompatible)
      ? publicValue.openAiCompatible
      : {};
    if (
      openAi.protocol === "responses" ||
      openAi.protocol === "chat-completions"
    ) {
      next.openAiCompatible.protocol = openAi.protocol;
    }
    if (typeof openAi.baseUrl === "string")
      next.openAiCompatible.baseUrl = openAi.baseUrl;
    if (typeof openAi.model === "string")
      next.openAiCompatible.model = openAi.model;
    if (typeof openAi.prompt === "string")
      next.openAiCompatible.prompt = openAi.prompt;

    const mathpix = isRecord(publicValue.mathpix) ? publicValue.mathpix : {};
    if (typeof mathpix.baseUrl === "string")
      next.mathpix.baseUrl = mathpix.baseUrl;
    if (typeof mathpix.appId === "string") next.mathpix.appId = mathpix.appId;

    const simpleTex = isRecord(publicValue.simpleTex)
      ? publicValue.simpleTex
      : {};
    if (simpleTex.model === "standard" || simpleTex.model === "turbo") {
      next.simpleTex.model = simpleTex.model;
    }
  }

  const secrets = readJson(window.sessionStorage, SECRET_CONFIGURATION_KEY);
  if (isRecord(secrets)) {
    if (typeof secrets.openAiApiKey === "string")
      next.openAiCompatible.apiKey = secrets.openAiApiKey;
    if (typeof secrets.mathpixAppKey === "string")
      next.mathpix.appKey = secrets.mathpixAppKey;
    if (typeof secrets.paddleAccessToken === "string")
      next.paddleOcr.accessToken = secrets.paddleAccessToken;
    if (typeof secrets.simpleTexAccessToken === "string")
      next.simpleTex.accessToken = secrets.simpleTexAccessToken;
  }
  return next;
}

export function saveWebOcrConfiguration(
  configuration: WebOcrConfiguration,
): void {
  if (typeof window === "undefined") return;
  const normalized = validateConfiguration(configuration, false);
  window.localStorage.setItem(
    PUBLIC_CONFIGURATION_KEY,
    JSON.stringify({
      activeProvider: normalized.activeProvider,
      openAiCompatible: {
        protocol: normalized.openAiCompatible.protocol,
        baseUrl: normalized.openAiCompatible.baseUrl,
        model: normalized.openAiCompatible.model,
        prompt: normalized.openAiCompatible.prompt,
      },
      mathpix: {
        baseUrl: normalized.mathpix.baseUrl,
        appId: normalized.mathpix.appId,
      },
      paddleOcr: {
        model: normalized.paddleOcr.model,
      },
      simpleTex: {
        model: normalized.simpleTex.model,
      },
    }),
  );
  window.sessionStorage.setItem(
    SECRET_CONFIGURATION_KEY,
    JSON.stringify({
      openAiApiKey: normalized.openAiCompatible.apiKey,
      mathpixAppKey: normalized.mathpix.appKey,
      paddleAccessToken: normalized.paddleOcr.accessToken,
      simpleTexAccessToken: normalized.simpleTex.accessToken,
    }),
  );
}

function normalizeBaseUrl(value: string, label: string): string {
  const trimmed = value.trim().replace(/\/+$/, "");
  let url: URL;
  try {
    url = new URL(trimmed);
  } catch {
    throw new Error(`${label}不是有效网址`);
  }
  const secure =
    url.protocol === "https:" ||
    (url.protocol === "http:" &&
      (url.hostname === "127.0.0.1" || url.hostname === "localhost") &&
      typeof window !== "undefined" &&
      window.location.protocol === "http:");
  if (!secure) throw new Error(`${label}必须使用 HTTPS`);
  if (url.username || url.password) {
    throw new Error(`${label}不能包含用户名或密码`);
  }
  return trimmed;
}

function validateConfiguration(
  value: WebOcrConfiguration,
  requireSecret: boolean,
): WebOcrConfiguration {
  const next: WebOcrConfiguration = {
    activeProvider: value.activeProvider,
    openAiCompatible: {
      protocol: value.openAiCompatible.protocol,
      baseUrl: normalizeBaseUrl(
        value.openAiCompatible.baseUrl,
        "OpenAI 兼容接口地址",
      ),
      model: value.openAiCompatible.model.trim(),
      prompt: value.openAiCompatible.prompt.trim(),
      apiKey: value.openAiCompatible.apiKey.trim(),
    },
    mathpix: {
      baseUrl: normalizeBaseUrl(value.mathpix.baseUrl, "Mathpix 地址"),
      appId: value.mathpix.appId.trim(),
      appKey: value.mathpix.appKey.trim(),
    },
    paddleOcr: {
      model: "PaddleOCR-VL-1.6",
      accessToken: value.paddleOcr.accessToken.trim(),
    },
    simpleTex: {
      model: value.simpleTex.model,
      accessToken: value.simpleTex.accessToken.trim(),
    },
  };
  if (!requireSecret) return next;

  if (next.activeProvider === "openai-compatible") {
    if (!next.openAiCompatible.model)
      throw new Error("请填写 OpenAI 兼容模型名称");
    if (!next.openAiCompatible.prompt)
      throw new Error("请填写公式识别提示词");
    if (!next.openAiCompatible.apiKey)
      throw new Error("请填写 OpenAI 兼容 API Key");
  } else if (next.activeProvider === "mathpix") {
    if (!next.mathpix.appId) throw new Error("请填写 Mathpix app_id");
    if (!next.mathpix.appKey) throw new Error("请填写 Mathpix app_key");
  } else if (next.activeProvider === "paddleocr") {
    if (!next.paddleOcr.accessToken)
      throw new Error("请填写 PaddleOCR Access Token");
  } else if (!next.simpleTex.accessToken) {
    throw new Error("请填写 SimpleTex UAT");
  }
  return next;
}

function appendEndpoint(baseUrl: string, route: string): string {
  const normalizedRoute = route.replace(/^\/+/, "");
  return baseUrl.toLowerCase().endsWith(`/${normalizedRoute.toLowerCase()}`)
    ? baseUrl
    : `${baseUrl}/${normalizedRoute}`;
}

async function responseJson(response: Response, label: string): Promise<unknown> {
  const text = await response.text();
  let value: unknown;
  try {
    value = JSON.parse(text);
  } catch {
    value = null;
  }
  if (!response.ok) {
    const detail =
      isRecord(value) &&
      (typeof value.message === "string"
        ? value.message
        : isRecord(value.error) && typeof value.error.message === "string"
          ? value.error.message
          : typeof value.error === "string"
            ? value.error
            : "");
    throw new Error(
      `${label}返回 HTTP ${response.status}${detail ? `：${detail}` : ""}`,
    );
  }
  if (value === null) throw new Error(`${label}返回的不是 JSON`);
  return value;
}

async function fetchWithTimeout(
  input: RequestInfo | URL,
  init: RequestInit,
  timeoutMs: number,
): Promise<Response> {
  const controller = new AbortController();
  const timer = window.setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(input, { ...init, signal: controller.signal });
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") {
      throw new Error("OCR 请求超时");
    }
    if (error instanceof TypeError) {
      throw new Error(
        "浏览器无法直接连接该 OCR 接口。请确认接口允许当前网站跨域访问；VisualTeX 不会把密钥或图片转发到自有服务器。",
      );
    }
    throw error;
  } finally {
    window.clearTimeout(timer);
  }
}

function stripFences(value: string): string {
  return value
    .trim()
    .replace(/^\`\`\`(?:json|latex|tex)?\s*/i, "")
    .replace(/\s*\`\`\`$/, "")
    .trim();
}

function stripOuterMath(value: string): string {
  let current = stripFences(value).trim();
  const wrappers: Array<[string, string]> = [
    ["$$", "$$"],
    ["\\[", "\\]"],
    ["\\(", "\\)"],
    ["$", "$"],
  ];
  for (const [start, end] of wrappers) {
    if (
      current.startsWith(start) &&
      current.endsWith(end) &&
      current.length > start.length + end.length
    ) {
      current = current.slice(start.length, -end.length).trim();
      break;
    }
  }
  return current;
}

function uniqueFormulas(values: string[]): string[] {
  const result: string[] = [];
  for (const value of values.map(stripOuterMath)) {
    if (value && value !== "[EMPTY]" && !result.includes(value)) {
      result.push(value);
    }
  }
  return result;
}

function parseFormulaText(text: string): string[] {
  const normalized = stripFences(text);
  try {
    const parsed = JSON.parse(normalized);
    if (isRecord(parsed) && Array.isArray(parsed.formulas)) {
      return uniqueFormulas(
        parsed.formulas.flatMap((item) =>
          typeof item === "string"
            ? [item]
            : isRecord(item) && typeof item.latex === "string"
              ? [item.latex]
              : [],
        ),
      );
    }
    if (isRecord(parsed) && typeof parsed.latex === "string") {
      return uniqueFormulas([parsed.latex]);
    }
  } catch {
    // Some compatible endpoints ignore structured output and return plain LaTeX.
  }
  const display = [
    ...normalized.matchAll(/\$\$([\s\S]*?)\$\$/g),
    ...normalized.matchAll(/\\\[([\s\S]*?)\\\]/g),
  ].map((match) => match[1]);
  return uniqueFormulas(display.length ? display : [normalized]);
}

function extractOpenAiText(value: unknown): string {
  if (!isRecord(value)) return "";
  if (typeof value.output_text === "string") return value.output_text;
  if (Array.isArray(value.output)) {
    for (const item of value.output) {
      if (!isRecord(item) || !Array.isArray(item.content)) continue;
      for (const content of item.content) {
        if (isRecord(content) && typeof content.text === "string") {
          return content.text;
        }
      }
    }
  }
  const choices = Array.isArray(value.choices) ? value.choices : [];
  const first = choices[0];
  if (isRecord(first) && isRecord(first.message)) {
    const content = first.message.content;
    if (typeof content === "string") return content;
    if (Array.isArray(content)) {
      const text = content
        .filter(isRecord)
        .map((part) => (typeof part.text === "string" ? part.text : ""))
        .join("");
      if (text) return text;
    }
  }
  return "";
}

const FORMULA_SCHEMA = {
  type: "object",
  additionalProperties: false,
  required: ["formulas"],
  properties: {
    formulas: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        required: ["latex"],
        properties: { latex: { type: "string" } },
      },
    },
  },
};

async function recognizeOpenAi(
  file: File,
  configuration: WebOcrConfiguration,
): Promise<{ model: string; formulas: string[] }> {
  const selected = configuration.openAiCompatible;
  const dataUrl = await fileToDataUrl(file);
  const route =
    selected.protocol === "responses" ? "responses" : "chat/completions";
  const endpoint = appendEndpoint(selected.baseUrl, route);
  const body =
    selected.protocol === "responses"
      ? {
          model: selected.model,
          input: [
            {
              role: "user",
              content: [
                { type: "input_text", text: selected.prompt },
                { type: "input_image", image_url: dataUrl },
              ],
            },
          ],
          text: {
            format: {
              type: "json_schema",
              name: "visualtex_formula_ocr",
              strict: true,
              schema: FORMULA_SCHEMA,
            },
          },
          max_output_tokens: 4096,
        }
      : {
          model: selected.model,
          messages: [
            {
              role: "user",
              content: [
                { type: "text", text: selected.prompt },
                { type: "image_url", image_url: { url: dataUrl } },
              ],
            },
          ],
          response_format: {
            type: "json_schema",
            json_schema: {
              name: "visualtex_formula_ocr",
              strict: true,
              schema: FORMULA_SCHEMA,
            },
          },
          temperature: 0,
        };
  const value = await responseJson(
    await fetchWithTimeout(
      endpoint,
      {
        method: "POST",
        headers: {
          "content-type": "application/json",
          authorization: `Bearer ${selected.apiKey}`,
        },
        body: JSON.stringify(body),
      },
      120_000,
    ),
    "OpenAI 兼容 OCR",
  );
  const text = extractOpenAiText(value);
  const formulas = parseFormulaText(text);
  if (!formulas.length) throw new Error("OCR API 没有返回可用公式");
  return { model: selected.model, formulas };
}

async function recognizeMathpix(
  file: File,
  configuration: WebOcrConfiguration,
): Promise<{ model: string; formulas: string[] }> {
  const dataUrl = await fileToDataUrl(file);
  if (new TextEncoder().encode(dataUrl).byteLength > 2 * 1024 * 1024) {
    throw new Error("Mathpix 的 base64 图片上限为 2 MB，请裁紧或压缩图片");
  }
  const endpoint = appendEndpoint(configuration.mathpix.baseUrl, "v3/text");
  const value = await responseJson(
    await fetchWithTimeout(
      endpoint,
      {
        method: "POST",
        headers: {
          "content-type": "application/json",
          app_id: configuration.mathpix.appId,
          app_key: configuration.mathpix.appKey,
        },
        body: JSON.stringify({
          src: dataUrl,
          formats: ["latex_styled", "text"],
          math_inline_delimiters: ["$", "$"],
          rm_spaces: true,
          metadata: { improve_mathpix: false },
        }),
      },
      120_000,
    ),
    "Mathpix OCR",
  );
  if (!isRecord(value)) throw new Error("Mathpix 返回格式不正确");
  if (typeof value.error === "string" && value.error.trim()) {
    throw new Error(`Mathpix OCR 失败：${value.error}`);
  }
  const text =
    typeof value.latex_styled === "string"
      ? value.latex_styled
      : typeof value.text === "string"
        ? value.text
        : "";
  const formulas = parseFormulaText(text);
  if (!formulas.length) throw new Error("Mathpix 没有返回可用公式");
  return { model: "Mathpix Text API", formulas };
}

async function recognizeSimpleTex(
  file: File,
  configuration: WebOcrConfiguration,
): Promise<{ model: string; formulas: string[] }> {
  const selected = configuration.simpleTex;
  const form = new FormData();
  form.append("file", file, file.name || "visualtex-formula.png");
  const value = await responseJson(
    await fetchWithTimeout(
      SIMPLETEX_ENDPOINTS[selected.model],
      {
        method: "POST",
        headers: { "x-visualtex-ocr-token": selected.accessToken },
        body: form,
      },
      45_000,
    ),
    "SimpleTex OCR",
  );
  if (!isRecord(value) || value.status !== true) {
    const message =
      isRecord(value) && typeof value.message === "string"
        ? value.message
        : "未知 API 错误";
    throw new Error(`SimpleTex OCR 失败：${message}`);
  }
  const res = isRecord(value.res) ? value.res : {};
  const latex = typeof res.latex === "string" ? res.latex : "";
  const formulas = parseFormulaText(latex);
  if (!formulas.length) throw new Error("SimpleTex 没有返回可用公式");
  return { model: `SimpleTex ${selected.model}`, formulas };
}

function collectPaddleFormulas(value: unknown, output: string[]): void {
  if (Array.isArray(value)) {
    value.forEach((item) => collectPaddleFormulas(item, output));
    return;
  }
  if (!isRecord(value)) return;
  const label =
    typeof value.block_label === "string"
      ? value.block_label.toLowerCase()
      : typeof value.blockLabel === "string"
        ? value.blockLabel.toLowerCase()
        : "";
  const content =
    typeof value.block_content === "string"
      ? value.block_content
      : typeof value.blockContent === "string"
        ? value.blockContent
        : "";
  if ((label.includes("formula") || label.includes("equation")) && content) {
    output.push(content);
    return;
  }
  const recognized =
    typeof value.rec_formula === "string"
      ? value.rec_formula
      : typeof value.recFormula === "string"
        ? value.recFormula
        : "";
  if (recognized) output.push(recognized);
  Object.values(value).forEach((child) => collectPaddleFormulas(child, output));
}

function extractMarkdownFormulas(markdown: string): string[] {
  const values = [
    ...markdown.matchAll(/\$\$([\s\S]*?)\$\$/g),
    ...markdown.matchAll(/\\\[([\s\S]*?)\\\]/g),
    ...markdown.matchAll(/\\\(([\s\S]*?)\\\)/g),
  ].map((match) => match[1]);
  if (values.length) return uniqueFormulas(values);
  return uniqueFormulas(
    [...markdown.matchAll(/(^|[^$])\$([^$\n]+)\$/g)].map(
      (match) => match[2],
    ),
  );
}

function parsePaddleResult(value: unknown): string[] {
  if (!isRecord(value)) return [];
  const result = isRecord(value.result) ? value.result : {};
  const pages = Array.isArray(result.layoutParsingResults)
    ? result.layoutParsingResults
    : [];
  const candidates: string[] = [];
  for (const page of pages) {
    if (!isRecord(page)) continue;
    collectPaddleFormulas(page.prunedResult, candidates);
  }
  if (!candidates.length) {
    for (const page of pages) {
      if (!isRecord(page) || !isRecord(page.markdown)) continue;
      if (typeof page.markdown.text === "string") {
        candidates.push(...extractMarkdownFormulas(page.markdown.text));
      }
    }
  }
  return uniqueFormulas(candidates);
}

function parsePaddleResultText(text: string): string[] {
  try {
    const formulas = parsePaddleResult(JSON.parse(text.trim()));
    if (formulas.length) return formulas;
  } catch {
    // Paddle can return newline-delimited page JSON instead of one JSON value.
  }
  const formulas: string[] = [];
  for (const line of text.split(/\r?\n/).map((item) => item.trim()).filter(Boolean)) {
    try {
      formulas.push(...parsePaddleResult(JSON.parse(line)));
    } catch {
      // Ignore non-JSON diagnostic lines in an NDJSON result.
    }
  }
  return uniqueFormulas(formulas);
}

async function recognizePaddle(
  file: File,
  configuration: WebOcrConfiguration,
  onProgress?: (progress: WebOcrProgress) => void,
): Promise<{ model: string; formulas: string[] }> {
  const form = new FormData();
  form.append("model", "PaddleOCR-VL-1.6");
  form.append(
    "optionalPayload",
    JSON.stringify({
      useDocOrientationClassify: false,
      useDocUnwarping: false,
      useLayoutDetection: true,
      useChartRecognition: false,
      showFormulaNumber: false,
      prettifyMarkdown: false,
    }),
  );
  form.append("file", file, file.name || "visualtex-formula.png");
  const accessToken = configuration.paddleOcr.accessToken
    .replace(/^Bearer\s+/i, "")
    .trim();
  const headers = {
    "x-visualtex-ocr-token": accessToken,
  };
  const submitted = await responseJson(
    await fetchWithTimeout(
      PADDLE_JOBS_URL,
      { method: "POST", headers, body: form },
      20_000,
    ),
    "PaddleOCR 任务提交",
  );
  if (!isRecord(submitted)) throw new Error("PaddleOCR 提交结果格式不正确");
  if (typeof submitted.code === "number" && submitted.code !== 0) {
    throw new Error(
      `PaddleOCR 任务提交失败：${String(submitted.message ?? submitted.code)}`,
    );
  }
  const data = isRecord(submitted.data) ? submitted.data : {};
  const jobId = typeof data.jobId === "string" ? data.jobId.trim() : "";
  if (!jobId) throw new Error("PaddleOCR 没有返回 jobId");

  onProgress?.({
    stage: "queued",
    messageZh: "图片已提交，正在等待 PaddleOCR 处理…",
    messageEn: "Image uploaded; waiting for PaddleOCR…",
  });
  const started = Date.now();
  while (Date.now() - started < 120_000) {
    await new Promise((resolve) => window.setTimeout(resolve, 1000));
    const statusValue = await responseJson(
      await fetchWithTimeout(
        `${PADDLE_JOBS_URL}/${encodeURIComponent(jobId)}`,
        { method: "GET", headers },
        8_000,
      ),
      "PaddleOCR 任务状态",
    );
    if (!isRecord(statusValue)) throw new Error("PaddleOCR 状态格式不正确");
    if (typeof statusValue.code === "number" && statusValue.code !== 0) {
      throw new Error(
        `PaddleOCR 状态请求失败：${String(statusValue.message ?? statusValue.code)}`,
      );
    }
    const statusData = isRecord(statusValue.data) ? statusValue.data : {};
    const state = typeof statusData.state === "string" ? statusData.state : "";
    if (state === "pending") continue;
    if (state === "running") {
      onProgress?.({
        stage: "recognizing",
        messageZh: "PaddleOCR 正在识别公式…",
        messageEn: "PaddleOCR is recognizing the formula…",
      });
      continue;
    }
    if (state === "failed") {
      throw new Error(
        `PaddleOCR 识别失败：${String(statusData.errorMsg ?? "未知错误")}`,
      );
    }
    if (state !== "done") {
      throw new Error(`PaddleOCR 返回未知任务状态：${state || "空"}`);
    }
    const resultText =
      typeof statusData.visualtexResultText === "string"
        ? statusData.visualtexResultText
        : "";
    if (!resultText.trim()) {
      throw new Error("PaddleOCR 完成后没有返回可读取的 JSON 结果");
    }
    onProgress?.({
      stage: "result",
      messageZh: "PaddleOCR 识别完成，正在读取结果…",
      messageEn: "PaddleOCR finished; reading the result…",
    });
    const formulas = parsePaddleResultText(resultText);
    if (!formulas.length) throw new Error("PaddleOCR 没有返回可用公式");
    return { model: "PaddleOCR-VL-1.6", formulas };
  }
  throw new Error("PaddleOCR 在 120 秒内没有完成");
}

export async function recognizeFormulaWithWebApi(
  file: File,
  configuration: WebOcrConfiguration,
  onProgress?: (progress: WebOcrProgress) => void,
): Promise<WebOcrResult> {
  if (!file.type.startsWith("image/"))
    throw new Error("请选择 PNG、JPEG、WebP、BMP 或 TIFF 图片");
  if (file.size <= 0) throw new Error("图片文件为空");
  if (file.size > 20 * 1024 * 1024) throw new Error("图片不能超过 20 MB");

  const normalized = validateConfiguration(configuration, true);
  onProgress?.({
    stage: "upload",
    messageZh: "正在将图片直接发送到所选 OCR 服务…",
    messageEn: "Sending the image directly to the selected OCR service…",
  });
  const started = performance.now();
  const recognized =
    normalized.activeProvider === "openai-compatible"
      ? await recognizeOpenAi(file, normalized)
      : normalized.activeProvider === "mathpix"
        ? await recognizeMathpix(file, normalized)
        : normalized.activeProvider === "paddleocr"
          ? await recognizePaddle(file, normalized, onProgress)
          : await recognizeSimpleTex(file, normalized);
  return {
    provider: normalized.activeProvider,
    model: recognized.model,
    elapsedMs: Math.round(performance.now() - started),
    formulas: recognized.formulas,
  };
}

export function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("无法读取图片"));
    reader.onload = () =>
      typeof reader.result === "string"
        ? resolve(reader.result)
        : reject(new Error("无法读取图片"));
    reader.readAsDataURL(file);
  });
}
