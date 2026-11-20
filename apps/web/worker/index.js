const SIMPLETEX_ENDPOINTS = {
  standard: "https://server.simpletex.cn/api/latex_ocr",
  turbo: "https://server.simpletex.cn/api/latex_ocr_turbo",
};
const PADDLE_JOBS_URL =
  "https://paddleocr.aistudio-app.com/api/v2/ocr/jobs";
const OPENAI_ENDPOINTS = {
  responses: "https://api.openai.com/v1/responses",
  "chat-completions": "https://api.openai.com/v1/chat/completions",
};

function securityHeaders(extra = {}) {
  return {
    "cache-control": "no-store, max-age=0",
    pragma: "no-cache",
    "x-content-type-options": "nosniff",
    ...extra,
  };
}

function jsonResponse(value, status = 200) {
  return new Response(JSON.stringify(value), {
    status,
    headers: securityHeaders({ "content-type": "application/json; charset=utf-8" }),
  });
}

function upstreamResponse(response, body) {
  return new Response(body, {
    status: response.status,
    headers: securityHeaders({
      "content-type":
        response.headers.get("content-type") ?? "application/json; charset=utf-8",
    }),
  });
}

function sameOriginRequest(request) {
  const origin = request.headers.get("origin");
  if (!origin) return request.headers.get("sec-fetch-site") === "same-origin";
  try {
    return new URL(origin).origin === new URL(request.url).origin;
  } catch {
    return false;
  }
}

function requestBodyIsWithinLimit(request, maximumBytes = 22 * 1024 * 1024) {
  const length = Number(request.headers.get("content-length") ?? "0");
  return !Number.isFinite(length) || length <= 0 || length <= maximumBytes;
}

async function readLimitedBody(response, maximumBytes) {
  const length = Number(response.headers.get("content-length") ?? "0");
  if (Number.isFinite(length) && length > maximumBytes) {
    throw new Error("OCR upstream response is too large");
  }
  const body = await response.arrayBuffer();
  if (body.byteLength > maximumBytes) {
    throw new Error("OCR upstream response is too large");
  }
  return body;
}

function rejectUpstreamRedirect(response, label) {
  if (response.status >= 300 && response.status < 400) {
    throw new Error(`${label} unexpectedly redirected`);
  }
}

function readRelayAccessToken(request) {
  const supplied = request.headers.get("x-visualtex-ocr-token")?.trim() ?? "";
  return supplied.replace(/^Bearer\s+/i, "").trim();
}

function paddleUnauthorizedResponse() {
  return jsonResponse(
    {
      error:
        "PaddleOCR Access Token 无效或已过期，请从 AI Studio 重新复制完整 Token",
    },
    401,
  );
}

async function proxySimpleTex(request, url) {
  if (request.method !== "POST") {
    return jsonResponse({ error: "Method not allowed" }, 405);
  }
  const model = url.searchParams.get("model");
  const endpoint = SIMPLETEX_ENDPOINTS[model];
  if (!endpoint) return jsonResponse({ error: "Unsupported SimpleTex model" }, 400);
  const token = request.headers.get("x-visualtex-ocr-token")?.trim();
  if (!token) return jsonResponse({ error: "SimpleTex token is required" }, 400);
  if (!requestBodyIsWithinLimit(request)) {
    return jsonResponse({ error: "OCR image upload is too large" }, 413);
  }

  const contentType = request.headers.get("content-type") ?? "";
  if (!contentType.toLowerCase().startsWith("multipart/form-data")) {
    return jsonResponse({ error: "Expected multipart/form-data" }, 415);
  }

  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      token,
      "content-type": contentType,
      accept: "application/json",
    },
    body: request.body,
    redirect: "manual",
  });
  rejectUpstreamRedirect(response, "SimpleTex");
  const body = await readLimitedBody(response, 2 * 1024 * 1024);
  return upstreamResponse(response, body);
}

async function proxyOpenAi(request, protocol) {
  if (request.method !== "POST") {
    return jsonResponse({ error: "Method not allowed" }, 405);
  }
  const endpoint = OPENAI_ENDPOINTS[protocol];
  if (!endpoint) {
    return jsonResponse({ error: "Unsupported OpenAI protocol" }, 400);
  }
  const accessToken = readRelayAccessToken(request);
  if (!accessToken) {
    return jsonResponse({ error: "OpenAI API Key is required" }, 400);
  }
  if (!requestBodyIsWithinLimit(request, 30 * 1024 * 1024)) {
    return jsonResponse({ error: "OpenAI OCR request is too large" }, 413);
  }
  const contentType = request.headers.get("content-type") ?? "";
  if (!contentType.toLowerCase().startsWith("application/json")) {
    return jsonResponse({ error: "Expected application/json" }, 415);
  }

  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      authorization: `Bearer ${accessToken}`,
      "content-type": "application/json",
      accept: "application/json",
    },
    body: request.body,
    redirect: "manual",
  });
  rejectUpstreamRedirect(response, "OpenAI");
  if (response.status === 401) {
    return jsonResponse(
      {
        error:
          "OpenAI API Key 无效或已过期，请检查当前账号和项目的密钥",
      },
      401,
    );
  }
  const body = await readLimitedBody(response, 4 * 1024 * 1024);
  return upstreamResponse(response, body);
}

async function proxyPaddleSubmit(request) {
  if (request.method !== "POST") {
    return jsonResponse({ error: "Method not allowed" }, 405);
  }
  const accessToken = readRelayAccessToken(request);
  if (!accessToken) {
    return jsonResponse({ error: "PaddleOCR access token is required" }, 400);
  }
  if (!requestBodyIsWithinLimit(request)) {
    return jsonResponse({ error: "OCR image upload is too large" }, 413);
  }
  const contentType = request.headers.get("content-type") ?? "";
  if (!contentType.toLowerCase().startsWith("multipart/form-data")) {
    return jsonResponse({ error: "Expected multipart/form-data" }, 415);
  }

  const response = await fetch(PADDLE_JOBS_URL, {
    method: "POST",
    headers: {
      authorization: `Bearer ${accessToken}`,
      "content-type": contentType,
      accept: "application/json",
    },
    body: request.body,
    redirect: "manual",
  });
  rejectUpstreamRedirect(response, "PaddleOCR task submission");
  if (response.status === 401) return paddleUnauthorizedResponse();
  const body = await readLimitedBody(response, 2 * 1024 * 1024);
  return upstreamResponse(response, body);
}

async function proxyPaddleStatus(request, jobId) {
  if (request.method !== "GET") {
    return jsonResponse({ error: "Method not allowed" }, 405);
  }
  if (!/^[A-Za-z0-9_-]{1,200}$/.test(jobId)) {
    return jsonResponse({ error: "Invalid PaddleOCR job id" }, 400);
  }
  const accessToken = readRelayAccessToken(request);
  if (!accessToken) {
    return jsonResponse({ error: "PaddleOCR access token is required" }, 400);
  }

  const response = await fetch(
    `${PADDLE_JOBS_URL}/${encodeURIComponent(jobId)}`,
    {
      method: "GET",
      headers: {
        authorization: `Bearer ${accessToken}`,
        accept: "application/json",
      },
      redirect: "manual",
    },
  );
  rejectUpstreamRedirect(response, "PaddleOCR task status");
  if (response.status === 401) return paddleUnauthorizedResponse();
  const statusBody = await readLimitedBody(response, 2 * 1024 * 1024);
  if (!response.ok) return upstreamResponse(response, statusBody);

  let statusValue;
  try {
    statusValue = JSON.parse(new TextDecoder().decode(statusBody));
  } catch {
    return upstreamResponse(response, statusBody);
  }

  if (statusValue?.data?.state !== "done") {
    return jsonResponse(statusValue, response.status);
  }
  const resultUrlValue = statusValue?.data?.resultUrl?.jsonUrl;
  if (typeof resultUrlValue !== "string" || !resultUrlValue.trim()) {
    return jsonResponse(
      { error: "PaddleOCR completed without a JSON result URL" },
      502,
    );
  }

  let resultUrl;
  try {
    resultUrl = new URL(resultUrlValue);
  } catch {
    return jsonResponse({ error: "PaddleOCR returned an invalid result URL" }, 502);
  }
  if (
    resultUrl.protocol !== "https:" ||
    resultUrl.username ||
    resultUrl.password
  ) {
    return jsonResponse({ error: "PaddleOCR returned an unsafe result URL" }, 502);
  }

  const resultResponse = await fetch(resultUrl, {
    method: "GET",
    headers: { accept: "application/json, application/x-ndjson, text/plain" },
    redirect: "follow",
  });
  const resultBody = await readLimitedBody(resultResponse, 16 * 1024 * 1024);
  if (!resultResponse.ok) return upstreamResponse(resultResponse, resultBody);

  statusValue.data.visualtexResultText = new TextDecoder().decode(resultBody);
  delete statusValue.data.resultUrl;
  return jsonResponse(statusValue);
}

async function handleApiRequest(request) {
  if (!sameOriginRequest(request)) {
    return jsonResponse({ error: "Cross-origin OCR relay requests are forbidden" }, 403);
  }
  const url = new URL(request.url);
  if (url.pathname === "/api/ocr/simpletex") {
    return proxySimpleTex(request, url);
  }
  if (url.pathname === "/api/ocr/openai/responses") {
    return proxyOpenAi(request, "responses");
  }
  if (url.pathname === "/api/ocr/openai/chat-completions") {
    return proxyOpenAi(request, "chat-completions");
  }
  if (url.pathname === "/api/ocr/paddle/jobs") {
    return proxyPaddleSubmit(request);
  }
  const paddleStatus = url.pathname.match(
    /^\/api\/ocr\/paddle\/jobs\/([A-Za-z0-9_-]{1,200})$/,
  );
  if (paddleStatus) return proxyPaddleStatus(request, paddleStatus[1]);
  return jsonResponse({ error: "Unknown API route" }, 404);
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname.startsWith("/api/")) {
      try {
        return await handleApiRequest(request);
      } catch (error) {
        const message =
          error instanceof Error ? error.message : "OCR relay request failed";
        return jsonResponse({ error: message }, 502);
      }
    }
    return env.ASSETS.fetch(request);
  },
};
