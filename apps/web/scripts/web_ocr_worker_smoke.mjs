import assert from "node:assert/strict";
import worker from "../worker/index.js";

const originalFetch = globalThis.fetch;

async function runRequest(path, init = {}) {
  const request = new Request(`https://visualtex.example${path}`, {
    ...init,
    headers: {
      origin: "https://visualtex.example",
      ...(init.headers ?? {}),
    },
  });
  return worker.fetch(request, {
    ASSETS: {
      fetch: () => new Response("asset"),
    },
  });
}

try {
  {
    const response = await worker.fetch(
      new Request("https://visualtex.example/api/ocr/simpletex?model=standard", {
        method: "POST",
        headers: {
          origin: "https://evil.example",
          "content-type": "multipart/form-data; boundary=test",
          "x-visualtex-ocr-token": "secret",
        },
        body: "--test--",
      }),
      { ASSETS: { fetch: () => new Response("asset") } },
    );
    assert.equal(response.status, 403);
  }

  {
    let forwardedUrl = "";
    let forwardedToken = "";
    let forwardedRedirect = "";
    globalThis.fetch = async (input, init) => {
      forwardedUrl = String(input);
      forwardedToken = new Headers(init?.headers).get("token") ?? "";
      forwardedRedirect = init?.redirect ?? "";
      return Response.json({
        status: true,
        res: { latex: "\\frac{1}{2}" },
      });
    };
    const response = await runRequest("/api/ocr/simpletex?model=standard", {
      method: "POST",
      headers: {
        "content-type": "multipart/form-data; boundary=test",
        "x-visualtex-ocr-token": "session-token",
      },
      body: "--test--",
    });
    assert.equal(response.status, 200);
    assert.equal(forwardedUrl, "https://server.simpletex.cn/api/latex_ocr");
    assert.equal(forwardedToken, "session-token");
    assert.equal(forwardedRedirect, "manual");
    assert.equal(response.headers.get("cache-control"), "no-store, max-age=0");
  }

  {
    const upstreamCalls = [];
    globalThis.fetch = async (input, init) => {
      upstreamCalls.push({
        url: String(input),
        authorization: new Headers(init?.headers).get("authorization"),
        redirect: init?.redirect,
      });
      return Response.json({
        output_text: '{"formulas":[{"latex":"\\\\sqrt{x}"}]}',
      });
    };
    for (const route of ["responses", "chat-completions"]) {
      const response = await runRequest(`/api/ocr/openai/${route}`, {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-visualtex-ocr-token": "Bearer openai-session-token",
        },
        body: "{}",
      });
      assert.equal(response.status, 200);
    }
    assert.deepEqual(upstreamCalls, [
      {
        url: "https://api.openai.com/v1/responses",
        authorization: "Bearer openai-session-token",
        redirect: "manual",
      },
      {
        url: "https://api.openai.com/v1/chat/completions",
        authorization: "Bearer openai-session-token",
        redirect: "manual",
      },
    ]);
  }

  {
    let forwarded;
    globalThis.fetch = async (input, init) => {
      forwarded = {
        url: String(input),
        authorization: new Headers(init?.headers).get("authorization"),
        redirect: init?.redirect,
      };
      return Response.json({
        code: 0,
        data: { jobId: "job_123" },
      });
    };
    const response = await runRequest("/api/ocr/paddle/jobs", {
      method: "POST",
      headers: {
        "x-visualtex-ocr-token": "Bearer paddle-session-token",
        "content-type": "multipart/form-data; boundary=test",
      },
      body: "--test--",
    });
    assert.equal(response.status, 200);
    assert.deepEqual(forwarded, {
      url: "https://paddleocr.aistudio-app.com/api/v2/ocr/jobs",
      authorization: "Bearer paddle-session-token",
      redirect: "manual",
    });
  }

  {
    globalThis.fetch = async () =>
      new Response(null, {
        status: 302,
        headers: { location: "https://redirect.example/jobs" },
      });
    const response = await runRequest("/api/ocr/paddle/jobs", {
      method: "POST",
      headers: {
        "x-visualtex-ocr-token": "paddle-session-token",
        "content-type": "multipart/form-data; boundary=test",
      },
      body: "--test--",
    });
    assert.equal(response.status, 502);
    assert.match(
      String((await response.json()).error),
      /task submission unexpectedly redirected/,
    );
  }

  {
    globalThis.fetch = async () => new Response(null, { status: 401 });
    const response = await runRequest("/api/ocr/paddle/jobs", {
      method: "POST",
      headers: {
        "x-visualtex-ocr-token": "expired-token",
        "content-type": "multipart/form-data; boundary=test",
      },
      body: "--test--",
    });
    assert.equal(response.status, 401);
    assert.match(
      String((await response.json()).error),
      /Access Token 无效或已过期/,
    );
  }

  {
    const upstreamCalls = [];
    globalThis.fetch = async (input, init) => {
      upstreamCalls.push({
        url: String(input),
        authorization: new Headers(init?.headers).get("authorization"),
        redirect: init?.redirect,
      });
      if (upstreamCalls.length === 1) {
        return Response.json({
          code: 0,
          data: {
            state: "done",
            resultUrl: { jsonUrl: "https://result.example/paddle.json" },
          },
        });
      }
      return Response.json({
        result: {
          layoutParsingResults: [
            {
              prunedResult: [
                {
                  block_label: "formula",
                  block_content: "\\sqrt{x}",
                },
              ],
            },
          ],
        },
      });
    };
    const response = await runRequest("/api/ocr/paddle/jobs/job_123", {
      method: "GET",
      headers: { "x-visualtex-ocr-token": "paddle-session-token" },
    });
    assert.equal(response.status, 200);
    const value = await response.json();
    assert.equal(typeof value.data.visualtexResultText, "string");
    assert.equal("resultUrl" in value.data, false);
    assert.deepEqual(upstreamCalls, [
      {
        url: "https://paddleocr.aistudio-app.com/api/v2/ocr/jobs/job_123",
        authorization: "Bearer paddle-session-token",
        redirect: "manual",
      },
      {
        url: "https://result.example/paddle.json",
        authorization: null,
        redirect: "follow",
      },
    ]);
  }

  {
    const response = await runRequest("/api/ocr/not-allowed", {
      method: "POST",
    });
    assert.equal(response.status, 404);
  }

  console.log("Web OCR relay smoke test passed");
} finally {
  globalThis.fetch = originalFetch;
}
