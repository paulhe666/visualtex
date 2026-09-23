import { randomUUID } from "node:crypto";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";

process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

const base = "https://127.0.0.1:43127";
const installPath = resolve(process.env.APPDATA ?? "", "com.visualtex.studio/office/install.json");
const install = JSON.parse(await readFile(installPath, "utf8"));
if (typeof install.installToken !== "string" || install.installToken.length !== 64) {
  throw new Error("Office install token is missing or invalid.");
}
const headers = {
  "content-type": "application/json",
  "X-VisualTeX-Install-Token": install.installToken,
};
const sleep = (ms) => new Promise((resolveSleep) => setTimeout(resolveSleep, ms));

async function raw(path, options = {}) {
  try {
    const response = await fetch(`${base}${path}`, options);
    const text = await response.text();
    let body = text;
    if (text) {
      try { body = JSON.parse(text); } catch {}
    }
    return { ok: response.ok, status: response.status, body };
  } catch (error) {
    return { ok: false, status: 0, error: String(error?.stack ?? error) };
  }
}

async function request(path, options = {}) {
  const result = await raw(path, options);
  if (!result.ok) throw new Error(`${options.method ?? "GET"} ${path} -> ${result.status}: ${JSON.stringify(result.body ?? result.error)}`);
  return result.body;
}

const healthBefore = await raw("/health");
if (!healthBefore.ok || healthBefore.body?.ok !== true) {
  throw new Error("Existing companion must be healthy before the create-race repro.");
}

async function createSession(index) {
  const lineId = randomUUID();
  return request("/api/v1/sessions", {
    method: "POST",
    headers,
    body: JSON.stringify({
      mode: "create",
      host: "word",
      formulaId: randomUUID(),
      sourceDocumentId: null,
      sourceObjectId: null,
      title: `P13 window-create race ${index}`,
      lines: [{ id: lineId, latex: index === 0 ? "x^2+y^2" : index === 1 ? "\\frac{a+b}{c+d}" : "\\sqrt{x_1^2+x_2^2}" }],
      activeLineId: lineId,
      codeFormat: "latex",
      displayMode: index % 2 ? "block" : "inline",
      objectMode: index === 1 ? "wordOmml" : "nativeOle",
      numbered: false,
      mathTypeNumberPosition: "right",
      fontSizePt: 10.5,
      originalMetadata: null,
      autoCommitOnClose: false,
    }),
  });
}

const sessions = await Promise.all([0, 1, 2].map(createSession));
const ids = sessions.map((session) => session.id);
if (ids.some((id) => typeof id !== "string" || !id)) {
  throw new Error("Session creation returned an invalid id.");
}

const actions = [
  ["prewarm", "/api/v1/app/converter/prewarm", {}],
  ["single", `/api/v1/app/sessions/${encodeURIComponent(ids[0])}/convert`, {}],
  ["batch", "/api/v1/app/converter/convert-batch", { sessionIds: ids.slice(1) }],
];
const startedAt = Date.now();
const results = await Promise.all(actions.map(async ([name, path, body]) => ({
  name,
  ...(await raw(path, {
    method: "POST",
    headers,
    body: JSON.stringify(body),
  })),
})));

await sleep(1500);
const healthAfter = await raw("/health");
const states = [];
if (healthAfter.ok) {
  for (const id of ids) {
    states.push({ id, ...(await raw(`/api/v1/sessions/${encodeURIComponent(id)}`, { headers })) });
  }
}

console.log(JSON.stringify({
  status: healthAfter.ok && healthAfter.body?.ok === true ? "COMPANION_ALIVE" : "COMPANION_UNHEALTHY",
  startedAt,
  elapsedMs: Date.now() - startedAt,
  ids,
  results,
  healthAfter,
  states: states.map(({ id, ok, status, body, error }) => ({
    id,
    ok,
    httpStatus: status,
    sessionStatus: body?.status ?? null,
    error: body?.error ?? error ?? null,
  })),
}, null, 2));
