import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

const port = 17000 + (process.pid % 1000);
const debugPort = 22000 + (process.pid % 1000);
const baseUrl = `http://127.0.0.1:${port}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const profile = await mkdtemp(join(tmpdir(), "visualtex-physics-kernel-"));
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15_000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return response;
    } catch {
      // The local preview or Chrome is still starting.
    }
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

const preview = spawn(process.execPath, [
  "node_modules/vite/bin/vite.js", "preview", "--host", "127.0.0.1",
  "--port", String(port), "--strictPort",
], { cwd: process.cwd(), stdio: "ignore" });
let chrome;
let socket;

try {
  await waitFor(baseUrl);
  chrome = spawn(chromePath, [
    "--headless=new", "--disable-gpu", "--no-first-run",
    "--no-default-browser-check", `--remote-debugging-port=${debugPort}`,
    `--user-data-dir=${profile}`, baseUrl,
  ], { stdio: "ignore" });
  const targetResponse = await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
  const targets = await targetResponse.json();
  const target = targets.find((item) => item.type === "page" && item.url.startsWith(baseUrl));
  assert.ok(target, "No VisualTeX preview target was found");

  socket = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => {
    socket.addEventListener("open", resolve, { once: true });
    socket.addEventListener("error", reject, { once: true });
  });
  await sleep(500);
  const result = await new Promise((resolve, reject) => {
    socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      if (message.id !== 1) return;
      if (message.error || message.result?.exceptionDetails) {
        reject(new Error(message.error?.message ?? message.result.exceptionDetails.text));
      } else resolve(message.result.result.value);
    });
    socket.send(JSON.stringify({
      id: 1,
      method: "Runtime.evaluate",
      params: {
        awaitPromise: true,
        returnByValue: true,
        expression: `(async () => {
          await customElements.whenDefined("math-field");
          const samples = [
            "\\\\div", "\\\\divisionsymbol", "\\\\quantity{x}",
            "\\\\pqty{x}", "\\\\derivative{f}{x}",
            "\\\\partialderivative{f}{x}",
            "\\\\pmqty{1&0\\\\\\\\0&1}",
            "\\\\qq{if}", "A\\\\qq{if}B", "\\\\vqty{x}",
            "\\\\absolutevalue{x}", "\\\\vev{A}",
            "\\\\expectationvalue{A}{p}",
            "\\\\matrixelement{p}{A}{q}",
            "\\\\outerproduct{p}{q}",
            "\\\\flatfrac{a}{b}", "\\\\qty{x}"
          ];
          const fields = samples.map((latex) => {
            const field = document.createElement("math-field");
            document.body.append(field);
            field.setValue(latex, { mode: "math", format: "latex" });
            return { latex, field };
          });
          await new Promise((resolve) => setTimeout(resolve, 100));
          return fields.map(({ latex, field }) => {
            const value = field.value;
            let editedValue = "";
            if (latex.startsWith("\\\\pmqty")) {
              field.position = 2;
              field.insert("2", { mode: "math", format: "latex" });
              editedValue = field.value;
            }
            const result = {
              latex,
              value,
              editedValue,
              error: Boolean(field.shadowRoot?.querySelector(".ML__error")),
              array: Boolean(field._mathfield?.model?.atoms?.some(
                (atom) => atom.type === "array",
              )),
              renderedText: field.shadowRoot?.querySelector(".ML__base")?.textContent ?? "",
            };
            field.remove();
            return result;
          });
        })()`,
      },
    }));
  });

  const bySource = new Map(result.map((item) => [item.latex, item]));
  for (const item of result) {
    if (item.latex === String.raw`\qty{x}`) {
      assert.equal(item.error, true, "Conflicting siunitx \\qty was unexpectedly added");
      continue;
    }
    assert.equal(item.error, false, `${item.latex} failed in the built MathLive kernel`);
    assert.equal(item.value, item.latex, `${item.latex} did not preserve the source command`);
  }
  assert.match(bySource.get(String.raw`\div`)?.renderedText ?? "", /÷/, "\\div lost its division sign");
  assert.match(bySource.get(String.raw`\divisionsymbol`)?.renderedText ?? "", /÷/, "\\divisionsymbol is not division");
  for (const [source, contents] of [
    [String.raw`\vqty{x}`, /x/],
    [String.raw`\absolutevalue{x}`, /x/],
    [String.raw`\vev{A}`, /A/],
    [String.raw`\expectationvalue{A}{p}`, /A/],
    [String.raw`\matrixelement{p}{A}{q}`, /A/],
    [String.raw`\outerproduct{p}{q}`, /pq|p.*q/],
    [String.raw`\flatfrac{a}{b}`, /a\s*\/\s*b/],
    [String.raw`A\qq{if}B`, /AifB/],
  ]) {
    assert.match(
      bySource.get(source)?.renderedText ?? "",
      contents,
      `${source} lost an argument in the built MathLive kernel`,
    );
  }
  assert.equal(
    bySource.get(String.raw`\pmqty{1&0\\0&1}`)?.array,
    true,
    "\\pmqty was not parsed as a real tabular matrix",
  );
  assert.match(
    (bySource.get(String.raw`\pmqty{1&0\\0&1}`)?.editedValue ?? "").replace(/\s+/g, ""),
    /^\\pmqty\{[^&]*2[^&]*&0\\\\0&1\}$/,
    "Editing a matrix cell lost its column or row separators",
  );
  console.log("Built MathLive physics kernel browser regression passed.");
} finally {
  socket?.close();
  chrome?.kill();
  preview.kill();
  for (let attempt = 0; attempt < 5; attempt += 1) {
    try {
      await sleep(150);
      await rm(profile, { recursive: true, force: true });
      break;
    } catch (error) {
      if (attempt === 4) throw error;
    }
  }
}
