import { execFileSync, spawn } from "node:child_process";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { join } from "node:path";
import process from "node:process";

const portOffset = process.pid % 1000;
const vitePort = 6400 + portOffset;
const debugPort = 11400 + portOffset;
const baseUrl = `http://127.0.0.1:${vitePort}`;
const chromeProfile = `/tmp/visualtex-omml-smoke-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";

const sleep = (milliseconds) =>
  new Promise((resolve) => setTimeout(resolve, milliseconds));

async function waitFor(url, timeoutMs = 20_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return response;
    } catch {
      // Retry while the local process starts.
    }
    await sleep(100);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

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

function expect(condition, message) {
  if (!condition) throw new Error(message);
}

function expectIncludes(value, fragment, message) {
  expect(value.includes(fragment), `${message}\nMissing: ${fragment}\nActual: ${value}`);
}

async function main() {
  const vite = spawn(
    process.execPath,
    [
      "node_modules/vite/bin/vite.js",
      "--host",
      "127.0.0.1",
      "--port",
      String(vitePort),
      "--strictPort",
    ],
    { cwd: process.cwd(), stdio: "ignore" },
  );
  let chrome;
  let client;
  let docxDirectory;

  try {
    await waitFor(baseUrl);
    chrome = spawn(
      chromePath,
      [
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`,
        `--user-data-dir=${chromeProfile}`,
        baseUrl,
      ],
      { stdio: "ignore" },
    );

    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const pages = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const page = pages.find(
      (item) => item.type === "page" && item.url?.startsWith(baseUrl),
    );
    if (!page?.webSocketDebuggerUrl) {
      throw new Error("Chrome did not expose a debuggable page.");
    }

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");

    const formulas = [
      String.raw`\frac{a}{b}+dddc`,
      String.raw`\int_a^b x\,dy`,
      String.raw`\sum_{b}^{a}xc`,
      String.raw`\sqrt{x}`,
      String.raw`x_i^2`,
      String.raw`\begin{pmatrix}a&b\\c&d\end{pmatrix}`,
    ];
    const expression = `
      (async () => {
        const module = await import(${JSON.stringify(`${baseUrl}/src/office/omml/latexToOmml.ts`)});
        return ${JSON.stringify(formulas)}.map((latex) =>
          module.latexLinesToOmml([latex], 'block')
        );
      })()
    `;
    const evaluation = await client.send("Runtime.evaluate", {
      expression,
      awaitPromise: true,
      returnByValue: true,
    });
    if (evaluation.exceptionDetails) {
      throw new Error(
        evaluation.exceptionDetails.exception?.description ??
          evaluation.exceptionDetails.text ??
          "OMML browser evaluation failed.",
      );
    }
    const results = evaluation.result?.value;
    expect(Array.isArray(results) && results.length === formulas.length, "OMML regression did not return every formula.");

    const [fraction, integral, sum, root, scripts, matrix] = results;
    expectIncludes(fraction, "<m:f>", "Fraction must use a structural OMML fraction node.");
    expectIncludes(fraction, "<m:num>", "Fraction must preserve its numerator.");
    expectIncludes(fraction, "<m:den>", "Fraction must preserve its denominator.");
    expect(!fraction.includes("<m:t>/</m:t>"), "Fraction must not degrade to a slash glyph.");
    expect(fraction.indexOf("</m:f>") < fraction.lastIndexOf("<m:t>"), "Fraction must remain before the trailing expression.");

    expectIncludes(integral, '<m:chr m:val="∫"/>', "Integral must use an OMML n-ary operator.");
    expectIncludes(integral, "<m:sub>", "Integral must preserve the lower limit.");
    expectIncludes(integral, "<m:sup>", "Integral must preserve the upper limit.");
    expectIncludes(integral, "<m:e>", "Integral must preserve the integrand as its body.");

    expectIncludes(sum, '<m:chr m:val="∑"/>', "Summation must use an OMML n-ary operator.");
    expectIncludes(sum, '<m:limLoc m:val="undOvr"/>', "Summation limits must use above-and-below placement.");

    expectIncludes(root, "<m:rad>", "Square root must use an OMML radical node.");
    expectIncludes(scripts, "<m:sSubSup>", "Combined scripts must use an OMML subscript/superscript node.");
    expectIncludes(matrix, "<m:m>", "Matrix must use an OMML matrix node.");
    expect((matrix.match(/<m:mr>/g) ?? []).length === 2, "Matrix must preserve both rows.");

    const docxExpression = `
      (async () => {
        const module = await import(${JSON.stringify(`${baseUrl}/src/office/omml/latexToOmml.ts`)});
        const artifacts = module.latexLinesToOmmlArtifacts([String.raw\`\\frac{a}{b}\`], 'inline');
        return {
          ommlBase64: artifacts.ommlBase64,
          ommlDocxBase64: artifacts.ommlDocxBase64,
        };
      })()
    `;
    const docxEvaluation = await client.send("Runtime.evaluate", {
      expression: docxExpression,
      awaitPromise: true,
      returnByValue: true,
    });
    if (docxEvaluation.exceptionDetails) {
      throw new Error(docxEvaluation.exceptionDetails.exception?.description ?? "DOCX generation failed.");
    }
    const artifacts = docxEvaluation.result?.value;
    const docxBase64 = artifacts?.ommlDocxBase64;
    const ommlBase64 = artifacts?.ommlBase64;
    expect(typeof docxBase64 === "string" && docxBase64.length > 100, "OMML DOCX export is missing.");
    expect(typeof ommlBase64 === "string" && ommlBase64.length > 100, "OMML Base64URL export is missing.");
    docxDirectory = await mkdtemp("/tmp/visualtex-omml-docx-");
    const docxPath = join(docxDirectory, "fraction.docx");
    await writeFile(docxPath, Buffer.from(docxBase64, "base64url"));
    execFileSync("/usr/bin/unzip", ["-tqq", docxPath]);
    const documentXml = execFileSync(
      "/usr/bin/unzip",
      ["-p", docxPath, "word/document.xml"],
      { encoding: "utf8" },
    );
    expectIncludes(documentXml, "<m:f>", "Generated DOCX must contain the structural fraction.");

    const persistentRoot = process.env.VISUALTEX_WORD_REGRESSION_ROOT;
    if (persistentRoot) {
      const nativeDocuments = join(persistentRoot, "NativeDocuments");
      const tests = join(persistentRoot, "Tests");
      await mkdir(nativeDocuments, { recursive: true, mode: 0o700 });
      await mkdir(tests, { recursive: true, mode: 0o700 });
      await writeFile(
        join(nativeDocuments, "11111111-1111-4111-8111-111111111111.docx"),
        Buffer.from(docxBase64, "base64url"),
        { mode: 0o600 },
      );
      await writeFile(
        join(tests, "word-native-regression-omml.txt"),
        ommlBase64,
        { encoding: "utf8", mode: 0o600 },
      );
      console.log(`Word native regression fixture written to ${persistentRoot}.`);
    }

    console.log("Word structural OMML regression passed.");
  } finally {
    client?.close();
    if (chrome && chrome.exitCode === null) {
      const exited = new Promise((resolve) => chrome.once("exit", resolve));
      chrome.kill("SIGTERM");
      await Promise.race([exited, sleep(2_000)]);
    }
    vite.kill("SIGTERM");
    await rm(chromeProfile, { recursive: true, force: true });
    if (docxDirectory) await rm(docxDirectory, { recursive: true, force: true });
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.stack : error);
  process.exitCode = 1;
});
