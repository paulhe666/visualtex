import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const portOffset = process.pid % 900;
const vitePort = 7200 + portOffset;
const debugPort = 13200 + portOffset;
const baseUrl = `http://127.0.0.1:${vitePort}`;
const chromeProfile = `/tmp/visualtex-toolbar-omml-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";

const addedCommandIds = [
  "left-parenthesis-only",
  "right-parenthesis-only",
  "left-bracket-only",
  "right-bracket-only",
  "left-brace-only",
  "right-brace-only",
  "mixed-paren-bracket",
  "mixed-bracket-paren",
  "double-brackets",
  "moustache-pair",
  "corner-pair",
  "baraccent",
  "vecaccent",
  "dddotaccent",
  "ddddotaccent",
  "cases-three",
  "bigcap-limits",
  "bigcup-limits",
  "bigsqcup-limits",
  "bigvee-limits",
  "bigwedge-limits",
  "bigodot-limits",
  "bigoplus-limits",
  "bigotimes-limits",
  "biguplus-limits",
  "matrix-dots",
  "diagonal-matrix",
  "augmented-matrix",
  "block-matrix",
  "plusminus",
  "minusplus",
  "multiply",
  "division",
  "centerdot",
  "asterisk",
  "starop",
  "circleop",
  "bulletop",
  "diamondop",
  "triangleop",
  "triangledownop",
  "bigtriangleupop",
  "bigtriangledownop",
  "ominus",
  "oslash",
  "odot",
  "circledast",
  "circledcirc",
  "circleddash",
  "boxplus",
  "boxminus",
  "boxtimes",
  "boxdot",
  "uplus",
  "sqcap",
  "sqcup",
  "amalg",
  "wreath",
  "divideontimes",
  "ltimes",
  "rtimes",
  "leftthreetimes",
  "rightthreetimes",
  "intercal",
  "approxeq",
  "eqsim",
  "backsim",
  "backsimeq",
  "fallingdotseq",
  "risingdotseq",
  "eqcirc",
  "circeq",
  "bumpeq",
  "Bumpeq",
  "between",
  "pitchfork",
  "therefore",
  "because",
  "sqsubset",
  "sqsupset",
  "sqsubseteq",
  "sqsupseteq",
  "Subset",
  "Supset",
  "subseteqq",
  "supseteqq",
  "nsubseteq",
  "nsupseteq",
  "Uparrow",
  "Downarrow",
  "Updownarrow",
  "upuparrows",
  "downdownarrows",
  "leftleftarrows",
  "rightrightarrows",
  "leftrightarrows",
  "Lleftarrow",
  "Rrightarrow",
  "leftarrowtail",
  "rightarrowtail",
  "looparrowleft",
  "looparrowright",
  "curvearrowleft",
  "curvearrowright",
  "leftrightsquigarrow",
  "dashleftarrow",
  "dashrightarrow",
  "Lsh",
  "Rsh",
  "multimap",
  "nleftarrow",
  "nrightarrow",
  "nleftrightarrow",
  "nLeftarrow",
  "nRightarrow",
  "nLeftrightarrow",
  "plus",
  "minus",
  "colonop",
  "ldots",
  "cdots",
  "vdots",
  "ddots",
  "nless",
  "ngtr",
  "nleq",
  "ngeq",
  "leqslant",
  "geqslant",
  "lessdot",
  "gtrdot",
  "lessgtr",
  "gtrless",
  "lesseqgtr",
  "gtreqless",
  "triangleleft",
  "triangleright",
  "trianglelefteq",
  "trianglerighteq",
  "vartriangleleft",
  "vartriangleright",
  "ntriangleleft",
  "ntriangleright",
  "ntrianglelefteq",
  "ntrianglerighteq",
  "square-symbol",
  "blacksquare-symbol",
  "lozenge-symbol",
  "blacklozenge-symbol",
  "bigcirc-symbol",
  "blacktriangle-symbol",
  "blacktriangledown-symbol",
  "blacktriangleleft-symbol",
  "blacktriangleright-symbol",
  "ni",
  "notni",
  "smallsetminus",
  "upharpoonleft",
  "upharpoonright",
  "downharpoonleft",
  "downharpoonright",
  "mapsfrom",
  "longmapsfrom",
  "leadsto",
  "bm-bold-symbol",
  "math-bold-italic",
];

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

    const expression = `
      (async () => {
        const registryModule = await import(${JSON.stringify(`${baseUrl}/src/autocomplete/commandRegistry.ts`)});
        const ommlModule = await import(${JSON.stringify(`${baseUrl}/src/office/omml/latexToOmml.ts`)});
        const mathfieldHarness = await import(${JSON.stringify(`${baseUrl}/scripts/mathfieldNormalizationHarness.ts`)});
        const ids = ${JSON.stringify(addedCommandIds)};
        const results = [];
        for (const id of ids) {
          const command = registryModule.commandRegistry.find((item) => item.id === id);
          if (!command) {
            results.push({ id, ok: false, error: 'missing command' });
            continue;
          }
          let insertionLatex = command.insertTemplate;
          try {
            insertionLatex = command.insertTemplate.replace(/\\\\placeholder\{\}/g, '{x}');
            const normalizedLatex =
              mathfieldHarness.normalizeMathfieldLatexForRegression(insertionLatex);
            const previewOmml = ommlModule.latexLinesToOmml([command.previewLatex], 'block');
            const insertedOmml = ommlModule.latexLinesToOmml([normalizedLatex], 'block');
            results.push({
              id,
              ok:
                typeof previewOmml === 'string' &&
                previewOmml.includes('<m:oMath') &&
                !previewOmml.includes('merror') &&
                typeof insertedOmml === 'string' &&
                insertedOmml.includes('<m:oMath') &&
                !insertedOmml.includes('merror'),
              previewLatex: command.previewLatex,
              insertionLatex,
              normalizedLatex,
              previewOmmlLength: previewOmml.length,
              insertedOmmlLength: insertedOmml.length,
            });
          } catch (error) {
            results.push({
              id,
              ok: false,
              previewLatex: command.previewLatex,
              insertionLatex,
              normalizedLatex: '',
              error: error instanceof Error ? error.message : String(error),
            });
          }
        }
        return results;
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
          "Toolbar OMML evaluation failed.",
      );
    }

    const results = evaluation.result?.value;
    if (!Array.isArray(results)) {
      throw new Error("Toolbar OMML regression did not return an array.");
    }
    const failures = results.filter((result) => !result.ok);
    if (failures.length > 0) {
      throw new Error(
        `New toolbar commands failed Word OMML conversion:\n${JSON.stringify(failures, null, 2)}`,
      );
    }

    console.log(
      `Toolbar command Word OMML regression passed (${results.length} commands)`,
    );
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    vite.kill("SIGTERM");
    await sleep(600);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => {});
  }
}

await main();
