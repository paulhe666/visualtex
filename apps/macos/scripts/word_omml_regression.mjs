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
      String.raw`\differentialD x+\capitalDifferentialD y+\exponentialE^{\imaginaryI x}+\imaginaryJ`,
      String.raw`\mathrm{x}`,
      String.raw`\mathbf{A+1}`,
      String.raw`\mathit{x}`,
      String.raw`\boldsymbol{\alpha}`,
      String.raw`\mathbb{R}`,
      String.raw`\mathcal{G}`,
      String.raw`\mathscr{g}`,
      String.raw`\mathfrak{g}`,
      String.raw`\mathsf{x}`,
      String.raw`\mathtt{x}`,
      String.raw`\operatorname{sin}x`,
      String.raw`\sum_{b}^{a}xc`,
      String.raw`\sqrt{x}`,
      String.raw`x_i^2`,
      String.raw`\begin{pmatrix}a&b\\c&d\end{pmatrix}`,
      String.raw`\lim_{x\to0}f(x)`,
      String.raw`\binom{n}{k}`,
      String.raw`\begin{cases}x,&x>0\\-x,&x<0\end{cases}`,
      String.raw`\overline{x}+\hat{y}`,
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

    const [
      fraction,
      integral,
      uprightSymbols,
      roman,
      bold,
      italic,
      boldItalic,
      doubleStruck,
      calligraphic,
      scriptVariant,
      fraktur,
      sansSerif,
      monospace,
      operatorName,
      sum,
      root,
      scripts,
      matrix,
      limit,
      binomial,
      cases,
      accents,
    ] = results;
    expectIncludes(fraction, "<m:f>", "Fraction must use a structural OMML fraction node.");
    expectIncludes(fraction, "<m:num>", "Fraction must preserve its numerator.");
    expectIncludes(fraction, "<m:den>", "Fraction must preserve its denominator.");
    expect(!fraction.includes("<m:t>/</m:t>"), "Fraction must not degrade to a slash glyph.");
    expect(fraction.indexOf("</m:f>") < fraction.lastIndexOf("<m:t>"), "Fraction must remain before the trailing expression.");

    expectIncludes(integral, '<m:chr m:val="∫"/>', "Integral must use an OMML n-ary operator.");
    expectIncludes(integral, "<m:sub>", "Integral must preserve the lower limit.");
    expectIncludes(integral, "<m:sup>", "Integral must preserve the upper limit.");
    expectIncludes(integral, "<m:e>", "Integral must preserve the integrand as its body.");

    expect(!uprightSymbols.includes("differentialD"), "MathLive differentialD must not leak into OMML.");
    expect(!uprightSymbols.includes("capitalDifferentialD"), "MathLive capitalDifferentialD must not leak into OMML.");
    expect(!uprightSymbols.includes("exponentialE"), "MathLive exponentialE must not leak into OMML.");
    expect(!uprightSymbols.includes("imaginaryI"), "MathLive imaginaryI must not leak into OMML.");
    expect(!uprightSymbols.includes("imaginaryJ"), "MathLive imaginaryJ must not leak into OMML.");
    expect((uprightSymbols.match(/<m:nor\/>/g) ?? []).length >= 5, "Every canonical upright symbol must use explicit OMML normal style.");
    for (const symbol of ["d", "D", "e", "i", "j"]) {
      expectIncludes(uprightSymbols, `<m:t>${symbol}</m:t>`, `Upright ${symbol} must survive OMML conversion.`);
    }

    expectIncludes(roman, "<m:nor/>", "mathrm must retain Word's explicit upright marker.");
    expectIncludes(roman, '<m:scr m:val="roman"/>', "mathrm must use the OMML roman script.");
    expectIncludes(roman, '<m:sty m:val="p"/>', "mathrm must use plain OMML style.");

    expect((bold.match(/<m:sty m:val="b"\/>/g) ?? []).length >= 3, "mathbf must preserve bold style on identifiers, operators and numbers.");
    expectIncludes(bold, '<m:scr m:val="roman"/>', "mathbf must use the OMML roman script.");
    expectIncludes(italic, '<m:sty m:val="i"/>', "mathit must use italic OMML style.");
    expectIncludes(boldItalic, '<m:sty m:val="bi"/>', "boldsymbol must use bold-italic OMML style.");
    expectIncludes(boldItalic, "<m:t>α</m:t>", "boldsymbol must preserve Greek characters.");

    expectIncludes(doubleStruck, '<m:scr m:val="double-struck"/>', "mathbb must use the OMML double-struck script.");
    expectIncludes(calligraphic, '<m:scr m:val="script"/>', "mathcal must use the OMML script alphabet.");
    expectIncludes(scriptVariant, '<m:scr m:val="script"/>', "mathscr must use the OMML script alphabet.");
    expectIncludes(fraktur, '<m:scr m:val="fraktur"/>', "mathfrak must use the OMML fraktur alphabet.");
    expectIncludes(sansSerif, '<m:scr m:val="sans-serif"/>', "mathsf must use the OMML sans-serif alphabet.");
    expectIncludes(monospace, '<m:scr m:val="monospace"/>', "mathtt must use the OMML monospace alphabet.");

    expect(
      /<m:r><m:rPr><m:scr m:val="roman"\/><m:sty m:val="p"\/><\/m:rPr><m:t>sin<\/m:t><\/m:r>/.test(operatorName),
      "Multi-character operator names must default to upright roman OMML runs.",
    );

    expectIncludes(sum, '<m:chr m:val="∑"/>', "Summation must use an OMML n-ary operator.");
    expectIncludes(sum, '<m:limLoc m:val="undOvr"/>', "Summation limits must use above-and-below placement.");

    expectIncludes(root, "<m:rad>", "Square root must use an OMML radical node.");
    expectIncludes(scripts, "<m:sSubSup>", "Combined scripts must use an OMML subscript/superscript node.");
    expectIncludes(matrix, "<m:m>", "Matrix must use an OMML matrix node.");
    expect((matrix.match(/<m:mr>/g) ?? []).length === 2, "Matrix must preserve both rows.");

    expectIncludes(limit, "<m:limLow>", "Limit notation must use an OMML lower-limit node.");
    expectIncludes(binomial, '<m:type m:val="noBar"/>', "Binomial coefficients must use a no-bar fraction.");
    expectIncludes(binomial, "<m:d>", "Binomial coefficients must retain their scalable parentheses.");
    expectIncludes(cases, '<m:begChr m:val="{"/>', "Cases must retain their opening brace.");
    expectIncludes(cases, "<m:m>", "Cases must retain their row and column structure.");
    expectIncludes(accents, "<m:bar>", "Overline must use an OMML bar node.");
    expectIncludes(accents, "<m:acc>", "Hat accents must use an OMML accent node.");

    const multiLineExpression = `
      (async () => {
        const module = await import(${JSON.stringify(`${baseUrl}/src/office/omml/latexToOmml.ts`)});
        return module.latexLinesToOmml(['a=b', 'c=d'], 'block');
      })()
    `;
    const multiLineEvaluation = await client.send("Runtime.evaluate", {
      expression: multiLineExpression,
      awaitPromise: true,
      returnByValue: true,
    });
    if (multiLineEvaluation.exceptionDetails) {
      throw new Error(multiLineEvaluation.exceptionDetails.exception?.description ?? "Multi-line OMML generation failed.");
    }
    const multiLine = multiLineEvaluation.result?.value;
    expectIncludes(multiLine, "<m:eqArr>", "Multiple editor lines must use one OMML equation array.");
    expect((multiLine.match(/<m:e>/g) ?? []).length >= 2, "Equation arrays must preserve every editor line.");

    const docxExpression = `
      (async () => {
        const module = await import(${JSON.stringify(`${baseUrl}/src/office/omml/latexToOmml.ts`)});
        const artifacts = module.latexLinesToOmmlArtifacts([String.raw\`\\frac{a}{b}\`], 'inline');
        const fontArtifacts = module.latexLinesToOmmlArtifacts([
          String.raw\`\\mathrm{x}+\\mathbf{A+1}+\\mathit{x}+\\boldsymbol{\\alpha}+\\mathbb{R}+\\mathcal{G}+\\mathscr{g}+\\mathfrak{g}+\\mathsf{x}+\\mathtt{x}\`,
        ], 'inline');
        return {
          ommlBase64: artifacts.ommlBase64,
          ommlDocxBase64: artifacts.ommlDocxBase64,
          fontOmmlDocxBase64: fontArtifacts.ommlDocxBase64,
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
    const fontDocxBase64 = artifacts?.fontOmmlDocxBase64;
    expect(typeof docxBase64 === "string" && docxBase64.length > 100, "OMML DOCX export is missing.");
    expect(typeof ommlBase64 === "string" && ommlBase64.length > 100, "OMML Base64URL export is missing.");
    expect(typeof fontDocxBase64 === "string" && fontDocxBase64.length > 100, "Font-variant OMML DOCX export is missing.");
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

    const fontDocxPath = join(docxDirectory, "font-variants.docx");
    await writeFile(fontDocxPath, Buffer.from(fontDocxBase64, "base64url"));
    execFileSync("/usr/bin/unzip", ["-tqq", fontDocxPath]);
    const fontDocumentXml = execFileSync(
      "/usr/bin/unzip",
      ["-p", fontDocxPath, "word/document.xml"],
      { encoding: "utf8" },
    );
    for (const script of [
      "roman",
      "double-struck",
      "script",
      "fraktur",
      "sans-serif",
      "monospace",
    ]) {
      expectIncludes(
        fontDocumentXml,
        `<m:scr m:val="${script}"/>`,
        `Generated DOCX must preserve the ${script} OMML script.`,
      );
    }
    for (const style of ["p", "b", "i", "bi"]) {
      expectIncludes(
        fontDocumentXml,
        `<m:sty m:val="${style}"/>`,
        `Generated DOCX must preserve the ${style} OMML style.`,
      );
    }

    const roundtripRoot = process.env.VISUALTEX_OMML_ROUNDTRIP_ROOT;
    if (roundtripRoot) {
      await mkdir(roundtripRoot, { recursive: true, mode: 0o700 });
      const sourcePath = join(roundtripRoot, "font-variants-source.docx");
      await writeFile(sourcePath, Buffer.from(fontDocxBase64, "base64url"), {
        mode: 0o600,
      });
      console.log(`Font-variant Word round-trip source written to ${sourcePath}.`);
    }

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
