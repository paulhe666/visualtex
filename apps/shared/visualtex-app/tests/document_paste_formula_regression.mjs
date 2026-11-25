// "$$…$$" pasted into a document line is a display formula, whatever came with the text: plain text alone, or HTML that
// only shows the text (a chat message, a web page's paragraph, VS Code, a code block holding nothing but the formula).
// A formula with \begin{pmatrix} came in from a browser as the line's text: the paste took the HTML for a view of
// source text and let the editor read the HTML, which made it a line (2026-11-24). A code block of code stays code.
// Pasted into the last line, the formula gets a line below it for the caret: two formulas pasted one after the other
// are both kept (the first stayed selected and the second went over it), and what is typed next goes below them.
import assert from "node:assert/strict";
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const LATEX = String.raw`$$\hat{H} = \frac{\hbar}{2} \begin{pmatrix} \omega + \delta & \Omega e^{-i(\omega t - k\hat{x})} \\ \Omega e^{i(\omega t - k\hat{x})} & -\omega - \delta \end{pmatrix}$$`;
// the user's second one: aligned rows, \vert{}, \text
const ALIGNED = String.raw`$$\begin{aligned} \langle \uparrow, n \vert{} H^{\text{car}} \vert{} \downarrow, n \rangle &= \frac{\hbar\Omega}{2} \langle \uparrow \vert{} \sigma^+ \vert{} \downarrow \rangle \langle n \vert{} n \rangle = \frac{\hbar\Omega}{2} \\ \langle \uparrow, n-1 \vert{} H^{\text{red}} \vert{} \downarrow, n \rangle &= \frac{\hbar\eta\Omega}{2} \langle \uparrow \vert{} \sigma^+ \vert{} \downarrow \rangle \langle n-1 \vert{} \hat{a} \vert{} n \rangle = \frac{\sqrt{n}\hbar\eta\Omega}{2} \\ \langle \uparrow, n+1 \vert{} H^{\text{blue}} \vert{} \downarrow, n \rangle &= \frac{\hbar\eta\Omega}{2} \langle \uparrow \vert{} \sigma^+ \vert{} \downarrow \rangle \langle n+1 \vert{} \hat{a}^\dagger \vert{} n \rangle = \frac{\sqrt{n+1}\hbar\eta\Omega}{2} \end{aligned}$$`;
const escapeHtml = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const html = (inner) => `<meta charset="utf-8">${inner}`;
const FORMULA = ["P:first line", "DISPLAY", "P:"];
const chat = (text) => ({ "text/plain": text, "text/html": html(`<div style="white-space: pre-wrap;">${escapeHtml(text)}</div>`) });
const CASES = [
  ["plain text", { "text/plain": LATEX }, FORMULA],
  ["plain text, CRLF", { "text/plain": LATEX + "\r\n" }, FORMULA],
  ["a web page's paragraph", { "text/plain": LATEX, "text/html": html(`<p>${escapeHtml(LATEX)}</p>`) }, FORMULA],
  ["styled text", { "text/plain": LATEX, "text/html": html(`<span style="color: rgb(0,0,0); font-family: sans-serif;">${escapeHtml(LATEX)}</span>`) }, FORMULA],
  ["a chat message", { "text/plain": LATEX, "text/html": html(`<div style="white-space: pre-wrap; word-break: break-word;">${escapeHtml(LATEX)}</div>`) }, FORMULA],
  ["VS Code", { "text/plain": LATEX, "text/html": html(`<div style="color: #000; white-space: pre;"><div><span>${escapeHtml(LATEX)}</span></div></div>`), "vscode-editor-data": '{"mode":"markdown"}' }, FORMULA],
  ["a code block holding the formula", { "text/plain": LATEX, "text/html": html(`<pre><code>${escapeHtml(LATEX)}</code></pre>`) }, FORMULA],
  ["a code block of code", { "text/plain": "# install\nnpm i", "text/html": html("<pre><code># install\nnpm i</code></pre>") }, ["P:first line", "P:```", "P:# install", "P:npm i", "P:```"]],
  ["aligned rows, plain text", { "text/plain": ALIGNED }, FORMULA],
  ["aligned rows, a chat message", { "text/plain": ALIGNED, "text/html": html(`<div style="white-space: pre-wrap;">${escapeHtml(ALIGNED)}</div>`) }, FORMULA],
  ["aligned rows, a web page's paragraph", { "text/plain": ALIGNED, "text/html": html(`<p>${escapeHtml(ALIGNED)}</p>`) }, FORMULA],
  ["two formulas one after the other, then a key", [chat(LATEX), chat(ALIGNED), "z"], ["P:first line", "DISPLAY", "DISPLAY", "P:z"]],
  ["two plain-text formulas one after the other", [{ "text/plain": LATEX }, { "text/plain": ALIGNED }], ["P:first line", "DISPLAY", "DISPLAY", "P:"]],
];

for (const [name, steps, expected] of CASES) {
  await withAppBrowser(async ({ baseUrl, client }) => {
    const evaluate = async (expression) => {
      const result = await client.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
      if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
      return result.result.value;
    };
    for (let tries = 0; tries < 300 && !String(await evaluate("location.href")).startsWith(baseUrl); tries += 1) await sleep(50);
    await evaluate(`(() => {
      for (const k of ["visualtex.onboarding.v3.completed", "visualtex.office.macos.first-run.v1.completed", "visualtex.onboarding.macos.desktop.v1.2.0.completed", "visualtex.office.macos.native-first-run.v1.2.0.completed"]) localStorage.setItem(k, "true");
      localStorage.setItem("visualtex.document.tutorial-shown.v1", "1");
      localStorage.setItem("visualtex.workspace-content.v1", JSON.stringify({ state: { content: "document" }, version: 0 }));
      localStorage.setItem("visualtex.document.autosave.v1", JSON.stringify({ doc: { type: "doc", content: [
        { type: "paragraph", content: [{ type: "text", text: "first line" }] }, { type: "paragraph" } ] } }));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(800);
    for (let t = 0; t < 200 && !(await evaluate(`Boolean(document.querySelector(".vtx-document p")?.getBoundingClientRect().width)`)); t += 1) await sleep(50);
    await sleep(500);
    // the caret in the empty second line, then the paste
    const at = await evaluate(`(() => { const r = document.querySelectorAll(".vtx-document p")[1].getBoundingClientRect(); return { x: r.left + 4, y: r.top + r.height / 2 }; })()`);
    for (const type of ["mousePressed", "mouseReleased"]) await client.send("Input.dispatchMouseEvent", { type, x: at.x, y: at.y, button: "left", clickCount: 1 });
    await sleep(300);
    for (const step of Array.isArray(steps) ? steps : [steps]) {
      if (typeof step === "string") {
        await client.send("Input.insertText", { text: step });
        await sleep(400);
        continue;
      }
      await evaluate(`(() => {
        const dt = new DataTransfer();
        for (const [type, value] of Object.entries(${JSON.stringify(step)})) dt.setData(type, value);
        document.querySelector(".vtx-document").dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
      })()`);
      await sleep(1200);
    }
    // a display formula drawn (its LaTeX shown as text when the drawing fails), without the drawing's error marks
    const blocks = await evaluate(`[...document.querySelector(".vtx-document").children].map((el) => {
      if (!el.classList.contains("vtx-math-display")) return el.tagName + ":" + el.textContent;
      const box = el.querySelector(".vtx-formula-static");
      if (!box?.firstElementChild) return box?.textContent ? "DISPLAY AS TEXT" : "EMPTY DISPLAY";
      return box.querySelector("[class*=error], .ML__error") ? "DISPLAY WITH ERRORS" : "DISPLAY";
    })`);
    console.log(`${name}: ${JSON.stringify(blocks).slice(0, 160)}`);
    assert.deepEqual(blocks, expected, name);
  }, { windowSize: "1100,900" });
}

console.log("pasted $$…$$ is a display formula, a code block of code stays code: passed");
