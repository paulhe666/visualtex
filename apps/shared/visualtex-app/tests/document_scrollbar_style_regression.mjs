// Document mode keeps its layout when the scroll bar changes kind. macOS switches between overlay scroll bars (no
// room) and classic ones (15 px) on its own — a mouse attached or let go, Universal Control taking the pointer to
// another device — and Windows can do the same. Fitted to the room inside the scroll bar, every switch rescaled the
// page by about 2%: text and formulas grew or shrank, the pages reflowed and the view jumped. Here the canvas's
// scroll bar is taken away and given back (scrollbar-width) while the page is measured, with the page fitted to the
// width and at a zoom chosen with + / -.
import assert from "node:assert/strict";
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// a long document: sections, paragraphs with inline formulas, numbered display formulas
const content = [];
for (let section = 1; section <= 6; section += 1) {
  content.push({ type: "heading", attrs: { level: 1 }, content: [{ type: "text", text: `Section ${section}` }] });
  for (let paragraph = 1; paragraph <= 4; paragraph += 1) {
    content.push({
      type: "paragraph",
      content: [
        { type: "text", text: `Paragraph ${section}.${paragraph}: the energy of a mass ` },
        { type: "math_inline", attrs: { latex: "m" } },
        { type: "text", text: " on a spring of stiffness " },
        { type: "math_inline", attrs: { latex: "k" } },
        { type: "text", text: " is " },
        { type: "math_inline", attrs: { latex: "E=\\frac{1}{2}mv^2+\\frac{1}{2}kx^2" } },
        { type: "text", text: ", which stays the same while the mass moves back and forth, and its frequency is" },
      ],
    });
    content.push({ type: "math_display", attrs: { latex: `\\omega_{${section}${paragraph}}=\\sqrt{\\frac{k}{m}}`, numbered: true } });
  }
}

await withAppBrowser(async ({ baseUrl, client }) => {
  const evaluate = async (expression) => {
    const result = await client.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
    if (result.exceptionDetails) {
      throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text || "Runtime.evaluate failed");
    }
    return result.result.value;
  };
  for (let tries = 0; tries < 300 && !String(await evaluate("location.href")).startsWith(baseUrl); tries += 1) await sleep(50);

  // the page's width, place and font size, where a paragraph and a formula in view are, and their widths
  const measure = () => evaluate(`(() => {
    const canvas = document.querySelector(".mc-canvas");
    const sheet = document.querySelector(".vtx-sheet");
    const view = canvas.getBoundingClientRect();
    const inView = (node) => { const r = node.getBoundingClientRect(); return r.top > view.top + 40 && r.bottom < view.bottom - 40; };
    const paragraph = [...sheet.querySelectorAll(".vtx-document > p")].filter((node) => node.querySelector(".vtx-math-inline")).find(inView);
    const inline = paragraph && [...paragraph.querySelectorAll(".vtx-math-inline")].pop();
    const formula = [...sheet.querySelectorAll(".vtx-math-display")].find(inView);
    const box = (node) => { const r = node.getBoundingClientRect(); return { left: r.left, top: r.top, width: r.width, height: r.height }; };
    return {
      scrollbar: canvas.offsetWidth - canvas.clientWidth,
      sheetWidth: sheet.getBoundingClientRect().width,
      sheetLeft: sheet.getBoundingClientRect().left,
      fontSize: getComputedStyle(sheet).fontSize,
      scrollTop: canvas.scrollTop,
      paragraph: paragraph && { text: paragraph.textContent.slice(0, 20), ...box(paragraph) },
      inline: inline && box(inline),
      formula: formula && box(formula),
    };
  })()`);

  for (const [mode, fitWidth, zoom] of [["fitted to the width", "1", "1"], ["at zoom 0.75", "0", "0.75"]]) {
    await evaluate(`(() => {
      for (const key of [
        "visualtex.onboarding.v3.completed",
        "visualtex.office.macos.first-run.v1.completed",
        "visualtex.onboarding.macos.desktop.v1.2.0.completed",
        "visualtex.office.macos.native-first-run.v1.2.0.completed",
      ]) localStorage.setItem(key, "true");
      localStorage.setItem("visualtex.document.tutorial-shown.v1", "1");
      localStorage.setItem("visualtex.document.fit-width.v1", ${JSON.stringify(fitWidth)});
      localStorage.setItem("visualtex.document.zoom.v1", ${JSON.stringify(zoom)});
      localStorage.setItem("visualtex.workspace-content.v1", JSON.stringify({ state: { content: "document" }, version: 0 }));
      localStorage.setItem("visualtex.document.autosave.v1", JSON.stringify({ doc: { type: "doc", content: ${JSON.stringify(content)} } }));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(500);
    // the document open, its formulas drawn
    const ready = await evaluate(`new Promise((resolve) => {
      const started = Date.now();
      const poll = () => {
        const canvas = document.querySelector(".mc-canvas");
        const formulas = document.querySelectorAll(".vtx-sheet .vtx-math-inline, .vtx-sheet .vtx-math-display");
        const drawn = [...formulas].filter((node) => node.getBoundingClientRect().width > 0).length;
        if (canvas && formulas.length >= 72 && drawn === formulas.length && canvas.scrollHeight > canvas.clientHeight * 2) return resolve(true);
        if (Date.now() - started > 30000) return resolve(false);
        setTimeout(poll, 100);
      };
      poll();
    })`);
    assert.ok(ready, `the long document opens in document mode with its formulas drawn (${mode})`);
    await sleep(800);
    await evaluate(`document.querySelector(".mc-canvas").scrollTop = 1600`);
    await sleep(600);

    // a browser that draws only overlay bars (macOS' headless Chrome) gets a classic one, 15 px as macOS's
    if (await evaluate(`(() => { const c = document.querySelector(".mc-canvas"); return c.offsetWidth - c.clientWidth === 0; })()`)) {
      await evaluate(`(() => {
        const style = document.createElement("style");
        style.id = "classic-scroll-bar";
        style.textContent = ".mc-canvas::-webkit-scrollbar { width: 15px; background: #eee; } .mc-canvas::-webkit-scrollbar-thumb { background: #999; }";
        document.head.append(style);
      })()`);
      await sleep(900);
    }
    const classic = await measure();
    assert.ok(classic.paragraph && classic.inline && classic.formula, `a paragraph with formulas and a display formula are in view (${mode})`);
    await evaluate(`(() => {
      const style = document.createElement("style");
      style.id = "no-scroll-bar-room";
      style.textContent = ".mc-canvas { scrollbar-width: none !important; }";
      document.head.append(style);
    })()`);
    await sleep(900);
    const overlay = await measure();
    await evaluate(`document.getElementById("no-scroll-bar-room").remove()`);
    await sleep(900);
    const back = await measure();

    console.log(mode, JSON.stringify({ classic, overlay, back }));
    assert.notEqual(classic.scrollbar, overlay.scrollbar, `the scroll bar's room changes (the switch being simulated, ${mode})`);
    for (const [label, state] of [["without the scroll bar's room", overlay], ["with it again", back]]) {
      const where = `${label}, ${mode}`;
      assert.equal(state.fontSize, classic.fontSize, `the page's text size stays the same ${where}`);
      assert.ok(Math.abs(state.sheetWidth - classic.sheetWidth) < 0.5, `the page keeps its width ${where}: ${classic.sheetWidth} -> ${state.sheetWidth}`);
      assert.ok(Math.abs(state.sheetLeft - classic.sheetLeft) < 0.5, `the page stays where it is ${where}: ${classic.sheetLeft} -> ${state.sheetLeft}`);
      assert.equal(state.paragraph?.text, classic.paragraph.text, `the same paragraph is in view ${where}`);
      assert.ok(Math.abs(state.paragraph.top - classic.paragraph.top) < 1, `the paragraph does not move ${where}: ${classic.paragraph.top} -> ${state.paragraph.top}`);
      assert.ok(Math.abs(state.paragraph.height - classic.paragraph.height) < 0.5, `the paragraph keeps its lines ${where}`);
      assert.ok(Math.abs(state.inline.width - classic.inline.width) < 0.5, `the inline formula keeps its size ${where}: ${classic.inline.width} -> ${state.inline.width}`);
      assert.ok(Math.abs(state.inline.left - classic.inline.left) < 0.5 && Math.abs(state.inline.top - classic.inline.top) < 1, `the inline formula does not move ${where}`);
      assert.ok(Math.abs(state.formula.width - classic.formula.width) < 0.5, `the display formula keeps its size ${where}: ${classic.formula.width} -> ${state.formula.width}`);
      assert.ok(Math.abs(state.formula.top - classic.formula.top) < 1, `the display formula does not move ${where}`);
    }
  }
  console.log("document layout unchanged by the scroll bar's kind: passed");
}, { windowSize: "860,900" });
