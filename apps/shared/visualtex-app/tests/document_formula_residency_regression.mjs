// A document formula is the same size as a live editor as it is as a picture, so the document stands still while
// formulas near the screen get their editors and those far away go back to pictures (formulas.ts, Residency). The
// editor reused the formula mode's row sizing, whose 10.8 px floor drew smaller formulas (a page at 75%, a narrow
// window) about 4% larger than their pictures: formulas grew and shrank by themselves, lines rewrapped and the view
// moved, seconds after the last scroll. Same formulas near the screen (editors) and far below (pictures) are compared,
// then the view is left still past the editors' release.
import assert from "node:assert/strict";
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// a long document (about six pages at 75%): the same paragraph and display formula again and again
const content = [];
for (let section = 1; section <= 12; section += 1) {
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
    content.push({ type: "math_display", attrs: { latex: "\\omega=\\sqrt{\\frac{k}{m}}" } });
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

  await evaluate(`(() => {
    for (const key of [
      "visualtex.onboarding.v3.completed",
      "visualtex.office.macos.first-run.v1.completed",
      "visualtex.onboarding.macos.desktop.v1.2.0.completed",
      "visualtex.office.macos.native-first-run.v1.2.0.completed",
    ]) localStorage.setItem(key, "true");
    localStorage.setItem("visualtex.document.tutorial-shown.v1", "1");
    localStorage.setItem("visualtex.document.fit-width.v1", "0");
    localStorage.setItem("visualtex.document.zoom.v1", "0.75");
    localStorage.setItem("visualtex.workspace-content.v1", JSON.stringify({ state: { content: "document" }, version: 0 }));
    localStorage.setItem("visualtex.document.autosave.v1", JSON.stringify({ doc: { type: "doc", content: ${JSON.stringify(content)} } }));
  })()`);
  await client.send("Page.reload", { ignoreCache: true });
  await sleep(500);
  const ready = await evaluate(`new Promise((resolve) => {
    const started = Date.now();
    const poll = () => {
      const canvas = document.querySelector(".mc-canvas");
      const formulas = document.querySelectorAll(".vtx-sheet .vtx-math-inline, .vtx-sheet .vtx-math-display");
      const drawn = [...formulas].filter((node) => node.getBoundingClientRect().width > 0).length;
      if (canvas && formulas.length >= 192 && drawn === formulas.length && canvas.scrollHeight > canvas.clientHeight * 4) return resolve(true);
      if (Date.now() - started > 30000) return resolve(false);
      setTimeout(poll, 100);
    };
    poll();
  })`);
  assert.ok(ready, "the long document opens at 75% with its formulas drawn");
  // the editors near the screen started (the document stands still), those far below not
  const started = await evaluate(`new Promise((resolve) => {
    const began = Date.now();
    const poll = () => {
      const paragraphs = [...document.querySelectorAll(".vtx-sheet .vtx-document > p")].filter((node) => node.querySelector(".vtx-math-inline"));
      const live = paragraphs.filter((p) => [...p.querySelectorAll(".vtx-math-inline")].every((f) => f.classList.contains("is-live"))).length;
      if (live >= 4 && document.querySelector(".vtx-sheet .vtx-math-display.is-live")) return resolve(true);
      if (Date.now() - began > 20000) return resolve(false);
      setTimeout(poll, 100);
    };
    poll();
  })`);
  assert.ok(started, "the formulas near the screen get their editors");
  await sleep(1000);

  // each formula as an editor (near the top) and as a picture (far below): the same box
  const compare = await evaluate(`(() => {
    const paragraphs = [...document.querySelectorAll(".vtx-sheet .vtx-document > p")].filter((node) => node.querySelector(".vtx-math-inline"));
    const formulas = (p) => [...p.querySelectorAll(".vtx-math-inline")];
    const live = paragraphs.find((p) => formulas(p).every((f) => f.classList.contains("is-live")));
    const picture = [...paragraphs].reverse().find((p) => formulas(p).every((f) => !f.classList.contains("is-live")));
    const displays = [...document.querySelectorAll(".vtx-sheet .vtx-math-display")];
    const liveDisplay = displays.find((d) => d.classList.contains("is-live"));
    const pictureDisplay = [...displays].reverse().find((d) => !d.classList.contains("is-live"));
    const size = (node) => node && { width: node.getBoundingClientRect().width, height: node.getBoundingClientRect().height };
    return {
      fontSize: getComputedStyle(document.querySelector(".vtx-sheet")).fontSize,
      live: live && { paragraph: size(live), formulas: formulas(live).map(size) },
      picture: picture && { paragraph: size(picture), formulas: formulas(picture).map(size) },
      liveDisplay: size(liveDisplay),
      pictureDisplay: size(pictureDisplay),
    };
  })()`);
  console.log(JSON.stringify(compare));
  assert.equal(compare.fontSize, "10px", "the page's text is 10 px (formulas below the row editor's old 10.8 px floor)");
  assert.ok(compare.live && compare.picture && compare.liveDisplay && compare.pictureDisplay, "formulas with editors and formulas as pictures are both on the page");
  compare.live.formulas.forEach((formula, index) => {
    const picture = compare.picture.formulas[index];
    assert.ok(Math.abs(formula.width - picture.width) < 0.1 && Math.abs(formula.height - picture.height) < 0.1,
      `inline formula ${index + 1}: the editor is the picture's size (${picture.width}x${picture.height} -> ${formula.width}x${formula.height})`);
  });
  assert.ok(Math.abs(compare.live.paragraph.height - compare.picture.paragraph.height) < 0.1, "a paragraph is as tall with editors as with pictures");
  assert.ok(Math.abs(compare.liveDisplay.height - compare.pictureDisplay.height) < 0.1, `a display formula is as tall with its editor as as a picture (${compare.pictureDisplay.height} -> ${compare.liveDisplay.height})`);

  // the view moved down: the editors there start, those at the top are let go after a while; nothing moves meanwhile
  await evaluate(`document.querySelector(".mc-canvas").scrollTop = 2600`);
  await sleep(1500);
  const where = () => evaluate(`(() => {
    const canvas = document.querySelector(".mc-canvas");
    const view = canvas.getBoundingClientRect();
    const paragraph = [...document.querySelectorAll(".vtx-sheet .vtx-document > p")].find((p) => {
      const r = p.getBoundingClientRect();
      return p.querySelector(".vtx-math-inline") && r.top > view.top + 40 && r.bottom < view.bottom - 40;
    });
    return { scrollTop: canvas.scrollTop, text: paragraph?.textContent.slice(0, 16), top: paragraph?.getBoundingClientRect().top,
      released: [...document.querySelectorAll(".vtx-sheet .vtx-math-inline")].slice(0, 12).filter((f) => !f.classList.contains("is-live")).length };
  })()`);
  const before = await where();
  // the editors at the top are let go 4 s after they left (RELEASE_DELAY_MS)
  await evaluate(`new Promise((resolve) => {
    const began = Date.now();
    const poll = () => {
      const released = [...document.querySelectorAll(".vtx-sheet .vtx-math-inline")].slice(0, 12).every((f) => !f.classList.contains("is-live"));
      if (released || Date.now() - began > 20000) return resolve(released);
      setTimeout(poll, 100);
    };
    poll();
  })`);
  await sleep(1000);
  const after = await where();
  console.log(JSON.stringify({ before, after }));
  assert.ok(before.text, "a paragraph with formulas is in view");
  assert.equal(after.released, 12, "the formulas at the top went back to pictures (the release happened)");
  assert.equal(after.scrollTop, before.scrollTop, "the view was not scrolled by the release");
  assert.equal(after.text, before.text, "the same paragraph is in view");
  assert.ok(Math.abs(after.top - before.top) < 0.5, `the paragraph in view did not move: ${before.top} -> ${after.top}`);
  console.log("document formulas the same size as editors and as pictures, the document still: passed");
}, { windowSize: "860,900" });
