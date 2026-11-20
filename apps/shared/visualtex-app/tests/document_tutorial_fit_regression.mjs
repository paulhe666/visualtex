// Document mode's tutorial fits the window. In a low window the dialog used to scroll and hid its dots and buttons
// below the fold (a 900×640 window at 150% on Windows): now the page with the animation gives up height first, and in
// a very low window it goes, so the card's text and the buttons stay in view without scrolling. In a large window
// the page keeps its 2:1 shape.
import assert from "node:assert/strict";
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

await withAppBrowser(async ({ baseUrl, client }) => {
  const evaluate = async (expression) => {
    const result = await client.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
    if (result.exceptionDetails) {
      throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text || "Runtime.evaluate failed");
    }
    return result.result.value;
  };
  const viewport = (width, height) =>
    client.send("Emulation.setDeviceMetricsOverride", { width, height, deviceScaleFactor: 1, mobile: false });
  for (let tries = 0; tries < 300 && !String(await evaluate("location.href")).startsWith(baseUrl); tries += 1) await sleep(50);

  for (const language of ["cn", "en"]) {
    await evaluate(`(() => {
      for (const key of [
        "visualtex.onboarding.v3.completed",
        "visualtex.office.macos.first-run.v1.completed",
        "visualtex.onboarding.macos.desktop.v1.2.0.completed",
        "visualtex.office.macos.native-first-run.v1.2.0.completed",
      ]) localStorage.setItem(key, "true");
      localStorage.setItem("visualtex.document.tutorial-shown.v1", "1");
      localStorage.setItem("visualtex.workspace-content.v1", JSON.stringify({ state: { content: "document" }, version: 0 }));
      const editor = JSON.parse(localStorage.getItem("visualtex-editor") || '{"state":{},"version":0}');
      editor.state = { ...editor.state, language: ${JSON.stringify(language)} };
      localStorage.setItem("visualtex-editor", JSON.stringify(editor));
    })()`);
    await viewport(1280, 800);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(500);
    for (let tries = 0; tries < 100; tries += 1) {
      const state = await evaluate("document.readyState").catch(() => "navigating");
      if (state === "complete") break;
      await sleep(100);
    }
    const opened = await evaluate(`new Promise((resolve) => {
      const started = Date.now();
      const poll = () => {
        const button = document.querySelector(".vtx-tutorial-button");
        if (button && !document.querySelector(".vtx-tutorial-dialog")) button.click();
        if (document.querySelector(".vtx-tutorial-dialog .vtx-tutorial-stage img")) return resolve(true);
        if (Date.now() - started > 30000) return resolve(false);
        setTimeout(poll, 150);
      };
      poll();
    })`);
    assert.ok(opened, `the tutorial opens from document mode's toolbar with its first animation (${language})`);

    const cards = await evaluate(`document.querySelectorAll(".vtx-tutorial-dots button").length`);
    assert.ok(cards >= 3, `the tutorial has its cards (${cards})`);

    // [width, height, what the page does]: keeps its shape, gives up height, or goes
    for (const [width, height, page] of [
      [1280, 800, "2:1"],
      [720, 560, "2:1"],
      [600, 397, "lower"],
      [640, 350, "lower"],
      [640, 340, "hidden"],
      [640, 240, "hidden"],
      [500, 272, "hidden"],
    ]) {
      await viewport(width, height);
      for (let card = 0; card < cards; card += 1) {
        await evaluate(`document.querySelectorAll(".vtx-tutorial-dots button")[${card}].click()`);
        await sleep(250);
        const m = await evaluate(`(() => {
          const dialog = document.querySelector(".vtx-tutorial-dialog");
          const stage = dialog.querySelector(".vtx-tutorial-stage");
          const footer = dialog.querySelector(".vtx-tutorial-footer");
          const primary = dialog.querySelector(".vtx-tutorial-actions .primary-button");
          const d = dialog.getBoundingClientRect();
          const s = stage.getBoundingClientRect();
          const f = footer.getBoundingClientRect();
          const p = primary.getBoundingClientRect();
          return {
            dialog: [d.left, d.top, d.right, d.bottom],
            scrolls: dialog.scrollHeight > dialog.clientHeight + 1,
            stageShown: getComputedStyle(stage).display !== "none",
            stage: [s.width, s.height],
            footer: [f.top, f.bottom],
            primary: [p.left, p.top, p.right, p.bottom],
            hit: document.elementFromPoint((p.left + p.right) / 2, (p.top + p.bottom) / 2)?.closest(".primary-button") === primary,
          };
        })()`);
        const where = `${width}×${height}, card ${card + 1}, ${language}: ${JSON.stringify(m)}`;
        assert.ok(m.dialog[0] >= 0 && m.dialog[1] >= 0 && m.dialog[2] <= width + 0.5 && m.dialog[3] <= height + 0.5, `the dialog is inside the window (${where})`);
        assert.ok(!m.scrolls, `the dialog does not scroll (${where})`);
        assert.ok(m.footer[0] >= m.dialog[1] && m.footer[1] <= m.dialog[3] + 0.5, `the dots and buttons are in view (${where})`);
        assert.ok(m.hit, `the primary button can be clicked where it is drawn (${where})`);
        if (page === "hidden") {
          assert.ok(!m.stageShown, `in a very low window the page goes (${where})`);
        } else {
          assert.ok(m.stageShown, `the page is shown (${where})`);
          if (page === "2:1") assert.ok(Math.abs(m.stage[0] / m.stage[1] - 2) < 0.02, `the page keeps its 2:1 shape (${where})`);
          else assert.ok(m.stage[1] >= 110 && m.stage[0] / m.stage[1] > 2.05, `the page is lower but still there (${where})`);
        }
      }
    }
    await evaluate(`document.querySelector(".vtx-tutorial-close").click()`);
    await sleep(200);
  }
  console.log("document tutorial fits the window: passed");
}, { windowSize: "1280,800" });
