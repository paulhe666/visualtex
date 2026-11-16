// The candidate list in VisualTeX Anywhere's panel (the input method's formula and the pasting card) is the main window's,
// whatever the Settings' look, and whole: the panel is only as large as its content asks for, so a list sized from its
// window came out one row high (the kernel takes the list's height from its window) and a row of candidates 113 px wide
// (the looks cap the width at the viewport's), in a card of 129 x 69 px. The page sizes the list from the screen.
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

// the page's Tauri side: events in, commands recorded (anywhere_layout is what the native side makes the window from)
const STAND_IN = `
(() => {
  const handlers = new Map();
  let next = 1;
  window.__calls = [];
  window.__listeners = [];
  window.__TAURI_INTERNALS__ = {
    metadata: { currentWindow: { label: "visualtex-anywhere" }, currentWebview: { label: "visualtex-anywhere", windowLabel: "visualtex-anywhere" } },
    transformCallback(callback) { const id = next++; handlers.set(id, callback); return id; },
    unregisterCallback(id) { handlers.delete(id); },
    convertFileSrc: (path) => path,
    async invoke(command, args) {
      window.__calls.push([command, args]);
      if (command === "plugin:event|listen") { window.__listeners.push([args.event, args.handler]); return args.handler; }
      return null;
    },
  };
  window.__TAURI_EVENT_PLUGIN_INTERNALS__ = { unregisterListener() {} };
  window.__emit = (event, payload) => { for (const [name, id] of window.__listeners) if (name === event) handlers.get(id)?.({ event, id, payload }); };
})();
`;

const KEYS = { "\\": ["Backslash", 220], f: ["KeyF", 70], r: ["KeyR", 82] };

await withAppBrowser(
  async ({ baseUrl, client, evaluate }) => {
    await client.send("Page.addScriptToEvaluateOnNewDocument", { source: STAND_IN });
    const type = async (text) => {
      for (const ch of text) {
        const [code, vk] = KEYS[ch];
        const common = { key: ch, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk };
        await client.send("Input.dispatchKeyEvent", { type: "keyDown", ...common, text: ch, unmodifiedText: ch });
        await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
        await sleep(80);
      }
    };

    // the window as the card has it (129 x 69 px) on a larger screen: the window's size changes, the screen's does not
    const window_ = (width, height) =>
      client.send("Emulation.setDeviceMetricsOverride", { width, height, deviceScaleFactor: 1, mobile: false, screenWidth: 1600, screenHeight: 900 });

    for (const style of ["list", "card", "row", "bar"]) {
      await window_(129, 69);
      await client.send("Page.navigate", { url: `${baseUrl}/anywhere.html?transport=tauri&theme=light` });
      await sleep(1500);
      await evaluate(`(() => {
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        persisted.state = { ...(persisted.state || {}), candidateStyle: ${JSON.stringify(style)}, suggestionCount: 8 };
        localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
      })()`);
      await client.send("Page.reload", { ignoreCache: true });
      await sleep(2500);
      await evaluate(`window.__emit("visualtex-anywhere-activate", { generation: 1, caretHeight: 22, anchorSource: "caret", appName: "Test", accessibilityTrusted: true, resume: false, inputMethod: false, fontSize: 16, lineHeight: 22, lineDescent: 5, textColor: "#202020ff" })`);
      await sleep(1200);
      await type("\\fr");
      // the native side makes the window the size the page last reported, as it does for every report
      for (let round = 0; round < 6; round += 1) {
        await sleep(350);
        const layout = await evaluate(`(() => { const reports = window.__calls.filter(([command]) => command === "anywhere_layout"); return reports.length ? reports[reports.length - 1][1].layout : null; })()`);
        if (layout) await window_(Math.max(1, Math.round(layout.width)), Math.max(1, Math.round(layout.height)));
      }
      await sleep(500);
      const seen = await evaluate(`(() => {
        const popover = document.getElementById("mathlive-suggestion-popover");
        if (!popover || !popover.classList.contains("is-visible")) return null;
        const box = popover.getBoundingClientRect();
        const list = popover.querySelector("ul");
        const reports = window.__calls.filter(([command]) => command === "anywhere_layout");
        const layout = reports[reports.length - 1][1].layout;
        return {
          width: box.width, height: box.height, right: box.right, bottom: box.bottom,
          listClient: list.clientHeight, listScroll: list.scrollHeight, rows: popover.querySelectorAll("li").length,
          layoutWidth: layout.width, layoutHeight: layout.height,
          windowWidth: innerWidth, windowHeight: innerHeight,
        };
      })()`);
      assert(seen, `${style}: the candidate list is not showing`);
      assert(seen.rows >= 6, `${style}: ${seen.rows} candidates`);
      if (style === "list" || style === "card") {
        // a column of candidates, as tall as the kernel makes it (320 px) or as its candidates are
        const wanted = Math.min(300, seen.listScroll);
        assert(seen.listClient >= wanted, `${style}: the list is ${seen.listClient}px of ${seen.listScroll}px high (a window of 129 x 69 px made it one row)`);
        assert(seen.width >= 300, `${style}: the list is ${seen.width}px wide`);
      } else {
        // one row of candidates: as wide as they are, on the screen's width and not the window's
        assert(seen.width >= 600, `${style}: the row of candidates is ${seen.width}px wide (a window of 129 px made it 113)`);
        assert(seen.listClient >= 60, `${style}: the row is ${seen.listClient}px high`);
      }
      // the window holds the whole list
      assert(seen.layoutWidth >= seen.right && seen.layoutHeight >= seen.bottom, `${style}: the panel (${seen.layoutWidth} x ${seen.layoutHeight}) does not hold the list (to ${seen.right} x ${seen.bottom})`);
      assert(seen.windowWidth >= seen.right && seen.windowHeight >= seen.bottom, `${style}: the window (${seen.windowWidth} x ${seen.windowHeight}) cuts the list (to ${seen.right} x ${seen.bottom})`);
    }

    console.log("Anywhere's candidate list (the Settings' look, whole in a window as small as the card's: list, cards, horizontal, AxMath style): passed");
  },
  { path: "about:blank", windowSize: "400,300" },
);
