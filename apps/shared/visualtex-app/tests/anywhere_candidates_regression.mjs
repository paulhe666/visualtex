// The candidate list in VisualTeX Anywhere's panel (the input method's formula and the pasting card) is the main window's,
// whatever the Settings' look, and whole: the panel is only as large as its content asks for, so a list sized from its
// window came out one row high (the kernel takes the list's height from its window) and a row of candidates 113 px wide
// (the looks cap the width at the viewport's), in a card of 129 x 69 px. The page sizes the list from the screen.
// The list opens under the line, over it only where the screen has no room under it. Where the panel stands comes from the
// native side's frame event (macOS: WebKit gives a panel's window.screenY as the screen's height, so a page that asked its
// window put every list above the line, with the whole screen free below).
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
  // what WebKit tells a panel on macOS: no place of its own on the screen
  Object.defineProperty(window, "screenX", { get: () => 0 });
  Object.defineProperty(window, "screenY", { get: () => 956 });
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
      client.send("Emulation.setDeviceMetricsOverride", { width, height, deviceScaleFactor: 1, mobile: false, screenWidth: 1600, screenHeight: 906 });

    // One session of the card: the page opened with the look `style`, the panel at `frameY` on a 1600 x 906 screen (menu bar
    // 33 px), \fr typed, and the list as it stands then.
    const session = async (style, frameY) => {
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
      await evaluate(`window.__emit("visualtex-anywhere-frame", { x: 150, y: ${frameY}, width: 129, height: 69, visibleLeft: 0, visibleTop: 33, visibleRight: 1600, visibleBottom: 906 })`);
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
      return evaluate(`(() => {
        const popover = document.getElementById("mathlive-suggestion-popover");
        if (!popover || !popover.classList.contains("is-visible")) return null;
        const box = popover.getBoundingClientRect();
        const strip = document.querySelector(".anywhere-strip").getBoundingClientRect();
        const list = popover.querySelector("ul");
        const reports = window.__calls.filter(([command]) => command === "anywhere_layout");
        const layout = reports[reports.length - 1][1].layout;
        return {
          width: box.width, height: box.height, top: box.top, right: box.right, bottom: box.bottom,
          stripTop: strip.top, stripBottom: strip.bottom,
          listClient: list.clientHeight, listScroll: list.scrollHeight, rows: popover.querySelectorAll("li").length,
          layoutWidth: layout.width, layoutHeight: layout.height,
          windowWidth: innerWidth, windowHeight: innerHeight,
        };
      })()`);
    };

    // the panel in the upper part of the screen: the list whole, under the line, in each look
    for (const style of ["list", "card", "row", "bar"]) {
      const seen = await session(style, 330);
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
      // under the line, though the window itself says it stands at the screen's bottom (956)
      const middle = (seen.stripTop + seen.stripBottom) / 2;
      assert(seen.top >= middle, `${style}: the list (from ${seen.top}) is not under the formula (the line at ${middle}) with the whole screen free below it`);
    }

    // the panel low on the screen (the list does not fit under it, does over it): the list opens over the line, whole
    const low = await session("list", 700);
    assert(low, "low on the screen: the candidate list is not showing");
    const lowMiddle = (low.stripTop + low.stripBottom) / 2;
    assert(low.bottom <= lowMiddle, `low on the screen: the list (to ${low.bottom}) is not over the formula (the line at ${lowMiddle})`);
    assert(low.listClient >= Math.min(300, low.listScroll), `low on the screen: the list is ${low.listClient}px of ${low.listScroll}px high`);

    console.log("Anywhere's candidate list (the Settings' look, whole in a window as small as the card's: list, cards, horizontal, AxMath style; under the line, over it where there is no room under it): passed");
  },
  { path: "about:blank", windowSize: "400,300" },
);
