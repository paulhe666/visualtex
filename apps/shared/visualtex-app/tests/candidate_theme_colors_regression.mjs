// The candidate boxes take each theme's colours (Settings > Formula editor > Candidates, every look): every theme has
// its own, readable ones (dark themes dark bands — purple and green used to get the light theme's blue bands under
// light text), the formula editor's box and the document's "/" menu draw with them, a custom theme's follow its own
// accent, and Settings shows the box in the chosen look under the looks' buttons.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const app = JSON.parse(await readFile("package.json", "utf8")).name.replace(/^visualtex-/, "");
const PANEL = app === "windows" ? "#visualtex-native-input-suggestion-popover" : "#mathlive-suggestion-popover";
const THEMES = ["light", "beige", "dark", "purple", "green", "codex", "notion", "one", "proof", "raycast", "rose-pine", "solarized", "vercel", "vscode-plus", "xcode", "custom"];
const DARK = new Set(["dark", "purple", "green"]);
const PROPERTIES = ["band", "band-current", "picture-current", "head", "head-hover", "head-current", "head-text", "cell", "cell-current", "mark", "text"];

const rgb = (value) => {
  const match = String(value).match(/rgba?\(([\d.]+),\s*([\d.]+),\s*([\d.]+)/);
  if (match) return match.slice(1, 4).map(Number);
  const hex = String(value).trim().match(/^#([0-9a-f]{6})$/i);
  if (hex) return [0, 2, 4].map((i) => parseInt(hex[1].slice(i, i + 2), 16));
  throw new Error(`not a colour: ${value}`);
};
const luminance = (value) => {
  const [r, g, b] = rgb(value).map((c) => {
    const v = c / 255;
    return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
};
const contrast = (a, b) => {
  const [high, low] = [luminance(a), luminance(b)].sort((x, y) => y - x);
  return (high + 0.05) / (low + 0.05);
};
const same = (a, b) => rgb(a).every((c, i) => Math.abs(c - rgb(b)[i]) <= 1);

await withAppBrowser(async ({ baseUrl, client }) => {
  const evaluate = async (expression) => {
    const result = await client.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
    return result.result.value;
  };
  const key = async (value, code, vk, modifiers = 0) => {
    const common = { key: value, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers };
    const text = value.length === 1 && !modifiers ? { text: value, unmodifiedText: value } : {};
    await client.send("Input.dispatchKeyEvent", { type: value.length === 1 && !modifiers ? "keyDown" : "rawKeyDown", ...common, ...text });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
    await sleep(40);
  };
  const typeText = async (text) => {
    for (const ch of text) ch === "\\" ? await key(ch, "Backslash", 220) : await key(ch, `Key${ch.toUpperCase()}`, ch.toUpperCase().charCodeAt(0));
  };
  const open = async ({ theme, style = "list", content = "formula", custom = null }) => {
    await evaluate(`(() => {
      for (const k of ["visualtex.onboarding.v3.completed", "visualtex.office.macos.first-run.v1.completed", "visualtex.onboarding.macos.desktop.v1.2.0.completed", "visualtex.office.macos.native-first-run.v1.2.0.completed"]) localStorage.setItem(k, "true");
      localStorage.setItem("visualtex.document.tutorial-shown.v1", "1");
      localStorage.removeItem("visualtex.active-theme");
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      persisted.state = { ...(persisted.state || {}), theme: ${JSON.stringify(theme)}, candidateStyle: ${JSON.stringify(style)},
        lines: [{ id: "line-1", latex: "" }], activeLineId: "line-1", usage: {}, personalize: false, suggestionCount: 6 };
      localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
      localStorage.setItem("visualtex.workspace-content.v1", JSON.stringify({ state: { content: ${JSON.stringify(content)} }, version: 0 }));
      localStorage.setItem("visualtex.document.autosave.v1", JSON.stringify({ doc: { type: "doc", content: [{ type: "paragraph", content: [{ type: "text", text: "Text " }] }] } }));
      const custom = ${JSON.stringify(custom)};
      if (custom) localStorage.setItem("visualtex.custom-theme.v1", JSON.stringify(custom)); else localStorage.removeItem("visualtex.custom-theme.v1");
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(1600);
    for (let t = 0; t < 100 && !(await evaluate(`document.documentElement.dataset.theme === ${JSON.stringify(theme)}`)); t += 1) await sleep(50);
  };
  const candidateColours = () => evaluate(`(() => {
    const style = getComputedStyle(document.documentElement);
    const out = {};
    for (const name of ${JSON.stringify(PROPERTIES)}) out[name] = style.getPropertyValue("--cand-" + name).trim();
    out.surface = style.getPropertyValue("--surface").trim();
    return out;
  })()`);
  // a colour as the browser computes it (custom properties may hold var()/color-mix())
  const resolve = (value) => evaluate(`(() => { const probe = document.createElement("span"); probe.style.color = ${JSON.stringify(value)};
    document.body.append(probe); const colour = getComputedStyle(probe).color; probe.remove(); return colour; })()`);
  const openPanel = async () => {
    for (let t = 0; t < 100 && !(await evaluate(`Boolean(document.querySelector("math-field.visual-mathfield")?.shadowRoot)`)); t += 1) await sleep(50);
    await evaluate(`(() => { const f = document.querySelector("math-field.visual-mathfield");
      f.focus(); f.position = f.lastOffset; f.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true }); })()`);
    await sleep(150);
    await typeText("\\fr");
    for (let t = 0; t < 60 && !(await evaluate(`Boolean(document.querySelector("${PANEL}.is-visible li.ML__popover__current"))`)); t += 1) await sleep(50);
    await sleep(250);
  };

  // every theme's own colours, readable
  const seen = new Map();
  for (const theme of THEMES) {
    await open({ theme });
    const raw = await candidateColours();
    const colours = {};
    for (const [name, value] of Object.entries(raw)) colours[name] = await resolve(value);
    const label = (name) => `${theme}: ${name}`;
    assert.ok(contrast(colours.text, colours.band) >= 4.5, `${label("text on the band")} ${contrast(colours.text, colours.band).toFixed(2)}`);
    assert.ok(contrast(colours.text, colours["band-current"]) >= 4.5, `${label("text on the current band")} ${contrast(colours.text, colours["band-current"]).toFixed(2)}`);
    assert.ok(contrast(colours.text, colours.cell) >= 4.5, `${label("text on a picture cell")}`);
    assert.ok(contrast(colours.text, colours["cell-current"]) >= 4.5, `${label("text on the current picture cell")}`);
    assert.ok(contrast(colours["head-text"], colours.head) >= 7, `${label("header text on its band")} ${contrast(colours["head-text"], colours.head).toFixed(2)}`);
    assert.ok(contrast(colours["head-text"], colours["head-current"]) >= 4.5, `${label("header text on the current band")} ${contrast(colours["head-text"], colours["head-current"]).toFixed(2)}`);
    assert.ok(contrast(colours.mark, colours["band-current"]) >= 3, `${label("numbers on the current band")} ${contrast(colours.mark, colours["band-current"]).toFixed(2)}`);
    if (DARK.has(theme)) assert.ok(luminance(colours.band) < 0.1 && luminance(colours["band-current"]) < 0.15, `${label("a dark theme's bands are dark")} ${colours.band}`);
    else assert.ok(luminance(colours.band) > 0.6, `${label("a light theme's band is light")} ${colours.band}`);
    seen.set(theme, colours);
  }
  // each named theme its own colours (custom starts as Codex's)
  for (const theme of THEMES.filter((t) => t !== "light" && t !== "custom")) {
    assert.ok(!same(seen.get(theme)["band-current"], seen.get("light")["band-current"]), `${theme} has its own current band, not the light theme's`);
    assert.ok(!same(seen.get(theme).head, seen.get("light").head), `${theme} has its own header band, not the light theme's`);
  }

  // the formula editor's box and the document's menu draw with them
  for (const theme of ["light", "beige", "purple", "raycast"]) {
    const colours = seen.get(theme);
    await open({ theme, style: "list" });
    await openPanel();
    const list = await evaluate(`(() => { const li = document.querySelector("${PANEL} li.ML__popover__current"); if (!li) return null;
      return { image: getComputedStyle(li).backgroundImage, text: getComputedStyle(li.querySelector(".ML__popover__latex")).color }; })()`);
    assert.ok(list, `${theme}: the formula editor's candidate box opens`);
    const band = list.image.match(/rgba?\([^)]*\)/)?.[0];
    assert.ok(band && same(band, colours["band-current"]), `${theme}: the default list's current band is the theme's (${band} vs ${colours["band-current"]})`);
    assert.ok(same(list.text, colours.text), `${theme}: the commands are the theme's candidate text`);

    await open({ theme, style: "row" });
    await openPanel();
    const head = await evaluate(`(() => { const li = document.querySelector("${PANEL} li.ML__popover__current"); if (!li) return null;
      const other = [...document.querySelectorAll("${PANEL} li")].find((n) => n !== li);
      return { current: getComputedStyle(li.querySelector(".ML__popover__index")).backgroundColor,
        resting: other && getComputedStyle(other.querySelector(".ML__popover__index")).backgroundColor,
        cell: other && getComputedStyle(other.querySelector(".ML__popover__command")).backgroundColor }; })()`);
    assert.ok(head && same(head.current, colours["head-current"]), `${theme}: the horizontal look's current header is the theme's`);
    assert.ok(head.resting && same(head.resting, colours.head), `${theme}: the horizontal look's header band is the theme's`);
    assert.ok(same(head.cell, colours.cell), `${theme}: the horizontal look's picture cell is the theme's`);

    await open({ theme, style: "bar", content: "document" });
    for (let attempt = 0; attempt < 3 && !(await evaluate(`Boolean(document.querySelector(".vtx-cmd-menu .vtx-cmd-item.is-current"))`)); attempt += 1) {
      for (let t = 0; t < 100 && !(await evaluate(`Boolean(document.querySelector(".vtx-document p")?.getBoundingClientRect().width)`)); t += 1) await sleep(50);
      await sleep(300);
      const pos = await evaluate(`(() => { const p = document.querySelector(".vtx-document p"); const r = p.getBoundingClientRect(); return { x: r.left + 200, y: r.top + r.height / 2 }; })()`);
      for (const type of ["mousePressed", "mouseReleased"]) await client.send("Input.dispatchMouseEvent", { type, x: pos.x, y: pos.y, button: "left", clickCount: 1 });
      await sleep(200);
      if (attempt) await key("Backspace", "Backspace", 8);
      await key("/", "Slash", 191);
      for (let t = 0; t < 40 && !(await evaluate(`Boolean(document.querySelector(".vtx-cmd-menu .vtx-cmd-item.is-current"))`)); t += 1) await sleep(50);
    }
    const menu = await evaluate(`(() => { const item = document.querySelector(".vtx-cmd-menu .vtx-cmd-item.is-current"); if (!item) return null;
      return { head: getComputedStyle(item.querySelector(".vtx-cmd-copy")).backgroundColor, cell: getComputedStyle(item.querySelector(".vtx-cmd-preview")).backgroundColor }; })()`);
    assert.ok(menu && same(menu.head, colours["head-current"]) && same(menu.cell, colours["cell-current"]), `${theme}: the document's "/" menu (AxMath look) draws with the theme's colours`);
    // the "/" typed goes again (the document is saved as it is when the page reloads)
    await key("Escape", "Escape", 27);
    await key("Backspace", "Backspace", 8);
    await sleep(300);
  }

  // a custom theme's candidate colours follow its own accent
  const codex = seen.get("codex");
  await open({ theme: "custom", custom: { version: 1, mode: "light", colors: { accent: "#C2410C", accentHover: "#9A3412", accentSoft: "#FFEDD5", background: "#FFFFFF", surface: "#FFFFFF", foreground: "#1F2937" } } });
  const red = {};
  for (const [name, value] of Object.entries(await candidateColours())) red[name] = await resolve(value);
  const [r, , b] = rgb(red["band-current"]);
  assert.ok(r > b + 15, `a custom theme with an orange accent gets orange bands, not Codex's blue (${red["band-current"]})`);
  assert.ok(!same(red.head, codex.head), "a custom theme's header band follows its accent");
  assert.ok(contrast(red["head-text"], red["head-current"]) >= 4.5 && contrast(red.text, red["band-current"]) >= 4.5, "a custom theme's candidate colours stay readable");

  // Settings: the box in the chosen look under the looks' buttons, in the theme's colours
  await open({ theme: "beige", style: "list" });
  await key(",", "Comma", 188, 2);
  await sleep(800);
  for (let t = 0; t < 40 && !(await evaluate(`Boolean(document.querySelector('[data-settings-page="editor"]'))`)); t += 1) await sleep(50);
  await evaluate(`document.querySelector('[data-settings-page="editor"]')?.click()`);
  await sleep(700);
  const looks = [["list", "默认", "Default"], ["card", "卡片", "Cards"], ["row", "横排", "Horizontal"], ["bar", "AxMath 风格", "AxMath style"]];
  const beige = seen.get("beige");
  for (const [look, zh, en] of looks) {
    const clicked = await evaluate(`(() => { const button = [...document.querySelectorAll(".settings-inline-choice button")].find((b) => [${JSON.stringify(zh)}, ${JSON.stringify(en)}].includes(b.textContent.trim()));
      if (!button) return false; button.click(); return true; })()`);
    assert.ok(clicked, `Settings has the ${look} look's button`);
    await sleep(300);
    const preview = await evaluate(`(() => {
      const menu = document.querySelector(".candidate-style-preview .vtx-cmd-menu");
      if (!menu) return null;
      const buttons = document.querySelector(".settings-inline-choice").getBoundingClientRect();
      const box = menu.getBoundingClientRect();
      const current = menu.querySelector(".vtx-cmd-item.is-current");
      return { style: document.documentElement.dataset.candidateStyle, below: box.top >= buttons.bottom - 1, width: box.width, height: box.height,
        direction: getComputedStyle(menu.querySelector(".vtx-cmd-list")).flexDirection, items: menu.querySelectorAll(".vtx-cmd-item").length,
        image: getComputedStyle(current).backgroundImage, fill: getComputedStyle(current).backgroundColor,
        head: getComputedStyle(current.querySelector(".vtx-cmd-copy")).backgroundColor,
        formulas: [...menu.querySelectorAll(".vtx-cmd-formula")].filter((f) => f.getBoundingClientRect().width > 0).length };
    })()`);
    assert.ok(preview, `the ${look} look has a preview`);
    assert.equal(preview.style, look, `the ${look} button chose the ${look} look`);
    assert.ok(preview.below && preview.width > 200 && preview.height > 40, `the preview stands under the buttons (${look})`);
    assert.equal(preview.items, 4, `the preview shows the candidates (${look})`);
    assert.equal(preview.formulas, 4, `the preview draws the candidates' pictures (${look})`);
    if (look === "list") assert.ok(same(preview.image.match(/rgba?\([^)]*\)/)?.[0] ?? "", beige["band-current"]), "the preview's default list is in the theme's colours");
    if (look === "card") assert.ok(same(preview.fill, beige.band), "the preview's cards are in the theme's colours");
    if (look === "row" || look === "bar") {
      assert.equal(preview.direction, "row", `the preview's ${look} look is one row`);
      assert.ok(same(preview.head, beige["head-current"]), `the preview's ${look} look is in the theme's colours`);
    }
  }
  console.log("candidate boxes in every theme's colours, Settings preview: passed");
}, { windowSize: "1200,900" });
