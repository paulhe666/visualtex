// Records the tutorial demos from the real editor.
//
//   npm run dev -- --port 4310                       (or any running build)
//   CHROME_PATH=... node scripts/tutorial/record_tutorial.mjs --base http://localhost:4310 [--only typing] [--lang zh]
//
// Each lesson opens /editor?tutorial=<id> (the same start document as the
// practice editor), is driven with trusted CDP input, and is written to
// public/tutorial/<id>.<lang>.{mp4,jpg,json}. The JSON timeline holds the
// keys shown over the video on /tutorial. Needs ffmpeg with libx264.
import { spawn, spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const outDir = path.join(root, "public/tutorial");
const args = process.argv.slice(2);
const option = (name, fallback) => {
  const index = args.indexOf(`--${name}`);
  return index >= 0 ? args[index + 1] : fallback;
};
const base = option("base", "http://localhost:4310");
const only = option("only", null);
const onlyLang = option("lang", null);
const WIDTH = 1280;
const HEIGHT = 760;
const SCALE = 2;
const FPS = 30;

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// ---------------------------------------------------------------- CDP session
async function openBrowser() {
  if (!process.env.CHROME_PATH) throw new Error("Set CHROME_PATH to a Chrome / Chromium executable.");
  const port = 9300 + Math.floor(Math.random() * 600);
  const profile = mkdtempSync(path.join(tmpdir(), "vtx-tutorial-"));
  const chrome = spawn(process.env.CHROME_PATH, [
    "--headless=new", "--no-sandbox", `--remote-debugging-port=${port}`, "--hide-scrollbars",
    "--no-first-run", "--force-color-profile=srgb", `--user-data-dir=${profile}`, "about:blank",
  ], { stdio: "ignore" });
  let target;
  for (let i = 0; i < 100 && !target; i++) {
    try {
      target = (await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()).find((item) => item.type === "page");
    } catch {}
    if (!target) await sleep(100);
  }
  if (!target) throw new Error("Chrome did not start.");
  const ws = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((resolve) => ws.addEventListener("open", resolve));
  let nextId = 0;
  const pending = new Map();
  const listeners = new Set();
  ws.addEventListener("message", (event) => {
    const message = JSON.parse(event.data);
    if (message.id) {
      pending.get(message.id)?.(message);
      pending.delete(message.id);
    } else {
      for (const listener of listeners) listener(message);
    }
  });
  const send = (method, params = {}) => new Promise((resolve, reject) => {
    const id = ++nextId;
    pending.set(id, (message) => (message.error ? reject(new Error(`${method}: ${message.error.message}`)) : resolve(message.result)));
    ws.send(JSON.stringify({ id, method, params }));
  });
  return {
    send,
    on: (listener) => listeners.add(listener),
    off: (listener) => listeners.delete(listener),
    close: () => {
      chrome.kill();
      rmSync(profile, { recursive: true, force: true });
    },
  };
}

// --------------------------------------------------------------- page driver
function driver(cdp) {
  const evaluate = async (expression) => {
    const result = await cdp.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description ?? result.exceptionDetails.text);
    return result.result.value;
  };
  const MODS = { Alt: 1, Ctrl: 2, Meta: 4, Shift: 8 };
  const MOD_KEYS = { Alt: ["Alt", "AltLeft", 18], Ctrl: ["Control", "ControlLeft", 17], Meta: ["Meta", "MetaLeft", 91], Shift: ["Shift", "ShiftLeft", 16] };
  const NAMED = {
    Enter: ["Enter", 13, "\r"], Tab: ["Tab", 9], Escape: ["Escape", 27], Backspace: ["Backspace", 8], End: ["End", 35],
    "→": ["ArrowRight", 39], "←": ["ArrowLeft", 37], "↑": ["ArrowUp", 38], "↓": ["ArrowDown", 40],
  };
  let cursor = { x: WIDTH * 0.62, y: HEIGHT * 0.55 };

  const press = async (key, mods = []) => {
    const modifiers = mods.reduce((sum, mod) => sum | MODS[mod], 0);
    for (const mod of mods) {
      const [k, code, vk] = MOD_KEYS[mod];
      await cdp.send("Input.dispatchKeyEvent", { type: "rawKeyDown", key: k, code, windowsVirtualKeyCode: vk, modifiers });
    }
    let event;
    if (NAMED[key]) {
      const [name, vk, text] = NAMED[key];
      const plain = text && !mods.some((mod) => mod !== "Shift");
      event = { key: name, code: name, windowsVirtualKeyCode: vk, text: plain ? text : undefined };
    } else {
      const upper = key.toUpperCase();
      const code = /^[a-z]$/i.test(key) ? `Key${upper}` : /^\d$/.test(key) ? `Digit${key}` : "";
      event = { key, code, windowsVirtualKeyCode: /^[a-z0-9]$/i.test(key) ? upper.charCodeAt(0) : 0, text: mods.some((mod) => mod !== "Shift") ? undefined : key };
    }
    await cdp.send("Input.dispatchKeyEvent", { type: event.text ? "keyDown" : "rawKeyDown", modifiers, ...event, unmodifiedText: event.text });
    await cdp.send("Input.dispatchKeyEvent", { type: "keyUp", modifiers, key: event.key, code: event.code, windowsVirtualKeyCode: event.windowsVirtualKeyCode });
    for (const mod of [...mods].reverse()) {
      const [k, code, vk] = MOD_KEYS[mod];
      await cdp.send("Input.dispatchKeyEvent", { type: "keyUp", key: k, code, windowsVirtualKeyCode: vk, modifiers: 0 });
    }
  };

  const center = async (selector, { pick = "first", dx = 0.5 } = {}) => {
    const box = await evaluate(`(() => {
      const all = [...document.querySelectorAll(${JSON.stringify(selector)})].filter((el) => el.getClientRects().length);
      const el = ${pick === "last" ? "all.at(-1)" : typeof pick === "number" ? `all[${pick}]` : "all[0]"};
      if (!el) return null;
      const r = el.getBoundingClientRect();
      return { x: r.left + r.width * ${dx}, y: r.top + r.height / 2 };
    })()`);
    if (!box) throw new Error(`Nothing matches ${selector}`);
    return box;
  };

  // A drawn pointer, so clicks are visible in the recording.
  const installPointer = () => evaluate(`(() => {
    const el = document.createElement("div");
    el.id = "tutorial-pointer";
    el.innerHTML = '<svg width="22" height="26" viewBox="0 0 22 26"><path d="M2 2 L2 21 L7 16.5 L10.5 24 L14 22.5 L10.6 15.2 L17.5 15.2 Z" fill="#1d232b" stroke="#fff" stroke-width="1.6" stroke-linejoin="round"/></svg><span></span>';
    el.style.cssText = "position:fixed;left:0;top:0;z-index:2147483647;pointer-events:none;transform:translate(${cursor.x}px,${cursor.y}px);transition:transform 0.55s cubic-bezier(.4,0,.2,1)";
    const ring = el.querySelector("span");
    ring.style.cssText = "position:absolute;left:-13px;top:-13px;width:28px;height:28px;border-radius:50%;background:rgba(31,99,142,.28);transform:scale(0);opacity:0";
    document.documentElement.append(el);
    window.__tutorialPointer = (x, y) => { el.style.transform = "translate(" + x + "px," + y + "px)"; };
    window.__tutorialPointerClick = () => {
      ring.animate([{ transform: "scale(.4)", opacity: 1 }, { transform: "scale(1.5)", opacity: 0 }], { duration: 450, easing: "ease-out" });
    };
  })()`);

  const moveTo = async ({ x, y }) => {
    await evaluate(`window.__tutorialPointer(${x}, ${y})`);
    await cdp.send("Input.dispatchMouseEvent", { type: "mouseMoved", x, y });
    cursor = { x, y };
    await sleep(650);
  };

  const click = async (selector, { button = "left", ...options } = {}) => {
    const point = await center(selector, options);
    await moveTo(point);
    await evaluate("window.__tutorialPointerClick()");
    await cdp.send("Input.dispatchMouseEvent", { type: "mousePressed", ...point, button, clickCount: 1 });
    await cdp.send("Input.dispatchMouseEvent", { type: "mouseReleased", ...point, button, clickCount: 1 });
  };

  return { evaluate, press, click, moveTo, center, installPointer };
}

// ------------------------------------------------------------------ lessons
// Each action: { step, keys | type | click, label?, wait }. `step` is the index
// of the lesson step (src/tutorial/lessons.ts) the action demonstrates.
const field = (n = 0) => `.formula-line:nth-child(${n + 1}) math-field`;
const marker = (n) => `.formula-line:nth-child(${n + 1}) .formula-line-mode-toggle`;
const L = (zh, en) => ({ zh, en });

const LESSONS = {
  typing: {
    actions: [
      { step: 0, click: field(0), dx: 0.05, label: L("点击公式行", "Click the row") },
      { step: 0, type: "x^2" },
      { step: 1, keys: ["→"] },
      { step: 1, type: "+1=" },
      { step: 2, type: "\\frac", wait: 1300 },
      { step: 2, keys: ["Enter"] },
      { step: 3, type: "a" },
      { step: 3, keys: ["Tab"] },
      { step: 3, type: "b", wait: 1800 },
    ],
  },
  lines: {
    setup: [{ click: "[data-classic-bottom-view='source']" }],
    actions: [
      { step: 0, click: field(0), dx: 0.95, label: L("点到行尾", "Click the end of the row") },
      { step: 0, keys: ["End"], silent: true },
      { step: 0, keys: ["Enter"] },
      { step: 0, type: "p=mv" },
      { step: 1, keys: ["Ctrl", "Enter"] },
      { step: 1, type: "a^2" },
      { step: 1, keys: ["→"], silent: true },
      { step: 1, type: "+b^2" },
      { step: 2, keys: ["Alt", "Enter"] },
      { step: 2, type: "F=ma", wait: 900 },
      { step: 3, click: marker(1), label: L("点击 $$", "Click $$"), wait: 1300 },
      { step: 4, moveTo: ".classic-bottom-source, .latex-source-editor, .cm-content", wait: 2400 },
    ],
  },
  multiline: {
    setup: [{ click: "[data-classic-bottom-view='source']" }],
    actions: [
      { step: 0, click: field(0), dx: 0.05, label: L("点击公式行", "Click the row") },
      { step: 0, type: "a=b+c" },
      { step: 0, keys: ["Shift", "Enter"] },
      { step: 0, type: "=d", wait: 1500 },
      { step: 2, click: ".code-format-primary", label: L("LaTeX 格式", "LaTeX format") },
      { step: 2, click: "[data-latex-multiline='align']", label: L("选 align", "Choose align"), wait: 900 },
      { step: 2, keys: ["Escape"], silent: true },
      { step: 3, click: field(0), dx: 0.97, label: L("点到行尾", "Click the end of the row") },
      { step: 3, keys: ["Alt", "Enter"] },
      { step: 3, type: "y&=x^2" },
      { step: 3, keys: ["→"], silent: true },
      { step: 3, type: "+2x+1" },
      { step: 3, keys: ["Shift", "Enter"] },
      { step: 3, type: "&=(x+1)^2", wait: 2600 },
    ],
  },
  tools: {
    actions: [
      { step: 0, click: field(0), dx: 0.05, label: L("点击公式行", "Click the row") },
      { step: 0, click: "[data-command-id='sqrt']", label: L("平方根", "Square root") },
      { step: 0, type: "x" },
      { step: 0, keys: ["→"] },
      { step: 0, type: "+1", wait: 900 },
      { step: 1, keys: ["Enter"] },
      { step: 1, click: "[data-formula-tile-id='mass-energy']", label: L("点击磁贴", "Click a tile"), wait: 1200 },
      { step: 2, click: "[data-command-id='sqrt']", button: "right", label: L("右键", "Right-click"), wait: 2200 },
      { step: 2, keys: ["Escape"], silent: true },
    ],
  },
  copy: {
    setup: [{ click: "[data-classic-bottom-view='source']" }],
    actions: [
      { step: 0, click: "[data-classic-bottom-copy]", label: L("复制", "Copy"), wait: 1600 },
      { step: 2, click: ".code-format-primary", label: L("LaTeX 格式", "LaTeX format") },
      { step: 2, click: "[data-latex-display-wrapper='bracket']", label: L("行间 \\[…\\]", "Display \\[…\\]"), wait: 1000 },
      { step: 2, click: "[data-latex-inline-wrapper='paren']", label: L("行内 \\(…\\)", "Inline \\(…\\)"), wait: 1500 },
      { step: 2, keys: ["Escape"], silent: true },
      { step: 3, click: field(1), dx: 0.3, button: "right", label: L("右键公式", "Right-click a formula"), wait: 2200 },
      { step: 3, keys: ["Escape"], silent: true, wait: 600 },
    ],
  },
};

// ------------------------------------------------------------------- record
async function recordLesson(id, lang) {
  const lesson = LESSONS[id];
  const cdp = await openBrowser();
  const page = driver(cdp);
  const frames = [];
  const frameDir = mkdtempSync(path.join(tmpdir(), `vtx-${id}-`));
  try {
    await cdp.send("Page.enable");
    await cdp.send("Runtime.enable");
    await cdp.send("Emulation.setDeviceMetricsOverride", { width: WIDTH, height: HEIGHT, deviceScaleFactor: SCALE, mobile: false });
    await cdp.send("Browser.grantPermissions", { origin: new URL(base).origin, permissions: ["clipboardReadWrite", "clipboardSanitizedWrite"] });
    await cdp.send("Page.navigate", { url: `${base}/editor?tutorial=${id}&lang=${lang === "en" ? "en" : "cn"}` });
    await sleep(7000);
    for (const action of lesson.setup ?? []) {
      const point = await page.center(action.click);
      await cdp.send("Input.dispatchMouseEvent", { type: "mousePressed", ...point, button: "left", clickCount: 1 });
      await cdp.send("Input.dispatchMouseEvent", { type: "mouseReleased", ...point, button: "left", clickCount: 1 });
      await sleep(800);
    }
    await page.installPointer();
    await sleep(500);

    let index = 0;
    const onFrame = (message) => {
      if (message.method !== "Page.screencastFrame") return;
      const { data, metadata, sessionId } = message.params;
      const file = path.join(frameDir, `${String(index++).padStart(5, "0")}.png`);
      writeFileSync(file, Buffer.from(data, "base64"));
      frames.push({ file, time: metadata.timestamp });
      cdp.send("Page.screencastFrameAck", { sessionId }).catch(() => undefined);
    };
    cdp.on(onFrame);
    const started = Date.now() / 1000;
    await cdp.send("Page.startScreencast", { format: "png", everyNthFrame: 1 });
    await sleep(900);

    const events = [];
    const now = () => Date.now() / 1000 - started;
    for (const action of lesson.actions) {
      const t = now();
      if (action.click) {
        await page.click(action.click, { button: action.button, dx: action.dx, pick: action.pick });
        if (action.label) events.push({ t: now() - 0.15, dur: 1.4, step: action.step, kind: "click", label: action.label });
        // Keep the pointer off the formula being typed.
        if (action.click.includes("math-field") && action.button !== "right") await page.moveTo({ x: WIDTH * 0.42, y: HEIGHT * 0.45 });
      } else if (action.moveTo) {
        await page.moveTo(await page.center(action.moveTo, { dx: 0.3 }));
      } else if (action.keys) {
        const keys = action.keys;
        await page.press(keys.at(-1), keys.slice(0, -1));
        if (!action.silent) events.push({ t, dur: 1.2, step: action.step, kind: "keys", keys });
      } else if (action.type) {
        for (const char of action.type) {
          await page.press(char);
          await sleep(120);
        }
        events.push({ t, dur: now() - t + 0.7, step: action.step, kind: "type", text: action.type });
      }
      await sleep(action.wait ?? 650);
    }
    await sleep(500);
    await cdp.send("Page.stopScreencast");
    cdp.off(onFrame);
    await sleep(300);

    // Frames arrive only when the page repaints; hold each until the next.
    const first = frames[0].time;
    const duration = Math.max(now(), frames.at(-1).time - first + 0.5);
    const list = frames.map((frame, i) => {
      const next = i + 1 < frames.length ? frames[i + 1].time - first : duration;
      return `file '${frame.file}'\nduration ${Math.max(0.001, next - (frame.time - first)).toFixed(4)}`;
    });
    list.push(`file '${frames.at(-1).file}'`);
    const listFile = path.join(frameDir, "list.txt");
    writeFileSync(listFile, list.join("\n"));
    const name = `${id}.${lang}`;
    const mp4 = path.join(outDir, `${name}.mp4`);
    const ffmpeg = spawnSync("ffmpeg", [
      "-y", "-loglevel", "error", "-f", "concat", "-safe", "0", "-i", listFile,
      "-vf", `fps=${FPS},format=yuv420p`, "-c:v", "libx264", "-preset", "slow", "-crf", "24", "-tune", "animation",
      "-movflags", "+faststart", "-an", mp4,
    ], { stdio: "inherit" });
    if (ffmpeg.status !== 0) throw new Error(`ffmpeg failed for ${name}`);
    spawnSync("ffmpeg", ["-y", "-loglevel", "error", "-i", frames.at(-1).file, "-vf", "scale=iw/2:-1", "-q:v", "3", path.join(outDir, `${name}.jpg`)], { stdio: "inherit" });
    const offset = first - started;
    writeFileSync(
      path.join(outDir, `${name}.json`),
      `${JSON.stringify({
        duration: Number(duration.toFixed(2)),
        width: WIDTH,
        height: HEIGHT,
        events: events.map((event) => ({ ...event, t: Number(Math.max(0, event.t - offset).toFixed(2)), dur: Number(event.dur.toFixed(2)) })),
      }, null, 1)}\n`,
    );
    console.log(`${name}: ${frames.length} frames, ${duration.toFixed(1)} s`);
  } finally {
    cdp.close();
    rmSync(frameDir, { recursive: true, force: true });
  }
}

mkdirSync(outDir, { recursive: true });
for (const id of Object.keys(LESSONS)) {
  if (only && only !== id) continue;
  for (const lang of ["zh", "en"]) {
    if (onlyLang && onlyLang !== lang) continue;
    await recordLesson(id, lang);
  }
}
