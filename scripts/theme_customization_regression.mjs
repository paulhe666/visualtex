import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 650;
const previewPort = 18600 + offset;
const debugPort = 24400 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const officeUrl = `${baseUrl}/office-native-dialog.html?session=00000000-0000-4000-8000-000000000000&theme=codex`;
const chromeProfile = `/tmp/visualtex-theme-customization-${process.pid}`;
const chromePath = (process.env.CHROME_PATH || "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome");
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while Vite or Chrome starts.
    }
    await sleep(80);
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

  async evaluate(expression) {
    const result = await this.send("Runtime.evaluate", {
      expression,
      awaitPromise: true,
      returnByValue: true,
    });
    if (result.exceptionDetails) {
      throw new Error(
        result.exceptionDetails.exception?.description ||
          result.exceptionDetails.text ||
          "Runtime.evaluate failed",
      );
    }
    return result.result.value;
  }

  close() {
    this.socket?.close();
  }
}

async function pageTarget(predicate) {
  const targets = await (
    await fetch(`http://127.0.0.1:${debugPort}/json/list`)
  ).json();
  return targets.find((target) => target.type === "page" && predicate(target.url));
}

async function connectPage(target) {
  const client = new CdpClient(target.webSocketDebuggerUrl);
  await client.connect();
  await client.send("Runtime.enable");
  await client.send("Page.enable");
  return client;
}

async function waitForExpression(client, expression, timeoutMs = 10000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    const value = await client.evaluate(expression);
    if (value) return value;
    await sleep(60);
  }
  throw new Error(`Timed out waiting for expression: ${expression}`);
}

async function main() {
  const preview = spawn(
    process.execPath,
    [
      "node_modules/vite/bin/vite.js",
      "preview",
      "--host",
      "127.0.0.1",
      "--port",
      String(previewPort),
      "--strictPort",
    ],
    { cwd: process.cwd(), stdio: "ignore" },
  );
  let chrome;
  let mainClient;
  let officeClient;

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
        "--window-size=1400,1100",
        baseUrl,
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);

    const mainTarget = await waitForExpression(
      {
        evaluate: async () => pageTarget((url) => url === `${baseUrl}/` || url === baseUrl),
      },
      "true",
    ).catch(() => undefined);
    const resolvedMainTarget =
      mainTarget || (await pageTarget((url) => url === `${baseUrl}/` || url === baseUrl));
    assert.ok(resolvedMainTarget, "Main VisualTeX target must exist");
    mainClient = await connectPage(resolvedMainTarget);

    await waitForExpression(mainClient, "Boolean(document.querySelector('.settings-toggle'))");
    await mainClient.evaluate(`(() => {
      document.querySelector('.settings-toggle')?.click();
      return true;
    })()`);
    await waitForExpression(
      mainClient,
      "Boolean(document.querySelector('[data-interface-customization-trigger]'))",
    );
    await mainClient.evaluate(`(() => {
      document.querySelector('[data-interface-customization-trigger]')?.click();
      return true;
    })()`);
    await waitForExpression(mainClient, "Boolean(document.querySelector('[data-theme-studio]'))");

    const presetCount = await mainClient.evaluate(
      "document.querySelectorAll('[data-theme-preset]').length",
    );
    assert.equal(presetCount, 15, "Theme studio must expose all built-in themes as peers");

    const expectedBuiltInThemes = [
      ["codex", "#339cff", "#ffffff", "light"],
      ["notion", "#4981d2", "#ffffff", "light"],
      ["one", "#586ef6", "#fafafa", "light"],
      ["proof", "#4b745f", "#f5f3ee", "light"],
      ["rose-pine", "#cb8681", "#f9f4ee", "light"],
      ["solarized", "#ae8b2d", "#fcf6e5", "light"],
      ["vercel", "#2d69f6", "#fafafa", "light"],
      ["vscode-plus", "#3478c6", "#f3f3f3", "light"],
      ["xcode", "#0f0ef5", "#f5f5f5", "light"],
      ["raycast", "#ed6e69", "#ffffff", "light"],
    ];
    for (const [themeId, expectedAccent, expectedBackground, expectedMode] of expectedBuiltInThemes) {
      await mainClient.evaluate(`(() => {
        document.querySelector('[data-theme-preset="${themeId}"]')?.click();
        return true;
      })()`);
      await sleep(70);
      const applied = await mainClient.evaluate(`(() => ({
        theme: document.documentElement.dataset.theme,
        palette: document.documentElement.dataset.themePalette,
        accent: getComputedStyle(document.documentElement).getPropertyValue('--accent').trim().toLowerCase(),
        background: getComputedStyle(document.documentElement).getPropertyValue('--bg').trim().toLowerCase(),
        mode: getComputedStyle(document.documentElement).colorScheme,
        tile: getComputedStyle(document.documentElement).getPropertyValue('--tile-bg').trim().toLowerCase(),
        syntaxKeyword: getComputedStyle(document.documentElement).getPropertyValue('--syntax-keyword').trim().toLowerCase(),
        customStored: localStorage.getItem('visualtex.custom-theme.v1'),
      }))()`);
      assert.equal(applied.theme, themeId);
      assert.equal(applied.palette, themeId);
      assert.equal(applied.accent, expectedAccent);
      assert.equal(applied.background, expectedBackground);
      assert.equal(applied.mode, expectedMode);
      assert.ok(applied.tile, `${themeId} must define tile colors`);
      assert.ok(applied.syntaxKeyword, `${themeId} must define syntax colors`);
      assert.equal(applied.customStored, null, "Selecting a preset must not create a customization overlay");
    }

    await mainClient.evaluate(`(() => {
      document.querySelector('[data-theme-preset="proof"]')?.click();
      return true;
    })()`);
    await sleep(90);
    const proofPurity = await mainClient.evaluate(`(() => {
      const rootStyle = getComputedStyle(document.documentElement);
      const backgroundOf = (selector) => {
        const element = document.querySelector(selector);
        return element ? getComputedStyle(element).backgroundColor.toLowerCase() : null;
      };
      return {
        surface: rootStyle.getPropertyValue('--surface').trim().toLowerCase(),
        formulaSurface: rootStyle.getPropertyValue('--formula-surface').trim().toLowerCase(),
        paper: rootStyle.getPropertyValue('--bg-paper').trim().toLowerCase(),
        input: rootStyle.getPropertyValue('--input-bg').trim().toLowerCase(),
        tile: rootStyle.getPropertyValue('--tile-bg').trim().toLowerCase(),
        menu: rootStyle.getPropertyValue('--menu-bg').trim().toLowerCase(),
        settings: backgroundOf('.settings-dialog'),
        toolbar: backgroundOf('.formula-toolbar'),
        editor: backgroundOf('.editor-surface'),
        subdialog: backgroundOf('.settings-subdialog'),
        template: backgroundOf('.template-button'),
      };
    })()`);
    assert.equal(proofPurity.surface, '#f5f3ee');
    assert.equal(proofPurity.formulaSurface, '#f5f3ee');
    assert.equal(proofPurity.paper, '#f5f3ee');
    assert.equal(proofPurity.input, '#f5f3ee');
    assert.equal(proofPurity.tile, '#f5f3ee');
    assert.equal(proofPurity.menu, '#f5f3ee');
    const forbiddenDefaultBackgrounds = new Set([
      'rgb(255, 255, 255)',
      'rgb(242, 244, 246)',
      'rgb(247, 248, 250)',
    ]);
    for (const [area, value] of Object.entries({
      settings: proofPurity.settings,
      toolbar: proofPurity.toolbar,
      editor: proofPurity.editor,
      subdialog: proofPurity.subdialog,
      template: proofPurity.template,
    })) {
      assert.ok(value, `Proof ${area} must exist for purity validation`);
      assert.ok(
        !forbiddenDefaultBackgrounds.has(value),
        `Proof ${area} must not fall back to a VisualTeX default background: ${value}`,
      );
    }

    await mainClient.evaluate(`(() => {
      document.querySelector('[data-theme-preset="vscode-plus"]')?.click();
      return true;
    })()`);
    await sleep(70);
    const vscodeSyntax = await mainClient.evaluate(`(() => {
      const style = getComputedStyle(document.documentElement);
      return {
        keyword: style.getPropertyValue('--syntax-keyword').trim().toLowerCase(),
        fn: style.getPropertyValue('--syntax-function').trim().toLowerCase(),
        string: style.getPropertyValue('--syntax-string').trim().toLowerCase(),
      };
    })()`);
    assert.deepEqual(vscodeSyntax, {
      keyword: "#af00db",
      fn: "#795e26",
      string: "#a31515",
    });

    await mainClient.evaluate(`(() => {
      document.querySelector('[data-theme-preset="raycast"]')?.click();
      return true;
    })()`);
    await sleep(70);

    await mainClient.evaluate(`(() => {
      const input = document.querySelector('[data-theme-color-setting="accent"]');
      if (!(input instanceof HTMLInputElement)) return false;
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
      setter?.call(input, '#123456');
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    })()`);
    await sleep(120);
    const customState = await mainClient.evaluate(`(() => ({
      theme: document.documentElement.dataset.theme,
      accent: getComputedStyle(document.documentElement).getPropertyValue('--accent').trim().toLowerCase(),
      stored: JSON.parse(localStorage.getItem('visualtex.custom-theme.v1') || 'null'),
    }))()`);
    assert.equal(customState.theme, "custom", "Editing a color must switch to the peer Custom theme");
    assert.equal(customState.accent, "#123456", "Color picker changes must apply immediately");
    assert.equal(customState.stored?.colors?.accent?.toLowerCase(), "#123456");

    await mainClient.evaluate(`(() => {
      document.querySelector('[data-interface-customization-close]')?.click();
      return true;
    })()`);
    await waitForExpression(mainClient, "!document.querySelector('[data-theme-studio]')");
    await mainClient.evaluate(`(() => {
      document.querySelector('[data-theme-choice="beige"]')?.click();
      return true;
    })()`);
    await sleep(120);
    const baseTheme = await mainClient.evaluate(`(() => ({
      theme: document.documentElement.dataset.theme,
      accent: getComputedStyle(document.documentElement).getPropertyValue('--accent').trim().toLowerCase(),
      storedCustomAccent: JSON.parse(localStorage.getItem('visualtex.custom-theme.v1') || 'null')?.colors?.accent?.toLowerCase(),
    }))()`);
    assert.equal(baseTheme.theme, "beige");
    assert.equal(baseTheme.accent, "#785536");
    assert.equal(baseTheme.storedCustomAccent, "#123456", "Built-in themes must ignore, but preserve, Custom colors");

    await mainClient.evaluate(`(() => {
      document.querySelector('.settings-toggle')?.click();
      return true;
    })()`);
    await waitForExpression(
      mainClient,
      "Boolean(document.querySelector('[data-interface-customization-trigger]'))",
    );
    await mainClient.evaluate(`(() => {
      document.querySelector('[data-interface-customization-trigger]')?.click();
      return true;
    })()`);
    await waitForExpression(
      mainClient,
      `Boolean(document.querySelector('[data-theme-preset="codex"]'))`,
    );
    await mainClient.evaluate(`(() => {
      document.querySelector('[data-theme-preset="codex"]')?.click();
      return true;
    })()`);
    await sleep(120);

    await mainClient.send("Target.createTarget", { url: officeUrl });
    let officeTarget;
    for (let attempt = 0; attempt < 80 && !officeTarget; attempt += 1) {
      officeTarget = await pageTarget((url) => url.startsWith(officeUrl));
      if (!officeTarget) await sleep(50);
    }
    assert.ok(officeTarget, "Office formula editor target must exist");
    officeClient = await connectPage(officeTarget);
    await sleep(250);
    const officeInherited = await officeClient.evaluate(`(() => ({
      theme: document.documentElement.dataset.theme,
      accent: getComputedStyle(document.documentElement).getPropertyValue('--accent').trim().toLowerCase(),
      background: getComputedStyle(document.documentElement).getPropertyValue('--bg').trim().toLowerCase(),
    }))()`);
    assert.deepEqual(officeInherited, {
      theme: "codex",
      accent: "#339cff",
      background: "#ffffff",
    });

    await mainClient.evaluate(`(() => {
      document.querySelector('[data-theme-preset="rose-pine"]')?.click();
      return true;
    })()`);
    await sleep(180);
    const officeSynchronized = await officeClient.evaluate(`(() => ({
      theme: document.documentElement.dataset.theme,
      accent: getComputedStyle(document.documentElement).getPropertyValue('--accent').trim().toLowerCase(),
      background: getComputedStyle(document.documentElement).getPropertyValue('--bg').trim().toLowerCase(),
    }))()`);
    assert.deepEqual(officeSynchronized, {
      theme: "rose-pine",
      accent: "#cb8681",
      background: "#f9f4ee",
    });

    await mainClient.evaluate(`(() => {
      document.querySelector('[data-theme-use-custom]')?.click();
      return true;
    })()`);
    await sleep(180);
    const officeCustom = await officeClient.evaluate(`(() => ({
      theme: document.documentElement.dataset.theme,
      accent: getComputedStyle(document.documentElement).getPropertyValue('--accent').trim().toLowerCase(),
    }))()`);
    assert.deepEqual(officeCustom, {
      theme: "custom",
      accent: "#123456",
    });

    process.stdout.write("Peer theme switching, Custom theme editing, persistence, and Office synchronization regression passed\n");
  } finally {
    mainClient?.close();
    officeClient?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(240);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => undefined);
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
