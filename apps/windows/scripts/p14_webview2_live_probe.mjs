import assert from "node:assert/strict";

const debugPort = Number(process.argv[2] ?? "9223");
const focusCore = process.argv.includes("--focus-core");
const probeSystemFont = process.argv.includes("--probe-system-font");
if (!Number.isInteger(debugPort) || debugPort < 1 || debugPort > 65535) {
  throw new Error("Usage: node p14_webview2_live_probe.mjs [debugPort] [--focus-core] [--probe-system-font]");
}

const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
const target = targets.find((item) =>
  item.type === "page" && typeof item.url === "string" && item.url.includes("runtime=vsto-desktop"));
assert.ok(target, "No real Office vsto-desktop WebView2 target is exposed on the requested debug port.");

class Client {
  id = 0;
  pending = new Map();
  events = [];
  async connect(url) {
    this.ws = new WebSocket(url);
    await new Promise((resolve, reject) => {
      this.ws.addEventListener("open", resolve, { once: true });
      this.ws.addEventListener("error", reject, { once: true });
    });
    this.ws.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      if (message.id) {
        const pending = this.pending.get(message.id);
        if (!pending) return;
        this.pending.delete(message.id);
        message.error ? pending.reject(new Error(message.error.message)) : pending.resolve(message.result);
        return;
      }
      this.events.push(message);
    });
  }
  send(method, params = {}) {
    const id = ++this.id;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.ws.send(JSON.stringify({ id, method, params }));
    });
  }
  async evaluate(expression) {
    const response = await this.send("Runtime.evaluate", {
      expression,
      awaitPromise: true,
      returnByValue: true,
    });
    if (response.exceptionDetails) {
      throw new Error(response.exceptionDetails.exception?.description ?? response.exceptionDetails.text);
    }
    return response.result.value;
  }
  close() { this.ws?.close(); }
}

const client = new Client();
await client.connect(target.webSocketDebuggerUrl);
try {
  await client.send("Runtime.enable");
  await client.send("Log.enable");
  if (focusCore) {
    const focused = await client.evaluate(`(() => {
      const input = document.querySelector('.vt-core-input');
      if (!input) return false;
      input.focus();
      return document.activeElement === input;
    })()`);
    assert.equal(focused, true, "Real WebView2 core input could not receive focus.");
  }
  const systemFont = probeSystemFont
    ? await client.evaluate(`window.__TAURI_INTERNALS__?.invoke?.('read_system_font_asset', { fontFamilies: ['Cambria Math'] }).then((asset) => ({
        requestedFamily: asset.requestedFamily,
        resolvedFamily: asset.resolvedFamily,
        faceIndex: asset.faceIndex,
        sha256: asset.sha256,
        byteLength: asset.byteLength,
        hasMathTable: asset.hasMathTable,
      }))`)
    : null;
  const state = await client.evaluate(`(() => ({
    readyState: document.readyState,
    title: document.title,
    location: location.href,
    rootChildren: document.querySelector('#root')?.children.length ?? -1,
    rootHtml: document.querySelector('#root')?.innerHTML.slice(0, 4000) ?? '',
    bodyText: document.body?.innerText ?? '',
    coreEditors: document.querySelectorAll('[data-engine="visualtex-core"]').length,
    coreInputs: document.querySelectorAll('.vt-core-input').length,
    mathFields: document.querySelectorAll('math-field').length,
    officeState: document.querySelector('.office-dialog-state')?.textContent ?? null,
    primaryActions: [...document.querySelectorAll('[data-office-primary-action]')].map((el) => ({
      text: el.textContent,
      disabled: el.disabled,
    })),
    scripts: [...document.scripts].map((script) => script.src || '[inline]'),
    resources: [...performance.getEntriesByType('resource')]
      .map((entry) => entry.name)
      .filter((name) => /visualtex_engine_wasm|latinmodern-math|CoreMathEditor|dialog-/.test(name)),
  }))()`);
  const logs = client.events
    .filter((event) => event.method === "Runtime.exceptionThrown" || event.method === "Log.entryAdded" || event.method === "Runtime.consoleAPICalled")
    .slice(-50);
  console.log(JSON.stringify({
    status: "PASS",
    target: { id: target.id, title: target.title, url: target.url },
    state,
    systemFont,
    focusedCoreInput: focusCore,
    observedEvents: logs,
  }, null, 2));
} finally {
  client.close();
}
