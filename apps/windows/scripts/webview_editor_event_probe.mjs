const debugPort = Number(process.env.VISUALTEX_WEBVIEW_DEBUG_PORT || 17777);
const action = process.argv[2] || "read";

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

  close() {
    this.socket?.close();
  }
}

const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
const page = targets.find((target) => target.type === "page" && target.url.includes("localhost:1420"));
if (!page) throw new Error("VisualTeX WebView2 debug target not found");

const client = new CdpClient(page.webSocketDebuggerUrl);
await client.connect();
await client.send("Runtime.enable");

const evaluate = async (expression) => {
  const result = await client.send("Runtime.evaluate", {
    expression,
    awaitPromise: true,
    returnByValue: true,
  });
  if (result.exceptionDetails) {
    throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
  }
  return result.result.value;
};

if (action === "install") {
  await evaluate(`(() => {
    window.__visualTexEditorProbeEvents = [];
    if (window.__visualTexEditorProbeInstalled) return;
    window.__visualTexEditorProbeInstalled = true;
    const snapshot = (event) => {
      const field = document.querySelector("math-field");
      let persisted = null;
      try { persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "null"); } catch {}
      window.__visualTexEditorProbeEvents.push({
        type: event.type,
        key: event.key ?? null,
        code: event.code ?? null,
        inputType: event.inputType ?? null,
        data: event.data ?? null,
        defaultPrevented: event.defaultPrevented,
        mode: field?.mode ?? null,
        value: field?.value ?? null,
        raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
          .filter((node) => !node.classList.contains("ML__suggestion"))
          .map((node) => node.textContent ?? "").join(""),
        position: field?.position ?? null,
        selection: field?.selection ?? null,
        pendingNativeSuggestion: field?.dataset.pendingNativeSuggestion ?? "",
        source: [...document.querySelectorAll(".source-panel .cm-line")]
          .map((line) => line.textContent ?? ""),
        store: persisted?.state?.lines ?? [],
      });
    };
    for (const type of ["keydown", "keyup", "beforeinput", "input", "compositionstart", "compositionend"])
      window.addEventListener(type, snapshot, true);
  })()`);
  console.log("VisualTeX WebView2 editor event probe installed");
} else {
  const result = await evaluate(`(() => {
    const field = document.querySelector("math-field");
    let persisted = null;
    try { persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "null"); } catch {}
    return {
      events: ${action === "state" ? "[]" : "window.__visualTexEditorProbeEvents ?? []"},
      field: field ? {
        mode: field.mode,
        value: field.value,
        bounds: (() => {
          const rect = field.getBoundingClientRect();
          return { left: rect.left, top: rect.top, width: rect.width, height: rect.height };
        })(),
        raw: [...(field.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
          .filter((node) => !node.classList.contains("ML__suggestion"))
          .map((node) => node.textContent ?? "").join(""),
        position: field.position,
        selection: field.selection,
        pendingNativeSuggestion: field.dataset.pendingNativeSuggestion ?? "",
      } : null,
      source: [...document.querySelectorAll(".source-panel .cm-line")]
        .map((line) => line.textContent ?? ""),
      store: persisted?.state?.lines ?? [],
      inputBehavior: persisted?.state?.inputBehavior ?? null,
      nativeCandidates: [...document.querySelectorAll("#visualtex-native-input-suggestion-popover li[data-command]")]
        .map((item) => ({ command: item.dataset.command ?? "", current: item.classList.contains("ML__popover__current") })),
      visualTexCandidateVisible: Boolean(document.querySelector(".suggestion-popup")),
    };
  })()`);
  console.log(JSON.stringify(result, null, 2));
}

client.close();
