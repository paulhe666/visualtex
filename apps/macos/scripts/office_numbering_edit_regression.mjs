import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
const offset = process.pid % 500;
const port = 9900 + offset, debugPort = 18000 + offset;
const baseUrl = `http://127.0.0.1:${port}`;
const sessionId = "11111111-2222-4333-8444-555555555555";
const profile = `/private/tmp/visualtex-number-toggle-${process.pid}`;
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
async function waitFor(test) {
  const deadline = Date.now() + 20000;
  do { const value = await test(); if (value) return value; await sleep(50); } while (Date.now() < deadline);
  throw new Error("Timed out waiting for numbering editor regression");
}
const preview = spawn(process.execPath, ["node_modules/vite/bin/vite.js", "preview", "--host", "127.0.0.1", "--port", String(port), "--strictPort"], { stdio: "ignore" });
let chrome, socket;
try {
  await waitFor(async () => { try { return (await fetch(baseUrl)).ok; } catch {} });
  chrome = spawn("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", ["--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check", `--remote-debugging-port=${debugPort}`, `--user-data-dir=${profile}`, baseUrl], { stdio: "ignore" });
  const targets = await waitFor(async () => { try { return await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json(); } catch {} });
  socket = new WebSocket(targets.find(t => t.type === "page").webSocketDebuggerUrl);
  await new Promise(resolve => socket.addEventListener("open", resolve, { once: true }));
  let id = 0; const pending = new Map();
  socket.addEventListener("message", event => { const msg = JSON.parse(event.data); const call = pending.get(msg.id); if (call) { pending.delete(msg.id); msg.error ? call.reject(new Error(msg.error.message)) : call.resolve(msg.result); } });
  const send = (method, params = {}) => new Promise((resolve, reject) => { pending.set(++id, { resolve, reject }); socket.send(JSON.stringify({ id, method, params })); });
  const evaluate = async expression => { const result = await send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true }); if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text); return result.result.value; };
  await send("Runtime.enable");
  await send("Page.enable");
  await send("Page.addScriptToEvaluateOnNewDocument", { source: `(() => {
    const q = new URLSearchParams(location.search);
    const callbacks = new Map(), listeners = new Map(); let serial = 0;
    const probe = window.__numberingProbe = { commands: [], writes: [], commits: [], closes: 0, cancels: 0 };
    let session = {
      id: ${JSON.stringify(sessionId)}, mode: "edit", host: "word", nativeEquation: q.get("native") === "1",
      formulaId: "22222222-3333-4444-8555-666666666666", sourceDocumentId: "probe", sourceObjectId: "probe",
      title: "Numbering edit regression", lines: [{ id: "line-1", latex: "x=\\\\frac{1}{2}" }], activeLineId: "line-1",
      codeFormat: "raw", displayMode: "block", inlineImageMathStyle: "text", numbered: q.get("initial") === "1",
      fontSizePt: 14, exportWidth: 0, exportHeight: 0, exportResult: null, originalMetadata: null,
      dirty: false, status: "editing", autoCommitOnClose: q.get("close") !== "cancel", explicitCancel: false,
      error: null, createdAt: Date.now(), updatedAt: Date.now(), expiresAt: Date.now() + 600000
    };
    localStorage.setItem("visualtex-office-word-create-numbered", "false");
    window.__TAURI_EVENT_PLUGIN_INTERNALS__ = { unregisterListener: (event, id) => { listeners.delete(id); } };
    window.__TAURI_INTERNALS__ = {
      metadata: { currentWindow: { label: "office-editor" }, currentWebview: { label: "office-editor" } },
      transformCallback: callback => { callbacks.set(++serial, callback); return serial; },
      invoke: async (command, args = {}) => {
        probe.commands.push(command);
        if (command === "plugin:event|listen") { listeners.set(++serial, args); return serial; }
        if (command === "plugin:event|unlisten") return;
        if (command === "get_macos_offline_office_editor_activation") return { sessionId: session.id, host: "word", generation: 1, receivedEpochMs: Date.now() };
        if (command === "get_macos_offline_office_session") return { ...session };
        if (command === "update_macos_offline_office_session") {
          // Force an old draft to remain in flight when Apply or Close fires.
          await new Promise(resolve => setTimeout(resolve, 180));
          session = { ...session, ...args.patch }; probe.writes.push({ status: session.status, numbered: session.numbered }); return { ...session };
        }
        if (command === "commit_macos_offline_office_session") {
          session = { ...session, ...args.patch, status: "completed" };
          probe.writes.push({ status: "completed", numbered: session.numbered });
          probe.commits.push({ patch: args.patch, started: args.applyStartedEpochMs, finished: Date.now() }); return { ...session };
        }
        if (command === "cancel_macos_offline_office_session") { probe.cancels++; session.status = "cancelled"; return { ...session }; }
        if (command === "close_macos_offline_office_editor_window") { probe.closes++; return; }
        return null;
      }
    };
    probe.close = () => { for (const [id, listener] of listeners) { if (listener.event === "tauri://close-requested") callbacks.get(listener.handler)({ id, event: listener.event, payload: null }); } };
  })();` });
  for (const native of [false, true]) for (const [initial, close] of [[false, "apply"], [true, "apply"], [true, "settled"], [false, "commit"], [true, "cancel"]]) {
    await send("Page.navigate", { url: `${baseUrl}/office-native-dialog.html?transport=tauri&officeHost=word&native=${+native}&initial=${+initial}&close=${close}` });
    await waitFor(() => evaluate("Boolean(document.querySelector('[data-office-numbered]') && document.querySelector('math-field'))")).catch(async error => { console.error(await evaluate("({text: document.body.innerText, probe: window.__numberingProbe})")); throw error; });
    const enabled = await evaluate("!document.querySelector('[data-office-numbered]').disabled"); assert.equal(enabled, true);
    // Each click yields to React. The last click and Apply/Close are adjacent,
    // while prior saves and PNG work are still outstanding.
    await evaluate(`(async () => {
      for (let i = 0; i < 3; i++) { document.querySelector('[data-office-numbered]').click(); await new Promise(resolve => setTimeout(resolve, 0)); }
      ${close === "settled" ? "await new Promise(resolve => setTimeout(resolve, 1000));" : ""}
      ${["apply", "settled"].includes(close) ? "document.querySelector('[data-office-primary-action]').click();" : "window.__numberingProbe.close();"}
    })()`);
    await waitFor(() => evaluate("window.__numberingProbe.closes"));
    await sleep(600);
    const result = await evaluate("JSON.parse(JSON.stringify(window.__numberingProbe))");
    if (close === "cancel") {
      assert.equal(result.commits.length, 0); assert.equal(result.cancels, 1);
    } else {
      assert.equal(result.commits.length, 1, "Rapid toggle must commit exactly once");
      const patch = result.commits[0].patch;
      assert.equal(patch.numbered, !initial); assert.equal(patch.dirty, true);
      assert.ok(patch.exportResult.pngBase64 && patch.exportResult.ommlDocxBase64, "Final commit must await complete current exports");
      const { unzipSync, strFromU8 } = await import("fflate");
      const omml = strFromU8(unzipSync(Buffer.from(patch.exportResult.ommlDocxBase64, "base64url"))["word/document.xml"]);
      assert.equal(omml.includes("<m:eqArr>"), !initial, "OMML numbering shell must match the final checkbox");
      assert.deepEqual(result.writes.at(-1), { status: "completed", numbered: !initial }, "No stale draft may overwrite the applied numbering");
    }
    console.log(`PASS ${native ? "OMML" : "image"}: ${initial ? "disable" : "enable"} numbering, rapid toggle + ${close}, delayed drafts`);
  }
} finally {
  socket?.close(); chrome?.kill(); preview.kill();
  await sleep(150); await rm(profile, { recursive: true, force: true });
}
