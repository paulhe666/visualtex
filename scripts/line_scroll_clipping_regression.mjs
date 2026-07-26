import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 700;
const previewPort = 9700 + offset;
const debugPort = 15700 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}/editor`;
const profile = `/tmp/visualtex-line-layout-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeout = 15000) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    try {
      if ((await fetch(url)).ok) return;
    } catch {}
    await sleep(70);
  }
  throw new Error(`Timeout waiting for ${url}`);
}

class Cdp {
  constructor(url) {
    this.url = url;
    this.id = 1;
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
    const id = this.id++;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }
  close() {
    this.socket?.close();
  }
}

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
let client;

try {
  await waitFor(baseUrl);
  chrome = spawn(
    chromePath,
    [
      "--headless=new",
      "--disable-gpu",
      "--no-first-run",
      `--remote-debugging-port=${debugPort}`,
      `--user-data-dir=${profile}`,
      "--window-size=1400,900",
      baseUrl,
    ],
    { stdio: "ignore" },
  );
  await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
  const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
  const page = targets.find((target) => target.type === "page" && target.url.startsWith(baseUrl));
  assert.ok(page, "VisualTeX page target missing");
  client = new Cdp(page.webSocketDebuggerUrl);
  await client.connect();
  await client.send("Runtime.enable");
  await client.send("Page.enable");

  const evaluate = async (expression) => {
    const result = await client.send("Runtime.evaluate", {
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
  };

  await evaluate(`(() => {
    for (const key of [
      "visualtex.onboarding.v3.completed",
      "visualtex.office.macos.first-run.v1.completed",
      "visualtex.onboarding.macos.desktop.v1.2.0.completed",
      "visualtex.office.macos.native-first-run.v1.2.0.completed",
    ]) localStorage.setItem(key, "true");
  })()`);

  let caseId = 0;
  const load = async (lines, activeIndex = 0, zoom = 1) => {
    caseId += 1;
    const ids = lines.map((_, index) => `layout-${caseId}-${index}`);
    await evaluate(`(() => {
      const key = "visualtex-editor";
      const persisted = JSON.parse(localStorage.getItem(key) || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        title: "layout regression",
        lines: ${JSON.stringify(lines)}.map((latex, index) => ({
          id: ${JSON.stringify(ids)}[index],
          latex,
        })),
        activeLineId: ${JSON.stringify(ids[activeIndex])},
        sourceOpen: false,
        zoom: ${JSON.stringify(zoom)},
        inputBehavior: {
          autoExitSuperscript: true,
          autoExitSubscript: true,
          autoExitAccent: true,
          autoExitWrapperCommand: true,
          showStructuredCommandSuggestions: true,
          showOtherCommandSuggestions: true,
        },
      };
      localStorage.setItem(key, JSON.stringify(persisted));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(500);
    await evaluate(`new Promise((resolve) => {
      const poll = () => {
        const fields = document.querySelectorAll("math-field");
        if (fields.length !== ${lines.length}) return setTimeout(poll, 25);
        resolve(true);
      };
      poll();
    })`);
    return ids;
  };

  const focus = async (index, prefix = null) => {
    return evaluate(`(() => {
      const field = document.querySelectorAll("math-field")[${index}];
      field.focus({ preventScroll: true });
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      let position = field.lastOffset;
      if (${JSON.stringify(prefix)} !== null) {
        const target = ${JSON.stringify(prefix)}.replace(/\\s+/g, "");
        const candidates = [];
        for (let offset = 0; offset <= field.lastOffset; offset += 1) {
          const serialized = field
            .getValue(0, offset, "latex")
            .replace(/\\s+/g, "")
            .replace(/\\{([A-Za-z0-9])\\}/g, "$1");
          candidates.push({ offset, serialized });
        }
        const match = candidates.find((item) => item.serialized === target);
        if (!match) return { value: field.value, position: -1, lastOffset: field.lastOffset, candidates };
        position = match.offset;
      }
      field.selection = { ranges: [[position, position]], direction: "none" };
      field.position = position;
      return { value: field.value, position, lastOffset: field.lastOffset };
    })()`);
  };

  const dispatchKey = async (index, key, code) => {
    await evaluate(`(() => {
      const field = document.querySelectorAll("math-field")[${index}];
      const event = new KeyboardEvent("keydown", {
        key: ${JSON.stringify(key)},
        code: ${JSON.stringify(code)},
        bubbles: true,
        composed: true,
        cancelable: true,
      });
      field.dispatchEvent(event);
      return event.defaultPrevented;
    })()`);
  };

  const probeOperation = async ({ operation, anchorId = null }) => {
    return evaluate(`new Promise((resolve) => {
      const scroll = document.querySelector(".editor-pane-scroll");
      const samples = [];
      const originalRows = Array.from(document.querySelectorAll(".formula-line"));
      const originalFields = Array.from(document.querySelectorAll("math-field"));
      const nodeIds = new WeakMap();
      let nextNodeId = 1;
      const nodeId = (node) => {
        if (!node) return null;
        if (!nodeIds.has(node)) nodeIds.set(node, nextNodeId++);
        return nodeIds.get(node);
      };
      let createdRow = null;
      let createdField = null;
      const read = (label) => {
        const rows = Array.from(document.querySelectorAll(".formula-line"));
        const fields = Array.from(document.querySelectorAll("math-field"));
        if (!createdRow && rows.length > originalRows.length) {
          createdRow = rows.find((row) => !originalRows.includes(row)) || null;
          createdField = fields.find((field) => !originalFields.includes(field)) || null;
        }
        const anchorId = ${JSON.stringify(anchorId)};
        const anchor = anchorId
          ? document.querySelector('[data-line-id="' + anchorId + '"]')
          : null;
        const activeField = rows
          .find((row) => row.classList.contains("is-active"))
          ?.querySelector("math-field") ?? null;
        const deep = (() => {
          let current = document.activeElement;
          while (current?.shadowRoot?.activeElement) current = current.shadowRoot.activeElement;
          return current;
        })();
        const caretNodes = Array.from(
          activeField?.shadowRoot?.querySelectorAll(
            ".ML__caret, .ML__text-caret, .ML__latex-caret, .visualtex-structural-placeholder-caret"
          ) ?? [],
        ).map((node) => {
          const rect = node.getBoundingClientRect();
          const style = getComputedStyle(node);
          const pseudo = getComputedStyle(node, "::after");
          return {
            nodeId: nodeId(node),
            className: node.className,
            width: rect.width,
            height: rect.height,
            display: style.display,
            visibility: style.visibility,
            opacity: style.opacity,
            pseudoVisibility: pseudo.visibility,
            pseudoOpacity: pseudo.opacity,
            pseudoAnimation: pseudo.animationName,
          };
        });
        samples.push({
          label,
          scrollTop: scroll?.scrollTop ?? 0,
          scrollHeight: scroll?.scrollHeight ?? 0,
          clientHeight: scroll?.clientHeight ?? 0,
          anchorTop: anchor?.getBoundingClientRect().top ?? null,
          rowCount: rows.length,
          rowIds: rows.map((row) => row.dataset.lineId),
          rowTops: rows.map((row) => row.getBoundingClientRect().top),
          rowHeights: rows.map((row) => row.getBoundingClientRect().height),
          fieldHeights: fields.map((field) => field.getBoundingClientRect().height),
          activeRows: rows.filter((row) => row.classList.contains("is-active")).map((row) => row.dataset.lineId),
          backgrounds: rows.map((row) => getComputedStyle(row).backgroundColor),
          createdRowStable: !createdRow || (createdRow.isConnected && rows.includes(createdRow)),
          createdFieldStable: !createdField || (createdField.isConnected && fields.includes(createdField)),
          deepTag: deep?.tagName || "",
          deepPart: deep?.getAttribute?.("part") || "",
          deepNodeId: nodeId(deep),
          activeFieldId: activeField?.closest(".formula-line")?.dataset.lineId ?? null,
          activeFieldNodeId: nodeId(activeField),
          activeFieldPosition: activeField?.position ?? null,
          activeFieldSelection: activeField?.selection ?? null,
          activeFieldFocused: activeField === document.activeElement || activeField?.matches(":focus-within") || false,
          caretNodes,
        });
      };
      read("before");
      ${operation}
      queueMicrotask(() => read("microtask"));
      requestAnimationFrame(() => {
        read("raf1");
        requestAnimationFrame(() => {
          read("raf2");
          requestAnimationFrame(() => read("raf3"));
        });
      });
      setTimeout(() => read("40ms"), 40);
      setTimeout(() => read("90ms"), 90);
      setTimeout(() => read("180ms"), 180);
      setTimeout(() => resolve(samples), 230);
    })`);
  };

  const sampleIntegral = async (operation) => {
    return evaluate(`new Promise((resolve) => {
      const samples = [];
      const read = (label) => {
        const fields = Array.from(document.querySelectorAll("math-field"));
        const field = fields[1] || null;
        const scroll = document.querySelector(".editor-pane-scroll");
        const viewport = scroll?.getBoundingClientRect();
        const fieldRect = field?.getBoundingClientRect();
        const operators = Array.from(field?.shadowRoot?.querySelectorAll(
          ".ML__op-group, .ML__mop, .ML__vlist, [data-atom-id]"
        ) || []).map((node) => node.getBoundingClientRect()).filter((rect) => rect.height > 0);
        const top = operators.length ? Math.min(...operators.map((rect) => rect.top)) : null;
        const bottom = operators.length ? Math.max(...operators.map((rect) => rect.bottom)) : null;
        samples.push({
          label,
          values: fields.map((item) => item.value),
          fieldHeight: fieldRect?.height ?? 0,
          fieldTop: fieldRect?.top ?? null,
          fieldBottom: fieldRect?.bottom ?? null,
          operatorTop: top,
          operatorBottom: bottom,
          clippedTop: top !== null && fieldRect ? top < fieldRect.top - 0.5 : null,
          clippedBottom: bottom !== null && fieldRect ? bottom > fieldRect.bottom + 0.5 : null,
          overflow: field ? getComputedStyle(field).overflow : "",
          rowHeight: field?.closest(".formula-line")?.getBoundingClientRect().height ?? 0,
          scrollTop: scroll?.scrollTop ?? 0,
          scrollHeight: scroll?.scrollHeight ?? 0,
          viewportTop: viewport?.top ?? null,
          viewportBottom: viewport?.bottom ?? null,
          viewportClippedTop: fieldRect && viewport ? fieldRect.top < viewport.top : null,
          viewportClippedBottom: fieldRect && viewport ? fieldRect.bottom > viewport.bottom : null,
        });
      };
      read("before");
      ${operation}
      queueMicrotask(() => read("microtask"));
      requestAnimationFrame(() => {
        read("raf1");
        requestAnimationFrame(() => {
          read("raf2");
          requestAnimationFrame(() => read("raf3"));
        });
      });
      setTimeout(() => read("40ms"), 40);
      setTimeout(() => read("90ms"), 90);
      setTimeout(() => read("180ms"), 180);
      setTimeout(() => resolve(samples), 230);
    })`);
  };

  const assertVisibleCaret = (samples, label) => {
    for (const sample of samples.slice(1)) {
      assert.equal(
        sample.activeFieldFocused,
        true,
        `${label}: target field lost focus at ${sample.label}: ${JSON.stringify(samples)}`,
      );
      assert.equal(
        sample.deepPart,
        "keyboard-sink",
        `${label}: keyboard sink is not focused at ${sample.label}: ${JSON.stringify(samples)}`,
      );
      assert.ok(
        sample.caretNodes.some(
          (caret) =>
            caret.height > 0 &&
            caret.display !== "none" &&
            caret.visibility !== "hidden" &&
            caret.opacity !== "0" &&
            caret.pseudoVisibility !== "hidden" &&
            caret.pseudoOpacity !== "0",
        ),
        `${label}: visible caret missing at ${sample.label}: ${JSON.stringify(samples)}`,
      );
    }
    const finalActiveLineId = samples.at(-1)?.activeFieldId;
    const activeFieldNodeIds = samples
      .slice(1)
      .filter((sample) => sample.activeFieldId === finalActiveLineId)
      .map((sample) => sample.activeFieldNodeId);
    assert.equal(
      new Set(activeFieldNodeIds).size,
      1,
      `${label}: target Mathfield was replaced: ${JSON.stringify(samples)}`,
    );
  };

  const assertStableViewport = (samples, label) => {
    const first = samples[0];
    for (const sample of samples.slice(1)) {
      assert.ok(
        Math.abs(sample.scrollTop - first.scrollTop) <= 1,
        `${label}: scrollTop moved at ${sample.label}: ${JSON.stringify(samples)}`,
      );
      if (first.anchorTop !== null) {
        assert.ok(
          Math.abs(sample.anchorTop - first.anchorTop) <= 1,
          `${label}: anchor moved at ${sample.label}: ${JSON.stringify(samples)}`,
        );
      }
    }
  };

  // Top-of-document deletion baseline.
  await load(["a+b", ""] , 1);
  await focus(1);
  await evaluate(`window.__visualtexLateStructuralKeydown = 0`);
  const topDelete = await probeOperation({
    operation: `{
      const field = document.querySelectorAll("math-field")[1];
      field.addEventListener("keydown", () => { window.__visualtexLateStructuralKeydown += 1; }, { capture: true, once: true });
      field.dispatchEvent(new KeyboardEvent("keydown", { key: "Backspace", code: "Backspace", bubbles: true, composed: true, cancelable: true }));
    }`,
  });
  assert.equal(await evaluate(`window.__visualtexLateStructuralKeydown`), 0, "empty-row Backspace leaked to a later same-field listener");
  console.log("[top-delete-caret]", JSON.stringify(topDelete.map((sample) => ({
    label: sample.label,
    activeFieldId: sample.activeFieldId,
    position: sample.activeFieldPosition,
    selection: sample.activeFieldSelection,
    focused: sample.activeFieldFocused,
    deepPart: sample.deepPart,
    caretNodes: sample.caretNodes,
  }))));
  assert.equal(topDelete.at(-1).rowCount, 1);
  assert.equal(topDelete.at(-1).activeRows.length, 1);
  assertStableViewport(topDelete, "top delete");
  assertVisibleCaret(topDelete, "top delete");
  const settledDeleteBackgrounds = topDelete.slice(1).map(
    (sample) => sample.backgrounds[0],
  );
  assert.equal(
    new Set(settledDeleteBackgrounds).size,
    1,
    `top delete active background animated: ${JSON.stringify(topDelete)}`,
  );

  // Scrolled add/delete baseline with a stable line above the edited line.
  const manyLines = Array.from({ length: 18 }, (_, index) =>
    index % 4 === 0 ? `x_${index}+\\frac{a}{b}` : `x_${index}+y`,
  );
  const ids = await load(manyLines, 11);
  await evaluate(`(() => {
    const scroll = document.querySelector(".editor-pane-scroll");
    const anchor = document.querySelector('[data-line-id="${ids[6]}"]');
    scroll.scrollTop += anchor.getBoundingClientRect().top - scroll.getBoundingClientRect().top - 35;
  })()`);
  await focus(11);
  await evaluate(`window.__visualtexLateStructuralKeydown = 0`);
  const scrolledAdd = await probeOperation({
    anchorId: ids[6],
    operation: `{
      const field = document.querySelectorAll("math-field")[11];
      field.addEventListener("keydown", () => { window.__visualtexLateStructuralKeydown += 1; }, { capture: true, once: true });
      field.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", code: "Enter", bubbles: true, composed: true, cancelable: true }));
    }`,
  });
  assert.equal(await evaluate(`window.__visualtexLateStructuralKeydown`), 0, "Enter leaked to a later same-field listener");
  assert.equal(scrolledAdd.at(-1).rowCount, 19);
  assertStableViewport(scrolledAdd, "scrolled add");
  assertVisibleCaret(scrolledAdd, "scrolled add");
  const newIndex = await evaluate(`document.querySelectorAll("math-field").length - 1`);
  const newFieldIndex = await evaluate(`Array.from(document.querySelectorAll("math-field")).findIndex((field) => field.value === "" && field.closest(".formula-line")?.classList.contains("is-active"))`);
  assert.ok(newFieldIndex >= 0, `new empty field missing; count=${newIndex + 1}`);
  const scrolledDelete = await probeOperation({
    anchorId: ids[6],
    operation: `document.querySelectorAll("math-field")[${newFieldIndex}].dispatchEvent(new KeyboardEvent("keydown", { key: "Backspace", code: "Backspace", bubbles: true, composed: true, cancelable: true }));`,
  });
  assert.equal(scrolledDelete.at(-1).rowCount, 18);
  assertStableViewport(scrolledDelete, "scrolled empty delete");
  assertVisibleCaret(scrolledDelete, "scrolled empty delete");

  // Split and merge in the middle of a scrolled document.
  const splitIds = await load(manyLines, 11);
  await evaluate(`(() => {
    const scroll = document.querySelector(".editor-pane-scroll");
    const anchor = document.querySelector('[data-line-id="${splitIds[6]}"]');
    scroll.scrollTop += anchor.getBoundingClientRect().top - scroll.getBoundingClientRect().top - 35;
  })()`);
  const splitFocus = await focus(11, "x_11+");
  assert.ok(splitFocus.position > 0 && splitFocus.position < splitFocus.lastOffset, JSON.stringify(splitFocus));
  const scrolledSplit = await probeOperation({
    anchorId: splitIds[6],
    operation: `document.querySelectorAll("math-field")[11].dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", code: "Enter", bubbles: true, composed: true, cancelable: true }));`,
  });
  assert.equal(scrolledSplit.at(-1).rowCount, 19);
  assertStableViewport(scrolledSplit, "scrolled split");
  assertVisibleCaret(scrolledSplit, "scrolled split");
  const splitValues = await evaluate(`Array.from(document.querySelectorAll("math-field")).slice(11, 13).map((field) => field.value)`);
  assert.match(splitValues[0], /x_\{?11\}?\+?/);
  assert.equal(splitValues[1], "y");

  await focus(12, "");
  await evaluate(`window.__visualtexLateStructuralKeydown = 0`);
  const scrolledMerge = await probeOperation({
    anchorId: splitIds[6],
    operation: `{
      const field = document.querySelectorAll("math-field")[12];
      field.addEventListener("keydown", () => { window.__visualtexLateStructuralKeydown += 1; }, { capture: true, once: true });
      field.dispatchEvent(new KeyboardEvent("keydown", { key: "Backspace", code: "Backspace", bubbles: true, composed: true, cancelable: true }));
    }`,
  });
  assert.equal(await evaluate(`window.__visualtexLateStructuralKeydown`), 0, "merge Backspace leaked to a later same-field listener");
  console.log("[merge-caret]", JSON.stringify(scrolledMerge.map((sample) => ({
    label: sample.label,
    activeFieldId: sample.activeFieldId,
    position: sample.activeFieldPosition,
    selection: sample.activeFieldSelection,
    focused: sample.activeFieldFocused,
    deepPart: sample.deepPart,
    caretNodes: sample.caretNodes,
  }))));
  assert.equal(scrolledMerge.at(-1).rowCount, 18);
  assertStableViewport(scrolledMerge, "scrolled merge");
  assertVisibleCaret(scrolledMerge, "scrolled merge");
  const mergedValue = await evaluate(`document.querySelectorAll("math-field")[11].value`);
  assert.match(mergedValue, /x_\{?11\}?\+?y/);

  // Split immediately before a tall integral.
  const integral = "a+\\int_{0}^{1}\\frac{f(x)}{g(x)}\\,\\mathrm{d}x";
  await load([integral], 0);
  const focused = await focus(0, "a+");
  assert.ok(focused.position < focused.lastOffset, JSON.stringify(focused));
  const integralSplit = await sampleIntegral(`document.querySelector("math-field").dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", code: "Enter", bubbles: true, composed: true, cancelable: true }));`);
  const integralVisible = integralSplit.slice(1);
  assert.deepEqual(integralVisible.at(-1).values, [
    "a+",
    "\\int_0^1\\frac{f(x)}{g(x)}\\,\\mathrm{d}x",
  ]);
  assert.ok(integralVisible.every((sample) => !sample.clippedTop && !sample.clippedBottom));
  assert.ok(integralVisible.every((sample) => !sample.viewportClippedTop && !sample.viewportClippedBottom));
  assert.ok(
    Math.max(...integralVisible.map((sample) => sample.fieldTop)) -
      Math.min(...integralVisible.map((sample) => sample.fieldTop)) <= 1,
    `integral row moved after split: ${JSON.stringify(integralSplit)}`,
  );
  assert.ok(
    Math.max(...integralVisible.map((sample) => sample.fieldHeight)) -
      Math.min(...integralVisible.map((sample) => sample.fieldHeight)) <= 1,
    `integral height changed after split: ${JSON.stringify(integralSplit)}`,
  );

  console.log("Line scroll, deletion, merge, and integral clipping regression passed");
} finally {
  client?.close();
  chrome?.kill("SIGTERM");
  preview.kill("SIGTERM");
  await sleep(220);
  await rm(profile, { recursive: true, force: true });
}
