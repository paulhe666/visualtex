import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 650;
const vitePort = 21100 + offset;
const debugPort = 27100 + offset;
const baseUrl = `http://127.0.0.1:${vitePort}`;
const officeUrl = `${baseUrl}/office-native-dialog.html?session=00000000-0000-4000-8000-000000000000&theme=light`;
const chromeProfile = `/tmp/visualtex-custom-symbol-runtime-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      if ((await fetch(url)).ok) return;
    } catch {
      // Retry while local services start.
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
      const timeout = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`CDP ${method} timed out`));
      }, 15000);
      this.pending.set(id, {
        resolve: (value) => {
          clearTimeout(timeout);
          resolve(value);
        },
        reject: (error) => {
          clearTimeout(timeout);
          reject(error);
        },
      });
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

async function findPage(debugPort, predicate) {
  const targets = await (
    await fetch(`http://127.0.0.1:${debugPort}/json/list`)
  ).json();
  return targets.find((target) => target.type === "page" && predicate(target.url));
}

async function waitUntil(client, expression, timeoutMs = 12000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const value = await client.evaluate(expression);
    if (value) return value;
    await sleep(60);
  }
  throw new Error(`Timed out waiting for ${expression}`);
}

async function connectTarget(target) {
  const client = new CdpClient(target.webSocketDebuggerUrl);
  await client.connect();
  await client.send("Runtime.enable");
  await client.send("Page.enable");
  return client;
}

async function dispatchKey(
  client,
  { key, code, keyCode, text = "", modifiers = 0 },
) {
  const common = {
    key,
    code,
    modifiers,
    windowsVirtualKeyCode: keyCode,
    nativeVirtualKeyCode: keyCode,
  };
  await client.send("Input.dispatchKeyEvent", {
    type: "keyDown",
    ...common,
    ...(text ? { text, unmodifiedText: text } : {}),
  });
  await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
  await sleep(45);
}

async function typeRawCommandPrefix(client, command) {
  await dispatchKey(client, {
    key: "\\",
    code: "Backslash",
    keyCode: 220,
    text: "\\",
  });
  for (const character of command) {
    const upper = character.toUpperCase();
    await dispatchKey(client, {
      key: character,
      code: `Key${upper}`,
      keyCode: upper.charCodeAt(0),
      text: character,
    });
  }
}

async function main() {
  const vite = spawn(
    process.execPath,
    [
      "node_modules/vite/bin/vite.js",
      "--host",
      "127.0.0.1",
      "--port",
      String(vitePort),
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
        "--window-size=1400,1000",
        baseUrl,
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const mainTarget = await findPage(
      debugPort,
      (url) => url === baseUrl || url === `${baseUrl}/`,
    );
    assert.ok(mainTarget, "Main page target must exist");
    mainClient = await connectTarget(mainTarget);
    await mainClient.send("Page.navigate", { url: baseUrl });
    await sleep(500);

    await mainClient.evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
      localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed", "true");
      localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true");
      localStorage.removeItem("visualtex.custom-symbols.v1");
      const key = "visualtex-editor";
      const persisted = JSON.parse(localStorage.getItem(key) || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        checkUpdatesOnStartup: false,
        latexCodeFormat: "raw",
      };
      localStorage.setItem(key, JSON.stringify(persisted));
      return true;
    })()`);
    await mainClient.send("Page.reload", { ignoreCache: true });
    await sleep(550);
    await mainClient.send("Page.reload", { ignoreCache: true });
    await sleep(550);
    await waitUntil(mainClient, `Boolean(document.querySelector("math-field"))`);
    process.stdout.write("[custom-symbol-runtime] main ready\n");

    const optionGuardProbe = await mainClient.evaluate(`(async () => {
      const compatibility = await import("/src/editor/mathLiveOptionCompatibility.ts");
      const probe = document.createElement("math-field");
      probe.value = "x";
      document.body.appendChild(probe);
      await new Promise((resolve) => requestAnimationFrame(() => resolve(true)));
      const original = probe._setOptions.bind(probe);
      let calls = 0;
      probe._setOptions = (options) => {
        calls += 1;
        throw new TypeError("Cannot set properties of undefined (setting 'mode')");
      };
      compatibility.installMathLiveOptionMutationGuard(probe);
      let rejected = false;
      try {
        probe.smartFence = false;
      } catch (error) {
        rejected = error instanceof TypeError && /setting ['\"]mode['\"]/.test(error.message);
      }
      const result = {
        calls,
        rejected,
        smartFence: probe.smartFence,
        connected: probe.isConnected,
      };
      probe._setOptions = original;
      probe.remove();
      return result;
    })()`);
    assert.equal(optionGuardProbe.calls, 1, "MathLive option guard must not retry an uncommitted failure");
    assert.equal(optionGuardProbe.rejected, true, "MathLive option guard must rethrow an uncommitted missing-mode TypeError");
    assert.equal(optionGuardProbe.smartFence, true, "an uncommitted option mutation must leave the original option intact");
    process.stdout.write("[custom-symbol-runtime] strict MathLive option guard verified\n");

    await mainClient.send("Target.createTarget", { url: officeUrl });
    let officeTarget;
    for (let attempt = 0; attempt < 100 && !officeTarget; attempt += 1) {
      officeTarget = await findPage(debugPort, (url) => url.startsWith(officeUrl));
      if (!officeTarget) await sleep(50);
    }
    assert.ok(officeTarget, "Office page target must exist");
    officeClient = await connectTarget(officeTarget);
    await sleep(450);
    const officeInitial = await officeClient.evaluate(`(async () => {
      const registry = await import("/src/math/customSymbolRegistry.ts");
      return {
        revision: registry.getCustomSymbolRevision(),
        count: registry.readCustomSymbolLibrary().symbols.length,
        hasMathfield: Boolean(document.querySelector("math-field")),
      };
    })()`);
    assert.equal(officeInitial.count, 0);
    process.stdout.write("[custom-symbol-runtime] office registry ready\n");

    const crossWindowStorageKey = "visualtex.regression.cross-window-removal";
    const cachedStorageValue = await mainClient.evaluate(`(async () => {
      const storageModule = await import("/src/runtime/safeStorage.ts");
      window.__visualtexSafeStorage = storageModule.safeStorage;
      storageModule.safeStorage.setItemStrict(
        ${JSON.stringify(crossWindowStorageKey)},
        "present",
      );
      return storageModule.safeStorage.getItem(
        ${JSON.stringify(crossWindowStorageKey)},
      );
    })()`);
    assert.equal(cachedStorageValue, "present");
    const officeSawStorageValue = await officeClient.evaluate(`(() => {
      const key = ${JSON.stringify(crossWindowStorageKey)};
      const value = localStorage.getItem(key);
      localStorage.removeItem(key);
      return value;
    })()`);
    assert.equal(officeSawStorageValue, "present");
    await waitUntil(
      mainClient,
      `window.__visualtexSafeStorage?.getItem(${JSON.stringify(crossWindowStorageKey)}) === null`,
    );
    process.stdout.write("[custom-symbol-runtime] cross-window storage removal cache invalidation verified\n");

    const collisionChecks = await mainClient.evaluate(`(async () => {
      const registration = await import("/src/math/customSymbolRegistration.ts");
      const registry = await import("/src/math/customSymbolRegistry.ts");
      const editorStore = await import("/src/stores/editorStore.ts");
      const exportRuntime = await import("/src/export/runtime.ts");
      const officeRender = await import("/src/office/shared/formulaRenderArtifacts.ts");
      const applicationConfiguration = await import("/src/runtime/applicationConfiguration.ts");
      window.__visualtexCustomSymbolRegistration = registration;
      window.__visualtexCustomSymbolRegistry = registry;
      window.__visualtexEditorStore = editorStore;
      window.__visualtexExportRuntime = exportRuntime;
      window.__visualtexOfficeRender = officeRender;
      window.__visualtexApplicationConfiguration = applicationConfiguration;
      const checks = {};
      for (const command of ["alpha", "frac", "color", "hslash"]) {
        try {
          registration.assertMathLiveCustomSymbolCommandAvailable(command);
          checks[command] = false;
        } catch {
          checks[command] = true;
        }
      }
      try {
        registration.assertMathLiveCustomSymbolCommandAvailable("selfdefa");
        checks.selfdefa = true;
      } catch {
        checks.selfdefa = false;
      }
      return checks;
    })()`);
    assert.deepEqual(collisionChecks, {
      alpha: true,
      frac: true,
      color: true,
      hslash: true,
      selfdefa: true,
    });
    const importedCollisionFiltered = await mainClient.evaluate(`(() => {
      const registry = window.__visualtexCustomSymbolRegistry;
      const now = Date.now();
      localStorage.setItem(
        "visualtex.custom-symbols.v1",
        JSON.stringify({
          version: 1,
          symbols: [{
            id: "regression-import-hslash",
            command: "hslash",
            name: "Imported collision",
            role: "ordinary",
            limitsBehavior: "auto",
            metrics: { widthEm: 0.7, ascentEm: 0.6, descentEm: 0.1 },
            artwork: { shapes: [{ kind: "circle", cx: 350, cy: 350, r: 250, fill: true }] },
            ommlFallback: null,
            createdAt: now,
            updatedAt: now,
          }],
        }),
      );
      registry.refreshCustomSymbolLibraryFromStorage();
      const filtered = registry.readCustomSymbolLibrary().symbols.length === 0;
      localStorage.removeItem("visualtex.custom-symbols.v1");
      registry.refreshCustomSymbolLibraryFromStorage();
      return filtered;
    })()`);
    assert.equal(importedCollisionFiltered, true);
    process.stdout.write("[custom-symbol-runtime] collision guards verified\n");

    const mathLiveRoleProbe = await mainClient.evaluate(`(() => {
      const registry = window.__visualtexCustomSymbolRegistry;
      const Mathfield = customElements.get("math-field");
      const BS = String.fromCharCode(92);
      const roles = [
        ["ordinary", "roleordinary", "auto"],
        ["binary", "rolebinary", "auto"],
        ["relation", "rolerelation", "auto"],
        ["operator", "roleoperator", "limits"],
        ["open", "roleopen", "auto"],
        ["close", "roleclose", "auto"],
        ["punctuation", "rolepunctuation", "auto"],
      ];
      const result = {};
      for (const [role, command, limitsBehavior] of roles) {
        const field = new Mathfield();
        field.style.position = "fixed";
        field.style.left = "-10000px";
        document.body.append(field);
        const definition = {
          id: "runtime-role-" + role,
          command,
          name: role,
          role,
          limitsBehavior,
          metrics: { widthEm: 0.76, ascentEm: 0.64, descentEm: 0.1 },
          artwork: { shapes: [{ kind: "circle", cx: 380, cy: 370, r: 250, fill: true }] },
          ommlFallback: null,
          createdAt: 0,
          updatedAt: 0,
        };
        field.macros = {
          ...field.macros,
          [command]: registry.customSymbolMathLiveMacroDefinition(definition),
        };
        const source = role === "operator"
          ? BS + command + "_i^j"
          : "A+" + BS + command + "+B";
        field.setValue(source, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        let macroEnd = -1;
        for (let offset = 1; offset <= field.lastOffset; offset += 1) {
          if (field.getElementInfo(offset)?.latex?.trim() === BS + command) {
            macroEnd = offset;
            break;
          }
        }
        let atomicDistance = 0;
        if (macroEnd > 0) {
          field.position = macroEnd;
          field.executeCommand("moveToPreviousChar");
          atomicDistance = macroEnd - field.position;
        }
        result[role] = {
          value: field.value,
          lastOffset: field.lastOffset,
          macroEnd,
          atomicDistance,
        };
        field.remove();
      }
      return result;
    })()`);
    for (const role of [
      "ordinary",
      "binary",
      "relation",
      "operator",
      "open",
      "close",
      "punctuation",
    ]) {
      assert.ok(mathLiveRoleProbe[role]?.macroEnd > 0, `${role} MathLive macro missing`);
      assert.ok(
        mathLiveRoleProbe[role]?.atomicDistance > 1,
        `${role} custom symbol must remain one captureSelection atom`,
      );
    }
    assert.match(mathLiveRoleProbe.operator.value, /\\roleoperator_\{i\}\^\{j\}/);
    process.stdout.write("[custom-symbol-runtime] MathLive roles verified\n");

    await mainClient.evaluate(`(() => {
      const BS = String.fromCharCode(92);
      const state = window.__visualtexEditorStore.useEditorStore.getState();
      const lineId = state.lines[0]?.id;
      if (!lineId) throw new Error("No formula line available for runtime sync test");
      state.replaceFormulaLine(lineId, "A+" + BS + "selfdefa+B");
      return true;
    })()`);
    await waitUntil(
      mainClient,
      `document.querySelector("math-field")?.value?.includes("selfdefa")`,
    );
    const unresolvedBefore = await mainClient.evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        value: field.value,
        classCount: field.shadowRoot?.querySelectorAll(
          ".visualtex-custom-symbol-regression-live-selfdefa",
        ).length || 0,
      };
    })()`);
    assert.match(unresolvedBefore.value, /\\selfdefa/);
    assert.equal(unresolvedBefore.classCount, 0);
    await mainClient.evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.position = 3;
      field.selection = { ranges: [[3, 3]], direction: "none" };
      return field.getValue(0, field.position, "latex");
    })()`);
    process.stdout.write("[custom-symbol-runtime] unresolved seed ready\n");

    const registered = await mainClient.evaluate(`(() => {
      const registration = window.__visualtexCustomSymbolRegistration;
      const now = Date.now();
      const library = registration.registerCustomSymbolSafely({
        id: "regression-live-selfdefa",
        command: "selfdefa",
        name: "Live custom symbol",
        role: "binary",
        limitsBehavior: "auto",
        metrics: { widthEm: 0.84, ascentEm: 0.64, descentEm: 0.1 },
        artwork: {
          shapes: [
            { kind: "circle", cx: 420, cy: 370, r: 255, fill: false, strokeWidth: 76 },
            { kind: "line", x1: 130, y1: 370, x2: 710, y2: 370, fill: false, strokeWidth: 76, lineCap: "round" },
          ],
        },
        ommlFallback: "\\\\oplus",
        createdAt: now,
        updatedAt: now,
      });
      return library.symbols.map((symbol) => symbol.command);
    })()`);
    assert.deepEqual(registered, ["selfdefa"]);
    process.stdout.write("[custom-symbol-runtime] registered\n");
    await sleep(600);
    const registrationDebug = await mainClient.evaluate(`(() => {
      const field = document.querySelector("math-field");
      const registry = window.__visualtexCustomSymbolRegistry;
      return {
        revision: registry.getCustomSymbolRevision(),
        commands: registry.readCustomSymbolLibrary().symbols.map((symbol) => symbol.command),
        fieldValue: field?.value || "",
        lastOffset: field?.lastOffset ?? -1,
        prefixAtThree: field?.lastOffset >= 3 ? field.getValue(0, 3, "latex") : "",
        prefixAtEleven: field?.lastOffset >= 11 ? field.getValue(0, 11, "latex") : "",
        hasFieldMacro: Boolean(field?.macros?.selfdefa),
        shadowStyleHasSymbol: field?.shadowRoot
          ?.getElementById("visualtex-custom-symbol-runtime-shadow-style")
          ?.textContent?.includes("regression-live-selfdefa") || false,
        shadowStyleHasPrototype: field?.shadowRoot
          ?.getElementById("visualtex-custom-symbol-runtime-shadow-style")
          ?.textContent?.includes("visualtex-custom-symbol-vtxtestsymbol") || false,
        globalStyleHasSymbol: document
          .getElementById("visualtex-custom-symbol-runtime-style")
          ?.textContent?.includes("regression-live-selfdefa") || false,
        globalStyleHasPrototype: document
          .getElementById("visualtex-custom-symbol-runtime-style")
          ?.textContent?.includes("visualtex-custom-symbol-vtxtestsymbol") || false,
      };
    })()`);
    process.stdout.write(`[custom-symbol-runtime] registration debug ${JSON.stringify(registrationDebug)}\n`);
    assert.equal(registrationDebug.lastOffset, 13, "Runtime registration must rebuild the MathLive macro model");
    assert.equal(registrationDebug.shadowStyleHasSymbol, true);
    assert.equal(
      registrationDebug.shadowStyleHasPrototype,
      false,
      "Each Mathfield must only duplicate CSS for symbols used by that formula",
    );
    assert.equal(registrationDebug.globalStyleHasSymbol, true);
    assert.equal(
      registrationDebug.globalStyleHasPrototype,
      true,
      "The single global preview stylesheet must retain the complete active symbol library",
    );
    const mappedCaret = await mainClient.evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        position: field.position,
        prefix: field.getValue(0, field.position, "latex"),
      };
    })()`);
    assert.equal(mappedCaret.position, 11);
    assert.equal(mappedCaret.prefix, "A+\\selfdefa");
    await mainClient.send("Target.activateTarget", { targetId: mainTarget.id });
    await sleep(220);

    await waitUntil(
      mainClient,
      `document.querySelector("math-field")?.shadowRoot?.querySelectorAll(".visualtex-custom-symbol-regression-live-selfdefa").length === 1`,
    );
    const resolvedAfter = await mainClient.evaluate(`(() => {
      const field = document.querySelector("math-field");
      const style = field.shadowRoot?.getElementById("visualtex-custom-symbol-runtime-shadow-style");
      return {
        value: field.value,
        classCount: field.shadowRoot?.querySelectorAll(
          ".visualtex-custom-symbol-regression-live-selfdefa",
        ).length || 0,
        style: style?.textContent || "",
      };
    })()`);
    assert.match(resolvedAfter.value, /\\selfdefa/);
    assert.equal(resolvedAfter.classCount, 1);
    assert.match(resolvedAfter.style, /regression-live-selfdefa/);
    const runtimeGeometry = await mainClient.evaluate(`(() => {
      const field = document.querySelector("math-field");
      const element = field?.shadowRoot?.querySelector(".visualtex-custom-symbol-regression-live-selfdefa");
      if (!(element instanceof HTMLElement)) return null;
      const rect = element.getBoundingClientRect();
      const rule = element.querySelector(".ML__rule");
      const ruleRect = rule?.getBoundingClientRect() ?? null;
      const ruleStyle = rule ? getComputedStyle(rule) : null;
      const own = getComputedStyle(element);
      const descendants = Array.from(element.querySelectorAll("*")).map((child) => {
        const childRect = child.getBoundingClientRect();
        return {
          tag: child.tagName,
          cls: child.getAttribute("class") || "",
          rect: { x: childRect.x, y: childRect.y, width: childRect.width, height: childRect.height },
        };
      }).filter((item) => item.rect.width > 0 || item.rect.height > 0);
      const ancestors = [];
      let ancestor = rule;
      while (ancestor && ancestor instanceof HTMLElement) {
        const style = getComputedStyle(ancestor);
        ancestors.push({
          tag: ancestor.tagName,
          cls: ancestor.getAttribute("class") || "",
          opacity: style.opacity,
          visibility: style.visibility,
          display: style.display,
          color: style.color,
          filter: style.filter,
        });
        if (ancestor === element) break;
        ancestor = ancestor.parentElement;
      }
      return {
        rect: { x: rect.x, y: rect.y, width: rect.width, height: rect.height },
        ownPosition: own.position,
        ownVerticalAlign: own.verticalAlign,
        ruleRect: ruleRect ? { x: ruleRect.x, y: ruleRect.y, width: ruleRect.width, height: ruleRect.height } : null,
        ruleStyle: ruleStyle ? {
          width: ruleStyle.width,
          height: ruleStyle.height,
          borderTop: ruleStyle.borderTopWidth,
          borderRight: ruleStyle.borderRightWidth,
          borderBottom: ruleStyle.borderBottomWidth,
          borderLeft: ruleStyle.borderLeftWidth,
          paddingTop: ruleStyle.paddingTop,
          paddingRight: ruleStyle.paddingRight,
          paddingBottom: ruleStyle.paddingBottom,
          paddingLeft: ruleStyle.paddingLeft,
          marginTop: ruleStyle.marginTop,
          marginBottom: ruleStyle.marginBottom,
          boxSizing: ruleStyle.boxSizing,
          verticalAlign: ruleStyle.verticalAlign,
          color: ruleStyle.color,
          opacity: ruleStyle.opacity,
          visibility: ruleStyle.visibility,
          borderTopColor: ruleStyle.borderTopColor,
          borderRightColor: ruleStyle.borderRightColor,
          backgroundColor: ruleStyle.backgroundColor,
          maskImage: ruleStyle.maskImage,
          webkitMaskImage: ruleStyle.webkitMaskImage,
        } : null,
        descendants,
        ancestors,
      };
    })()`);
    process.stdout.write(`[custom-symbol-runtime] runtime geometry ${JSON.stringify(runtimeGeometry)}\n`);
    assert.ok(runtimeGeometry?.ruleRect);
    assert.equal(runtimeGeometry.ruleStyle?.opacity, "1");
    assert.equal(runtimeGeometry.ruleStyle?.borderTopColor, "rgba(0, 0, 0, 0)");
    assert.equal(runtimeGeometry.ruleStyle?.borderRightColor, "rgba(0, 0, 0, 0)");
    assert.notEqual(runtimeGeometry.ruleStyle?.backgroundColor, "rgba(0, 0, 0, 0)");
    assert.match(
      runtimeGeometry.ruleStyle?.maskImage || runtimeGeometry.ruleStyle?.webkitMaskImage || "",
      /data:image\/svg\+xml/,
      "Custom-symbol artwork must be painted directly on the real MathLive metric-rule box",
    );
    assert.ok(
      runtimeGeometry.ancestors.every((ancestor) => ancestor.opacity === "1"),
      `Custom-symbol phantom wrappers must be visible: ${JSON.stringify(runtimeGeometry.ancestors)}`,
    );
    process.stdout.write("[custom-symbol-runtime] main Mathfield refreshed\n");

    const searchResult = await mainClient.evaluate(`(async () => {
      const search = await import("/src/autocomplete/CommandSearchEngine.ts");
      return search.searchCommands("selfdefa", {}, false, 5).map((command) => ({
        id: command.id,
        command: command.command,
        preview: command.previewLatex,
      }));
    })()`);
    assert.equal(searchResult[0]?.id, "custom-symbol:regression-live-selfdefa");
    assert.equal(searchResult[0]?.command, "\\selfdefa");
    process.stdout.write("[custom-symbol-runtime] runtime search verified\n");

    await mainClient.evaluate(`(() => {
      const state = window.__visualtexEditorStore.useEditorStore.getState();
      state.setInputBehavior("showOtherCommandSuggestions", true);
      const lineId = state.lines[0]?.id;
      if (!lineId) throw new Error("No formula line available for suggestion test");
      state.replaceFormulaLine(lineId, "");
      return true;
    })()`);
    await waitUntil(
      mainClient,
      `document.querySelector("math-field")?.value === ""`,
    );
    const emptyFieldStyleState = await mainClient.evaluate(`(() => {
      const fieldStyle = document.querySelector("math-field")?.shadowRoot
        ?.getElementById("visualtex-custom-symbol-runtime-shadow-style")
        ?.textContent || "";
      const globalStyle = document
        .getElementById("visualtex-custom-symbol-runtime-style")
        ?.textContent || "";
      return {
        fieldHasSelfdefa: fieldStyle.includes("regression-live-selfdefa"),
        fieldHasPrototype: fieldStyle.includes("visualtex-custom-symbol-vtxtestsymbol"),
        globalHasSelfdefa: globalStyle.includes("regression-live-selfdefa"),
        globalHasPrototype: globalStyle.includes("visualtex-custom-symbol-vtxtestsymbol"),
      };
    })()`);
    assert.deepEqual(emptyFieldStyleState, {
      fieldHasSelfdefa: false,
      fieldHasPrototype: false,
      globalHasSelfdefa: true,
      globalHasPrototype: true,
    });
    process.stdout.write("[custom-symbol-runtime] per-field CSS filtering verified\n");
    await mainClient.evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      return true;
    })()`);
    await typeRawCommandPrefix(mainClient, "selfd");
    const suggestionPreview = await waitUntil(
      mainClient,
      `(() => {
        const items = Array.from(document.querySelectorAll(
          "#mathlive-suggestion-popover li[data-command], #visualtex-native-input-suggestion-popover li[data-command]",
        ));
        const item = items.find(
          (candidate) => candidate.dataset.command === "\\\\selfdefa",
        );
        if (!item) return null;
        return {
          command: item.dataset.command || "",
          previewHtml: item.querySelector(".ML__popover__command")?.innerHTML || "",
          source: item.closest("#visualtex-native-input-suggestion-popover")
            ? "visualtex-mirror"
            : "mathlive-native",
        };
      })()`,
    );
    assert.equal(suggestionPreview.command, "\\selfdefa");
    assert.match(
      suggestionPreview.previewHtml,
      /visualtex-custom-symbol-regression-live-selfdefa/,
      "Native runtime command suggestion must preserve the custom symbol artwork host",
    );
    await dispatchKey(mainClient, {
      key: "Escape",
      code: "Escape",
      keyCode: 27,
    });
    await mainClient.evaluate(`(() => {
      const BS = String.fromCharCode(92);
      const state = window.__visualtexEditorStore.useEditorStore.getState();
      const lineId = state.lines[0]?.id;
      state.replaceFormulaLine(lineId, "A+" + BS + "selfdefa+B");
      return true;
    })()`);
    await waitUntil(
      mainClient,
      `document.querySelector("math-field")?.value?.includes("selfdefa")`,
    );
    process.stdout.write("[custom-symbol-runtime] suggestion UI and MathPreview verified\n");

    await waitUntil(
      officeClient,
      `(async () => {
        const registry = await import("/src/math/customSymbolRegistry.ts");
        return registry.readCustomSymbolLibrary().symbols.some(
          (symbol) => symbol.command === "selfdefa"
        );
      })()`,
    );
    const officeSynchronized = await officeClient.evaluate(`(async () => {
      const registry = await import("/src/math/customSymbolRegistry.ts");
      return {
        revision: registry.getCustomSymbolRevision(),
        commands: registry.readCustomSymbolLibrary().symbols.map((symbol) => symbol.command),
      };
    })()`);
    assert.ok(officeSynchronized.revision > officeInitial.revision);
    assert.deepEqual(officeSynchronized.commands, ["selfdefa"]);
    process.stdout.write("[custom-symbol-runtime] office synchronized\n");

    const beforeUpdateStyle = resolvedAfter.style;
    await mainClient.evaluate(`(() => {
      const registration = window.__visualtexCustomSymbolRegistration;
      registration.updateCustomSymbolSafely("regression-live-selfdefa", {
        metrics: { widthEm: 0.92, ascentEm: 0.68, descentEm: 0.1 },
        artwork: {
          shapes: [
            { kind: "circle", cx: 460, cy: 390, r: 205, fill: false, strokeWidth: 82 },
            { kind: "line", x1: 165, y1: 390, x2: 755, y2: 390, fill: false, strokeWidth: 82, lineCap: "round" },
            { kind: "path", operation: "erase", d: "M430 315C455 335 480 355 505 375", fill: false, strokeWidth: 34, lineCap: "round", lineJoin: "round" },
          ],
        },
      });
      return true;
    })()`);
    const updatedStyle = await waitUntil(
      mainClient,
      `(() => {
        const field = document.querySelector("math-field");
        const text = field?.shadowRoot?.getElementById("visualtex-custom-symbol-runtime-shadow-style")?.textContent || "";
        return text && text !== ${JSON.stringify(beforeUpdateStyle)} ? text : "";
      })()`,
    );
    assert.match(updatedStyle, /regression-live-selfdefa/);
    assert.match(
      updatedStyle,
      /viewBox%3D%220%200%20920%20780%22/,
      "Live metric/artwork updates must refresh the mask to the updated 0.92em × 0.78em rule box",
    );
    assert.match(
      updatedStyle,
      /mask-composite:\s*subtract/,
      "Runtime erase rendering must subtract a dedicated alpha mask instead of nesting an SVG mask inside the CSS mask",
    );
    assert.match(updatedStyle, /-webkit-mask-composite:\s*source-out/);
    assert.equal(
      /%3Cmask(?:%20|%3E)/i.test(updatedStyle),
      false,
      "Runtime CSS mask data must stay flat so un-erased vector edges are not softened by a second mask-compositing pass",
    );
    const runtimeEraseMaskState = await mainClient.evaluate(`(() => {
      const field = document.querySelector("math-field");
      const element = field?.shadowRoot?.querySelector(".visualtex-custom-symbol-regression-live-selfdefa");
      const rule = element?.querySelector(".ML__rule");
      if (!rule) return null;
      const style = getComputedStyle(rule);
      return {
        maskImage: style.maskImage,
        webkitMaskImage: style.webkitMaskImage,
        maskComposite: style.maskComposite,
        webkitMaskComposite: style.webkitMaskComposite,
      };
    })()`);
    assert.ok(runtimeEraseMaskState);
    const computedMaskImages = runtimeEraseMaskState.maskImage || runtimeEraseMaskState.webkitMaskImage || "";
    assert.ok(
      (computedMaskImages.match(/data:image\/svg\+xml/g) || []).length >= 2,
      "Runtime erase rendering must expose separate paint and erase alpha masks",
    );
    assert.ok(
      [runtimeEraseMaskState.maskComposite, runtimeEraseMaskState.webkitMaskComposite]
        .some((value) => /subtract|source-out/.test(value || "")),
      `Runtime erase composition must be subtractive: ${JSON.stringify(runtimeEraseMaskState)}`,
    );
    process.stdout.write("[custom-symbol-runtime] live update and crisp erase-mask composition verified\n");

    const exportState = await mainClient.evaluate(`(() => {
      const runtime = window.__visualtexExportRuntime;
      const officeRender = window.__visualtexOfficeRender;
      const registration = window.__visualtexCustomSymbolRegistration;
      const registry = window.__visualtexCustomSymbolRegistry;
      const BS = String.fromCharCode(92);
      const svg = runtime.latexToSvg(BS + "selfdefa", {
        displayMode: false,
        fontSizePt: 12,
        paddingPx: 0,
        background: "transparent",
      });
      const word = officeRender.renderOfficeFormulaArtifacts({
        lines: [{ id: "selfdefa-word", latex: BS + "selfdefa" }],
        codeFormat: "raw",
        displayMode: "inline",
        host: "word",
        includeWordOmml: true,
      });
      const now = Date.now();
      registration.registerCustomSymbolSafely({
        id: "regression-live-selfdefb",
        command: "selfdefb",
        name: "Live custom symbol without Word fallback",
        role: "ordinary",
        limitsBehavior: "auto",
        metrics: { widthEm: 0.72, ascentEm: 0.62, descentEm: 0.08 },
        artwork: {
          shapes: [
            { kind: "rect", x: 90, y: 90, width: 540, height: 540, rx: 80, fill: false, strokeWidth: 70 },
            { kind: "line", x1: 180, y1: 360, x2: 540, y2: 360, fill: false, strokeWidth: 70, lineCap: "round" },
          ],
        },
        ommlFallback: null,
        createdAt: now,
        updatedAt: now,
      });
      let wordNoFallbackRejected = false;
      let wordNoFallbackMessage = "";
      try {
        officeRender.renderOfficeFormulaArtifacts({
          lines: [{ id: "selfdefb-word", latex: BS + "selfdefb" }],
          codeFormat: "raw",
          displayMode: "inline",
          host: "word",
          includeWordOmml: true,
        });
      } catch (error) {
        wordNoFallbackRejected = true;
        wordNoFallbackMessage = error instanceof Error ? error.message : String(error);
      }
      const powerpoint = officeRender.renderOfficeFormulaArtifacts({
        lines: [{ id: "selfdefb-powerpoint", latex: BS + "selfdefb" }],
        codeFormat: "raw",
        displayMode: "inline",
        host: "powerpoint",
        includeWordOmml: false,
      });
      registry.deleteCustomSymbol("regression-live-selfdefb");
      return {
        hasArtwork: svg.svg.includes('data-visualtex-custom-symbol="regression-live-selfdefa"'),
        hasUpdatedCircle: /<circle\\b[^>]*cx="460"[^>]*r="205"/.test(svg.svg),
        mathMl: runtime.latexToMathMl(BS + "selfdefa", false),
        wordHasOmml: Boolean(word.omml?.omml?.includes("<m:oMath")),
        wordSvgUsesArtwork: word.svg.svg.includes('data-visualtex-custom-symbol="regression-live-selfdefa"'),
        wordNoFallbackRejected,
        wordNoFallbackMessage,
        powerpointUsesArtwork: powerpoint.svg.svg.includes('data-visualtex-custom-symbol="regression-live-selfdefb"'),
        powerpointOmmlIsNull: powerpoint.omml === null,
      };
    })()`);
    assert.equal(exportState.hasArtwork, true);
    assert.equal(exportState.hasUpdatedCircle, true);
    assert.match(exportState.mathMl, /&#x2295;/);
    assert.equal(exportState.wordHasOmml, true);
    assert.equal(exportState.wordSvgUsesArtwork, true);
    assert.equal(exportState.wordNoFallbackRejected, true);
    assert.match(exportState.wordNoFallbackMessage, /Word OMML|vtxtestsymbol|selfdefb|resolve/i);
    assert.equal(exportState.powerpointUsesArtwork, true);
    assert.equal(exportState.powerpointOmmlIsNull, true);
    process.stdout.write("[custom-symbol-runtime] export and Office rendering verified\n");

    const configurationSnapshot = await mainClient.evaluate(`(async () => {
      const editorStore = await import("/src/stores/editorStore.ts");
      window.__visualtexEditorStore = editorStore;
      editorStore.useEditorStore.setState({
        usage: {
          cases: {
            commandId: "cases",
            useCount: 17,
            lastUsedAt: 1720000000000,
            recentUses: [1719999999000, 1720000000000],
            acceptedPrefixes: { b: 5, begin: 8 },
            contextCounts: { candidate: 13, toolbar: 4 },
            pinned: true,
          },
        },
      });
      const configuration = await window.__visualtexApplicationConfiguration.buildVisualTexConfiguration();
      window.__visualtexCustomSymbolConfigurationSnapshot = configuration;
      const raw = configuration.storage["visualtex.custom-symbols.v1"] || "";
      return {
        hasStorageEntry: Boolean(raw),
        commands: raw
          ? JSON.parse(raw).symbols.map((symbol) => symbol.command)
          : [],
        casesUsage: configuration.usage?.cases ?? null,
      };
    })()`);
    assert.equal(configurationSnapshot.hasStorageEntry, true);
    assert.deepEqual(configurationSnapshot.commands, ["selfdefa"]);
    assert.deepEqual(configurationSnapshot.casesUsage, {
      commandId: "cases",
      useCount: 17,
      lastUsedAt: 1720000000000,
      recentUses: [1719999999000, 1720000000000],
      acceptedPrefixes: { b: 5, begin: 8 },
      contextCounts: { candidate: 13, toolbar: 4 },
      pinned: true,
    });
    process.stdout.write("[custom-symbol-runtime] configuration export verified\n");

    const caretBeforeDelete = await mainClient.evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.position = 11;
      field.selection = { ranges: [[11, 11]], direction: "none" };
      return field.getValue(0, field.position, "latex");
    })()`);
    assert.equal(caretBeforeDelete, "A+\\selfdefa");

    await mainClient.evaluate(`(() => {
      window.__visualtexEditorStore.useEditorStore.setState({ usage: {} });
      window.__visualtexCustomSymbolRegistry.deleteCustomSymbol("regression-live-selfdefa");
      return true;
    })()`);
    await waitUntil(
      mainClient,
      `document.querySelector("math-field")?.shadowRoot?.querySelectorAll(".visualtex-custom-symbol-regression-live-selfdefa").length === 0`,
    );
    const afterDelete = await mainClient.evaluate(`(() => ({
      value: document.querySelector("math-field")?.value || "",
      hasCustomStyle: document.querySelector("math-field")?.shadowRoot
        ?.getElementById("visualtex-custom-symbol-runtime-shadow-style")
        ?.textContent?.includes("regression-live-selfdefa") || false,
    }))()`);
    assert.match(afterDelete.value, /\\selfdefa/);
    assert.equal(afterDelete.hasCustomStyle, false);
    const deletedCaret = await mainClient.evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        position: field.position,
        lastOffset: field.lastOffset,
        prefix: field.getValue(0, field.position, "latex"),
      };
    })()`);
    assert.equal(deletedCaret.lastOffset, 5);
    assert.equal(deletedCaret.position, 3);
    assert.equal(deletedCaret.prefix, "A+\\selfdefa");
    process.stdout.write("[custom-symbol-runtime] main delete verified\n");

    await waitUntil(
      officeClient,
      `(async () => {
        const registry = await import("/src/math/customSymbolRegistry.ts");
        return !registry.readCustomSymbolLibrary().symbols.some(
          (symbol) => symbol.command === "selfdefa"
        );
      })()`,
    );
    process.stdout.write("[custom-symbol-runtime] office delete verified\n");

    await mainClient.evaluate(`(async () => {
      await window.__visualtexApplicationConfiguration.applyVisualTexConfiguration(
        window.__visualtexCustomSymbolConfigurationSnapshot,
      );
      return true;
    })()`);
    await waitUntil(
      mainClient,
      `document.querySelector("math-field")?.shadowRoot?.querySelectorAll(".visualtex-custom-symbol-regression-live-selfdefa").length === 1`,
    );
    await waitUntil(
      officeClient,
      `(async () => {
        const registry = await import("/src/math/customSymbolRegistry.ts");
        return registry.readCustomSymbolLibrary().symbols.some(
          (symbol) => symbol.command === "selfdefa"
        );
      })()`,
    );
    const restoredConfiguration = await mainClient.evaluate(`(() => ({
      value: document.querySelector("math-field")?.value || "",
      commands: window.__visualtexCustomSymbolRegistry
        .readCustomSymbolLibrary().symbols.map((symbol) => symbol.command),
      casesUsage: window.__visualtexEditorStore.useEditorStore.getState().usage.cases ?? null,
    }))()`);
    assert.match(restoredConfiguration.value, /\\selfdefa/);
    assert.deepEqual(restoredConfiguration.commands, ["selfdefa"]);
    assert.deepEqual(restoredConfiguration.casesUsage, configurationSnapshot.casesUsage);
    process.stdout.write("[custom-symbol-runtime] configuration restore verified\n");

    await mainClient.evaluate(`(() => {
      window.__visualtexCustomSymbolRegistry.deleteCustomSymbol("regression-live-selfdefa");
      return true;
    })()`);
    await waitUntil(
      officeClient,
      `(async () => {
        const registry = await import("/src/math/customSymbolRegistry.ts");
        return registry.readCustomSymbolLibrary().symbols.length === 0;
      })()`,
    );

    console.log(
      "Custom symbol live registration, Mathfield refresh, search, update/delete, export, Office rendering, configuration round trip, and Office-window synchronization regression passed",
    );
  } finally {
    mainClient?.close();
    officeClient?.close();
    chrome?.kill("SIGTERM");
    vite.kill("SIGTERM");
    await sleep(240);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => undefined);
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
