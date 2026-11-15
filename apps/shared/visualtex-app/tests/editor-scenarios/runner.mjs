// Editor scenarios in the built frontend of the app this runs from (its dist, served by `vite preview`), each in a
// fresh headless Chrome driven over CDP (testing/browser_test_runtime.mjs). One file per scenario in ./scenarios: `names` (the scenarios it runs),
// `run(ctx)`, and optionally `storage()` (editor state the page starts with), `localStorageItems()` (other
// localStorage entries it starts with), `platforms` (the apps it applies to, e.g. ["windows"]; skipped elsewhere)
// and `beforeNavigation(ctx)`.
//   node ../shared/visualtex-app/tests/editor-scenarios/runner.mjs <scenario>   one scenario
//   node ../shared/visualtex-app/tests/editor-scenarios/runner.mjs --all        every scenario, one process each
//   node ../shared/visualtex-app/tests/editor-scenarios/runner.mjs --list
import { spawnSync } from "node:child_process";
import { readdirSync, readFileSync } from "node:fs";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { BLANK_PAGE_PATH, withAppBrowser } from "../../testing/browser_test_runtime.mjs";

const scenarioDir = new URL("./scenarios/", import.meta.url);
const registry = new Map();
for (const file of readdirSync(fileURLToPath(scenarioDir)).filter((name) => name.endsWith(".mjs")).sort()) {
  const module = await import(new URL(file, scenarioDir).href);
  for (const name of module.names) {
    if (registry.has(name)) throw new Error(`scenario ${name} is defined twice`);
    registry.set(name, module);
  }
}

// the app this runs in: apps/windows or apps/macos
const app = JSON.parse(readFileSync("package.json", "utf8")).name.replace(/^visualtex-/, "");
const appliesHere = (module) => !module.platforms || module.platforms.includes(app);

const argument = process.argv[2];
if (argument === "--list") {
  console.log([...registry.keys()].join("\n"));
  process.exit(0);
}
if (argument === "--all") {
  const failed = [];
  let ran = 0;
  for (const [name, module] of registry) {
    if (!appliesHere(module)) {
      console.log(`=== ${name} (${module.platforms.join(", ")} only: skipped) ===`);
      continue;
    }
    console.log(`=== ${name} ===`);
    ran += 1;
    const result = spawnSync(process.execPath, [fileURLToPath(import.meta.url), name], { stdio: "inherit" });
    if (result.status !== 0) failed.push(name);
  }
  console.log(failed.length ? `failed: ${failed.join(", ")}` : `All ${ran} ${app} editor scenarios passed`);
  process.exit(failed.length ? 1 : 0);
}
const scenario = argument;
const scenarioModule = registry.get(scenario);
if (!scenarioModule) {
  console.error(`Usage: node ../shared/visualtex-app/tests/editor-scenarios/runner.mjs <${[...registry.keys()].join("|")}|--all|--list>`);
  process.exit(1);
}
if (!appliesHere(scenarioModule)) {
  console.log(`${scenario} applies to the ${scenarioModule.platforms.join(", ")} app only: skipped in ${app}`);
  process.exit(0);
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const canonicalLatex = (value) => value.replace(/\s+/g, "").replace(/\{([A-Za-z0-9])\}/g, "$1");

async function main() {
  await withAppBrowser(async ({ baseUrl, client, evaluate }) => {
    const waitForEvaluation = async (expression, description, timeoutMs = 12000) => {
      const started = Date.now();
      let lastValue;
      while (Date.now() - started < timeoutMs) {
        try {
          lastValue = await evaluate(expression);
        } catch (error) {
          // a reload the scenario just started replaces the document under the evaluation
          if (!/Inspected target navigated or closed|Execution context was destroyed|Cannot find context|Promise was collected/.test(String(error))) throw error;
          lastValue = { ready: false, navigating: true };
        }
        if (lastValue?.ready) return lastValue;
        await sleep(50);
      }
      const runtimeEvents = client.events.map((event) => ({
        method: event.method,
        exception: event.params?.exceptionDetails?.exception?.description ?? event.params?.exceptionDetails?.text ?? null,
        console: event.params?.args?.map((arg) => arg.value ?? arg.description ?? "") ?? null,
      }));
      throw new Error(`Timed out waiting for ${description}: ${JSON.stringify({ lastValue, runtimeEvents })}`);
    };

    // a key with its text when it types one (no modifier, or Shift); a shortcut sends no text
    const key = async (value, code, virtualKeyCode, modifiers = 0) => {
      const common = { key: value, code, modifiers, windowsVirtualKeyCode: virtualKeyCode, nativeVirtualKeyCode: virtualKeyCode };
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        ...common,
        ...(value.length === 1 && (modifiers === 0 || modifiers === 8) ? { text: value, unmodifiedText: value } : {}),
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(45);
    };

    const typeText = async (text) => {
      for (const character of text) {
        const code = character === "\\" ? "Backslash" : `Key${character.toUpperCase()}`;
        const virtualKeyCode = character === "\\" ? 220 : character.toUpperCase().charCodeAt(0);
        await key(character, code, virtualKeyCode);
      }
    };

    // a real pointer click in the middle of an element, once nothing covers it
    const clickSelectorWithPointer = async (selector) => {
      const point = await waitForEvaluation(`(() => {
        const element = document.querySelector(__SELECTOR__);
        if (!(element instanceof HTMLElement)) return { ready: false };
        element.scrollIntoView({ block: "nearest", inline: "nearest" });
        const rect = element.getBoundingClientRect();
        const x = rect.left + rect.width / 2;
        const y = rect.top + rect.height / 2;
        const hit = document.elementFromPoint(x, y);
        return {
          ready:
            rect.width > 0 &&
            rect.height > 0 &&
            x >= 0 && x <= window.innerWidth &&
            y >= 0 && y <= window.innerHeight &&
            Boolean(hit && (hit === element || element.contains(hit))),
          x,
          y,
        };
      })()`.replace("__SELECTOR__", JSON.stringify(selector)), `pointer target ${selector}`);
      await client.send("Input.dispatchMouseEvent", { type: "mousePressed", x: point.x, y: point.y, button: "left", buttons: 1, clickCount: 1 });
      await client.send("Input.dispatchMouseEvent", { type: "mouseReleased", x: point.x, y: point.y, button: "left", buttons: 0, clickCount: 1 });
      await sleep(80);
    };

    const ctx = { client, evaluate, waitForEvaluation, key, typeText, clickSelectorWithPointer, sleep, scenario, baseUrl, canonicalLatex };
    await scenarioModule.beforeNavigation?.(ctx);

    // The start state goes into localStorage on a blank page of the app's origin, so it is in place before the app
    // first loads (written into a running app, the app's own next save could overwrite it before a reload).
    await client.send("Page.navigate", { url: `${baseUrl}${BLANK_PAGE_PATH}` });
    await waitForEvaluation(`(() => ({
      ready: location.pathname === ${JSON.stringify(BLANK_PAGE_PATH)} && document.readyState === "complete",
    }))()`, "blank page of the app's origin");
    const storage = scenarioModule.storage?.() ?? {};
    const localStorageItems = scenarioModule.localStorageItems?.() ?? {};
    await evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
      localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed", "true");
      localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true");
      const line = { id: crypto.randomUUID(), latex: "" };
      localStorage.setItem("visualtex-editor", JSON.stringify({
        state: { lines: [line], activeLineId: line.id, checkUpdatesOnStartup: false, ...${JSON.stringify(storage)} },
        version: 0,
      }));
      for (const [key, value] of Object.entries(${JSON.stringify(localStorageItems)})) localStorage.setItem(key, value);
    })()`);
    await client.send("Page.navigate", { url: baseUrl });
    await waitForEvaluation(`(() => ({
      ready: location.pathname === "/" && Boolean(document.querySelector("math-field")),
    }))()`, "formula field");
    await evaluate(`(() => {
      const laterButton = [...document.querySelectorAll('.office-first-run-backdrop button')]
        .find((button) => /Later|稍后处理/.test(button.textContent || ''));
      if (laterButton instanceof HTMLElement) laterButton.click();
      return true;
    })()`);
    // the app's own first saves after it starts
    await sleep(400);

    const ensurePointerTarget = async (x, y) => {
      await waitForEvaluation(`(() => {
        const later = [...document.querySelectorAll('.office-first-run-backdrop button')]
          .find(button => /Later|稍后处理/.test(button.textContent || ''));
        later?.click();
        return { ready: document.elementFromPoint(${x}, ${y}) === document.querySelector('math-field') };
      })()`, "unobstructed formula before pointer drag");
    };

    // the active row's formula field (the first one when no row is active)
    const focusField = async () => {
      await waitForEvaluation(`(() => {
        const field =
          document.querySelector(".formula-line.is-active math-field") ??
          document.querySelector("math-field");
        if (!field?.isConnected) return { ready: false };
        field.focus();
        field.position = field.lastOffset;
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        return { ready: field.isConnected && field.hasFocus() };
      })()`, "stable focused formula field");
      await sleep(80);
    };

    const clearField = async () => {
      await waitForEvaluation(`(() => {
        const field =
          document.querySelector(".formula-line.is-active math-field") ??
          document.querySelector("math-field");
        if (!field?.isConnected) return { ready: false };
        field.setValue("", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.position = field.lastOffset;
        field.dispatchEvent(new InputEvent("input", { bubbles: true, composed: true, inputType: "insertText" }));
        return { ready: field.isConnected && field.value === "" };
      })()`, "stable empty formula field");
      await sleep(100);
      await focusField();
    };

    await scenarioModule.run({ ...ctx, ensurePointerTarget, focusField, clearField });
  }, { path: "about:blank" });
}

// a browser or preview server that fails to start is retried; a scenario's own failure is not
async function runWithTransientBrowserRetry() {
  for (let attempt = 0; attempt < 3; attempt += 1) {
    try {
      await main();
      return;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      const transientTargetFailure =
        message.includes("Inspected target navigated or closed") || message.includes("The browser did not open");
      if (!transientTargetFailure || attempt >= 2) throw error;
      console.warn(`Transient browser/preview startup failure; retrying ${scenario} (${attempt + 2}/3).`);
      await sleep(700);
    }
  }
}

runWithTransientBrowserRetry().catch((error) => {
  console.error(error instanceof Error ? error.stack : error);
  process.exitCode = 1;
});
