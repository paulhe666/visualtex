import assert from "node:assert/strict";
import process from "node:process";
import { withAppBrowser } from "../../shared/visualtex-app/testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitUntil(client, expression, timeoutMs = 10000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const value = await client.evaluate(expression);
    if (value) return value;
    await sleep(50);
  }
  throw new Error(`Timed out waiting for ${expression}`);
}

async function dispatchCharacter(client, key, code, keyCode, modifiers = 0) {
  const common = {
    key,
    code,
    windowsVirtualKeyCode: keyCode,
    nativeVirtualKeyCode: keyCode,
    modifiers,
  };
  await client.send("Input.dispatchKeyEvent", { type: "rawKeyDown", ...common });
  await client.send("Input.dispatchKeyEvent", {
    type: "char",
    ...common,
    text: key,
    unmodifiedText: key,
  });
  await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
  await sleep(70);
}

async function dispatchShortcut(client, key, code, keyCode, modifiers) {
  const common = {
    key,
    code,
    windowsVirtualKeyCode: keyCode,
    nativeVirtualKeyCode: keyCode,
    modifiers,
  };
  await client.send("Input.dispatchKeyEvent", { type: "rawKeyDown", ...common });
  await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
}

async function main() {
  await withAppBrowser(async ({ baseUrl, client }) => {
    await client.send("Page.addScriptToEvaluateOnNewDocument", {
      source: `(() => {
        try {
          localStorage.setItem("visualtex.onboarding.v3.completed", "true");
          localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
          localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed", "true");
          localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true");
          const key = "visualtex-editor";
          let persisted;
          try { persisted = JSON.parse(localStorage.getItem(key) || "null"); }
          catch { persisted = null; }
          if (!persisted || typeof persisted !== "object") persisted = { state: {}, version: 0 };
          persisted.state = {
            ...(persisted.state || {}),
            checkUpdatesOnStartup: false,
          };
          localStorage.setItem(key, JSON.stringify(persisted));
        } catch {}
      })();`,
    });
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(2500);
    await waitUntil(
      client,
      `document.readyState === "complete" && Boolean(document.querySelector("math-field"))`,
      30000,
    );

    const reset = async ({ title = "History Test", lines, inputBehavior = {} }) => {
      await client.evaluate(`(() => {
        localStorage.setItem("visualtex.onboarding.v3.completed", "true");
        localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
        localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed", "true");
        localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true");
        localStorage.removeItem("visualtex.custom-symbols.v1");
        let persisted;
        try { persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "null"); }
        catch { persisted = null; }
        if (!persisted || typeof persisted !== "object") persisted = { state: {}, version: 0 };
        persisted.state = {
          ...(persisted.state || {}),
          title: ${JSON.stringify(title)},
          lines: ${JSON.stringify(lines)},
          activeLineId: ${JSON.stringify(lines[0].id)},
          sourceOpen: false,
          checkUpdatesOnStartup: false,
          inputBehavior: {
            autoExitSuperscript: true,
            autoExitSubscript: true,
            autoExitAccent: true,
            autoExitWrapperCommand: true,
            showStructuredCommandSuggestions: true,
            showOtherCommandSuggestions: false,
            ...${JSON.stringify(inputBehavior)},
          },
        };
        delete persisted.state.latex;
        localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
      })()`);
      await client.send("Page.reload", { ignoreCache: true });
      await sleep(550);
      await waitUntil(client, `document.querySelectorAll("math-field").length === ${lines.length}`);
    };

    process.stdout.write("[custom-symbol-noninterference] reset script case\n");
    await reset({
      lines: [{ id: "script-line", latex: "" }],
      inputBehavior: { autoExitSuperscript: false, autoExitSubscript: false },
    });
    await client.evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
    })()`);
    await dispatchCharacter(client, "x", "KeyX", 88);
    await dispatchCharacter(client, "^", "Digit6", 54, 8);
    await dispatchCharacter(client, "2", "Digit2", 50);
    const scriptState = await client.evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        value: field.value,
        position: field.position,
        lastOffset: field.lastOffset,
        customStyles: Array.from(field.shadowRoot?.querySelectorAll("style") || [])
          .filter((style) => style.id?.includes("custom-symbol") || style.textContent?.includes("visualtex-custom-symbol-"))
          .map((style) => style.textContent || ""),
      };
    })()`);
    assert.match(scriptState.value, /^x(?:\^2|\^\{2\})$/);
    assert.notEqual(scriptState.position, scriptState.lastOffset);
    assert.equal(
      scriptState.customStyles.some((css) => css.includes("visualtex-custom-symbol-vtxtestsymbol")),
      false,
      "A normal formula must not receive prototype/user symbol CSS",
    );
    process.stdout.write("[custom-symbol-noninterference] ordinary script input verified\n");

    console.log(
      "Custom symbol non-interference regression passed for ordinary superscript input and per-field CSS isolation",
    );
  }, { path: "about:blank", windowSize: "1450,980" });
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
