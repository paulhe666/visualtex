import { withAppBrowser } from "../../shared/visualtex-app/testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function main() {
  await withAppBrowser(async ({ baseUrl, client }) => {
    await client.send("Page.addScriptToEvaluateOnNewDocument", {
      source: `(() => {
        const replace = (target, key, value) => {
          try {
            Object.defineProperty(target, key, {
              configurable: true,
              writable: true,
              value,
            });
          } catch {}
        };
        replace(Array.prototype, "at", undefined);
        replace(String.prototype, "replaceAll", undefined);
        replace(globalThis, "queueMicrotask", undefined);
        replace(globalThis, "ResizeObserver", undefined);
        if (globalThis.crypto) replace(globalThis.crypto, "randomUUID", undefined);
        for (const method of ["getItem", "setItem", "removeItem"]) {
          replace(Storage.prototype, method, function () {
            throw new DOMException("Storage disabled by startup regression", "SecurityError");
          });
        }
      })();`,
    });
    await client.send("Page.navigate", { url: baseUrl });

    const started = Date.now();
    let state;
    while (Date.now() - started < 15000) {
      const result = await client.send("Runtime.evaluate", {
        expression: `(() => {
          const root = document.getElementById("root");
          const shell = document.querySelector(".app-shell");
          const text = root?.innerText ?? "";
          return {
            ready: Boolean(shell),
            rootChildren: root?.children.length ?? 0,
            text: text.slice(0, 500),
            compatibility: {
              arrayAt: typeof Array.prototype.at,
              replaceAll: typeof String.prototype.replaceAll,
              queueMicrotask: typeof globalThis.queueMicrotask,
              resizeObserver: typeof globalThis.ResizeObserver,
              nativeUuid: typeof globalThis.crypto?.randomUUID,
            },
          };
        })()`,
        returnByValue: true,
      });
      state = result.result.value;
      if (state?.ready) break;
      await sleep(80);
    }

    if (!state?.ready) {
      const events = client.events.map((event) => ({
        method: event.method,
        exception:
          event.params?.exceptionDetails?.exception?.description ??
          event.params?.exceptionDetails?.text ??
          null,
        console:
          event.params?.args?.map((arg) => arg.value ?? arg.description ?? "") ??
          null,
      }));
      throw new Error(
        `VisualTeX did not mount under degraded WebView APIs: ${JSON.stringify({ state, events })}`,
      );
    }

    for (const name of [
      "arrayAt",
      "replaceAll",
      "queueMicrotask",
      "resizeObserver",
    ]) {
      const value = state.compatibility[name];
      if (value !== "function") {
        throw new Error(`Compatibility API ${name} was not restored: ${value}`);
      }
    }
    if (state.rootChildren < 1 || state.text.includes("界面加载失败")) {
      throw new Error(`Unexpected startup fallback: ${JSON.stringify(state)}`);
    }

    console.log(JSON.stringify(state, null, 2));
    console.log("Startup browser regression passed");
  }, { path: "about:blank" });
}

await main();
