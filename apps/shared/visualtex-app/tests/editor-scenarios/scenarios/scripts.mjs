// superscripts and subscripts exit on their own

export const names = ["scripts"];

export async function run({ clearField, client, evaluate, focusField, key, sleep, waitForEvaluation }) {
  const setInputBehavior = async (autoExitSuperscript, autoExitSubscript) => {
    await evaluate(`(() => {
          const storageKey = "visualtex-editor";
          const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
          persisted.state = {
            ...(persisted.state || {}),
            inputBehavior: {
              ...(persisted.state?.inputBehavior || {}),
              autoExitSuperscript: ${autoExitSuperscript},
              autoExitSubscript: ${autoExitSubscript},
            },
          };
          localStorage.setItem(storageKey, JSON.stringify(persisted));
          location.reload();
        })()`);
    await sleep(650);
    await waitForEvaluation(
      `(() => ({ ready: Boolean(document.querySelector("math-field")) }))()`,
      "formula field after script-setting reload",
    );
    await focusField();
    await clearField();
  };

  const scriptKey = async (character, code, virtualKeyCode) => {
    const common = {
      key: character,
      code,
      windowsVirtualKeyCode: virtualKeyCode,
      nativeVirtualKeyCode: virtualKeyCode,
    };
    await client.send("Input.dispatchKeyEvent", {
      type: "keyDown",
      ...common,
      text: character,
      unmodifiedText: character,
    });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
    await sleep(60);
  };

  const runCase = async ({
    name,
    autoExitSuperscript,
    autoExitSubscript,
    scriptCharacter,
    scriptCode,
    scriptVirtualKeyCode,
    expected,
  }) => {
    await setInputBehavior(autoExitSuperscript, autoExitSubscript);
    await key("x", "KeyX", 88);
    await scriptKey(scriptCharacter, scriptCode, scriptVirtualKeyCode);
    await key("a", "KeyA", 65);
    await key("b", "KeyB", 66);
    return await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const markers = [...(field?.shadowRoot?.querySelectorAll(
            ".ML__placeholder-selected, .ML__selected, .ML__caret",
          ) ?? [])];
          return {
            ready: field?.value === ${JSON.stringify(expected)},
            name: ${JSON.stringify(name)},
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
            markerAncestors: markers.map((marker) => {
              const chain = [];
              let node = marker;
              while (node && chain.length < 7) {
                const bounds = node.getBoundingClientRect();
                chain.push({
                  tag: node.tagName,
                  className: node.className || "",
                  text: node.textContent || "",
                  top: bounds.top,
                  height: bounds.height,
                });
                node = node.parentElement;
              }
              return chain;
            }),
            contentHtml:
              field?.shadowRoot?.querySelector('[part="content"]')?.innerHTML ?? "",
          };
        })()`, `script independence: ${name}`, 3500);
  };

  const cases = [];
  cases.push(await runCase({
    name: "superscript disabled while subscript enabled",
    autoExitSuperscript: false,
    autoExitSubscript: true,
    scriptCharacter: "^",
    scriptCode: "Digit6",
    scriptVirtualKeyCode: 54,
    expected: "x^{ab}",
  }));
  cases.push(await runCase({
    name: "superscript enabled while subscript disabled",
    autoExitSuperscript: true,
    autoExitSubscript: false,
    scriptCharacter: "^",
    scriptCode: "Digit6",
    scriptVirtualKeyCode: 54,
    expected: "x^{a}b",
  }));
  cases.push(await runCase({
    name: "subscript disabled while superscript enabled",
    autoExitSuperscript: true,
    autoExitSubscript: false,
    scriptCharacter: "_",
    scriptCode: "Minus",
    scriptVirtualKeyCode: 189,
    expected: "x_{ab}",
  }));
  cases.push(await runCase({
    name: "subscript enabled while superscript disabled",
    autoExitSuperscript: false,
    autoExitSubscript: true,
    scriptCharacter: "_",
    scriptCode: "Minus",
    scriptVirtualKeyCode: 189,
    expected: "x_{a}b",
  }));

  console.log(JSON.stringify({ cases }, null, 2));
  console.log("Targeted independent script auto-exit regression passed");
  return;
}
