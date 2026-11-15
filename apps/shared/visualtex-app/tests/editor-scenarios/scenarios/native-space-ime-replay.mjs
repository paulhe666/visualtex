// a candidate confirmed with Space by an input method is not inserted twice

export const names = ["native-space-ime-replay"];

export async function run({ clearField, evaluate, focusField, key, sleep, typeText, waitForEvaluation }) {
  const readAlphaCandidateState = () => evaluate(`(() => {
        const field = document.querySelector("math-field");
        const panel = document.getElementById("mathlive-suggestion-popover");
        const current = panel?.querySelector("li.ML__popover__current[data-command]");
        const raw = [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
          .filter((node) => !node.classList.contains("ML__suggestion"))
          .map((node) => node.textContent ?? "")
          .join("");
        return {
          ready:
            raw === "\\\\al" &&
            [...(panel?.querySelectorAll("li[data-command]") ?? [])]
              .some((item) => item.dataset.command === "\\\\alpha"),
          raw,
          selected: current?.dataset.command ?? "",
          candidates: [...(panel?.querySelectorAll("li[data-command]") ?? [])]
            .map((item) => item.dataset.command ?? ""),
        };
      })()`);

  const prepareAlpha = async () => {
    await clearField();
    await focusField();
    await typeText("\\al");
    let state = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const panel = document.getElementById("mathlive-suggestion-popover");
          const raw = [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join("");
          const candidates = [...(panel?.querySelectorAll("li[data-command]") ?? [])]
            .map((item) => item.dataset.command ?? "");
          return {
            ready: raw === "\\\\al" && candidates.includes("\\\\alpha"),
            raw,
            candidates,
          };
        })()`, "native alpha candidate list before replay probe");
    for (let index = 0; index < 12; index += 1) {
      state = await readAlphaCandidateState();
      if (state.selected === "\\alpha") return state;
      await key("ArrowDown", "ArrowDown", 40);
    }
    throw new Error(`Could not select alpha for replay probe: ${JSON.stringify(state)}`);
  };

  const commitOnce = async () => {
    await key(" ", "Space", 32);
    return waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const normalized = (field?.value ?? "").replaceAll(" ", "");
          return {
            ready: normalized === "\\\\alpha",
            value: field?.value ?? "",
            normalized,
          };
        })()`, "single alpha before replay injection");
  };

  const readReplayState = () => evaluate(`(() => {
        const field = document.querySelector("math-field");
        const normalized = (field?.value ?? "").replaceAll(" ", "");
        return {
          value: field?.value ?? "",
          normalized,
          alphaCount: normalized.split("\\\\alpha").length - 1,
          raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
          mode: field?.mode ?? "",
        };
      })()`);

  const initial = await prepareAlpha();
  const committed = await commitOnce();
  const keypressReplay = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const sink = field?.shadowRoot?.querySelector('[part="keyboard-sink"]');
        if (!sink) return false;
        for (let index = 0; index < 2; index += 1) {
          sink.dispatchEvent(new KeyboardEvent("keypress", {
            key: " ",
            code: "Space",
            keyCode: 32,
            charCode: 32,
            which: 32,
            bubbles: true,
            composed: false,
            cancelable: true,
          }));
        }
        return true;
      })()`);
  await sleep(80);
  const afterKeypressReplay = await readReplayState();

  await prepareAlpha();
  await commitOnce();
  const inputReplay = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const sink = field?.shadowRoot?.querySelector('[part="keyboard-sink"]');
        if (!sink) return false;
        sink.dispatchEvent(new InputEvent("input", {
          inputType: "insertText",
          data: " ",
          bubbles: true,
          composed: false,
          cancelable: true,
        }));
        return true;
      })()`);
  await sleep(80);
  const afterInputReplay = await readReplayState();

  await prepareAlpha();
  await commitOnce();
  const internalSpaceReplay = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const sink = field?.shadowRoot?.querySelector('[part="keyboard-sink"]');
        if (!sink) return false;
        const down = new KeyboardEvent("keydown", {
          key: " ",
          code: "Space",
          keyCode: 32,
          which: 32,
          bubbles: true,
          composed: false,
          cancelable: true,
        });
        const keydownAllowed = sink.dispatchEvent(down);
        const press = new KeyboardEvent("keypress", {
          key: " ",
          code: "Space",
          keyCode: 32,
          charCode: 32,
          which: 32,
          bubbles: true,
          composed: false,
          cancelable: true,
        });
        const keypressAllowed = sink.dispatchEvent(press);
        return {
          keydownAllowed,
          keydownPrevented: down.defaultPrevented,
          keypressAllowed,
          keypressPrevented: press.defaultPrevented,
        };
      })()`);
  await sleep(80);
  const afterInternalSpaceReplay = await readReplayState();

  console.log(JSON.stringify({
    initial,
    committed,
    keypressReplay,
    afterKeypressReplay,
    inputReplay,
    afterInputReplay,
    internalSpaceReplay,
    afterInternalSpaceReplay,
  }, null, 2));
  console.log("Targeted native Space IME replay probe completed");
  return;
}
