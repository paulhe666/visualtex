// a composition whose end the input method never reported is recovered

export const names = ["ime-stale-recovery"];

export async function run({ clearField, evaluate }) {
  await clearField();
  const state = await evaluate(`(async () => {
        const field = document.querySelector("math-field");
        const sink = field?.shadowRoot?.querySelector('[part="keyboard-sink"]');
        if (!field || !(sink instanceof HTMLElement)) return { ready: false };

        const frame = () =>
          new Promise((resolve) =>
            requestAnimationFrame(() => requestAnimationFrame(resolve))
          );
        const setFormula = (latex = "q") => {
          field.setValue(latex, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceAll",
            selectionMode: "after",
            silenceNotifications: true,
          });
          field.mode = "math";
          field.position = field.lastOffset;
          field.focus();
          sink.focus({ preventScroll: true });
        };
        const snapshot = () => {
          const model = field._mathfield?.model;
          let compositionAtoms = -1;
          if (model) {
            compositionAtoms = Array.from(
              { length: model.lastOffset + 1 },
              (_, index) => model.at(index),
            ).filter((atom) => atom?.type === "composition").length;
          }
          return {
            value: field.value,
            mode: field.mode,
            position: field.position,
            selection: field.selection,
            compositionAtoms,
            sinkFocused: field.shadowRoot?.activeElement === sink,
          };
        };
        const composition = (type, data = "") => {
          sink.dispatchEvent(
            new CompositionEvent(type, {
              bubbles: true,
              composed: true,
              data,
            }),
          );
        };
        const keyboard = (type, key, code) => {
          sink.dispatchEvent(
            new KeyboardEvent(type, {
              bubbles: true,
              composed: true,
              key,
              code,
            }),
          );
        };
        const hostCompositionEnds = [];
        const hostCompositionEndListener = (event) => {
          hostCompositionEnds.push(event.data ?? "");
        };
        field.addEventListener(
          "compositionend",
          hostCompositionEndListener,
        );
        const ordinaryInput = (data) => {
          sink.dispatchEvent(
            new InputEvent("input", {
              bubbles: true,
              composed: true,
              inputType: "insertText",
              data,
              isComposing: false,
            }),
          );
        };
        const typeOrdinary = async (data, code) => {
          keyboard("keydown", data, code);
          ordinaryInput(data);
          keyboard("keyup", data, code);
          await frame();
        };

        // 1. Missing compositionend followed by a definite ordinary key must
        // cancel the abandoned composition and process that key normally.
        setFormula();
        composition("compositionstart", "");
        composition("compositionupdate", "中");
        await frame();
        const staleBeforeOrdinaryKey = snapshot();
        await typeOrdinary("x", "KeyX");
        const staleAfterOrdinaryKey = snapshot();

        // 2. Some WebView2 paths omit compositionend but still send the final
        // non-composing input carrying the committed text.
        setFormula();
        composition("compositionstart", "");
        composition("compositionupdate", "中");
        ordinaryInput("中");
        await frame();
        const missingEndFinalInput = snapshot();
        await typeOrdinary("y", "KeyY");
        const missingEndAfterNextKey = snapshot();

        // 3. A second compositionstart means the previous transaction was
        // abandoned. It must be rolled back before the new one starts.
        setFormula();
        composition("compositionstart", "");
        composition("compositionupdate", "中");
        composition("compositionstart", "");
        composition("compositionupdate", "文");
        composition("compositionend", "文");
        await frame();
        const repeatedStartCommit = snapshot();
        await typeOrdinary("z", "KeyZ");
        const repeatedStartAfterNextKey = snapshot();

        // 4. Losing focus during composition must not leave MathLive locked.
        setFormula();
        composition("compositionstart", "");
        composition("compositionupdate", "中");
        sink.blur();
        await frame();
        sink.focus({ preventScroll: true });
        await typeOrdinary("w", "KeyW");
        const blurRecovery = snapshot();

        // 5. MathLive's explicit cancelComposition() must clear the internal
        // transaction itself instead of relying on blur to emit compositionend.
        setFormula();
        composition("compositionstart", "");
        composition("compositionupdate", "中");
        field._mathfield?.keyboardDelegate?.cancelComposition();
        await frame();
        sink.focus({ preventScroll: true });
        await typeOrdinary("v", "KeyV");
        const explicitCancelRecovery = snapshot();

        // 6. Stress the exact input-method-switch shape: composition begins,
        // Shift is pressed to switch Chinese/English, compositionend is lost,
        // then an ordinary Latin key arrives. Repeat enough times to catch a
        // sticky internal composition flag or stale CompositionAtom buildup.
        const switchStressFailures = [];
        for (let index = 0; index < 60; index += 1) {
          setFormula();
          composition("compositionstart", "");
          composition("compositionupdate", "中");
          keyboard("keydown", "Shift", "ShiftLeft");
          keyboard("keyup", "Shift", "ShiftLeft");
          await typeOrdinary("a", "KeyA");
          const cycle = snapshot();
          if (
            cycle.value !== "qa" ||
            cycle.mode !== "math" ||
            cycle.compositionAtoms !== 0
          ) {
            switchStressFailures.push({ index, cycle });
            break;
          }
          if (index % 10 === 9) await frame();
        }
        const switchStressFinal = snapshot();

        field.removeEventListener(
          "compositionend",
          hostCompositionEndListener,
        );

        return {
          ready: true,
          hostCompositionEnds,
          switchStressFailures,
          switchStressFinal,
          staleBeforeOrdinaryKey,
          staleAfterOrdinaryKey,
          missingEndFinalInput,
          missingEndAfterNextKey,
          repeatedStartCommit,
          repeatedStartAfterNextKey,
          blurRecovery,
          explicitCancelRecovery,
        };
      })()`);

  if (!state?.ready) {
    throw new Error(`IME stale recovery fixture failed to initialize: ${JSON.stringify(state)}`);
  }
  const expected = {
    staleAfterOrdinaryKey: "qx",
    missingEndFinalInput: "\\text{中}q",
    missingEndAfterNextKey: "\\text{中}qy",
    repeatedStartCommit: "\\text{文}q",
    repeatedStartAfterNextKey: "\\text{文}qz",
    blurRecovery: "qw",
    explicitCancelRecovery: "qv",
  };
  // Depending on root insertion order, a committed composition can serialize
  // immediately before or after the pre-existing root glyph. Accept either,
  // but never a lingering composition atom or a swallowed next key.
  const committedTextMatches = (value, glyph, tail = "") =>
    value === `q\\text{${glyph}}${tail}` ||
    value === `\\text{${glyph}}q${tail}`;
  if (
    state.staleBeforeOrdinaryKey.compositionAtoms < 1 ||
    state.hostCompositionEnds.filter((data) => data === "中").length !== 1 ||
    state.hostCompositionEnds.filter((data) => data === "").length !== 0 ||
    state.staleAfterOrdinaryKey.value !== expected.staleAfterOrdinaryKey ||
    state.staleAfterOrdinaryKey.compositionAtoms !== 0 ||
    state.staleAfterOrdinaryKey.mode !== "math" ||
    !committedTextMatches(state.missingEndFinalInput.value, "中") ||
    state.missingEndFinalInput.compositionAtoms !== 0 ||
    !committedTextMatches(state.missingEndAfterNextKey.value, "中", "y") ||
    !committedTextMatches(state.repeatedStartCommit.value, "文") ||
    state.repeatedStartCommit.value.includes("中") ||
    !committedTextMatches(state.repeatedStartAfterNextKey.value, "文", "z") ||
    state.blurRecovery.value !== expected.blurRecovery ||
    state.blurRecovery.compositionAtoms !== 0 ||
    state.explicitCancelRecovery.value !== expected.explicitCancelRecovery ||
    state.explicitCancelRecovery.compositionAtoms !== 0 ||
    state.switchStressFailures.length !== 0 ||
    state.switchStressFinal.value !== "qa" ||
    state.switchStressFinal.mode !== "math" ||
    state.switchStressFinal.compositionAtoms !== 0
  ) {
    throw new Error(
      `MathLive IME stale transaction recovery failed: ${JSON.stringify(state)}`,
    );
  }
  console.log(JSON.stringify(state, null, 2));
  console.log("Targeted stale IME composition recovery regression passed");
  return;
}
