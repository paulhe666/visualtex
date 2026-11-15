// structures confirmed from the native popover, then typed into

export const names = ["native-structure-input-under","native-structure-input-multi","native-structure-input-core"];

export async function run({ clearField, evaluate, key, scenario, sleep, typeText, waitForEvaluation }) {
  const cases =
    scenario === "native-structure-input-over"
      ? [
          ["overarc", 1, "\\overarc{a}"],
          ["overline", 1, "\\overline{a}"],
          ["overbrace", 1, "\\overbrace{a}"],
          ["overgroup", 1, "\\overgroup{a}"],
          ["overparen", 1, "\\overparen{a}"],
          ["overleftarrow", 1, "\\overleftarrow{a}"],
          ["overrightarrow", 1, "\\overrightarrow{a}"],
          ["overleftrightarrow", 1, "\\overleftrightarrow{a}"],
          ["overleftharpoon", 1, "\\overleftharpoon{a}"],
          ["overrightharpoon", 1, "\\overrightharpoon{a}"],
          ["overlinesegment", 1, "\\overlinesegment{a}"],
        ]
      : scenario === "native-structure-input-under"
        ? [
            ["underarc", 1, "\\underarc{a}"],
            ["underline", 1, "\\underline{a}"],
            ["underbrace", 1, "\\underbrace{a}"],
            ["undergroup", 1, "\\undergroup{a}"],
            ["underparen", 1, "\\underparen{a}"],
            ["underleftarrow", 1, "\\underleftarrow{a}"],
            ["underrightarrow", 1, "\\underrightarrow{a}"],
            ["underleftrightarrow", 1, "\\underleftrightarrow{a}"],
            ["underlinesegment", 1, "\\underlinesegment{a}"],
          ]
        : scenario === "native-structure-input-core"
          ? [
              ["sqrt", 1, "\\sqrt{a}"],
              ["frac", 2, "\\frac{a}{"],
              ["dfrac", 2, "\\dfrac{a}{"],
              ["tfrac", 2, "\\tfrac{a}{"],
              ["binom", 2, "\\binom{a}{"],
            ]
          : [
            ["overset", 2, "\\overset{a}{"],
            ["underset", 2, "\\underset{a}{"],
            ["stackrel", 2, "\\stackrel{a}{"],
            ["stackbin", 2, "\\stackbin{a}{"],
            ["overunderset", 3, "\\overset{a}{\\underset{"],
          ];

  const states = [];
  for (const [command, placeholderCount, expectedFragment] of cases) {
    await clearField();
    await typeText(`\\${command}`);
    await key(" ", "Space", 32);
    const confirmedCommand =
      command === "overunderset" ? "\\overset" : `\\${command}`;
    const pending = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const value = field?.value ?? "";
          return {
            ready:
              Boolean(field?.isConnected) &&
              value.includes(${JSON.stringify(confirmedCommand)}) &&
              !field.shadowRoot?.querySelector(".ML__raw-latex"),
            value,
            selection: field?.selection ?? null,
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, `confirmed \\${command} structure`);
    const actualPlaceholderCount =
      pending.value.match(/\\placeholder\{\}/g)?.length ?? 0;
    if (actualPlaceholderCount !== placeholderCount) {
      throw new Error(
        `\\${command} created ${actualPlaceholderCount} placeholders instead of ${placeholderCount}: ${pending.value}`,
      );
    }

    await typeText("a");
    const typed = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const value = field?.value ?? "";
          return {
            ready: Boolean(field?.isConnected) && value.includes("a"),
            value,
            selection: field?.selection ?? null,
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, `typed content inside \\${command}`);
    if (!typed.value.includes(expectedFragment)) {
      throw new Error(
        `\\${command} did not place the first character in its first LaTeX argument: ${typed.value}`,
      );
    }
    if (typed.value.includes(`\\${command}{}a`)) {
      throw new Error(`\\${command} left an empty argument before typed content`);
    }
    let finalTypedValue = typed.value;
    if (command === "sqrt") {
      await typeText("b");
      const continued = await waitForEvaluation(`(() => {
            const field = document.querySelector("math-field");
            return {
              ready: field?.value === "\\\\sqrt{ab}",
              value: field?.value ?? "",
            };
          })()`, "continued typing inside the radical");
      finalTypedValue = continued.value;
    }
    states.push({ command, pending: pending.value, typed: finalTypedValue });
  }

  const protectedStructureStates = {};
  if (scenario === "native-structure-input-core") {
    await evaluate(`(() => {
          const storageKey = "visualtex-editor";
          const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
          persisted.state = {
            ...(persisted.state || {}),
            inputBehavior: {
              ...(persisted.state?.inputBehavior || {}),
              autoExitSuperscript: true,
              autoExitSubscript: true,
            },
          };
          localStorage.setItem(storageKey, JSON.stringify(persisted));
          location.reload();
        })()`);
    await sleep(650);
    await waitForEvaluation(
      `(() => ({ ready: Boolean(document.querySelector("math-field")) }))()`,
      "formula field after protected-structure setting reload",
    );

    const setPlaceholderValue = async (latex, description) => {
      await waitForEvaluation(`(() => {
            const field = document.querySelector("math-field");
            if (!field?.isConnected) return { ready: false };
            field.setValue(${JSON.stringify(latex)}, {
              mode: "math",
              format: "latex",
              insertionMode: "replaceAll",
              selectionMode: "placeholder",
              silenceNotifications: true,
            });
            field.focus();
            field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
            return {
              ready: !field.selectionIsCollapsed,
              value: field.value,
              selection: field.selection,
            };
          })()`, description);
      await sleep(100);
    };

    const readStructuredCaret = async () =>
      evaluate(`(() => {
            const field = document.querySelector("math-field");
            const markers = [...(field?.shadowRoot?.querySelectorAll(
              ".ML__placeholder-selected, .ML__selected, .ML__caret",
            ) ?? [])];
            const marker = markers.find((candidate) =>
              candidate.closest(".ML__op-group, .ML__mfrac"),
            );
            const operator = marker?.closest(".ML__op-group");
            const fraction = marker?.closest(".ML__mfrac");
            const container = operator || fraction;
            const markerBox = (marker?.parentElement || marker)?.getBoundingClientRect();
            const containerBox = container?.getBoundingClientRect();
            const region =
              markerBox && containerBox
                ? markerBox.top + markerBox.height / 2 <
                  containerBox.top + containerBox.height / 2
                  ? "upper"
                  : "lower"
                : null;
            return {
              value: field?.value ?? "",
              position: field?.position ?? -1,
              lastOffset: field?.lastOffset ?? -1,
              inOperator: Boolean(operator),
              inFraction: Boolean(fraction),
              region,
            };
          })()`);

    await setPlaceholderValue(
      "\\int_{\\placeholder{}}^{n} f",
      "integral lower-limit placeholder",
    );
    await typeText("ab");
    const operatorContinuous = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const normalized = (field?.value ?? "").replaceAll(" ", "");
          const marker = [...(field?.shadowRoot?.querySelectorAll(
            ".ML__placeholder-selected, .ML__selected, .ML__caret",
          ) ?? [])].find((candidate) => candidate.closest(".ML__op-group"));
          return {
            ready: normalized.includes("\\\\int_{ab}^{n}") && Boolean(marker),
            value: field?.value ?? "",
          };
        })()`, "continuous typing inside an integral limit");
    await key("ArrowUp", "ArrowUp", 38);
    const operatorUpper = await readStructuredCaret();
    if (!operatorUpper.inOperator || operatorUpper.region !== "upper") {
      throw new Error(
        `ArrowUp left the integral limits: ${JSON.stringify(operatorUpper)}`,
      );
    }
    await key("ArrowDown", "ArrowDown", 40);
    const operatorLower = await readStructuredCaret();
    if (!operatorLower.inOperator || operatorLower.region !== "lower") {
      throw new Error(
        `ArrowDown left the integral limits: ${JSON.stringify(operatorLower)}`,
      );
    }

    await setPlaceholderValue(
      "\\sqrt{x^{\\placeholder{}}+q}",
      "superscript placeholder inside a radical",
    );
    await typeText("ab");
    const radicalNestedScript = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const normalized = (field?.value ?? "").replaceAll(" ", "");
          return {
            ready: normalized.includes("\\\\sqrt{x^{a}b+q}"),
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, "ordinary superscript auto-exits inside a radical");

    await setPlaceholderValue(
      "\\sqrt{x_{\\placeholder{}}+q}",
      "subscript placeholder inside a radical",
    );
    await typeText("ab");
    const radicalNestedSubscript = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const normalized = (field?.value ?? "").replaceAll(" ", "");
          return {
            ready: normalized.includes("\\\\sqrt{x_{a}b+q}"),
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, "ordinary subscript auto-exits inside a radical");

    await setPlaceholderValue(
      "\\int_0^1 x^{\\placeholder{}}\\,\\mathrm{d}x",
      "superscript placeholder inside an integral body",
    );
    await typeText("ab");
    const integralBodyNestedScript = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const normalized = (field?.value ?? "").replaceAll(" ", "");
          return {
            ready: normalized.includes("x^{a}b"),
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, "ordinary superscript auto-exits inside an integral body");

    await setPlaceholderValue(
      "\\int_{x^{\\placeholder{}}}^{n} f",
      "nested superscript inside an integral limit",
    );
    await typeText("ab");
    const integralLimitNestedScript = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const normalized = (field?.value ?? "").replaceAll(" ", "");
          const marker = [...(field?.shadowRoot?.querySelectorAll(
            ".ML__placeholder-selected, .ML__selected, .ML__caret",
          ) ?? [])].find((candidate) => candidate.closest(".ML__op-group"));
          return {
            ready: normalized.includes("\\\\int_{x^{a}b}^{n}") && Boolean(marker),
            value: field?.value ?? "",
          };
        })()`, "nested superscript auto-exits but remains inside an integral limit");

    await setPlaceholderValue(
      "\\frac{x^{\\placeholder{}}+y}{z}+q",
      "superscript placeholder inside a fraction",
    );
    await typeText("ab");
    const fractionScript = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const normalized = (field?.value ?? "").replaceAll(" ", "");
          const marker = [...(field?.shadowRoot?.querySelectorAll(
            ".ML__placeholder-selected, .ML__selected, .ML__caret",
          ) ?? [])].find((candidate) => candidate.closest(".ML__mfrac"));
          return {
            ready: normalized.includes("x^{a}b") && Boolean(marker),
            value: field?.value ?? "",
          };
        })()`, "ordinary superscript auto-exits inside a fraction");

    await setPlaceholderValue(
      "x^{\\placeholder{}}",
      "one-time superscript auto-exit placeholder",
    );
    await typeText("a");
    const firstScriptExit = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          return {
            ready:
              field?.value === "x^{a}" &&
              field.position === field.lastOffset,
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, "first character auto-exits its superscript");
    await key("ArrowLeft", "ArrowLeft", 37);
    const manualScriptReentry = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const marker = [...(field?.shadowRoot?.querySelectorAll(
            ".ML__placeholder-selected, .ML__selected, .ML__caret",
          ) ?? [])].find((candidate) => candidate.closest(".ML__msubsup"));
          return {
            ready: Boolean(marker),
            value: field?.value ?? "",
            position: field?.position ?? -1,
          };
        })()`, "manual re-entry into the consumed superscript");
    await typeText("bc");
    const retainedAfterManualReentry = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const marker = [...(field?.shadowRoot?.querySelectorAll(
            ".ML__placeholder-selected, .ML__selected, .ML__caret",
          ) ?? [])].find((candidate) => candidate.closest(".ML__msubsup"));
          return {
            ready:
              field?.value === "x^{abc}" &&
              Boolean(marker) &&
              field.position !== field.lastOffset,
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, "manual re-entry disables repeated auto-exit for that superscript");

    await setPlaceholderValue(
      "x^{abc}+y^{\\placeholder{}}",
      "a separate superscript after consuming the first one",
    );
    await typeText("d");
    const separateScriptStillAutoExits = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          return {
            ready:
              field?.value === "x^{abc}+y^{d}" &&
              field.position === field.lastOffset,
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, "a different superscript still performs its first auto-exit");

    await setPlaceholderValue(
      "x_{\\placeholder{}}",
      "one-time subscript auto-exit placeholder",
    );
    await typeText("a");
    const firstSubscriptExit = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          return {
            ready:
              field?.value === "x_{a}" &&
              field.position === field.lastOffset,
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, "first character auto-exits its subscript");
    await key("ArrowLeft", "ArrowLeft", 37);
    await typeText("bc");
    const retainedSubscriptAfterManualReentry = await waitForEvaluation(`(() => {
          const field = document.querySelector("math-field");
          const marker = [...(field?.shadowRoot?.querySelectorAll(
            ".ML__placeholder-selected, .ML__selected, .ML__caret",
          ) ?? [])].find((candidate) => candidate.closest(".ML__msubsup"));
          return {
            ready:
              field?.value === "x_{abc}" &&
              Boolean(marker) &&
              field.position !== field.lastOffset,
            value: field?.value ?? "",
            position: field?.position ?? -1,
            lastOffset: field?.lastOffset ?? -1,
          };
        })()`, "manual re-entry disables repeated auto-exit for that subscript");

    await setPlaceholderValue(
      "\\frac{n}{\\placeholder{}}+q",
      "fraction denominator placeholder",
    );
    await key("ArrowUp", "ArrowUp", 38);
    const fractionUpper = await readStructuredCaret();
    if (!fractionUpper.inFraction || fractionUpper.region !== "upper") {
      throw new Error(
        `ArrowUp left the fraction: ${JSON.stringify(fractionUpper)}`,
      );
    }
    await key("ArrowDown", "ArrowDown", 40);
    const fractionLower = await readStructuredCaret();
    if (!fractionLower.inFraction || fractionLower.region !== "lower") {
      throw new Error(
        `ArrowDown left the fraction: ${JSON.stringify(fractionLower)}`,
      );
    }

    Object.assign(protectedStructureStates, {
      operatorContinuous,
      operatorUpper,
      operatorLower,
      radicalNestedScript,
      radicalNestedSubscript,
      integralBodyNestedScript,
      integralLimitNestedScript,
      fractionScript,
      firstScriptExit,
      manualScriptReentry,
      retainedAfterManualReentry,
      separateScriptStillAutoExits,
      firstSubscriptExit,
      retainedSubscriptAfterManualReentry,
      fractionUpper,
      fractionLower,
    });
  }

  console.log(JSON.stringify({ states, protectedStructureStates }, null, 2));
  console.log(`${scenario} passed`);
  return;
}
