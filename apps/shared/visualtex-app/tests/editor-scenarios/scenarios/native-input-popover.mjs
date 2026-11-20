// the native candidate popover

export const names = ["native-input-popover"];

// Windows shows its own stable copy of the kernel's candidate list (mathLiveKernelCommands.ts); macOS the kernel's list
export const platforms = ["windows"];

// state the page starts with (merged into the editor's persisted state)
export const storage = () => ({ inputBehavior: { autoEscapeShortcuts: true, autoExitSuperscript: true, autoExitSubscript: true, autoExitAccent: true, autoExitWrapperCommand: true, showStructuredCommandSuggestions: false, showOtherCommandSuggestions: false } });

export async function run({ clearField, evaluate, key, sleep, typeText, waitForEvaluation }) {
  await clearField();
  await key("Enter", "Enter", 13);
  await waitForEvaluation(`(() => {
        const fields = [...document.querySelectorAll("math-field")];
        return {
          ready: fields.length === 2 && fields[1]?.hasFocus?.(),
          lineCount: fields.length,
          focusedIndex: fields.findIndex((field) => field.hasFocus?.()),
        };
      })()`, "second formula line before native suggestion navigation");
  await key("ArrowUp", "ArrowUp", 38);
  await waitForEvaluation(`(() => {
        const fields = [...document.querySelectorAll("math-field")];
        return {
          ready: fields.length === 2 && fields[0]?.hasFocus?.(),
          focusedIndex: fields.findIndex((field) => field.hasFocus?.()),
        };
      })()`, "first formula line before native suggestion navigation");
  // Row navigation reapplies focus and selection at 0 ms and 80 ms. Wait
  // until those callbacks settle, otherwise the test can type `\\` first,
  // have the caret reset in front of it, and accidentally produce `be\\`.
  await sleep(140);
  await typeText("\\be");
  const initial = await waitForEvaluation(`(() => {
        const stable = document.getElementById("visualtex-native-input-suggestion-popover");
        const source = document.getElementById("mathlive-suggestion-popover");
        const field = document.querySelectorAll("math-field")[0];
        const bounds = stable?.getBoundingClientRect();
        const commands = [...(stable?.querySelectorAll("li[data-command]") ?? [])]
          .map((item) => item.dataset.command ?? "");
        const sourceCommands = [...(source?.querySelectorAll("li[data-command]") ?? [])]
          .map((item) => item.dataset.command ?? "");
        const rawLatex = [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
          .filter((node) => !node.classList.contains("ML__suggestion"))
          .map((node) => node.textContent ?? "")
          .join("");
        const stableZIndex = stable
          ? Number.parseInt(getComputedStyle(stable).zIndex || "0", 10)
          : 0;
        const chromeZIndex = Math.max(
          0,
          ...[".formula-toolbar", ".source-panel", ".editor-pane-header"]
            .map((selector) => document.querySelector(selector))
            .filter(Boolean)
            .map((element) => Number.parseInt(getComputedStyle(element).zIndex || "0", 10) || 0),
        );
        return {
          ready:
            Boolean(stable?.classList.contains("is-visible")) &&
            commands.length >= 2 &&
            source?.dataset.visualtexInputPopoverSource === "true" &&
            document.querySelectorAll("math-field").length === 2 &&
            document.querySelectorAll("math-field")[0]?.hasFocus?.() &&
            !document.querySelector(".suggestion-popup") &&
            stableZIndex >= 190 &&
            stableZIndex > chromeZIndex,
          commands,
          sourceCommands,
          rawLatex,
          fieldMode: field?.mode ?? "",
          fieldValue: field?.value ?? "",
          sourceExists: Boolean(source),
          sourceVisible: source?.classList.contains("is-visible") ?? false,
          focusedIndex: [...document.querySelectorAll("math-field")]
            .findIndex((field) => field.hasFocus?.()),
          selected: stable?.querySelector("li.ML__popover__current")?.dataset.command ?? "",
          bounds: bounds
            ? { left: bounds.left, top: bounds.top, width: bounds.width, height: bounds.height }
            : null,
          sourceOpacity: source ? getComputedStyle(source).opacity : "",
          stableZIndex,
          chromeZIndex,
          customCandidateVisible: Boolean(document.querySelector(".suggestion-popup")),
        };
      })()`, "stable native input-selection popover for \\be");

  await evaluate(`(() => {
        const source = document.getElementById("mathlive-suggestion-popover");
        const stable = document.getElementById("visualtex-native-input-suggestion-popover");
        const field = document.querySelectorAll("math-field")[0];
        const beforeSelected = stable?.querySelector("li.ML__popover__current")?.dataset.command ?? "";
        const sourceItems = [...(source?.querySelectorAll("li[data-command]") ?? [])];
        const sourceIndex = sourceItems.findIndex((item) =>
          item.classList.contains("ML__popover__current"),
        );
        const downCommand = sourceItems[
          sourceIndex < 0 ? 0 : (sourceIndex + 1) % sourceItems.length
        ]?.dataset.command ?? "";
        const key = downCommand && downCommand !== beforeSelected ? "ArrowDown" : "ArrowUp";
        const direction = key === "ArrowDown" ? 1 : -1;
        const expectedCommand = sourceItems[
          sourceIndex < 0
            ? direction > 0 ? 0 : sourceItems.length - 1
            : (sourceIndex + direction + sourceItems.length) % sourceItems.length
        ]?.dataset.command ?? "";
        source?.classList.remove("is-visible");
        source?.setAttribute("aria-hidden", "true");
        stable?.classList.add("is-visible");
        stable?.setAttribute("aria-hidden", "false");
        field?.dispatchEvent(new KeyboardEvent("keydown", {
          key,
          code: key,
          bubbles: true,
          composed: true,
          cancelable: true,
        }));
        const selected = stable?.querySelector("li.ML__popover__current")?.dataset.command ?? "";
        const latestSource = document.getElementById("mathlive-suggestion-popover");
        latestSource?.classList.add("is-visible");
        latestSource?.setAttribute("aria-hidden", "false");
        window.__visualtexNativePriorityState = {
          beforeSelected,
          selected,
          expectedCommand,
          key,
        };
      })()`);
  const priorityState = await waitForEvaluation(`(() => {
        const state = window.__visualtexNativePriorityState;
        const lines = [...document.querySelectorAll(".formula-line")];
        const activeIndex = lines.findIndex((line) => line.classList.contains("is-active"));
        const source = document.getElementById("mathlive-suggestion-popover");
        const stable = document.getElementById("visualtex-native-input-suggestion-popover");
        return {
          ready: Boolean(
            state?.selected &&
            state.selected === state.expectedCommand &&
            activeIndex === 0
          ),
          beforeSelected: state?.beforeSelected ?? "",
          selected: state?.selected ?? "",
          expectedCommand: state?.expectedCommand ?? "",
          key: state?.key ?? "",
          activeIndex,
          sourceVisible: source?.classList.contains("is-visible") ?? false,
          stableVisible: stable?.classList.contains("is-visible") ?? false,
        };
      })()`, "native suggestion priority while source visibility is transient");

  await evaluate(`(() => {
        const node = document.getElementById("visualtex-native-input-suggestion-popover");
        const monitor = {
          node,
          removed: 0,
          hiddenTransitions: 0,
          ariaHiddenTransitions: 0,
          outerChildMutations: 0,
        };
        const observer = new MutationObserver((records) => {
          for (const record of records) {
            if (
              record.type === "childList" &&
              [...record.removedNodes].some(
                (removed) => removed === node || removed.contains?.(node),
              )
            ) {
              monitor.removed += 1;
            }
            if (record.target === node && record.type === "childList") {
              monitor.outerChildMutations += 1;
            }
            if (record.target === node && record.attributeName === "class") {
              if (!node.classList.contains("is-visible")) {
                monitor.hiddenTransitions += 1;
              }
            }
            if (record.target === node && record.attributeName === "aria-hidden") {
              if (node.getAttribute("aria-hidden") !== "false") {
                monitor.ariaHiddenTransitions += 1;
              }
            }
          }
        });
        observer.observe(document.body, {
          childList: true,
          subtree: true,
          attributes: true,
          attributeFilter: ["class", "aria-hidden"],
        });
        window.__visualtexNativeInputMonitor = monitor;
        window.__visualtexNativeInputObserver = observer;
      })()`);

  await key("ArrowDown", "ArrowDown", 40);
  const arrowState = await waitForEvaluation(`(() => {
        const monitor = window.__visualtexNativeInputMonitor;
        const stable = document.getElementById("visualtex-native-input-suggestion-popover");
        const bounds = stable?.getBoundingClientRect();
        const selected = stable?.querySelector("li.ML__popover__current")?.dataset.command ?? "";
        const fields = [...document.querySelectorAll("math-field")];
        return {
          ready:
            stable === monitor?.node &&
            stable?.classList.contains("is-visible") &&
            selected &&
            selected !== ${JSON.stringify(priorityState.selected)} &&
            fields.length === 2 &&
            fields[0]?.hasFocus?.() &&
            monitor.removed === 0 &&
            monitor.hiddenTransitions === 0 &&
            monitor.ariaHiddenTransitions === 0 &&
            Math.abs((bounds?.left ?? 0) - ${initial.bounds.left}) <= 1 &&
            Math.abs((bounds?.top ?? 0) - ${initial.bounds.top}) <= 1 &&
            Math.abs((bounds?.width ?? 0) - ${initial.bounds.width}) <= 1 &&
            Math.abs((bounds?.height ?? 0) - ${initial.bounds.height}) <= 1,
          sameNode: stable === monitor?.node,
          selected,
          focusedIndex: fields.findIndex((field) => field.hasFocus?.()),
          bounds: bounds
            ? { left: bounds.left, top: bounds.top, width: bounds.width, height: bounds.height }
            : null,
          removed: monitor?.removed ?? -1,
          hiddenTransitions: monitor?.hiddenTransitions ?? -1,
          ariaHiddenTransitions: monitor?.ariaHiddenTransitions ?? -1,
          customCandidateVisible: Boolean(document.querySelector(".suggestion-popup")),
        };
      })()`, "arrow key moves only the native input-selection highlight");

  await typeText("t");
  const refinedState = await waitForEvaluation(`(() => {
        const monitor = window.__visualtexNativeInputMonitor;
        const stable = document.getElementById("visualtex-native-input-suggestion-popover");
        const commands = [...(stable?.querySelectorAll("li[data-command]") ?? [])]
          .map((item) => item.dataset.command ?? "");
        return {
          ready:
            stable === monitor?.node &&
            stable?.classList.contains("is-visible") &&
            commands.some((command) => command === "\\\\beta") &&
            commands.every((command) => command.startsWith("\\\\bet")) &&
            monitor.removed === 0 &&
            monitor.hiddenTransitions === 0 &&
            monitor.ariaHiddenTransitions === 0 &&
            !document.querySelector(".suggestion-popup"),
          sameNode: stable === monitor?.node,
          commands,
          selected: stable?.querySelector("li.ML__popover__current")?.dataset.command ?? "",
          removed: monitor?.removed ?? -1,
          hiddenTransitions: monitor?.hiddenTransitions ?? -1,
          ariaHiddenTransitions: monitor?.ariaHiddenTransitions ?? -1,
          customCandidateVisible: Boolean(document.querySelector(".suggestion-popup")),
        };
      })()`, "\\be to \\bet updates inside one persistent input-selection frame");

  await key("Backspace", "Backspace", 8);
  const restoredState = await waitForEvaluation(`(() => {
        const monitor = window.__visualtexNativeInputMonitor;
        const stable = document.getElementById("visualtex-native-input-suggestion-popover");
        const commands = [...(stable?.querySelectorAll("li[data-command]") ?? [])]
          .map((item) => item.dataset.command ?? "");
        return {
          ready:
            stable === monitor?.node &&
            stable?.classList.contains("is-visible") &&
            commands.length >= ${initial.commands.length} &&
            monitor.removed === 0 &&
            monitor.hiddenTransitions === 0 &&
            monitor.ariaHiddenTransitions === 0,
          sameNode: stable === monitor?.node,
          commands,
          removed: monitor?.removed ?? -1,
          hiddenTransitions: monitor?.hiddenTransitions ?? -1,
          ariaHiddenTransitions: monitor?.ariaHiddenTransitions ?? -1,
        };
      })()`, "Backspace restores \\be suggestions without remounting the input-selection frame");

  await evaluate(`window.__visualtexNativeInputObserver?.disconnect()`);
  console.log(JSON.stringify({ initial, priorityState, arrowState, refinedState, restoredState }, null, 2));
  console.log("Targeted native input-selection popover regression passed");
  return;
}
