// moving between formula rows with the arrow keys

export const names = ["navigation"];

export async function run({ evaluate, key, sleep, waitForEvaluation }) {
  await evaluate(`(() => {
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        const first = { id: crypto.randomUUID(), latex: "\\\\alpha" };
        const second = { id: crypto.randomUUID(), latex: "\\\\beta" };
        persisted.state = {
          ...(persisted.state || {}),
          lines: [first, second],
          activeLineId: second.id,
          inputBehavior: {
            ...(persisted.state?.inputBehavior || {}),
            showOtherCommandSuggestions: true,
          },
        };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
        location.reload();
      })()`);
  await waitForEvaluation(`(() => ({
        ready: document.querySelectorAll("math-field").length === 2,
      }))()`, "two formula fields for navigation");
  await evaluate(`(() => {
        const field = document.querySelectorAll("math-field")[1];
        field.position = field.lastOffset;
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      })()`);
  await key("ArrowUp", "ArrowUp", 38);
  const switchedState = await waitForEvaluation(`(() => {
        const rows = [...document.querySelectorAll(".formula-line")];
        const fields = [...document.querySelectorAll("math-field")];
        const surface = document.querySelector(".multi-line-editor");
        const firstLineId = rows[0]?.dataset.lineId ?? "";
        return {
          ready:
            rows.length === 2 &&
            rows[0]?.classList.contains("is-active") &&
            fields[0]?.matches(":focus-within") &&
            surface?.dataset.activeLineId === firstLineId,
          candidateVisible: Boolean(document.querySelector(".suggestion-popup")),
          query: document.querySelector(".suggestion-popup")?.textContent ?? "",
          activeLineId: surface?.dataset.activeLineId ?? "",
          firstLineId,
          focusedIndex: fields.findIndex((field) => field.matches(":focus-within")),
        };
      })()`, "ArrowUp switches to first formula field");

  await key("Escape", "Escape", 27);
  const dismissedState = await waitForEvaluation(`(() => {
        const popup = document.querySelector(".suggestion-popup");
        const field = document.querySelectorAll("math-field")[0];
        return {
          ready: !popup,
          popupText: popup?.textContent ?? "",
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
          pendingNativeSuggestion: field?.dataset.pendingNativeSuggestion ?? "",
          activeTag: document.activeElement?.tagName ?? "",
          sinkFocused: field?.shadowRoot?.activeElement?.getAttribute?.("part") ?? "",
        };
      })()`, "Escape dismisses formula-line command candidate");
  await sleep(500);
  const stableDismissedState = await waitForEvaluation(`(() => ({
        ready: !document.querySelector(".suggestion-popup"),
        popupText: document.querySelector(".suggestion-popup")?.textContent ?? "",
      }))()`, "dismissed formula-line command candidate stays closed");

  // No command candidate shows on the row any more (the Windows app has no
  // "other command suggestions" setting): that Escape opened the empty
  // LaTeX box, and a second one closes it. (This used to pass only when a
  // stray 80 ms refocus after ArrowUp happened to close the box.)
  await key("Escape", "Escape", 27);
  await waitForEvaluation(`(() => ({
        ready: document.querySelectorAll("math-field")[0]?.mode === "math",
      }))()`, "second Escape closes the LaTeX box");

  await key("ArrowDown", "ArrowDown", 40);
  const returnedState = await waitForEvaluation(`(() => {
        const rows = [...document.querySelectorAll(".formula-line")];
        const fields = [...document.querySelectorAll("math-field")];
        const surface = document.querySelector(".multi-line-editor");
        const secondLineId = rows[1]?.dataset.lineId ?? "";
        return {
          ready:
            rows[1]?.classList.contains("is-active") &&
            fields[1]?.matches(":focus-within") &&
            surface?.dataset.activeLineId === secondLineId,
          candidateVisible: Boolean(document.querySelector(".suggestion-popup")),
          activeLineId: surface?.dataset.activeLineId ?? "",
          secondLineId,
          focusedIndex: fields.findIndex((field) => field.matches(":focus-within")),
        };
      })()`, "ArrowDown returns to second formula field");

  console.log(JSON.stringify({ switchedState, dismissedState, stableDismissedState, returnedState }, null, 2));
  console.log("Targeted formula-line navigation regression passed");
  return;
}
