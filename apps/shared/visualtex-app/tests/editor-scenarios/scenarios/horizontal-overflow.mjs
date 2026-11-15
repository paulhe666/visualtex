// a formula wider than its row scrolls inside it

export const names = ["horizontal-overflow"];

export async function run({ clearField, evaluate, waitForEvaluation }) {
  await clearField();
  await evaluate(`(() => {
        const field = document.querySelector(".formula-line.is-active math-field");
        const longTerm = Array.from({ length: 70 }, (_, index) =>
          "x_{" + index + "}^{2}+y_{" + index + "}^{2}",
        ).join("+");
        field.setValue(longTerm, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.position = field.lastOffset;
        return field.value;
      })()`);
  const overflowState = await waitForEvaluation(`(() => {
        const field = document.querySelector(".formula-line.is-active math-field");
        const host = field?.closest(".mathfield-host");
        const latex = field?.shadowRoot?.querySelector(".ML__latex");
        if (host && host.scrollWidth > host.clientWidth) {
          host.scrollLeft = host.scrollWidth;
        }
        return {
          ready:
            Boolean(field?.value.includes("x_{69}")) &&
            (latex?.getBoundingClientRect().width ?? 0) > 1000 &&
            getComputedStyle(host).overflowX === "auto" &&
            host.scrollWidth > host.clientWidth + 2 &&
            host.scrollLeft > 0,
          valueLength: field?.value.length ?? 0,
          formulaWidth: latex?.getBoundingClientRect().width ?? -1,
          clientWidth: host?.clientWidth ?? -1,
          scrollWidth: host?.scrollWidth ?? -1,
          scrollLeft: host?.scrollLeft ?? -1,
          overflowX: host ? getComputedStyle(host).overflowX : "",
        };
      })()`, "long formula horizontal scrollbar");
  console.log(JSON.stringify(overflowState, null, 2));
  console.log("Targeted horizontal formula overflow regression passed");
  return;
}
