// the look of a raw command's empty argument

export const names = ["raw-placeholder-visual"];

export async function run({ client, sleep, typeText, waitForEvaluation }) {
  const rawPlaceholderGeometry = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        if (!field?.isConnected) return { ready: false };
        field.setValue("\\\\frac{\\\\placeholder{}}{\\\\placeholder{}}", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "placeholder",
          silenceNotifications: true,
        });
        field.focus();
        const sink = field.shadowRoot?.querySelector('[part="keyboard-sink"]');
        sink?.focus({ preventScroll: true });
        const placeholder = field.shadowRoot?.querySelector(
          ".visualtex-structural-placeholder",
        );
        const bounds = placeholder?.getBoundingClientRect();
        return {
          ready:
            field.shadowRoot?.querySelectorAll(".visualtex-structural-placeholder").length === 2 &&
            Boolean(bounds && bounds.width > 0 && bounds.height > 0),
          x: bounds ? bounds.left + bounds.width / 2 : 0,
          y: bounds ? bounds.top + bounds.height / 2 : 0,
        };
      })()`, "fraction numerator placeholder before raw input");
  await client.send("Input.dispatchMouseEvent", {
    type: "mousePressed",
    x: rawPlaceholderGeometry.x,
    y: rawPlaceholderGeometry.y,
    button: "left",
    buttons: 1,
    clickCount: 1,
  });
  await client.send("Input.dispatchMouseEvent", {
    type: "mouseReleased",
    x: rawPlaceholderGeometry.x,
    y: rawPlaceholderGeometry.y,
    button: "left",
    buttons: 0,
    clickCount: 1,
  });
  await sleep(80);
  await typeText("\\the");

  const visualState = await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const root = field?.shadowRoot;
        const container = root?.querySelector(".ML__container");
        const rawNodes = [...(root?.querySelectorAll(".ML__raw-latex") ?? [])]
          .filter((node) => !node.classList.contains("ML__suggestion"));
        const rawText = rawNodes.map((node) => node.textContent ?? "").join("");
        const isTransparent = (value) =>
          value === "transparent" ||
          value === "rgba(0, 0, 0, 0)" ||
          /rgba\\([^)]*,\\s*0(?:\\.0+)?\\)$/.test(value);
        const inspected = container
          ? [container, ...container.querySelectorAll("*")]
          : [];
        const offenders = inspected.flatMap((node) => {
          if (!(node instanceof HTMLElement)) return [];
          if (
            node.classList.contains("visualtex-structural-placeholder") ||
            node.classList.contains("visualtex-structural-placeholder-caret") ||
            // an empty slot's marker, drawn by the layout engine
            node.classList.contains("ML__vtx-marker")
          ) {
            return [];
          }
          const bounds = node.getBoundingClientRect();
          const style = getComputedStyle(node);
          if (
            bounds.width < 8 ||
            bounds.height < 8 ||
            isTransparent(style.backgroundColor)
          ) {
            return [];
          }
          return [{
            classes: node.className,
            tag: node.tagName,
            backgroundColor: style.backgroundColor,
            width: bounds.width,
            height: bounds.height,
          }];
        });
        const latexNodes = [...(root?.querySelectorAll(".ML__latex") ?? [])].map((node) => ({
          backgroundColor: getComputedStyle(node).backgroundColor,
          boxShadow: getComputedStyle(node).boxShadow,
          outlineWidth: getComputedStyle(node).outlineWidth,
        }));
        const selection = root?.querySelector(".ML__selection");
        return {
          ready:
            rawText === "\\\\the" &&
            container?.classList.contains("has-visualtex-raw-latex-command") &&
            offenders.length === 0 &&
            latexNodes.every((item) => isTransparent(item.backgroundColor) && item.boxShadow === "none") &&
            (!selection || getComputedStyle(selection).display === "none"),
          rawText,
          rawClass: container?.classList.contains("has-visualtex-raw-latex-command") ?? false,
          offenders,
          latexNodes,
          selectionDisplay: selection ? getComputedStyle(selection).display : "missing",
          remainingPlaceholderCount: root?.querySelectorAll(".visualtex-structural-placeholder").length ?? -1,
          value: field?.value ?? "",
          mode: field?.mode ?? "",
        };
      })()`, "raw LaTeX input has no large gray selection background");

  console.log(JSON.stringify({ visualState }, null, 2));
  console.log("Targeted raw placeholder visual regression passed");
  return;
}
