// the export menu and its formats

export const names = ["export"];

export async function run({ evaluate, waitForEvaluation }) {
  await waitForEvaluation(`(() => ({
        ready: Boolean(
          document.querySelector(".document-title-area input") &&
          document.querySelector("math-field"),
        ),
      }))()`, "export editor controls");
  await evaluate(`(() => {
        const titleInput = document.querySelector(".document-title-area input");
        const field = document.querySelector("math-field");
        if (!(titleInput instanceof HTMLInputElement) || !field) return false;
        const valueSetter = Object.getOwnPropertyDescriptor(
          HTMLInputElement.prototype,
          "value",
        )?.set;
        valueSetter?.call(titleInput, "Export Test");
        titleInput.dispatchEvent(new Event("input", { bubbles: true }));
        titleInput.dispatchEvent(new Event("change", { bubbles: true }));
        field.setValue("\\\\frac{a}{b}+x^2", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.dispatchEvent(new InputEvent("input", {
          bubbles: true,
          composed: true,
          inputType: "insertText",
          data: "x",
        }));
        return true;
      })()`);
  await waitForEvaluation(`(() => ({
        ready:
          document.querySelector(".document-title-area input")?.value === "Export Test" &&
          document.querySelector("math-field")?.value === "\\\\frac{a}{b}+x^2",
        title: document.querySelector(".document-title-area input")?.value ?? "",
        value: document.querySelector("math-field")?.value ?? "",
      }))()`, "formula document prepared for export");

  await evaluate(`(() => {
        window.__visualtexCapturedExports = [];
        URL.revokeObjectURL = () => {};
        HTMLAnchorElement.prototype.click = function captureVisualTeXExport() {
          window.__visualtexCapturedExports.push({
            filename: this.download,
            href: this.href,
          });
        };
      })()`);

  const clickExportOption = async (label, expectedCount) => {
    await evaluate(`document.querySelector(".workspace-export-trigger")?.click()`);
    await waitForEvaluation(`(() => ({
          ready: Boolean(document.querySelector(".export-dialog")),
        }))()`, `export dialog opened for ${label}`);
    await evaluate(`(() => {
          const button = [...document.querySelectorAll(".export-format-option")]
            .find((candidate) => candidate.querySelector("strong")?.textContent?.trim() === ${JSON.stringify(label)});
          button?.click();
        })()`);
    await waitForEvaluation(`(() => ({
          ready: Boolean(
            [...document.querySelectorAll(".export-format-option")]
              .find((candidate) => candidate.querySelector("strong")?.textContent?.trim() === ${JSON.stringify(label)})
              ?.getAttribute("aria-checked") === "true",
          ),
        }))()`, `${label} export format selected`);
    await evaluate(`document.querySelector(".export-confirm-button")?.click()`);
    await waitForEvaluation(`(() => ({
          ready:
            (window.__visualtexCapturedExports?.length ?? 0) >= ${expectedCount} &&
            !document.querySelector(".export-dialog"),
          count: window.__visualtexCapturedExports?.length ?? 0,
        }))()`, `${label} export captured`, 15000);
  };

  await clickExportOption("Markdown", 1);
  const markdownState = await waitForEvaluation(`(async () => {
        const item = window.__visualtexCapturedExports?.[0];
        if (!item) return { ready: false };
        const text = await fetch(item.href).then((response) => response.text());
        return {
          ready:
            item.filename.endsWith(".md") &&
            text.includes("Export Test") &&
            text.includes("\\\\frac{a}{b}+x^2"),
          filename: item.filename,
          bytes: new TextEncoder().encode(text).length,
          text,
        };
      })()`, "Markdown Blob content");

  await clickExportOption("SVG", 2);
  const svgState = await waitForEvaluation(`(async () => {
        const item = window.__visualtexCapturedExports?.[1];
        if (!item) return { ready: false };
        const text = await fetch(item.href).then((response) => response.text());
        return {
          ready:
            item.filename.endsWith(".svg") &&
            text.startsWith("<svg") &&
            !text.includes("<foreignObject"),
          filename: item.filename,
          bytes: new TextEncoder().encode(text).length,
        };
      })()`, "self-contained SVG Blob content");

  await clickExportOption("PNG", 3);
  const pngState = await waitForEvaluation(`(async () => {
        const item = window.__visualtexCapturedExports?.[2];
        if (!item) return { ready: false };
        const bytes = new Uint8Array(
          await fetch(item.href).then((response) => response.arrayBuffer()),
        );
        const expected = [137, 80, 78, 71, 13, 10, 26, 10];
        return {
          ready:
            item.filename.endsWith(".png") &&
            bytes.length > expected.length &&
            expected.every((value, index) => bytes[index] === value),
          filename: item.filename,
          bytes: bytes.length,
          signature: [...bytes.slice(0, 8)],
        };
      })()`, "valid PNG Blob content", 10000);

  console.log(
    JSON.stringify(
      { markdownState, svgState, pngState },
      null,
      2,
    ),
  );
  console.log("Targeted export regression passed");
  return;
}
