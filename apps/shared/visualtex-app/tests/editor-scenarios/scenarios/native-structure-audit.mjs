// the structure candidates the native popover offers

export const names = ["native-structure-audit"];

export async function run({ clearField, evaluate, typeText, waitForEvaluation }) {
  const prefixes = ["\\\\over", "\\\\under", "\\\\stack"];
  const results = [];
  for (const prefix of prefixes) {
    await clearField();
    await typeText(prefix);
    const state = await waitForEvaluation(`(() => {
          const stable = document.getElementById(
            "mathlive-suggestion-popover",
          );
          const items = [...(stable?.querySelectorAll("li[data-command]") ?? [])]
            .map((item) => ({
              command: item.dataset.command ?? "",
              latex: item.querySelector(".ML__popover__latex")?.textContent?.trim() ?? "",
              previewText: item.querySelector(".ML__popover__command")?.textContent?.trim() ?? "",
              hasVisualPreview: item.classList.contains(
                "has-visualtex-command-preview",
              ),
              previewMarkup:
                item.querySelector(".ML__popover__command")?.innerHTML ?? "",
            }));
          return {
            ready: Boolean(stable?.classList.contains("is-visible")) && items.length > 0,
            items,
          };
        })()`, `${prefix} native structure suggestions`);
    results.push({ prefix, items: state.items });
  }
  await clearField();
  const modelAudit = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const templates = [
          "\\\\overset{\\\\placeholder{}}{\\\\placeholder{}}",
          "\\\\underset{\\\\placeholder{}}{\\\\placeholder{}}",
          "\\\\overunderset{\\\\placeholder{}}{\\\\placeholder{}}{\\\\placeholder{}}",
          "\\\\stackrel{\\\\placeholder{}}{\\\\placeholder{}}",
        ];
        return templates.map((template) => {
          field.setValue("", { mode: "math", format: "latex" });
          field.insert(template, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceSelection",
            selectionMode: "placeholder",
          });
          const initialSelection = field.selection;
          field.insert("ab", {
            mode: "math",
            format: "latex",
            insertionMode: "replaceSelection",
            selectionMode: "after",
          });
          const valueAfterTyping = field.value;
          field.setValue(template, { mode: "math", format: "latex" });
          const offsets = [];
          for (let offset = 0; offset <= field.lastOffset; offset += 1) {
            const info = field.getElementInfo(offset);
            offsets.push({
              offset,
              latex: info?.latex ?? "",
              depth: info?.depth ?? null,
            });
          }
          return {
            template,
            value: field.value,
            initialSelection,
            valueAfterTyping,
            selection: field.selection,
            position: field.position,
            lastOffset: field.lastOffset,
            offsets,
          };
        });
      })()`);
  console.log(JSON.stringify({ results, modelAudit }, null, 2));
  console.log("Native structure suggestion audit passed");
  return;
}
