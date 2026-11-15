// the painting of selections across formula rows

export const names = ["multi-line-selection"];

export async function run({ client, evaluate, sleep, waitForEvaluation }) {
  await evaluate(`(() => {
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        const lines = [
          { id: crypto.randomUUID(), latex: "abcdefghij" },
          { id: crypto.randomUUID(), latex: "klmnopqrst" },
          { id: crypto.randomUUID(), latex: "uvwxyzabcd" },
          { id: crypto.randomUUID(), latex: "efghijklmn" },
        ];
        persisted.state = {
          ...(persisted.state || {}),
          lines,
          activeLineId: lines[0].id,
          editorLayout: "standard",
          sidebarOpen: false,
        };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
      })()`);
  await client.send("Page.reload", { ignoreCache: true });
  await waitForEvaluation(`(() => ({
        ready: document.querySelectorAll("math-field").length === 4,
      }))()`, "four formula rows for multiline selection");

  const toolbar = await waitForEvaluation(`(() => {
        const selectors = [
          "[data-formula-selection-bold]",
          "[data-formula-selection-italic]",
          "[data-formula-selection-color]",
          "[data-formula-selection-background]",
        ];
        const buttons = selectors.map((selector) => document.querySelector(selector));
        const visible = buttons.every((button) => {
          if (!(button instanceof HTMLElement)) return false;
          const style = getComputedStyle(button);
          const rect = button.getBoundingClientRect();
          return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
        });
        return {
          ready: visible,
          count: buttons.filter(Boolean).length,
          visible,
          noHorizontalOverflow:
            document.documentElement.scrollWidth <= document.documentElement.clientWidth + 2,
        };
      })()`, "four selection-only formula formatting controls in the main editor");
  if (!toolbar.noHorizontalOverflow) {
    throw new Error(`Formula formatting controls caused horizontal overflow: ${JSON.stringify(toolbar)}`);
  }

  const drag = await waitForEvaluation(`(() => {
        const fields = [...document.querySelectorAll("math-field")];
        const firstContent = fields[0]?.shadowRoot?.querySelector('[part="content"]');
        const lastContent = fields.at(-1)?.shadowRoot?.querySelector('[part="content"]');
        const first = firstContent?.getBoundingClientRect();
        const last = lastContent?.getBoundingClientRect();
        return {
          ready: Boolean(first && last && first.width > 0 && last.width > 0),
          start: first ? { x: first.left + 2, y: first.top + first.height / 2 } : null,
          end: last ? { x: last.right - 2, y: last.top + last.height / 2 } : null,
        };
      })()`, "multiline selection drag geometry");

  await client.send("Input.dispatchMouseEvent", {
    type: "mousePressed",
    x: drag.start.x,
    y: drag.start.y,
    button: "left",
    buttons: 1,
    clickCount: 1,
  });
  await client.send("Input.dispatchMouseEvent", {
    type: "mouseMoved",
    x: drag.end.x,
    y: drag.end.y,
    button: "left",
    buttons: 1,
  });
  await client.send("Input.dispatchMouseEvent", {
    type: "mouseReleased",
    x: drag.end.x,
    y: drag.end.y,
    button: "left",
    buttons: 0,
    clickCount: 1,
  });

  const selectionState = await waitForEvaluation(`(() => {
        const rows = [...document.querySelectorAll(".formula-line")];
        const details = rows.map((row) => {
          const field = row.querySelector("math-field");
          const selections = [...(field?.shadowRoot?.querySelectorAll(".ML__selection") ?? [])]
            .filter((node) => {
              const style = getComputedStyle(node);
              const rect = node.getBoundingClientRect();
              return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
            });
          const rects = selections.map((node) => node.getBoundingClientRect());
          const union = rects.length
            ? {
                left: Math.min(...rects.map((rect) => rect.left)),
                right: Math.max(...rects.map((rect) => rect.right)),
                top: Math.min(...rects.map((rect) => rect.top)),
                bottom: Math.max(...rects.map((rect) => rect.bottom)),
              }
            : null;
          const fieldRect = field?.getBoundingClientRect();
          const rowRect = row.getBoundingClientRect();
          return {
            selectedClass: row.classList.contains("is-multi-line-selected"),
            selection: field ? JSON.parse(JSON.stringify(field.selection)) : null,
            visibleSelectionCount: selections.length,
            union,
            fieldRect: fieldRect
              ? { left: fieldRect.left, right: fieldRect.right, width: fieldRect.width }
              : null,
            rowWidth: rowRect.width,
            rowBackground: getComputedStyle(row).backgroundColor,
          };
        });
        const ready =
          details.length === 4 &&
          details.every((detail) =>
            detail.selectedClass &&
            detail.visibleSelectionCount > 0 &&
            detail.union &&
            detail.fieldRect &&
            detail.union.right > detail.union.left &&
            detail.union.left >= detail.fieldRect.left - 3 &&
            detail.union.right <= detail.fieldRect.right + 3 &&
            detail.union.right - detail.union.left < detail.rowWidth - 40
          );
        return { ready, details };
      })()`, "natural text-bounded highlight on every selected formula row");

  await sleep(300);
  const beforeMove = await evaluate(`(() =>
        [...document.querySelectorAll("math-field")].map((field) =>
          JSON.parse(JSON.stringify(field.selection)),
        )
      )()`);
  await client.send("Input.dispatchMouseEvent", {
    type: "mouseMoved",
    x: drag.end.x + 160,
    y: drag.end.y,
    button: "none",
    buttons: 0,
  });
  await sleep(120);
  const afterMove = await evaluate(`(() =>
        [...document.querySelectorAll("math-field")].map((field) =>
          JSON.parse(JSON.stringify(field.selection)),
        )
      )()`);
  if (JSON.stringify(afterMove) !== JSON.stringify(beforeMove)) {
    throw new Error(`Multiline selection changed after mouse release: ${JSON.stringify({ beforeMove, afterMove })}`);
  }

  console.log(JSON.stringify({ toolbar, selectionState }, null, 2));
  console.log("Targeted multiline selection visual regression passed");
  return;
}
