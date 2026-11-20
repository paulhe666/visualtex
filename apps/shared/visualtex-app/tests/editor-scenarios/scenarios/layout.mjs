// the header's file and edit actions, the export placement and the matrix size picker

export const names = ["layout"];

export async function run({ client, evaluate, waitForEvaluation }) {
  const toolbarOrder = await waitForEvaluation(`(() => {
      const group = document.querySelector(".canvas-tool-group");
      const exportMenu = group?.querySelector(".export-menu");
      const behavior = group?.querySelector(".input-behavior-menu");
      const children = [...(group?.children ?? [])];
      const fileActions = document.querySelector(".header-actions .file-actions");
      const editActions = document.querySelector(".header-actions .edit-actions");
      const titleInput = document.querySelector(".document-title-area input");
      const fileStyle = fileActions ? getComputedStyle(fileActions) : null;
      const fileBounds = fileActions?.getBoundingClientRect();
      const editBounds = editActions?.getBoundingClientRect();
      const titleBounds = titleInput?.getBoundingClientRect();
      return {
        ready: Boolean(behavior && fileActions && editActions && titleInput),
        exportIndex: children.indexOf(exportMenu),
        behaviorIndex: children.indexOf(behavior),
        fileBorderWidth: fileStyle?.borderTopWidth ?? "",
        fileBackground: fileStyle?.backgroundColor ?? "",
        fileLeftOffset: fileStyle?.left ?? "",
        titleRight: titleBounds?.right ?? 0,
        fileLeft: fileBounds?.left ?? 0,
        fileRight: fileBounds?.right ?? 0,
        editLeft: editBounds?.left ?? 0,
      };
    })()`, "unified export placement and shifted file actions");
  if (
    toolbarOrder.behaviorIndex < 0 ||
    (toolbarOrder.exportIndex >= 0 &&
      toolbarOrder.exportIndex >= toolbarOrder.behaviorIndex) ||
    toolbarOrder.fileBorderWidth !== "0px" ||
    (toolbarOrder.exportIndex >= 0 && toolbarOrder.fileLeftOffset !== "6px") ||
    toolbarOrder.titleRight > toolbarOrder.fileLeft ||
    toolbarOrder.editLeft - toolbarOrder.fileRight < 4
  ) {
    throw new Error(`Incorrect export/header placement: ${JSON.stringify(toolbarOrder)}`);
  }

  let exportMenuState = null;
  if (toolbarOrder.exportIndex >= 0) {
    await evaluate(`document.querySelector(".export-menu-trigger")?.click()`);
    exportMenuState = await waitForEvaluation(`(() => {
        const popover = document.querySelector(".export-menu-popover");
        const labels = [...document.querySelectorAll(".export-format-options strong")]
          .map((node) => node.textContent?.trim() ?? "");
        const pathSection = document.querySelector(".export-path-section");
        return {
          ready:
            Boolean(popover && pathSection) &&
            labels.join(",") === "Markdown,SVG,PNG",
          labels,
          pathText: pathSection?.textContent?.replace(/\\s+/g, " ").trim() ?? "",
        };
      })()`, "unified export menu options and path selector");
    await evaluate(`document.querySelector(".export-menu-trigger")?.click()`);
    await waitForEvaluation(`(() => ({
        ready: !document.querySelector(".export-menu-popover"),
      }))()`, "export menu closed before matrix test");
  }

  await evaluate(`document.querySelector('button[data-category="matrix"]').click()`);
  const gridState = await waitForEvaluation(`(() => ({
      // the 5×5 picker of the shared toolbar (Windows since 2026-10-14)
      ready: document.querySelectorAll(".matrix-size-cell").length === 25,
      cellCount: document.querySelectorAll(".matrix-size-cell").length,
    }))()`, "5 by 5 matrix grid");

  const matrixHoverPoint = await evaluate(`(() => {
      const cell = document.querySelector('.matrix-size-cell[data-matrix-rows="3"][data-matrix-columns="4"]');
      const rect = cell?.getBoundingClientRect();
      return rect
        ? { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 }
        : null;
    })()`);
  if (!matrixHoverPoint) throw new Error("Missing 3×4 matrix cell");
  await client.send("Input.dispatchMouseEvent", {
    type: "mouseMoved",
    x: Math.max(1, matrixHoverPoint.x - 30),
    y: matrixHoverPoint.y,
  });
  await client.send("Input.dispatchMouseEvent", {
    type: "mouseMoved",
    x: matrixHoverPoint.x,
    y: matrixHoverPoint.y,
  });
  await evaluate(`(() => {
      const cell = document.querySelector('.matrix-size-cell[data-matrix-rows="3"][data-matrix-columns="4"]');
      cell?.dispatchEvent(new PointerEvent("pointerover", {
        bubbles: true,
        composed: true,
        pointerId: 1,
        pointerType: "mouse",
        isPrimary: true,
      }));
      return true;
    })()`);
  const hoverState = await waitForEvaluation(`(() => ({
      ready:
        document.querySelector(".matrix-size-badge")?.textContent?.replace(/\\s+/g, " ").trim() === "3 × 4" &&
        document.querySelectorAll(".matrix-size-cell.is-previewed").length === 12,
      badge: document.querySelector(".matrix-size-badge")?.textContent?.replace(/\\s+/g, " ").trim() ?? "",
      previewedCount: document.querySelectorAll(".matrix-size-cell.is-previewed").length,
    }))()`, "matrix hover preview");

  await evaluate(`document.querySelector('.matrix-size-cell[data-matrix-rows="3"][data-matrix-columns="4"]').click()`);
  const selectedState = await waitForEvaluation(`(() => ({
      ready:
        document.querySelector('.matrix-size-cell[data-matrix-rows="3"][data-matrix-columns="4"]')?.classList.contains("is-selected-corner") &&
        document.querySelector(".matrix-insert-button")?.textContent?.includes("3 × 4"),
      selectedCorner: document.querySelector('.matrix-size-cell[data-matrix-rows="3"][data-matrix-columns="4"]')?.classList.contains("is-selected-corner") ?? false,
      insertLabel: document.querySelector(".matrix-insert-button")?.textContent ?? "",
    }))()`, "matrix selection");

  await evaluate(`document.querySelector(".matrix-insert-button").click()`);
  const insertionState = await waitForEvaluation(`(() => {
      const field = document.querySelector(".formula-line.is-active math-field");
      const value = field?.value ?? "";
      const body = value.match(/\\\\begin\\{bmatrix\\}([\\s\\S]*?)\\\\end\\{bmatrix\\}/)?.[1] ?? "";
      const root = field?.shadowRoot;
      const symbol = field?.placeholderSymbol || "▢";
      const placeholders = [...(root?.querySelectorAll(
        ".visualtex-structural-placeholder",
      ) ?? [])];
      const rawLeaves = [...(root?.querySelectorAll("[data-atom-id]") ?? [])]
        .filter(
          (node) =>
            (node.textContent || "").trim() === symbol &&
            !node.querySelector("[data-atom-id]") &&
            !node.classList.contains("visualtex-structural-placeholder"),
        );
      return {
        ready:
          value.includes("\\\\begin{bmatrix}") &&
          body.split(/\\\\\\\\/).length === 3 &&
          body.split(/\\\\\\\\/).every((row) => row.split("&").length === 4) &&
          placeholders.length === 12 &&
          rawLeaves.length === 0,
        value,
        placeholderCount: placeholders.length,
        rawPlaceholderCount: rawLeaves.length,
        placeholderClasses: [...new Set(placeholders.map((node) => node.className))],
      };
    })()`, "3 by 4 matrix insertion with VisualTeX placeholders");

  console.log(
    JSON.stringify(
      {
        toolbarOrder,
        exportMenuState,
        gridState,
        hoverState,
        selectedState,
        insertionState,
      },
      null,
      2,
    ),
  );
  console.log("Targeted layout regression passed");
}
