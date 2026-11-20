// common and custom formula tiles: insertion, adding, the context menu, deletion

export const names = ["formula-tiles"];

// state the page starts with (merged into the editor's persisted state)
export const storage = () => ({ editorLayout: "standard", sidebarOpen: true });

export async function run({ clearField, evaluate, waitForEvaluation }) {
  await evaluate(`(() => {
        localStorage.removeItem("visualtex-custom-formula-tiles");
        if (!document.querySelector(".formula-toolbar")) {
          document.querySelector(".sidebar-toggle")?.click();
        }
        return true;
      })()`);
  await waitForEvaluation(`(() => ({
        ready:
          Boolean(document.querySelector('[data-toolbar-view="tools"]')) &&
          Boolean(document.querySelector('[data-toolbar-view="tiles"]')) &&
          !document.querySelector(".formula-toolbar-actions") &&
          !document.querySelector(".add-formula-line"),
      }))()`, "formula tool and tile view tabs without legacy buttons");

  await evaluate(`document.querySelector('[data-toolbar-view="tiles"]').click()`);
  const commonTilesState = await waitForEvaluation(`(() => {
        const panel = document.querySelector(".formula-tiles-panel");
        const buttons = [...document.querySelectorAll(
          '.formula-tile-list > .formula-tile-item > .formula-tile-button',
        )];
        const previews = buttons.map((button) => {
          const host = button.querySelector(".formula-tile-preview");
          const content = host?.querySelector(".math-preview-fit-content");
          const hostBounds = host?.getBoundingClientRect();
          const contentBounds = content?.getBoundingClientRect();
          return {
            id: button.dataset.formulaTileId ?? "",
            fitReady: host?.dataset.fitReady === "true",
            scale: Number.parseFloat(host?.dataset.fitScale ?? "0"),
            buttonHeight: button.getBoundingClientRect().height,
            hostHeight: hostBounds?.height ?? 0,
            inside: Boolean(
              hostBounds &&
                contentBounds &&
                contentBounds.left >= hostBounds.left - 1 &&
                contentBounds.right <= hostBounds.right + 1 &&
                contentBounds.top >= hostBounds.top - 1 &&
                contentBounds.bottom <= hostBounds.bottom + 1,
            ),
          };
        });
        const heights = previews.map((preview) => Math.round(preview.hostHeight));
        return {
          ready:
            Boolean(panel) &&
            buttons.length === 10 &&
            previews.every(
              (preview) =>
                preview.fitReady &&
                preview.scale > 0 &&
                preview.inside &&
                preview.hostHeight >= 52 &&
                preview.buttonHeight >= preview.hostHeight,
            ) &&
            panel.scrollWidth <= panel.clientWidth + 1,
          count: buttons.length,
          heights,
          distinctHeightCount: new Set(heights).size,
          horizontalOverflow: panel
            ? panel.scrollWidth - panel.clientWidth
            : -1,
          previews,
        };
      })()`, "ten fitted common formula tiles");

  await clearField();
  await evaluate(`document.querySelector(
        '[data-formula-tile-id="quadratic-formula"]',
      ).click()`);
  const insertedTileState = await waitForEvaluation(`(() => {
        const field = document.querySelector(
          ".formula-line.is-active math-field",
        );
        return {
          ready:
            Boolean(field) &&
            field.value.includes("\\\\frac{-b\\\\pm\\\\sqrt{b^2-4ac}}{2a}"),
          value: field?.value ?? "",
        };
      })()`, "common formula tile insertion");

  await evaluate(`document.querySelector(
        '[data-tile-category="custom"]',
      ).click()`);
  await waitForEvaluation(`(() => ({
        ready: Boolean(document.querySelector('[data-open-custom-symbol-designer]')),
      }))()`, "custom symbol designer trigger");
  // The designer itself is covered by custom_symbol_designer_regression.mjs (both apps); its markup is
  // the shared custom-symbol-designer package's.

  // Registered custom symbols must expose a real, visible delete control in
  // the toolbar. The old top-right glyph could degrade into a '?' and there
  // was no reliable deletion path from the toolbar itself.
  await evaluate(`(() => {
        const key = 'visualtex.custom-symbols.v1';
        const value = {
          version: 1,
          symbols: [{
            id: 'qa-custom-symbol-delete',
            command: 'qaacustomglyph',
            name: 'QA custom glyph',
            role: 'ordinary',
            limitsBehavior: 'auto',
            metrics: { widthEm: 1, ascentEm: 1, descentEm: 0.1 },
            artwork: {
              shapes: [{
                kind: 'line', x1: 80, y1: 500, x2: 920, y2: 500,
                fill: false, strokeWidth: 80, lineCap: 'round',
              }],
            },
            ommlFallback: null,
            createdAt: Date.now(),
            updatedAt: Date.now(),
          }],
        };
        const oldValue = localStorage.getItem(key);
        const newValue = JSON.stringify(value);
        localStorage.setItem(key, newValue);
        window.dispatchEvent(new StorageEvent('storage', { key, oldValue, newValue }));
      })()`);
  const registeredDeleteState = await waitForEvaluation(`(() => {
        const item = document.querySelector('[data-registered-custom-symbol-toolbar-item="qa-custom-symbol-delete"]');
        const button = document.querySelector('[data-delete-registered-custom-symbol-toolbar="qa-custom-symbol-delete"]');
        const label = button?.textContent?.trim() ?? '';
        return {
          ready: Boolean(item && button && /删除|Delete/.test(label)),
          label,
          hasSvgIcon: Boolean(button?.querySelector('svg')),
        };
      })()`, "visible registered custom-symbol delete button");
  if (!registeredDeleteState.hasSvgIcon) {
    throw new Error(`Registered custom-symbol delete control has no SVG trash icon: ${JSON.stringify(registeredDeleteState)}`);
  }
  await evaluate(`document.querySelector('[data-delete-registered-custom-symbol-toolbar="qa-custom-symbol-delete"]')?.click()`);
  const firstDeleteClick = await waitForEvaluation(`(() => ({
        ready: document.querySelector('[data-delete-registered-custom-symbol-toolbar="qa-custom-symbol-delete"]')?.classList.contains('is-confirming') === true,
        stillPresent: Boolean(document.querySelector('[data-registered-custom-symbol-toolbar-item="qa-custom-symbol-delete"]')),
      }))()`, "custom-symbol delete confirmation");
  if (!firstDeleteClick.stillPresent) {
    throw new Error('Registered custom symbol was deleted without the required confirmation click');
  }
  await evaluate(`document.querySelector('[data-delete-registered-custom-symbol-toolbar="qa-custom-symbol-delete"]')?.click()`);
  await waitForEvaluation(`(() => {
        const stored = JSON.parse(localStorage.getItem('visualtex.custom-symbols.v1') || '{"symbols":[]}');
        return {
          ready: !document.querySelector('[data-registered-custom-symbol-toolbar-item="qa-custom-symbol-delete"]') &&
            !stored.symbols?.some?.((symbol) => symbol.id === 'qa-custom-symbol-delete'),
        };
      })()`, "delete registered custom symbol from toolbar");

  await waitForEvaluation(`(() => ({
        ready: Boolean(
          document.querySelector(".save-current-formula-tile:not(:disabled)"),
        ),
      }))()`, "enabled custom tile save button");
  await evaluate(`document.querySelector(
        ".save-current-formula-tile",
      ).click()`);
  const customTileState = await waitForEvaluation(`(() => {
        const button = document.querySelector(
          ".formula-tile-button.is-custom",
        );
        const stored = JSON.parse(
          localStorage.getItem("visualtex-custom-formula-tiles") || "{}",
        );
        const tiles = Array.isArray(stored.tiles) ? stored.tiles : [];
        const host = button?.querySelector(".formula-tile-preview");
        return {
          ready:
            Boolean(button) &&
            tiles.length === 1 &&
            tiles[0].latex.includes("\\\\frac{-b\\\\pm\\\\sqrt{b^2-4ac}}{2a}") &&
            host?.dataset.fitReady === "true",
          stored,
          latex: button?.dataset.formulaTileLatex ?? "",
          fitReady: host?.dataset.fitReady ?? "missing",
        };
      })()`, "persisted selected-line custom formula tile");
  await waitForEvaluation(`(() => ({
        ready: /已保存到|Saved to/.test(document.querySelector("[data-tile-save-notice]")?.textContent ?? ""),
      }))()`, "custom tile save notice");
  await evaluate(`document.querySelector(".save-current-formula-tile").click()`);
  await waitForEvaluation(`(() => {
        const stored = JSON.parse(localStorage.getItem("visualtex-custom-formula-tiles") || "{}");
        const notice = document.querySelector("[data-tile-save-notice]")?.textContent ?? "";
        return {
          ready: /磁贴已存在|Already a tile/.test(notice) && stored.tiles?.length === 1,
          notice,
          tiles: stored.tiles?.length,
        };
      })()`, "duplicate custom tile save says it already exists");

  await evaluate(`(() => {
        const button = document.querySelector(
          ".formula-tile-button.is-custom",
        );
        const field = document.querySelector(
          ".formula-line.is-active math-field",
        );
        if (!button || !field) return false;
        window.__visualtexFieldValueBeforeTileContextMenu = field.value;
        const bounds = button.getBoundingClientRect();
        button.dispatchEvent(new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          button: 2,
          buttons: 2,
          clientX: bounds.left + 24,
          clientY: bounds.top + 20,
        }));
        return true;
      })()`);
  const customTileContextState = await waitForEvaluation(`(() => {
        const menu = document.querySelector(".formula-tile-context-menu");
        const deleteButton = menu?.querySelector('[role="menuitem"]');
        const field = document.querySelector(
          ".formula-line.is-active math-field",
        );
        const stored = JSON.parse(
          localStorage.getItem("visualtex-custom-formula-tiles") || "{}",
        );
        const tiles = Array.isArray(stored.tiles) ? stored.tiles : [];
        return {
          ready:
            Boolean(menu && deleteButton) &&
            field?.value ===
              window.__visualtexFieldValueBeforeTileContextMenu &&
            tiles.length === 1,
          fieldValueBefore:
            window.__visualtexFieldValueBeforeTileContextMenu ?? "",
          fieldValue: field?.value ?? "",
          stored,
          menuText: deleteButton?.textContent?.trim() ?? "",
        };
      })()`, "custom tile right-click menu without insertion conflict");

  await evaluate(`document.querySelector(
        ".formula-tile-context-menu .formula-hotkey-context-action.is-danger",
      ).click()`);
  const deletedCustomTileState = await waitForEvaluation(`(() => {
        const stored = JSON.parse(
          localStorage.getItem("visualtex-custom-formula-tiles") || "{}",
        );
        const tiles = Array.isArray(stored.tiles) ? stored.tiles : [];
        return {
          ready:
            !document.querySelector(".formula-tile-button.is-custom") &&
            !document.querySelector(".formula-tile-context-menu") &&
            Boolean(document.querySelector(".custom-formula-section-empty")) &&
            tiles.length === 0,
          stored,
          emptyVisible: Boolean(document.querySelector(".custom-formula-section-empty")),
        };
      })()`, "right-click custom tile deletion persistence");

  console.log(
    JSON.stringify(
      {
        commonTilesState,
        insertedTileState,
        customTileState,
        customTileContextState,
        deletedCustomTileState,
      },
      null,
      2,
    ),
  );
  console.log("Targeted formula tiles regression passed");
  return;
}
