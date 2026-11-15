import assert from "node:assert/strict";
import process from "node:process";
import { withAppBrowser } from "../../shared/visualtex-app/testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const legacyCustomFormulaTiles = [
  String.raw`\beta_{\omega_1^2}`,
  String.raw`\beta`,
  String.raw`\int_b^a b`,
  String.raw`\int_b^a a`,
  String.raw`\int_b^a t`,
  String.raw`\frac{R}{Tf}\int_b^a t`,
  String.raw`\frac{R}{Tf}\int_b^a t\,dpq`,
  String.raw`\frac{R}{Tf}\int_b^a t\,dp`,
  "a^2+b^2=c^2",
];

async function main() {
  await withAppBrowser(async ({ baseUrl, client }) => {
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(700);

    const evaluate = async (expression) => {
      const result = await client.send("Runtime.evaluate", {
        expression,
        awaitPromise: true,
        returnByValue: true,
      });
      if (result.exceptionDetails) {
        throw new Error(
          result.exceptionDetails.exception?.description ||
            result.exceptionDetails.text ||
            "Runtime.evaluate failed",
        );
      }
      return result.result.value;
    };

    const reload = async () => {
      await client.send("Page.reload", { ignoreCache: true });
      await sleep(700);
      await evaluate(`new Promise((resolve) => {
        const done = () => document.querySelector("math-field") ? resolve(true) : setTimeout(done, 30);
        done();
      })`);
    };

    await evaluate(`(() => {
      localStorage.clear();
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem(
        "visualtex-custom-formula-tiles",
        ${JSON.stringify(JSON.stringify(legacyCustomFormulaTiles))},
      );
    })()`);
    await reload();

    const openContextMenu = async (selector) => {
      const target = await evaluate(`(() => {
        const button = document.querySelector(${JSON.stringify(selector)});
        if (!(button instanceof HTMLElement)) {
          return {
            ok: false,
            selectorFound: false,
            templateCount: document.querySelectorAll(".template-button").length,
            tileCount: document.querySelectorAll(".formula-tile-button").length,
            bodyText: document.body.textContent?.slice(0, 200) ?? "",
          };
        }
        const rect = button.getBoundingClientRect();
        return {
          ok: rect.width > 0 && rect.height > 0,
          selectorFound: true,
          x: rect.left + rect.width / 2,
          y: rect.top + rect.height / 2,
        };
      })()`);
      assert.equal(target.ok, true, `Context target unavailable for ${selector}: ${JSON.stringify(target)}`);
      const synthetic = await evaluate(`new Promise((resolve) => {
        const button = document.querySelector(${JSON.stringify(selector)});
        window.__visualtexContextDiagnostics = [];
        const record = (kind) => (event) => {
          const target = event.target;
          window.__visualtexContextDiagnostics.push({
            kind,
            tag: target?.tagName ?? null,
            className: target instanceof HTMLElement ? target.className : null,
            at: performance.now(),
          });
        };
        const onScroll = record("scroll");
        const onResize = record("resize");
        const onBlur = record("blur");
        const onPointer = record("pointerdown");
        window.addEventListener("scroll", onScroll, true);
        window.addEventListener("resize", onResize, true);
        window.addEventListener("blur", onBlur, true);
        document.addEventListener("pointerdown", onPointer, true);
        if (!(button instanceof HTMLElement)) {
          resolve({ dispatched: false, opened: false });
          return;
        }
        let appeared = false;
        const observer = new MutationObserver(() => {
          if (document.querySelector(".formula-hotkey-context-menu")) {
            appeared = true;
          }
        });
        observer.observe(document.getElementById("root") ?? document.body, {
          childList: true,
          subtree: true,
        });
        const dispatched = button.dispatchEvent(new MouseEvent("contextmenu", {
          bubbles: true,
          composed: true,
          cancelable: true,
          clientX: ${target.x},
          clientY: ${target.y},
          button: 2,
          buttons: 2,
        }));
        setTimeout(() => {
          observer.disconnect();
          window.removeEventListener("scroll", onScroll, true);
          window.removeEventListener("resize", onResize, true);
          window.removeEventListener("blur", onBlur, true);
          document.removeEventListener("pointerdown", onPointer, true);
          resolve({
            dispatched,
            appeared,
            opened: Boolean(document.querySelector(".formula-hotkey-context-menu")),
            diagnostics: window.__visualtexContextDiagnostics ?? [],
          });
        }, 120);
      })`);
      if (synthetic.opened) return;
      const hitBefore = await evaluate(`(() => {
        const hit = document.elementFromPoint(${target.x}, ${target.y});
        return {
          tag: hit?.tagName ?? null,
          className: hit instanceof HTMLElement ? hit.className : null,
          tileId: hit instanceof Element ? hit.closest(".formula-tile-button")?.getAttribute("data-formula-tile-id") ?? null : null,
          settingsOpen: Boolean(document.querySelector(".settings-dialog")),
        };
      })()`);
      await client.send("Input.dispatchMouseEvent", {
        type: "mousePressed",
        x: target.x,
        y: target.y,
        button: "right",
        buttons: 2,
        clickCount: 1,
      });
      await client.send("Input.dispatchMouseEvent", {
        type: "mouseReleased",
        x: target.x,
        y: target.y,
        button: "right",
        buttons: 0,
        clickCount: 1,
      });
      await sleep(80);
      const after = await evaluate(`(() => ({
        opened: Boolean(document.querySelector(".formula-hotkey-context-menu")),
        tileMenus: document.querySelectorAll(".formula-tile-context-menu").length,
        settingsOpen: Boolean(document.querySelector(".settings-dialog")),
      }))()`);
      assert.equal(after.opened, true, `Context menu did not open for ${selector}: ${JSON.stringify({ target, synthetic, hitBefore, after })}`);
    };

    // the context menu starts with the group actions; the hotkey action is found by its text
    const hotkeyAction = `[...document.querySelectorAll(".formula-hotkey-context-action")]
      .find((button) => /设置快捷键|修改快捷键|Set hotkey|Change hotkey/.test(button.textContent ?? ""))`;

    const assignCurrentContext = async (code, key, modifiers) => {
      const assigned = await evaluate(`new Promise((resolve) => {
        const setButton = ${hotkeyAction};
        if (!setButton) {
          resolve({ ok: false, step: "context-action" });
          return;
        }
        setButton.click();
        setTimeout(() => {
          const recorder = document.querySelector(".formula-hotkey-recorder-dialog");
          if (!recorder) {
            resolve({ ok: false, step: "recorder" });
            return;
          }
          document.dispatchEvent(new KeyboardEvent("keydown", {
            bubbles: true,
            cancelable: true,
            code: ${JSON.stringify(code)},
            key: ${JSON.stringify(key)},
            ctrlKey: ${Boolean(modifiers.ctrlKey)},
            altKey: ${Boolean(modifiers.altKey)},
            shiftKey: ${Boolean(modifiers.shiftKey)},
            metaKey: ${Boolean(modifiers.metaKey)},
          }));
          setTimeout(() => {
            const save = document.querySelector(".formula-hotkey-recorder-footer .primary-button");
            const disabled = save?.disabled ?? true;
            if (!save || disabled) {
              resolve({ ok: false, step: "save", disabled });
              return;
            }
            const cancel = document.querySelector(".formula-hotkey-recorder-footer .secondary-button");
            const keycap = document.querySelector(".formula-hotkey-capture-box kbd");
            const saveRect = save.getBoundingClientRect();
            const cancelRect = cancel?.getBoundingClientRect();
            const keycapStyle = keycap ? getComputedStyle(keycap) : null;
            const geometry = {
              buttonHeightDelta: cancelRect ? Math.abs(saveRect.height - cancelRect.height) : 999,
              keycapHeight: keycap?.getBoundingClientRect().height ?? 0,
              keycapDisplay: keycapStyle?.display ?? "",
              keycapLineHeight: keycapStyle?.lineHeight ?? "",
            };
            save.click();
            setTimeout(() => resolve({
              ok: true,
              geometry,
              bindings: JSON.parse(localStorage.getItem("visualtex-formula-hotkeys-v1") || "{}").state?.bindings?.length ?? 0,
            }), 30);
          }, 40);
        }, 50);
      })`);
      assert.equal(assigned.ok, true, JSON.stringify(assigned));
      assert.ok(assigned.geometry.buttonHeightDelta <= 0.5, JSON.stringify(assigned.geometry));
      assert.ok(assigned.geometry.keycapHeight >= 34, JSON.stringify(assigned.geometry));
      assert.ok(
        assigned.geometry.keycapDisplay === "inline-flex" ||
          assigned.geometry.keycapDisplay === "flex",
        JSON.stringify(assigned.geometry),
      );
      await sleep(120);
      return assigned;
    };

    const pressInFormula = async (code, key, modifiers) =>
      evaluate(`new Promise((resolve) => {
        const field = document.querySelector("math-field");
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        field.dispatchEvent(new KeyboardEvent("keydown", {
          bubbles: true,
          composed: true,
          cancelable: true,
          code: ${JSON.stringify(code)},
          key: ${JSON.stringify(key)},
          ctrlKey: ${Boolean(modifiers.ctrlKey)},
          altKey: ${Boolean(modifiers.altKey)},
          shiftKey: ${Boolean(modifiers.shiftKey)},
          metaKey: ${Boolean(modifiers.metaKey)},
        }));
        setTimeout(() => resolve(field.value), 120);
      })`);

    await openContextMenu('.template-button[data-command-id="frac"]');
    await assignCurrentContext("KeyF", "f", {
      ctrlKey: true,
      altKey: true,
    });
    const afterFraction = await pressInFormula("KeyF", "f", {
      ctrlKey: true,
      altKey: true,
    });
    assert.match(afterFraction, /\\frac/);

    await openContextMenu('.template-button[data-command-id="frac"]');
    const protectedState = await evaluate(`new Promise((resolve) => {
      ${hotkeyAction}?.click();
      setTimeout(() => {
        document.dispatchEvent(new KeyboardEvent("keydown", {
          bubbles: true,
          cancelable: true,
          code: "KeyS",
          key: "s",
          metaKey: true,
        }));
        setTimeout(() => {
          const save = document.querySelector(".formula-hotkey-recorder-footer .primary-button");
          resolve({
            disabled: save?.disabled ?? false,
            warning: document.querySelector(".formula-hotkey-message.is-danger")?.textContent ?? "",
          });
        }, 40);
      }, 50);
    })`);
    assert.equal(protectedState.disabled, true);
    assert.match(protectedState.warning, /保存|Save/);
    await evaluate(`document.querySelector(".formula-hotkey-recorder-dialog .dialog-header .icon-button")?.click()`);

    // the formula hotkeys are listed in Settings > Shortcuts > Editor
    const managerState = await evaluate(`new Promise((resolve) => {
      document.querySelector(".settings-toggle")?.click();
      setTimeout(() => {
        document.querySelector('[data-settings-page="shortcuts"]')?.click();
        setTimeout(() => document.querySelector('[data-shortcut-tab="editor"]')?.click(), 40);
        setTimeout(() => {
          const row = document.querySelector(".formula-hotkey-binding-row");
          const keycap = row?.querySelector(":scope > kbd");
          const rowRect = row?.getBoundingClientRect();
          const keycapRect = keycap?.getBoundingClientRect();
          resolve({
            rows: document.querySelectorAll(".formula-hotkey-binding-row").length,
            hotkey: keycap?.textContent ?? "",
            keycapCenterDelta: rowRect && keycapRect
              ? Math.abs(
                  rowRect.top + rowRect.height / 2 -
                  (keycapRect.top + keycapRect.height / 2),
                )
              : 999,
          });
        }, 160);
      }, 80);
    })`);
    assert.equal(managerState.rows, 10);
    assert.ok(managerState.hotkey);
    assert.ok(managerState.keycapCenterDelta <= 1, JSON.stringify(managerState));
    await evaluate(`document.querySelector(".settings-dialog .dialog-header .icon-button")?.click()`);
    await sleep(80);

    const expandedCategories = await evaluate(`new Promise(async (resolve) => {
      const result = {};
      const categories = [
        "common",
        "structure",
        "calculus",
        "matrix",
        "relation",
        "greek",
        "arrow",
        "physics",
        "set",
      ];
      for (const category of categories) {
        document.querySelector('[data-category="' + category + '"]')?.click();
        await document.fonts?.ready;
        await new Promise((done) => {
          const startedAt = Date.now();
          const waitForStablePreviews = () => {
            const previews = Array.from(
              document.querySelectorAll(".template-button .math-preview"),
            );
            const ready =
              previews.length > 0 &&
              previews.every((preview) =>
                preview.dataset.fitReady === "true" ||
                preview.dataset.fitReady === "static"
              );
            if (ready || Date.now() - startedAt > 2000) {
              requestAnimationFrame(() =>
                requestAnimationFrame(() => requestAnimationFrame(done)),
              );
              return;
            }
            setTimeout(waitForStablePreviews, 30);
          };
          waitForStablePreviews();
        });
        const buttons = Array.from(document.querySelectorAll(".template-button"));
        result[category] = {
          count: buttons.length,
          unifiedFitCount: buttons.filter((button) =>
            button.classList.contains("is-unified-fit")
          ).length,
          errorCount: buttons.reduce(
            (count, button) => count + button.querySelectorAll(".ML__error").length,
            0,
          ),
          ids: buttons.map((button) => button.dataset.commandId || ""),
          details: Object.fromEntries(buttons.map((button) => {
            const preview = button.querySelector(".math-preview");
            const content = preview?.querySelector(".math-preview-fit-content");
            const buttonRect = button.getBoundingClientRect();
            const previewRect = preview?.getBoundingClientRect();
            const contentRect = content?.getBoundingClientRect();
            const tolerance = 1.5;
            return [
              button.dataset.commandId || "",
              {
                width: Math.round(buttonRect.width),
                fontSize: Number.parseFloat(getComputedStyle(preview).fontSize),
                scale: Number.parseFloat(preview?.dataset.fitScale || "0"),
                fitReady:
                  preview?.dataset.fitReady === "true" ||
                  preview?.dataset.fitReady === "static",
                unifiedFit: button.classList.contains("is-unified-fit"),
                wide: button.classList.contains("is-wide-preview"),
                previewRect: previewRect
                  ? {
                      left: previewRect.left,
                      right: previewRect.right,
                      top: previewRect.top,
                      bottom: previewRect.bottom,
                      width: previewRect.width,
                      height: previewRect.height,
                    }
                  : null,
                contentRect: contentRect
                  ? {
                      left: contentRect.left,
                      right: contentRect.right,
                      top: contentRect.top,
                      bottom: contentRect.bottom,
                      width: contentRect.width,
                      height: contentRect.height,
                    }
                  : null,
                contained: Boolean(
                  previewRect &&
                  contentRect &&
                  contentRect.left >= previewRect.left - tolerance &&
                  contentRect.right <= previewRect.right + tolerance &&
                  contentRect.top >= previewRect.top - tolerance &&
                  contentRect.bottom <= previewRect.bottom + tolerance
                ),
              },
            ];
          })),
        };
      }
      resolve(result);
    })`);
    assert.ok(expandedCategories.relation.count >= 39, JSON.stringify(expandedCategories));
    assert.ok(expandedCategories.arrow.count >= 36, JSON.stringify(expandedCategories));
    assert.ok(expandedCategories.physics.count >= 45, JSON.stringify(expandedCategories));
    for (const [category, result] of Object.entries(expandedCategories)) {
      assert.equal(result.unifiedFitCount, result.count, JSON.stringify({ category, result }));
      assert.equal(result.errorCount, 0, JSON.stringify({ category, result }));
      for (const [id, detail] of Object.entries(result.details)) {
        assert.equal(detail.unifiedFit, true, JSON.stringify({ category, id, detail }));
        assert.equal(detail.fontSize, 24, JSON.stringify({ category, id, detail }));
        assert.equal(detail.fitReady, true, JSON.stringify({ category, id, detail }));
        const maximumScale = Math.min(1.55, Math.max(1, detail.width / 42));
        assert.ok(
          detail.scale > 0 && detail.scale <= maximumScale + 0.001,
          JSON.stringify({ category, id, detail, maximumScale }),
        );
        assert.equal(detail.contained, true, JSON.stringify({ category, id, detail }));
      }
    }
    assert.ok(expandedCategories.relation.ids.includes("triangleq"));
    assert.ok(expandedCategories.arrow.ids.includes("Longleftrightarrow"));
    assert.ok(expandedCategories.physics.ids.includes("outerproduct"));
    assert.ok(!expandedCategories.physics.ids.includes("time-ordering"));
    assert.equal(expandedCategories.physics.details.matrixelement.wide, true);
    assert.equal(expandedCategories.physics.details["expectation-operator"].wide, true);
    // two-argument products take a wide tile too (\outerproduct, as on macOS since 2026-09-25)
    assert.equal(expandedCategories.physics.details.outerproduct.wide, true);
    assert.equal(expandedCategories.physics.details.ket.wide, false);
    assert.ok(
      expandedCategories.physics.details.matrixelement.width >
        expandedCategories.physics.details.ket.width * 1.5,
      JSON.stringify(expandedCategories.physics.details),
    );
    for (const commandId of ["intplain", "int", "iint", "sum", "series", "prod"]) {
      assert.equal(
        expandedCategories.calculus.details[commandId].contained,
        true,
        JSON.stringify({ commandId, detail: expandedCategories.calculus.details[commandId] }),
      );
    }

    await evaluate(`(() => {
      document.querySelector('[data-toolbar-view="tiles"]')?.click();
      document.querySelector('[data-tile-category="common"]')?.click();
    })()`);
    await openContextMenu('.formula-tile-button[data-formula-tile-id="mass-energy"]');
    await assignCurrentContext("KeyE", "e", {
      ctrlKey: true,
      altKey: true,
    });
    const afterCommonTile = await pressInFormula("KeyE", "e", {
      ctrlKey: true,
      altKey: true,
    });
    assert.match(afterCommonTile, /E=mc\^2/);

    await evaluate(`document.querySelector('[data-tile-category="custom"]')?.click()`);
    await evaluate(`new Promise((resolve, reject) => {
      const startedAt = Date.now();
      const waitForTiles = () => {
        const count = document.querySelectorAll(".formula-tile-button.is-custom").length;
        if (count === 9) {
          resolve(true);
          return;
        }
        if (Date.now() - startedAt > 2000) {
          reject(new Error("Timed out waiting for custom formula tiles: " + count + "/9"));
          return;
        }
        setTimeout(waitForTiles, 30);
      };
      waitForTiles();
    })`);
    const customTileLayout = await evaluate(`(() => {
      const buttons = Array.from(document.querySelectorAll(".formula-tile-button.is-custom"));
      const rects = buttons.map((button) => button.getBoundingClientRect());
      const library = JSON.parse(
        localStorage.getItem("visualtex-custom-formula-tiles") || "{}"
      );
      const section = document.querySelector(".custom-formula-section");
      const grid = document.querySelector(".custom-formula-tile-grid");
      const list = document.querySelector(".formula-tile-list.is-custom");
      const sectionStyle = section ? getComputedStyle(section) : null;
      const listStyle = list ? getComputedStyle(list) : null;
      const gridRect = grid?.getBoundingClientRect();
      const betaButton = buttons.find(
        (button) =>
          button.dataset.formulaTileLatex === ${JSON.stringify(String.raw`\beta`)},
      );
      const betaRect = betaButton?.getBoundingClientRect();
      const firstTop = Math.min(...rects.map((rect) => Math.round(rect.top)));
      const firstRow = rects.filter((rect) => Math.abs(rect.top - firstTop) <= 1);
      return {
        count: buttons.length,
        widths: rects.map((rect) => Math.round(rect.width)),
        tops: rects.map((rect) => Math.round(rect.top)),
        weights: buttons.map((button) => Number(button.dataset.customTileWeight || 0)),
        scales: buttons.map((button) => Number(
          button.querySelector(".math-preview")?.dataset.fitScale || 1,
        )),
        betaWidth: Math.round(betaRect?.width || 0),
        betaWeight: Number(betaButton?.dataset.customTileWeight || 0),
        sectionWidth: Math.round(section?.getBoundingClientRect().width || 0),
        firstRowCount: firstRow.length,
        firstRowFill: gridRect && firstRow.length
          ? (Math.max(...firstRow.map((rect) => rect.right)) - gridRect.left) / gridRect.width
          : 0,
        sectionBorderWidth: sectionStyle?.borderTopWidth || "",
        listPaddingLeft: Number.parseFloat(listStyle?.paddingLeft || "999"),
        listPaddingRight: Number.parseFloat(listStyle?.paddingRight || "999"),
        // the seeded string array is migrated in memory and stored as version 3 at the first edit
        version: Array.isArray(library) ? "array" : library.version,
        storedTiles: Array.isArray(library) ? library.length : library.tiles?.length || 0,
      };
    })()`);
    assert.equal(customTileLayout.count, 9);
    assert.ok(["array", 3].includes(customTileLayout.version), JSON.stringify(customTileLayout));
    assert.equal(customTileLayout.storedTiles, 9);
    assert.ok(
      new Set(customTileLayout.widths).size >= 2,
      JSON.stringify(customTileLayout),
    );
    assert.ok(
      customTileLayout.widths.every(
        (width) => width <= customTileLayout.sectionWidth,
      ),
      JSON.stringify(customTileLayout),
    );
    assert.ok(
      customTileLayout.tops.some(
        (top, index) => customTileLayout.tops.indexOf(top) !== index,
      ),
      JSON.stringify(customTileLayout),
    );
    assert.ok(customTileLayout.firstRowCount >= 4, JSON.stringify(customTileLayout));
    assert.ok(customTileLayout.firstRowFill >= 0.82, JSON.stringify(customTileLayout));
    assert.equal(customTileLayout.betaWeight, 1, JSON.stringify(customTileLayout));
    assert.ok(customTileLayout.betaWidth <= 32, JSON.stringify(customTileLayout));
    assert.ok(
      customTileLayout.scales.every((scale) => scale >= 0.799 && scale <= 1.201),
      JSON.stringify(customTileLayout),
    );
    assert.equal(customTileLayout.sectionBorderWidth, "0px");
    assert.ok(customTileLayout.listPaddingLeft <= 4, JSON.stringify(customTileLayout));
    assert.ok(customTileLayout.listPaddingRight <= 4, JSON.stringify(customTileLayout));

    const renamedDefaultSection = await evaluate(`new Promise((resolve) => {
      document.querySelector(
        '.custom-formula-section:first-child .custom-formula-section-actions button[title="重命名分区"]',
      )?.click();
      setTimeout(() => {
        const input = document.querySelector(".custom-formula-section-editor input");
        const setter = Object.getOwnPropertyDescriptor(
          HTMLInputElement.prototype,
          "value",
        )?.set;
        setter?.call(input, "基础");
        input?.dispatchEvent(new Event("input", { bubbles: true }));
        setTimeout(() => {
          document.querySelector(".custom-formula-section-editor .icon-button")?.click();
          setTimeout(() => {
            const library = JSON.parse(
              localStorage.getItem("visualtex-custom-formula-tiles") || "{}"
            );
            resolve({
              name: library.sections?.[0]?.name || "",
              header: document.querySelector(
                ".custom-formula-section:first-child .custom-formula-section-select strong",
              )?.textContent || "",
            });
          }, 80);
        }, 30);
      }, 50);
    })`);
    assert.equal(renamedDefaultSection.name, "基础");
    assert.equal(renamedDefaultSection.header, "基础");

    const createdSection = await evaluate(`new Promise((resolve) => {
      document.querySelector(".create-formula-tile-section")?.click();
      setTimeout(() => {
        const input = document.querySelector(".custom-formula-section-editor input");
        const setter = Object.getOwnPropertyDescriptor(
          HTMLInputElement.prototype,
          "value",
        )?.set;
        setter?.call(input, "G1");
        input?.dispatchEvent(new Event("input", { bubbles: true }));
        setTimeout(() => {
          document.querySelector(".custom-formula-section-editor .icon-button")?.click();
          setTimeout(() => {
            const library = JSON.parse(
              localStorage.getItem("visualtex-custom-formula-tiles") || "{}"
            );
            resolve({
              sectionCount: document.querySelectorAll(".custom-formula-section").length,
              storedSections: library.sections?.length || 0,
              sectionName: library.sections?.at(-1)?.name || "",
            });
          }, 80);
        }, 30);
      }, 50);
    })`);
    assert.equal(createdSection.sectionCount, 2);
    assert.equal(createdSection.storedSections, 2);
    assert.equal(createdSection.sectionName, "G1");

    await openContextMenu('.formula-tile-button[data-formula-tile-latex="a^2+b^2=c^2"]');
    const customizedTile = await evaluate(`new Promise((resolve) => {
      document.querySelector('[data-custom-tile-color="#ca6f7b"]')?.click();
      const select = document.querySelector(".custom-tile-section-picker select");
      const option = select?.querySelector("option:last-child");
      const setter = Object.getOwnPropertyDescriptor(
        HTMLSelectElement.prototype,
        "value",
      )?.set;
      if (select && option) {
        setter?.call(select, option.value);
        select.dispatchEvent(new Event("change", { bubbles: true }));
      }
      setTimeout(() => {
        const tile = document.querySelector(
          '.formula-tile-button[data-formula-tile-latex="a^2+b^2=c^2"]',
        );
        const library = JSON.parse(
          localStorage.getItem("visualtex-custom-formula-tiles") || "{}"
        );
        const stored = library.tiles?.find((item) => item.latex === "a^2+b^2=c^2");
        resolve({
          color: stored?.color || "",
          sectionId: stored?.sectionId || "",
          renderedSectionId: tile?.closest(".custom-formula-section")?.dataset.customSectionId || "",
          hasColorClass: tile?.classList.contains("has-custom-color") || false,
        });
      }, 100);
    })`);
    assert.equal(customizedTile.color, "#ca6f7b");
    assert.ok(customizedTile.sectionId);
    assert.equal(customizedTile.sectionId, customizedTile.renderedSectionId);
    assert.equal(customizedTile.hasColorClass, true);
    await evaluate(`document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }))`);

    await openContextMenu('.formula-tile-button[data-formula-tile-latex="a^2+b^2=c^2"]');
    await assignCurrentContext("KeyP", "p", {
      ctrlKey: true,
      altKey: true,
    });
    const afterCustomTile = await pressInFormula("KeyP", "p", {
      ctrlKey: true,
      altKey: true,
    });
    assert.match(afterCustomTile, /a\^2\+b\^2=c\^2/);

    const deletedG1 = await evaluate(`new Promise((resolve) => {
      const section = Array.from(document.querySelectorAll(".custom-formula-section"))
        .find((item) => item.querySelector(".custom-formula-section-select strong")?.textContent === "G1");
      section?.querySelector(".custom-formula-section-actions button.is-danger")?.click();
      setTimeout(() => {
        const confirmationVisible = Boolean(
          section?.querySelector(".custom-formula-section-delete-copy"),
        );
        section?.querySelector(".custom-formula-section-actions button.is-danger")?.click();
        setTimeout(() => {
          const library = JSON.parse(
            localStorage.getItem("visualtex-custom-formula-tiles") || "{}"
          );
          const hotkeys = JSON.parse(
            localStorage.getItem("visualtex-formula-hotkeys-v1") || "{}"
          );
          const bindings = hotkeys.state?.bindings || [];
          resolve({
            confirmationVisible,
            sectionNames: (library.sections || []).map((item) => item.name),
            tileExists: (library.tiles || []).some(
              (item) => item.latex === "a^2+b^2=c^2",
            ),
            bindingExists: bindings.some(
              (binding) => binding.target?.command?.insertTemplate === "a^2+b^2=c^2",
            ),
          });
        }, 100);
      }, 50);
    })`);
    assert.equal(deletedG1.confirmationVisible, true);
    assert.deepEqual(deletedG1.sectionNames, ["基础"]);
    assert.equal(deletedG1.tileExists, false);
    assert.equal(deletedG1.bindingExists, false);

    const deletedDefaultSection = await evaluate(`new Promise((resolve) => {
      const section = document.querySelector(".custom-formula-section");
      section?.querySelector(".custom-formula-section-actions button.is-danger")?.click();
      setTimeout(() => {
        section?.querySelector(".custom-formula-section-actions button.is-danger")?.click();
        setTimeout(() => {
          const library = JSON.parse(
            localStorage.getItem("visualtex-custom-formula-tiles") || "{}"
          );
          resolve({
            sections: library.sections?.length || 0,
            tiles: library.tiles?.length || 0,
            renderedSections: document.querySelectorAll(".custom-formula-section").length,
            saveDisabled: document.querySelector(".save-current-formula-tile")?.disabled ?? false,
            saveCopy: document.querySelector(".save-current-formula-tile")?.textContent || "",
          });
        }, 100);
      }, 50);
    })`);
    assert.equal(deletedDefaultSection.sections, 0);
    assert.equal(deletedDefaultSection.tiles, 0);
    assert.equal(deletedDefaultSection.renderedSections, 0);
    assert.equal(deletedDefaultSection.saveDisabled, true);
    assert.match(deletedDefaultSection.saveCopy, /新建分区|Create a section/);

    console.log("Formula hotkey regression passed");
  }, { windowSize: "1440,1000" });
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
