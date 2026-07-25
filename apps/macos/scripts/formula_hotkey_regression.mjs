import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const portOffset = process.pid % 1000;
const previewPort = 7600 + portOffset;
const debugPort = 12600 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = `/tmp/visualtex-formula-hotkeys-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while the process starts.
    }
    await sleep(100);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

class CdpClient {
  constructor(url) {
    this.url = url;
    this.nextId = 1;
    this.pending = new Map();
  }

  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((resolve, reject) => {
      this.socket.addEventListener("open", resolve, { once: true });
      this.socket.addEventListener("error", reject, { once: true });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      if (!message.id) return;
      const pending = this.pending.get(message.id);
      if (!pending) return;
      this.pending.delete(message.id);
      if (message.error) pending.reject(new Error(message.error.message));
      else pending.resolve(message.result);
    });
  }

  send(method, params = {}) {
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }

  close() {
    this.socket?.close();
  }
}

async function main() {
  const preview = spawn(
    process.execPath,
    [
      "node_modules/vite/bin/vite.js",
      "preview",
      "--host",
      "127.0.0.1",
      "--port",
      String(previewPort),
      "--strictPort",
    ],
    { cwd: process.cwd(), stdio: "ignore" },
  );
  let chrome;
  let client;

  try {
    await waitFor(baseUrl);
    chrome = spawn(
      chromePath,
      [
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`,
        `--user-data-dir=${chromeProfile}`,
        "--window-size=1440,1000",
        baseUrl,
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const page = targets.find(
      (target) => target.type === "page" && target.url.startsWith(baseUrl),
    );
    if (!page) throw new Error("No VisualTeX Chrome page target found");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
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
        JSON.stringify([
          "\\beta_{\\omega_1^2}",
          "\\beta",
          "\\int_b^a b",
          "\\int_b^a a",
          "\\int_b^a t",
          "\\frac{R}{Tf}\\int_b^a t",
          "\\frac{R}{Tf}\\int_b^a t\\,dpq",
          "\\frac{R}{Tf}\\int_b^a t\\,dp",
          "a^2+b^2=c^2",
        ]),
      );
    })()`);
    await reload();

    const openContextMenu = async (selector) => {
      const opened = await evaluate(`new Promise((resolve) => {
        const button = document.querySelector(${JSON.stringify(selector)});
        if (!button) {
          resolve({
            ok: false,
            selectorFound: false,
            templateCount: document.querySelectorAll(".template-button").length,
            tileCount: document.querySelectorAll(".formula-tile-button").length,
            bodyText: document.body.textContent?.slice(0, 200) ?? "",
          });
          return;
        }
        button.dispatchEvent(new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          clientX: 420,
          clientY: 320,
        }));
        setTimeout(() => resolve({
          ok: Boolean(document.querySelector(".formula-hotkey-context-menu")),
          selectorFound: true,
        }), 50);
      })`);
      assert.equal(opened.ok, true, `Context menu did not open for ${selector}: ${JSON.stringify(opened)}`);
    };

    const assignCurrentContext = async (code, key, modifiers) => {
      const assigned = await evaluate(`new Promise((resolve) => {
        const setButton = document.querySelector(".formula-hotkey-context-action");
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
      document.querySelector(".formula-hotkey-context-action")?.click();
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

    const managerState = await evaluate(`new Promise((resolve) => {
      document.querySelector(".settings-toggle")?.click();
      setTimeout(() => {
        document.querySelector(".settings-hotkey-button")?.click();
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
        }, 80);
      }, 80);
    })`);
    assert.equal(managerState.rows, 1);
    assert.ok(managerState.hotkey);
    assert.ok(managerState.keycapCenterDelta <= 1, JSON.stringify(managerState));
    await evaluate(`document.querySelector(".formula-hotkey-manager-dialog .dialog-header .icon-button")?.click()`);

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
        await new Promise((done) => setTimeout(done, 100));
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
                fitReady: preview?.dataset.fitReady === "true",
                unifiedFit: button.classList.contains("is-unified-fit"),
                wide: button.classList.contains("is-wide-preview"),
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
        assert.ok(detail.scale > 0 && detail.scale <= 1, JSON.stringify({ category, id, detail }));
        assert.equal(detail.contained, true, JSON.stringify({ category, id, detail }));
      }
    }
    assert.ok(expandedCategories.relation.ids.includes("triangleq"));
    assert.ok(expandedCategories.arrow.ids.includes("Longleftrightarrow"));
    assert.ok(expandedCategories.physics.ids.includes("outerproduct"));
    assert.ok(!expandedCategories.physics.ids.includes("time-ordering"));
    assert.equal(expandedCategories.physics.details.matrixelement.wide, true);
    assert.equal(expandedCategories.physics.details["expectation-operator"].wide, true);
    assert.ok(
      expandedCategories.physics.details.matrixelement.width >
        expandedCategories.physics.details.outerproduct.width * 1.5,
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
    await sleep(180);
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
        (button) => button.dataset.formulaTileLatex === "\\beta",
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
        version: library.version,
        storedTiles: library.tiles?.length || 0,
      };
    })()`);
    assert.equal(customTileLayout.count, 9);
    assert.equal(customTileLayout.version, 3);
    assert.equal(customTileLayout.storedTiles, 9);
    assert.ok(
      new Set(customTileLayout.widths).size >= 2,
      JSON.stringify(customTileLayout),
    );
    assert.ok(
      customTileLayout.widths.every((width) => width < customTileLayout.sectionWidth),
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
      customTileLayout.scales.every((scale) => scale >= 0.899 && scale <= 1.101),
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
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(350);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => {});
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
