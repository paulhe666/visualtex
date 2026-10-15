// Real-browser audit of the command candidate panel: per-row geometry
// (previews spilling into neighbouring rows), missing glyphs (tofu / blank
// ink), "?" fallbacks, and the ranked command order for each query.
//
//   node scripts/candidate_panel_audit.mjs [--out=report.json] [--queries=\int,\sum]
//
// Needs a fresh `npm run build:desktop` (vite preview serves dist/).
import { spawn } from "node:child_process";
import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const argument = (name) =>
  process.argv.find((value) => value.startsWith(`--${name}=`))?.slice(name.length + 3);
const outPath = argument("out");
const shotDir = argument("shots");
const probeFile = argument("probe");
const offset = process.pid % 1000;
const previewPort = 18600 + offset;
const debugPort = 23600 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = createBrowserProfilePath("visualtex-candidate-audit");
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const defaultQueries = [
  ..."abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ".split("").map((c) => `\\${c}`),
  "\\in", "\\int", "\\ii", "\\oi", "\\su", "\\sum", "\\pr", "\\prod", "\\co",
  "\\li", "\\lim", "\\fr", "\\frac", "\\sq", "\\be", "\\beg", "\\ma", "\\math",
  "\\ov", "\\over", "\\un", "\\under", "\\x", "\\le", "\\left", "\\ri", "\\bi",
  "\\big", "\\to", "\\ar", "\\al", "\\de", "\\pa", "\\ve", "\\ha", "\\ce",
  "\\te", "\\op", "\\bo", "\\ca", "\\ne", "\\no", "\\si", "\\ph", "\\ps",
  "\\ra", "\\la", "\\Ri", "\\Le", "\\lo", "\\up", "\\do", "\\eq", "\\ap",
  "\\sim", "\\ci", "\\ti", "\\cd", "\\ld", "\\vd", "\\dd", "\\inf", "\\pm",
  "\\mp", "\\ge", "\\ll", "\\gg", "\\sub", "\\sup", "\\cup", "\\cap", "\\vee",
  "\\wedge", "\\ex", "\\fo", "\\ang", "\\tri", "\\par", "\\perp", "\\na",
];
const letters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
const sweepQueries = [...letters].flatMap((first) =>
  [..."abcdefghijklmnopqrstuvwxyz"].map((second) => `\\${first}${second}`),
);
const queries = argument("queries")?.split(",").filter(Boolean) ??
  (process.argv.includes("--sweep") ? [...defaultQueries, ...sweepQueries] : defaultQueries);

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
      const pending = message.id && this.pending.get(message.id);
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

async function waitFor(url, timeoutMs = 20_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      if ((await fetch(url)).ok) return;
    } catch {
      // still starting
    }
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

// Runs inside the page. Returns null until the visible panel matches `query`.
const inspectPanel = String.raw`(query) => {
  const panel = document.getElementById("visualtex-native-input-suggestion-popover");
  if (!panel?.classList.contains("is-visible")) return null;
  const field = document.querySelector("math-field.visual-mathfield");
  const raw = [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
    .filter((node) => !node.classList.contains("ML__suggestion"))
    .map((node) => node.textContent ?? "").join("");
  if (raw !== query) return null;
  const items = [...panel.querySelectorAll("li[data-command]")];
  if (!items.length) return null;

  const canvas = window.__vtAuditCanvas ??= Object.assign(document.createElement("canvas"), { width: 96, height: 96 });
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  const ink = (text, font) => {
    ctx.clearRect(0, 0, 96, 96);
    ctx.font = font;
    ctx.fillStyle = "#000";
    ctx.fillText(text, 16, 64);
    const data = ctx.getImageData(0, 0, 96, 96).data;
    let hash = 0, count = 0;
    for (let i = 3; i < data.length; i += 4) {
      if (data[i] > 0) { count++; hash = (hash * 31 + i * (data[i] >> 4)) | 0; }
    }
    return count ? count + ":" + hash : "0";
  };
  const notdefCache = window.__vtAuditNotdef ??= new Map();
  const notdef = (font) => {
    if (!notdefCache.has(font))
      notdefCache.set(font, new Set([ink("͸", font), ink("\u{10FFFD}", font), ink("￾", font)]));
    return notdefCache.get(font);
  };
  const missingGlyphs = (root) => {
    const found = [];
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const element = node.parentElement;
      if (!element) continue;
      const style = getComputedStyle(element);
      if (style.display === "none" || style.visibility === "hidden" || style.color === "transparent") continue;
      const font = style.fontStyle + " " + style.fontWeight + " 48px " + style.fontFamily;
      for (const ch of node.textContent ?? "") {
        if (/[\s​-‍⁠﻿]/u.test(ch)) continue;
        const rendered = ink(ch, font);
        if (rendered === "0") found.push({ ch, cp: ch.codePointAt(0).toString(16), font: style.fontFamily, kind: "blank" });
        else if (notdef(font).has(rendered)) found.push({ ch, cp: ch.codePointAt(0).toString(16), font: style.fontFamily, kind: "tofu" });
      }
    }
    return found;
  };
  // Ink extent of a preview, clipped by any overflow-clipping ancestor inside
  // it (stretchy arrows and braces are wide SVGs inside clipping boxes).
  const clippedRect = (node, root) => {
    let rect = node.getBoundingClientRect();
    let top = rect.top, bottom = rect.bottom, left = rect.left, right = rect.right;
    for (let parent = node.parentElement; parent && parent !== root.parentElement; parent = parent.parentElement) {
      const style = getComputedStyle(parent);
      if (style.overflowX !== "visible" || style.overflowY !== "visible") {
        const clip = parent.getBoundingClientRect();
        if (style.overflowX !== "visible") { left = Math.max(left, clip.left); right = Math.min(right, clip.right); }
        if (style.overflowY !== "visible") { top = Math.max(top, clip.top); bottom = Math.min(bottom, clip.bottom); }
      }
    }
    return { top, bottom, left, right, width: right - left, height: bottom - top };
  };
  const unionRect = (root) => {
    let top = Infinity, bottom = -Infinity, left = Infinity, right = -Infinity;
    for (const node of [root, ...root.querySelectorAll("*")]) {
      const style = getComputedStyle(node);
      if (style.display === "none" || style.visibility === "hidden") continue;
      const rect = clippedRect(node, root);
      if (rect.width < 0.5 || rect.height < 0.5) continue;
      top = Math.min(top, rect.top); bottom = Math.max(bottom, rect.bottom);
      left = Math.min(left, rect.left); right = Math.max(right, rect.right);
    }
    return { top, bottom, left, right };
  };
  const rows = items.map((item, index) => {
    const preview = item.querySelector(".ML__popover__command");
    const rect = item.getBoundingClientRect();
    const ink = preview ? unionRect(preview) : null;
    const text = (preview?.textContent ?? "").replace(/[​-‍⁠﻿]/g, "").trim();
    return {
      index,
      command: item.dataset.command ?? "",
      insertion: item.dataset.visualtexInsertionLatex ?? "",
      kind: preview?.dataset.visualtexPreviewKind ?? "native",
      previewLatex: preview?.dataset.visualtexPreview ?? "",
      itemTop: rect.top, itemBottom: rect.bottom, itemLeft: rect.left, itemRight: rect.right,
      inkTop: ink?.top, inkBottom: ink?.bottom, inkLeft: ink?.left, inkRight: ink?.right,
      overflowTop: ink ? Math.max(0, rect.top - ink.top) : 0,
      overflowBottom: ink ? Math.max(0, ink.bottom - rect.bottom) : 0,
      overflowRight: ink ? Math.max(0, ink.right - rect.right) : 0,
      error: Boolean(preview?.querySelector(".ML__error")),
      text,
      question: text.includes("?") || text.includes("�"),
      placeholders: (text.match(/[▢⬚]/g) ?? []).length +
        (preview?.querySelectorAll(".visualtex-native-preview-placeholder, .ML__placeholder").length ?? 0),
      inkWidth: ink ? ink.right - ink.left : 0,
      missing: preview ? missingGlyphs(preview) : [],
      selected: item.classList.contains("ML__popover__current"),
      number: item.querySelector(".ML__popover__index")?.textContent ?? "",
    };
  });
  const panelRect = panel.getBoundingClientRect();
  return { query, panelWidth: panelRect.width, panelHeight: panelRect.height, rows };
}`;

// Number keys pick from the numbered page like an input method; the commit is
// identical to Space on the same candidate.
async function runDigitChecks({ evaluate, key, typeText, clearAndFocus }) {
  const panelState = () =>
    evaluate(`(() => {
      const panel = document.getElementById("visualtex-native-input-suggestion-popover");
      const items = [...(panel?.querySelectorAll("li[data-command]") ?? [])];
      return {
        visible: Boolean(panel?.classList.contains("is-visible")),
        rows: items.map((item) => ({
          insertion: item.dataset.visualtexInsertionLatex ?? "",
          command: item.dataset.command ?? "",
          number: item.querySelector(".ML__popover__index")?.textContent ?? "",
          selected: item.classList.contains("ML__popover__current"),
        })),
      };
    })()`);
  const waitPanel = async () => {
    for (let i = 0; i < 60; i++) {
      const state = await panelState();
      if (state.visible && state.rows.length) {
        await sleep(120);
        return panelState();
      }
      await sleep(50);
    }
    throw new Error("candidate panel did not open");
  };
  const fieldValue = () =>
    evaluate(`document.querySelector("math-field.visual-mathfield").value`);
  const commit = async (query, moves, digit) => {
    await clearAndFocus();
    await typeText(query);
    let state = await waitPanel();
    for (let i = 0; i < moves; i++) await key("ArrowDown", "ArrowDown", 40);
    if (moves) {
      await sleep(150);
      state = await panelState();
    }
    const target = digit === " "
      ? state.rows.find((row) => row.selected)
      : state.rows.find((row) => row.number === digit);
    if (digit === " ") await key(" ", "Space", 32);
    else await key(digit, `Digit${digit}`, digit.charCodeAt(0));
    await sleep(250);
    return { state, target, value: await fieldValue(), panelAfter: (await panelState()).visible };
  };
  const normalize = (latex) => latex.replace(/\\placeholder\{\}/g, "").replace(/\s+/g, "");
  const results = [];
  const check = (name, ok, detail) => {
    results.push({ name, ok, detail });
    console.log(`${ok ? "PASS" : "FAIL"} ${name}${ok ? "" : " " + JSON.stringify(detail)}`);
  };

  // 1. Numbering of the first page.
  await clearAndFocus();
  await typeText("\\s");
  const first = await waitPanel();
  const numbers = first.rows.map((row) => row.number);
  check("first page numbered 1-9", numbers.slice(0, 9).join(",") === "1,2,3,4,5,6,7,8,9" && numbers.slice(9).every((n) => n === ""), numbers);

  // 2. Each digit commits the same thing Space commits on that row.
  for (const digit of ["1", "2", "5", "9"]) {
    const byDigit = await commit("\\s", 0, digit);
    const index = Number(digit) - 1;
    const bySpace = await commit("\\s", index, " ");
    check(`\\s + ${digit} equals Space on row ${digit}`,
      byDigit.target && byDigit.target.insertion === bySpace.target?.insertion &&
        normalize(byDigit.value) === normalize(bySpace.value) && byDigit.value !== "" &&
        !byDigit.value.includes(digit === "1" ? "\\s1" : "\\s" + digit) && !byDigit.panelAfter,
      { digit: byDigit.value, space: bySpace.value, target: byDigit.target, spaceTarget: bySpace.target, panelAfter: byDigit.panelAfter });
  }

  // 3. Arrowing onto the second page renumbers it; digits address that page.
  const paged = await commit("\\s", 10, "2");
  const pageNumbers = paged.state.rows.map((row) => row.number);
  const pagedSpace = await commit("\\s", 10, " ");
  const expectedSecond = await commit("\\s", 10, " ");
  check("second page numbered after arrowing", pageNumbers.slice(9, 18).join(",") === "1,2,3,4,5,6,7,8,9" && pageNumbers.slice(0, 9).every((n) => n === ""), pageNumbers);
  check("digit on second page picks row 11",
    paged.target?.insertion === paged.state.rows[10]?.insertion && normalize(paged.value) === normalize(pagedSpace.value),
    { digit: paged.value, space: pagedSpace.value, again: expectedSecond.value });

  // 4. A digit beyond a short page is swallowed and leaves the query alone.
  await clearAndFocus();
  await typeText("\\bet");
  const short = await waitPanel();
  await key("8", "Digit8", 56);
  await sleep(200);
  const afterShort = await panelState();
  const rawAfterShort = await evaluate(`[...document.querySelector("math-field.visual-mathfield").shadowRoot.querySelectorAll(".ML__raw-latex")].map((n) => n.textContent).join("")`);
  check("digit past the last candidate is ignored", short.rows.length < 8 && afterShort.visible && rawAfterShort === "\\bet", { rows: short.rows.length, rawAfterShort, visible: afterShort.visible });

  // 5. After Escape the digit is typed normally.
  await clearAndFocus();
  await typeText("x");
  await key("2", "Digit2", 50);
  await sleep(150);
  const plain = await fieldValue();
  check("digits outside the panel still type", plain === "x2", plain);

  // 6. Mouse click on a template candidate commits it.
  await clearAndFocus();
  await typeText("\\s");
  const clickState = await waitPanel();
  await evaluate(`document.querySelectorAll("#visualtex-native-input-suggestion-popover li[data-command]")[1].click()`);
  await sleep(250);
  const clicked = await fieldValue();
  check("click commits a template candidate", normalize(clicked) !== "" && !clicked.includes("\\s\\") && clicked.startsWith(clickState.rows[1].command), { clicked, row: clickState.rows[1] });

  // 7. An exactly typed command keeps its bare form ahead of shared variants.
  for (const [query, expected] of [["\\bigcap", "\\bigcap"], ["\\fint", "\\fint"]]) {
    const exact = await commit(query, 0, " ");
    check(`${query} + Space inserts the bare command`, exact.value.trim() === expected, exact);
  }

  if (results.some((result) => !result.ok)) process.exitCode = 1;
}

// LaTeX source panel (CodeMirror): candidate panel geometry while typing,
// snippet insertion with {} fields, Tab between fields, row-break guard.
async function runSourceChecks({ evaluate, key, typeText, client }) {
  const results = [];
  const check = (name, ok, detail) => {
    results.push({ name, ok });
    console.log(`${ok ? "PASS" : "FAIL"} ${name}${ok ? "" : " " + JSON.stringify(detail)}`);
  };
  const ready = await (async () => {
    for (let i = 0; i < 100; i++) {
      if (await evaluate(`Boolean(document.querySelector(".source-panel .cm-content"))`)) return true;
      await evaluate(`(() => {
        const toggle = document.querySelector('[data-classic-bottom-view="source"]:not(.is-active), .standard-source-toggle:not(.is-active)');
        toggle?.click();
        return Boolean(toggle);
      })()`);
      await sleep(100);
    }
    return false;
  })();
  check("source panel mounted", ready, null);
  if (!ready) {
    process.exitCode = 1;
    return;
  }
  const sourceText = () =>
    evaluate(`[...document.querySelectorAll(".source-panel .cm-content .cm-line")].map((line) => line.textContent).join("\\n")`);
  const reset = async () => {
    await evaluate(`(() => {
      const content = document.querySelector(".source-panel .cm-content");
      content.focus();
      document.execCommand("selectAll");
      document.execCommand("delete");
      return true;
    })()`);
    await key("Escape", "Escape", 27);
    await sleep(150);
  };
  const panel = () =>
    evaluate(`(() => {
      const tooltip = document.querySelector(".cm-tooltip-autocomplete");
      if (!tooltip) return null;
      const list = tooltip.querySelector("ul");
      // The visible box is the list; the tooltip is an invisible frame.
      const rect = (list ?? tooltip).getBoundingClientRect();
      const selection = window.getSelection();
      const range = selection.rangeCount ? selection.getRangeAt(0).cloneRange() : null;
      const caret = range ? (range.getClientRects()[0] ?? range.getBoundingClientRect()) : null;
      const line = document.querySelector(".source-panel .cm-activeLine")?.getBoundingClientRect();
      return {
        custom: tooltip.classList.contains("cm-vt-completion"),
        above: tooltip.classList.contains("cm-tooltip-above"),
        // Empty space inside the box: above the first row, or below the last.
        blankAbove: (() => {
          const first = tooltip.querySelector("li");
          return first && list.scrollTop === 0 ? first.getBoundingClientRect().top - rect.top : 0;
        })(),
        blankBelow: (() => {
          const rows = [...tooltip.querySelectorAll("li")];
          const lastRow = rows[rows.length - 1];
          const scrolledToEnd = list.scrollTop + list.clientHeight >= list.scrollHeight - 1;
          return lastRow && scrolledToEnd ? rect.bottom - lastRow.getBoundingClientRect().bottom : 0;
        })(),
        top: rect.top, bottom: rect.bottom, left: rect.left, height: rect.height, listHeight: list?.getBoundingClientRect().height ?? 0,
        caretTop: caret?.top ?? null, caretBottom: caret?.bottom ?? null,
        lineTop: line?.top ?? null, lineBottom: line?.bottom ?? null,
        labels: [...tooltip.querySelectorAll("li")].map((li) => li.querySelector(".cm-completionLabel")?.textContent ?? ""),
        previews: [...tooltip.querySelectorAll("li .cm-vt-completion-preview")].map((p) => p.getBoundingClientRect().width > 2),
        rowHeights: [...new Set([...tooltip.querySelectorAll("li")].map((li) => Math.round(li.getBoundingClientRect().height)))],
      };
    })()`);

  // 1. Panel while typing \fr -> \frac: own styling, previews, steady height,
  // never over the line being typed.
  await reset();
  // --push-lines=N moves the caret N lines down (towards the window bottom,
  // where the panel has to open above the line).
  for (let i = Number(argument("push-lines") ?? 0); i > 0; i--) await key("Enter", "Enter", 13);
  const heights = [];
  let last = null;
  for (const part of ["\\f", "r", "a", "c"]) {
    await typeText(part);
    await sleep(220);
    last = await panel();
    heights.push(last?.height ?? 0);
  }
  check("source panel uses the VisualTeX candidate list", Boolean(last?.custom), last);
  check("candidates show a rendered preview", Boolean(last?.previews.length) && last.previews.every(Boolean), last?.previews);
  check("rows have one fixed height", last?.rowHeights.length === 1, last?.rowHeights);
  check("no blank space in the box while typing", last && last.blankAbove <= 12 && last.blankBelow <= 12, { blankAbove: last?.blankAbove, blankBelow: last?.blankBelow, heights });
  const lineTop = last?.lineTop ?? last?.caretTop;
  const lineBottom = last?.lineBottom ?? last?.caretBottom;
  check(`panel does not cover the line being typed (${last?.above ? "above" : "below"} the line)`, last && (last.top >= lineBottom - 0.5 || last.bottom <= lineTop + 0.5), last);
  check("\\frac ranks first for \\frac", last?.labels[0] === "\\frac", last?.labels.slice(0, 6));

  // 1b. Narrowing a long list (\a) to a short one (\al) leaves no empty
  // space inside the box.
  await reset();
  for (let i = Number(argument("push-lines") ?? 0); i > 0; i--) await key("Enter", "Enter", 13);
  await typeText("\\a");
  await sleep(250);
  const wide = await panel();
  await typeText("l");
  await sleep(250);
  const narrow = await panel();
  if (argument("shot")) {
    const clip = await evaluate(`(() => {
      const rect = document.querySelector(".cm-tooltip-autocomplete > ul").getBoundingClientRect();
      return { x: Math.max(0, rect.left - 30), y: Math.max(0, rect.top - 40), width: rect.width + 60, height: rect.height + 100, scale: 1 };
    })()`);
    const shot = await client.send("Page.captureScreenshot", { format: "jpeg", quality: 60, clip });
    await writeFile(argument("shot"), Buffer.from(shot.data, "base64"));
  }
  if (process.argv.includes("--verbose")) {
    console.log(JSON.stringify(await evaluate(`(() => {
      const tooltip = document.querySelector(".cm-tooltip-autocomplete");
      const list = tooltip.querySelector("ul");
      const describe = (node) => {
        const style = getComputedStyle(node), rect = node.getBoundingClientRect();
        return { cls: node.className, inline: node.getAttribute("style"), top: rect.top, height: rect.height, display: style.display, position: style.position,
          minHeight: style.minHeight, maxHeight: style.maxHeight, cssHeight: style.height, overflow: style.overflowY, marginTop: style.marginTop, padding: style.padding, boxSizing: style.boxSizing };
      };
      return { tooltip: describe(tooltip), list: describe(list), scrollTop: list.scrollTop, scrollHeight: list.scrollHeight, first: describe(list.querySelector("li")), children: [...tooltip.children].map((c) => c.tagName + "." + c.className) };
    })()`), null, 1));
  }
  check(
    `no blank space in the box after narrowing (${wide?.labels.length} -> ${narrow?.labels.length} rows, ${narrow?.above ? "above" : "below"})`,
    narrow && narrow.blankAbove <= 12 && narrow.blankBelow <= 12,
    { blankAbove: narrow?.blankAbove, blankBelow: narrow?.blankBelow, height: narrow?.height, listHeight: narrow?.listHeight, rows: narrow?.labels.length, wideHeight: wide?.height },
  );
  await reset();
  await typeText("\\frac");
  await sleep(220);

  // 2. Enter inserts \frac{}{} with the caret in the first field; Tab moves on.
  await key("Enter", "Enter", 13);
  await sleep(150);
  await typeText("a");
  await key("Tab", "Tab", 9);
  await sleep(100);
  await typeText("b");
  await sleep(150);
  const frac = await sourceText();
  check("Enter inserts \\frac with {} fields and Tab reaches the second", frac.trim() === "\\frac{a}{b}", frac);

  // 3. Tab accepts too; one-argument command.
  await reset();
  await typeText("\\sqr");
  await sleep(220);
  await key("Tab", "Tab", 9);
  await sleep(120);
  await typeText("x");
  await sleep(150);
  const sqrt = await sourceText();
  check("Tab accepts \\sqrt into \\sqrt{x}", sqrt.trim() === "\\sqrt{x}", sqrt);

  // 4. Native command with arguments gets its fields as well.
  await reset();
  await typeText("\\overset");
  await sleep(220);
  await key("Enter", "Enter", 13);
  await sleep(120);
  await typeText("a");
  await key("Tab", "Tab", 9);
  await typeText("b");
  await sleep(150);
  const overset = await sourceText();
  check("\\overset inserts two fields", overset.trim() === "\\overset{a}{b}", overset);

  // 5. Zero-argument symbol inserts as is.
  await reset();
  await typeText("\\alph");
  await sleep(220);
  await key("Enter", "Enter", 13);
  await sleep(150);
  const alpha = await sourceText();
  check("\\alpha inserts bare", alpha.trim() === "\\alpha", alpha);

  // 6. A row break \\ does not open the panel.
  await reset();
  await typeText("a\\\\");
  await sleep(250);
  const rowBreak = await panel();
  check("row break does not open candidates", rowBreak === null, rowBreak);

  // 7. Ranking matches the visual editor's top candidates.
  await reset();
  await typeText("\\s");
  await sleep(250);
  const ranked = await panel();
  check("\\s ranks \\sum first", ranked?.labels[0] === "\\sum", ranked?.labels.slice(0, 8));
  await key("Escape", "Escape", 27);

  // 8. Number keys pick from the numbered page, like the visual editor.
  const numbered = () =>
    evaluate(`[...document.querySelectorAll(".cm-tooltip-autocomplete li")].map((li) => ({
      label: li.querySelector(".cm-completionLabel")?.textContent ?? "",
      number: li.querySelector(".cm-vt-completion-index")?.textContent ?? "",
    }))`);
  const pick = async (query, moves, finish) => {
    await reset();
    await typeText(query);
    await sleep(250);
    for (let i = 0; i < moves; i++) await key("ArrowDown", "ArrowDown", 40);
    await sleep(120);
    const rows = await numbered();
    if (finish === "Enter") await key("Enter", "Enter", 13);
    else if (finish === "Escape") await key("Escape", "Escape", 27);
    else await key(finish, `Digit${finish}`, finish.charCodeAt(0));
    await sleep(200);
    return { rows, text: (await sourceText()).trim(), panelAfter: await panel() };
  };
  const firstPage = await pick("\\s", 0, "Escape");
  const firstNumbers = firstPage.rows.map((row) => row.number);
  await reset();
  await typeText("\\s");
  await sleep(250);
  const numberLayout = await evaluate(`[...document.querySelectorAll(".cm-tooltip-autocomplete li")].slice(0, 9).map((li) => {
    const row = li.getBoundingClientRect();
    const number = li.querySelector(".cm-vt-completion-index").getBoundingClientRect();
    const label = li.querySelector(".cm-completionLabel").getBoundingClientRect();
    return { width: number.width, gap: label.left - number.right, offCenter: Math.abs((number.top + number.bottom) / 2 - (row.top + row.bottom) / 2) };
  })`);
  check("source: numbers sit in their own column left of the label",
    numberLayout.length === 9 && numberLayout.every((r) => r.width >= 10 && r.gap >= 4 && r.offCenter <= 2), numberLayout);
  await key("Escape", "Escape", 27);
  check("source: first page numbered 1-9",
    firstNumbers.slice(0, 9).join(",") === "1,2,3,4,5,6,7,8,9" && firstNumbers.slice(9).every((n) => n === ""), firstNumbers);
  for (const digit of ["1", "2", "5", "9"]) {
    const byDigit = await pick("\\s", 0, digit);
    const byEnter = await pick("\\s", Number(digit) - 1, "Enter");
    check(`source: \\s + ${digit} equals Enter on row ${digit}`,
      byDigit.text !== "" && byDigit.text === byEnter.text && !byDigit.text.endsWith(digit) && byDigit.panelAfter === null,
      { digit: byDigit.text, enter: byEnter.text, panelAfter: Boolean(byDigit.panelAfter) });
  }
  const secondPage = await pick("\\s", 10, "2");
  const secondEnter = await pick("\\s", 10, "Enter");
  const secondNumbers = secondPage.rows.map((row) => row.number);
  check("source: second page numbered after arrowing",
    secondNumbers.slice(9, 18).join(",") === "1,2,3,4,5,6,7,8,9" && secondNumbers.slice(0, 9).every((n) => n === ""), secondNumbers);
  check("source: digit on second page picks row 11",
    secondPage.text === secondEnter.text && secondPage.text.startsWith(secondPage.rows[10]?.label ?? "?"),
    { digit: secondPage.text, enter: secondEnter.text, row11: secondPage.rows[10]?.label });
  // The fuzzy match keeps short queries long (\bet has dozens of rows), so
  // take the first query whose list ends before row 9.
  let shortPage = null;
  let shortQuery = "";
  for (const query of ["\\varepsilo", "\\rightleftharpoo", "\\overbrac", "\\mathbb"]) {
    shortPage = await pick(query, 0, "9");
    shortQuery = query;
    if (shortPage.rows.length && shortPage.rows.length < 9) break;
  }
  check(`source: digit past the last candidate is ignored (${shortQuery})`,
    shortPage.rows.length && shortPage.rows.length < 9 && shortPage.text === shortQuery && shortPage.panelAfter !== null,
    { rows: shortPage.rows.length, text: shortPage.text, open: Boolean(shortPage.panelAfter) });
  await reset();
  await typeText("x2");
  await sleep(150);
  const plainDigits = (await sourceText()).trim();
  check("source: digits outside the panel still type", plainDigits === "x2", plainDigits);

  if (results.some((result) => !result.ok)) process.exitCode = 1;
}

async function main() {
  if (shotDir) await mkdir(shotDir, { recursive: true });
  const preview = spawn(
    process.execPath,
    ["node_modules/vite/bin/vite.js", "preview", "--host", "127.0.0.1", "--port", String(previewPort), "--strictPort",
      ...(argument("dist") ? ["--outDir", argument("dist")] : [])],
    { cwd: process.cwd(), stdio: "ignore" },
  );
  let chrome;
  let client;
  try {
    await waitFor(baseUrl);
    chrome = spawn(
      resolveChromiumExecutable(),
      [
        "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`, `--user-data-dir=${chromeProfile}`,
        "--window-size=1400,1000", baseUrl,
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
    const page = targets.find((t) => t.type === "page" && t.url.startsWith(baseUrl)) ?? targets.find((t) => t.type === "page");
    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");

    const evaluate = async (expression) => {
      const result = await client.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
      if (result.exceptionDetails)
        throw new Error(result.exceptionDetails.exception?.description ?? result.exceptionDetails.text);
      return result.result.value;
    };
    const key = async (value, code, vk) => {
      const common = { key: value, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk };
      const text = value.length === 1 ? { text: value, unmodifiedText: value } : {};
      await client.send("Input.dispatchKeyEvent", { type: value.length === 1 ? "keyDown" : "rawKeyDown", ...common, ...text });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(30);
    };
    const typeText = async (text) => {
      for (const ch of text) {
        if (ch === "\\") await key(ch, "Backslash", 220);
        else if (/[0-9]/.test(ch)) await key(ch, `Digit${ch}`, ch.charCodeAt(0));
        else await key(ch, `Key${ch.toUpperCase()}`, ch.toUpperCase().charCodeAt(0));
      }
    };

    await client.send("Page.navigate", { url: baseUrl });
    await sleep(1800);
    await evaluate(`(() => {
      for (const key of ["visualtex.onboarding.v3.completed", "visualtex.office.macos.first-run.v1.completed",
        "visualtex.onboarding.macos.desktop.v1.2.0.completed", "visualtex.office.macos.native-first-run.v1.2.0.completed"])
        localStorage.setItem(key, "true");
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      const id = crypto.randomUUID();
      persisted.state = { ...(persisted.state || {}), lines: [{ id, latex: "" }], activeLineId: id, usage: {}, personalize: false, suggestionCount: 24,
        sourceOpen: ${process.argv.includes("--source")} };
      localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(1500);
    const started = Date.now();
    while (Date.now() - started < 15_000) {
      if (await evaluate(`Boolean(document.querySelector("math-field.visual-mathfield")?.shadowRoot)`)) break;
      await sleep(50);
    }
    await evaluate(`document.fonts.ready.then(() => true)`);

    const clearAndFocus = async () => {
      await evaluate(`(() => {
        const field = document.querySelector("math-field.visual-mathfield");
        field.setValue("", { mode: "math", format: "latex", insertionMode: "replaceAll", selectionMode: "after", silenceNotifications: true });
        field.dispatchEvent(new InputEvent("input", { bubbles: true, composed: true, inputType: "insertText" }));
        return true;
      })()`);
      await sleep(80);
      await evaluate(`(() => {
        const field = document.querySelector("math-field.visual-mathfield");
        field.focus();
        field.position = field.lastOffset;
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        return field.hasFocus();
      })()`);
      await sleep(60);
    };

    if (process.argv.includes("--source")) {
      const viewportHeight = Number(argument("viewport-height") ?? 0);
      if (viewportHeight) {
        await client.send("Emulation.setDeviceMetricsOverride", { width: 1400, height: viewportHeight, deviceScaleFactor: 1, mobile: false });
        await sleep(300);
      }
      await runSourceChecks({ evaluate, key, typeText, client });
      return;
    }

    if (process.argv.includes("--digits")) {
      await runDigitChecks({ evaluate, key, typeText, clearAndFocus });
      return;
    }

    const report = [];
    for (const query of queries) {
      await clearAndFocus();
      await typeText(query);
      let state = null;
      const queryStarted = Date.now();
      while (Date.now() - queryStarted < 3_000) {
        state = await evaluate(`(${inspectPanel})(${JSON.stringify(query)})`);
        if (state) break;
        // A prefix without candidates leaves the raw command and no panel.
        if (Date.now() - queryStarted > 450 && (await evaluate(`(() => {
          const field = document.querySelector("math-field.visual-mathfield");
          const raw = [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "").join("");
          const panel = document.getElementById("visualtex-native-input-suggestion-popover");
          return raw === ${JSON.stringify(query)} && !panel?.classList.contains("is-visible");
        })()`))) break;
        await sleep(50);
      }
      if (state) {
        // Let decoration and web fonts settle, then measure the final frame.
        await sleep(200);
        state = (await evaluate(`(${inspectPanel})(${JSON.stringify(query)})`)) ?? state;
      }
      report.push(state ?? { query, rows: [] });
      if (process.argv.includes("--verbose")) console.log(`${report.length} ${query} ${state?.rows.length ?? 0}`);
      if (probeFile && state) {
        const probe = await readFile(probeFile, "utf8");
        console.log(query, JSON.stringify(await evaluate(`(${probe})()`), null, 1));
      }
      if (shotDir && state) {
        const clip = await evaluate(`(() => {
          const panel = document.getElementById("visualtex-native-input-suggestion-popover");
          const list = panel.querySelector("ul");
          list.style.maxHeight = "none";
          const rect = panel.getBoundingClientRect();
          return { x: rect.left, y: Math.max(0, rect.top), width: rect.width, height: Math.min(rect.height, 2400), scale: 1 };
        })()`);
        await client.send("Emulation.setDeviceMetricsOverride", { width: 1400, height: Math.ceil(clip.y + clip.height + 20), deviceScaleFactor: 1, mobile: false });
        await sleep(120);
        const shot = await client.send("Page.captureScreenshot", { format: "png", clip, captureBeyondViewport: true });
        const name = query.replace(/\\/g, "").replace(/[^A-Za-z0-9]/g, "_") + (/[A-Z]/.test(query) ? "_uc" : "");
        await writeFile(`${shotDir}/${name}.png`, Buffer.from(shot.data, "base64"));
        await client.send("Emulation.clearDeviceMetricsOverride");
        await evaluate(`document.querySelector("#visualtex-native-input-suggestion-popover ul").style.maxHeight = ""`);
      }
      await key("Escape", "Escape", 27);
    }

    const problems = [];
    for (const entry of report) {
      for (const row of entry.rows) {
        const issues = [];
        if (row.overflowTop > 1 || row.overflowBottom > 1) issues.push(`overflow v ${row.overflowTop.toFixed(1)}/${row.overflowBottom.toFixed(1)}`);
        if (row.overflowRight > 1) issues.push(`overflow right ${row.overflowRight.toFixed(1)}`);
        if (row.error) issues.push("parse error");
        if (row.question) issues.push("question mark");
        if (row.previewLatex === "unavailable") issues.push("no preview (label)");
        if (row.placeholders) issues.push(`placeholder boxes ${row.placeholders}`);
        if (row.inkWidth < 3) issues.push("blank preview");
        if (row.missing.length) issues.push("missing " + row.missing.map((m) => `${m.kind}:U+${m.cp}`).join(" "));
        if (issues.length) problems.push({ query: entry.query, command: row.command, kind: row.kind, issues });
      }
    }
    const unique = new Map();
    for (const problem of problems) {
      const id = problem.command + "|" + problem.issues.join(";");
      if (!unique.has(id)) unique.set(id, { ...problem, queries: [problem.query] });
      else unique.get(id).queries.push(problem.query);
    }
    const summary = {
      queries: report.length,
      emptyQueries: report.filter((entry) => !entry.rows.length).map((entry) => entry.query),
      problems: [...unique.values()],
    };
    if (outPath) await writeFile(outPath, JSON.stringify({ summary, report }, null, 2));
    console.log(JSON.stringify(summary, null, 2));
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(150);
    await rm(chromeProfile, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  }
}

await main();
