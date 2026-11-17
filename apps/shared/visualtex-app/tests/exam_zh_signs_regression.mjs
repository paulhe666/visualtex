// The congruence and similarity signs of Chinese geometry as exam-zh-symbols.sty draws them, and its spelling of them.
//
// Drawing: ≌ (\backcong, the character U+224C too), ∽ (\backsim) and their negations (\not\backcong, \not\backsim) are
// exam-zh's TikZ outlines (tex-layout/src/exam-zh-signs.generated.ts), the outline drawn for the size (10 pt in the
// formula, 7 pt in a subscript), set as relations: the same thick space on both sides as "=". ≅ (\cong) and ∼ (\sim)
// stay the TeX fonts' characters.
//
// Spelling (math/examZhSymbols.ts): the formulas keep the stored names; with the exam-zh setting on (the default of the
// Chinese interface) what the user types, pastes as text or edits in the source is read in exam-zh's spelling (\cong ≌,
// \cong* ≌, \sim ∽, \sim* ∼, \ncong, \nsim), and the text copied or shown in the source is written in it (≅ and the
// negated ≅ and ∼ as their characters). A copy's application/x-latex keeps the stored names, so pasting it inside the
// app keeps the signs; the toolbar tiles' tooltips name their commands in the same spelling, and follow the switch at
// once. With the setting off, and by default in the English interface, LaTeX is read and written as is.
import { readFileSync } from "node:fs";
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

const generated = readFileSync(new URL("../../tex-layout/src/exam-zh-signs.generated.ts", import.meta.url), "utf8");
function outline(name, size) {
  const block = new RegExp(`\\n    ${name}: \\[([\\s\\S]*?)\\n    \\],`).exec(generated)?.[1] ?? "";
  const path = new RegExp(`\\{ size: ${size}, [^}]*?path: "([^"]+)" \\}`).exec(block)?.[1];
  assert(path, `no ${size} pt outline of ${name} in exam-zh-signs.generated.ts`);
  return path;
}

const DRAWN_ROWS = [
  "a\\backcong b", // 0
  "a=b", // 1
  "a≌b", // 2
  "a\\cong b", // 3
  "a\\backsim b", // 4
  "a\\not\\backcong b", // 5
  "a\\not\\backsim b", // 6
  "a\\sim b", // 7
  "x_{a\\backcong b}", // 8
];
const STORED = "a\\backcong b\\cong c\\backsim d\\sim e\\not\\backcong f\\not\\backsim g";

await withAppBrowser(async ({ baseUrl, client }) => {
  const evaluate = async (expression) => {
    const result = await client.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
    return result.result.value;
  };
  for (let tries = 0; tries < 300 && !String(await evaluate("location.href")).startsWith(baseUrl); tries += 1) await sleep(50);

  // the app opened on the given rows; examZhSymbols undefined: the setting as a new user has it
  const open = async ({ rows, language = "cn", examZhSymbols }) => {
    await evaluate(`(() => {
      for (const k of ["visualtex.onboarding.v3.completed", "visualtex.office.macos.first-run.v1.completed", "visualtex.onboarding.macos.desktop.v1.2.0.completed", "visualtex.office.macos.native-first-run.v1.2.0.completed"]) localStorage.setItem(k, "true");
      localStorage.setItem("visualtex.document.tutorial-shown.v1", "1");
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      const state = { ...(persisted.state || {}), theme: "light", language: ${JSON.stringify(language)}, lines: ${JSON.stringify(rows)}.map((latex, i) => ({ id: "line-" + (i + 1), latex })), activeLineId: "line-1" };
      delete state.examZhSymbols;
      ${examZhSymbols === undefined ? "" : `state.examZhSymbols = ${examZhSymbols};`}
      persisted.state = state;
      localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
      localStorage.setItem("visualtex.workspace-content.v1", JSON.stringify({ state: { content: "formula" }, version: 0 }));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(1500);
    for (let t = 0; t < 200 && (await evaluate(`document.querySelectorAll("math-field.visual-mathfield").length`)) < rows.length; t += 1) await sleep(50);
    await sleep(800);
  };
  const field = (i) => `document.querySelectorAll("math-field.visual-mathfield")[${i}]`;
  const value = (i) => evaluate(`${field(i)}.getValue("latex")`);
  const focusEmpty = (i) =>
    evaluate(`(() => { const f = ${field(i)};
      f.setValue("", { mode: "math", format: "latex", insertionMode: "replaceAll", silenceNotifications: true });
      f.focus(); f.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true }); return true; })()`);
  const KEYS = { "\\": ["Backslash", 220], " ": ["Space", 32], "*": ["Digit8", 56] };
  const type = async (text) => {
    for (const ch of text) {
      const [code, vk] = KEYS[ch] ?? ["Key" + ch.toUpperCase(), ch.toUpperCase().charCodeAt(0)];
      const common = { key: ch, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk };
      await client.send("Input.dispatchKeyEvent", { type: "keyDown", ...common, text: ch, unmodifiedText: ch });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(30);
    }
    await sleep(150);
  };
  const typed = async (text) => {
    await focusEmpty(0);
    await type(text);
    return value(0);
  };
  const paste = (i, data) =>
    evaluate(`(() => { const f = ${field(i)}; const sink = f.shadowRoot.querySelector('[part="keyboard-sink"]');
      const clipboardData = new DataTransfer();
      for (const [type, text] of Object.entries(${JSON.stringify(data)})) clipboardData.setData(type, text);
      sink.dispatchEvent(new ClipboardEvent("paste", { clipboardData, bubbles: true, cancelable: true, composed: true }));
      return f.getValue("latex"); })()`);
  const copyWhole = (i) =>
    evaluate(`(() => { const f = ${field(i)}; f.focus(); f.executeCommand("selectAll");
      const sink = f.shadowRoot.querySelector('[part="keyboard-sink"]');
      const clipboardData = new DataTransfer();
      sink.dispatchEvent(new ClipboardEvent("copy", { clipboardData, bubbles: true, cancelable: true, composed: true }));
      return { text: clipboardData.getData("text/plain"), latex: clipboardData.getData("application/x-latex") }; })()`);
  // what a toolbar tile's tooltip names after its label (the tools view in front of the source panel)
  const tileCommands = async () => {
    await evaluate(`(() => { document.querySelector('[data-classic-bottom-view="tools"]:not(.is-active)')?.click(); return true; })()`);
    for (let t = 0; t < 40 && !(await evaluate(`Boolean(document.querySelector('[data-command-id="backcong"]'))`)); t += 1) await sleep(50);
    return evaluate(`Object.fromEntries(["cong", "backcong", "sim", "backsim"].map((id) => [id, (document.querySelector('[data-command-id="' + id + '"]')?.getAttribute("title") ?? "").split(" · ").slice(1).join(" · ")]))`);
  };
  const shortcut = async (key, code, vk, modifiers) => {
    for (const type of ["rawKeyDown", "keyUp"]) {
      await client.send("Input.dispatchKeyEvent", { type, key, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers });
    }
  };
  // the LaTeX source panel (Ctrl+\ from a formula opens it)
  const sourceText = async () => {
    if (!(await evaluate(`Boolean(document.querySelector(".cm-content"))`))) {
      await evaluate(`(() => { ${field(0)}.focus(); return true; })()`);
      await shortcut("\\", "Backslash", 220, 2);
      for (let t = 0; t < 60 && !(await evaluate(`Boolean(document.querySelector(".cm-content"))`)); t += 1) await sleep(50);
    }
    return evaluate(`document.querySelector(".cm-content")?.innerText ?? null`);
  };
  const editSource = async (text) => {
    await evaluate(`(() => { document.querySelector(".cm-content").focus(); return true; })()`);
    await sleep(150);
    await shortcut("a", "KeyA", 65, 2);
    await sleep(100);
    await client.send("Input.insertText", { text });
    await sleep(800);
  };

  // -- drawing ---------------------------------------------------------------------------------------------------
  await open({ rows: DRAWN_ROWS });
  const rows = await evaluate(`[...document.querySelectorAll("math-field.visual-mathfield")].map((f) => ({
    paths: [...f.shadowRoot.querySelectorAll("path")].map((p) => ({ d: p.getAttribute("d") || "", transform: p.getAttribute("transform") || "" })),
    glyphs: [...f.shadowRoot.querySelector(".ML__vtx").children]
      .filter((s) => /font:/.test(s.getAttribute("style") || "") && s.textContent)
      .map((s) => ({ text: s.textContent, left: parseFloat(/left:(-?[0-9.]+)em/.exec(s.getAttribute("style"))[1]) })),
  }))`);
  assert(rows.length === DRAWN_ROWS.length, `${rows.length} rows rendered, expected ${DRAWN_ROWS.length}`);
  const signPath = (row, name, size) => rows[row].paths.find((p) => p.d === outline(name, size));
  for (const [row, name, size] of [[0, "backcong", 10], [2, "backcong", 10], [4, "backsim", 10], [5, "notbackcong", 10], [6, "notbacksim", 10], [8, "backcong", 7]]) {
    assert(signPath(row, name, size), `${DRAWN_ROWS[row]}: not drawn as exam-zh's ${name} at ${size} pt (${rows[row].paths.length} outlines drawn)`);
  }
  for (const row of [3, 7]) {
    const exam = ["backcong", "notbackcong", "backsim", "notbacksim"].flatMap((name) => [10, 7, 5].map((size) => outline(name, size)));
    assert(!rows[row].paths.some((p) => exam.includes(p.d)), `${DRAWN_ROWS[row]} is drawn as an exam-zh sign`);
    assert(rows[row].glyphs.length >= 3, `${DRAWN_ROWS[row]}: the sign is not a font character (${rows[row].glyphs.length} characters)`);
  }
  // relations: they start where "=" starts (after a and a thick space), and b follows them after the same thick space
  const translateX = (p) => parseFloat(/translate\((-?[0-9.]+)/.exec(p.transform)[1]) / 100;
  const equals = rows[1].glyphs; // a, =, b
  const EQUALS_WIDTH = 0.77778;
  for (const [row, name, width] of [[0, "backcong", 1.12], [2, "backcong", 1.12], [4, "backsim", 1.16], [5, "notbackcong", 1.12], [6, "notbacksim", 1.16]]) {
    const x = translateX(signPath(row, name, 10));
    assert(Math.abs(x - equals[1].left) < 0.002, `${DRAWN_ROWS[row]}: the sign starts at ${x.toFixed(4)}em, = at ${equals[1].left.toFixed(4)}em`);
    const b = rows[row].glyphs[rows[row].glyphs.length - 1];
    const expected = equals[2].left + (width - EQUALS_WIDTH);
    assert(Math.abs(b.left - expected) < 0.003, `${DRAWN_ROWS[row]}: b at ${b.left.toFixed(4)}em, expected ${expected.toFixed(4)}em (the space after a relation)`);
  }

  // the toolbar's relation group offers ≌ after ≅ and ∽ after ∼, each named with its sign
  const click = (selector) =>
    evaluate(`(() => {
      const el = [...document.querySelectorAll(${JSON.stringify(selector)})].find((e) => e.getBoundingClientRect().width > 0);
      if (!el) return false;
      el.click();
      return true;
    })()`);
  assert(await click("[data-toolbar-edit-toggle]"), "the toolbar's edit button is not on the page");
  await sleep(500);
  assert(await click("[data-toolbar-add-item]"), "no + card in the toolbar's edit mode");
  await sleep(600);
  assert(await click('[data-toolbar-picker-source="relation"]'), "the picker has no relation category");
  await sleep(400);
  const cards = await evaluate(`[...document.querySelectorAll("[data-toolbar-item-picker] [data-toolbar-picker-item]")].map((c) => c.getAttribute("title"))`);
  for (const [after, card] of [["全等（≅）", "全等（≌）"], ["相似（∼）", "相似（∽）"]]) {
    assert(cards.includes(after) && cards.includes(card), `the relation category lacks ${after} or ${card} (${cards.slice(0, 14).join(" | ")} ...)`);
    assert(cards.indexOf(card) > cards.indexOf(after), `the ${card} card does not follow the ${after} card`);
  }
  console.log("exam-zh signs drawn: ≌ ∽ and their negations as exam-zh's outlines (10 pt, 7 pt in a subscript), relations like =");

  // -- exam-zh's spelling (the Chinese interface's default) --------------------------------------------------------
  await open({ rows: ["", STORED] });
  const examZhTyped = [
    ["\\cong ", "\\backcong"],
    ["\\sim ", "\\backsim"],
    ["\\ncong ", "\\not\\backcong"],
    ["\\nsim ", "\\not\\backsim"],
    ["\\cong* ", "\\backcong"],
    ["\\sim* ", "\\sim"],
    ["\\backcong ", "\\backcong"],
    ["\\backsim ", "\\backsim"],
  ];
  for (const [input, stored] of examZhTyped) {
    const result = await typed(input);
    assert(result === stored, `typed ${JSON.stringify(input)}: the formula is ${JSON.stringify(result)}, expected ${JSON.stringify(stored)}`);
  }
  // a typed word escapes as its command does
  const word = await typed("a cong b");
  assert(word.includes("\\backcong") && !/\\cong(?![a-z])/.test(word), `typed "a cong b": ${JSON.stringify(word)}`);
  await focusEmpty(0);
  const pastedText = await paste(0, { "text/plain": "A\\cong B,C\\sim D,E\\cong* F,G\\sim* H,I\\ncong J,K\\nsim L" });
  assert(
    pastedText === "A\\backcong B,C\\backsim D,E\\backcong F,G\\sim H,I\\not\\backcong J,K\\not\\backsim L",
    `pasted text is read in exam-zh's spelling: ${JSON.stringify(pastedText)}`,
  );
  await focusEmpty(0);
  const pastedLatex = await paste(0, { "application/x-latex": "A\\cong B\\sim C", "text/plain": "A\\cong B\\sim C" });
  assert(pastedLatex === "A\\cong B\\sim C", `a copy's application/x-latex keeps the stored names: ${JSON.stringify(pastedLatex)}`);
  const copied = await copyWhole(1);
  assert(copied.latex === STORED, `the copy's application/x-latex is not the stored formula: ${JSON.stringify(copied.latex)}`);
  for (const spelled of ["a\\cong b", "b≅ c", "c\\sim d", "d\\sim* e", "e\\ncong f", "f\\nsim g"]) {
    assert(copied.text.includes(spelled), `the copied text has no ${JSON.stringify(spelled)}: ${JSON.stringify(copied.text)}`);
  }
  assert(!/\\back|\\not/.test(copied.text), `the copied text keeps stored names: ${JSON.stringify(copied.text)}`);
  const source = await sourceText();
  assert(source && source.includes("a\\cong b≅ c\\sim d\\sim* e\\ncong f\\nsim g"), `the source panel is not in exam-zh's spelling: ${JSON.stringify(source)}`);
  await editSource("$$ a\\cong b\\cong* c\\sim d\\sim* e\\ncong f\\nsim g $$");
  const edited = await value(0);
  assert(
    edited === "a\\backcong b\\backcong c\\backsim d\\sim e\\not\\backcong f\\not\\backsim g",
    `the source edit is not read in exam-zh's spelling: ${JSON.stringify(edited)}`,
  );
  const tilesOn = await tileCommands();
  const expectOn = { cong: "≅", backcong: "\\cong", sim: "\\sim*", backsim: "\\sim" };
  for (const [id, command] of Object.entries(expectOn)) {
    assert(tilesOn[id] === command, `exam-zh spelling: the ${id} tile's tooltip names ${JSON.stringify(tilesOn[id])}, expected ${JSON.stringify(command)}`);
  }
  console.log("exam-zh spelling on: typing, quick escapes, pasted text, copied text, the source panel and the tiles' tooltips");

  // the switch (Settings > 公式编辑) takes effect at once: tooltips and typing follow it without a reload
  await open({ rows: ["", STORED] });
  assert((await tileCommands()).backcong === "\\cong", "a new page does not start in exam-zh's spelling");
  await shortcut(",", "Comma", 188, 2);
  for (let t = 0; t < 60 && !(await evaluate(`Boolean(document.querySelector('[data-settings-page="editor"]'))`)); t += 1) await sleep(50);
  await evaluate(`document.querySelector('[data-settings-page="editor"]')?.click()`);
  for (let t = 0; t < 60 && !(await evaluate(`Boolean(document.querySelector("[data-exam-zh-symbols-setting]"))`)); t += 1) await sleep(50);
  assert(await evaluate(`document.querySelector("[data-exam-zh-symbols-setting]").checked`), "the exam-zh switch is not on in the Chinese interface");
  await evaluate(`document.querySelector("[data-exam-zh-symbols-setting]").click()`);
  await sleep(300);
  assert(!(await evaluate(`document.querySelector("[data-exam-zh-symbols-setting]").checked`)), "the exam-zh switch did not turn off");
  await evaluate(`(() => { document.querySelector('[aria-label="关闭设置"], [aria-label="Close settings"]')?.click(); return true; })()`);
  for (let t = 0; t < 40 && (await evaluate(`Boolean(document.querySelector("[data-exam-zh-symbols-setting]"))`)); t += 1) await sleep(50);
  assert(!(await evaluate(`Boolean(document.querySelector("[data-exam-zh-symbols-setting]"))`)), "Settings did not close");
  const tilesSwitched = await tileCommands();
  assert(tilesSwitched.backcong === "\\backcong" && tilesSwitched.cong === "\\cong", `after the switch the tooltips still name ${JSON.stringify(tilesSwitched)}`);
  const typedSwitched = await typed("\\cong ");
  assert(typedSwitched === "\\cong", `after the switch typed \\cong gives ${JSON.stringify(typedSwitched)}`);
  console.log("exam-zh switch: turned off in Settings, tooltips and typing follow at once");

  // -- the setting off, and the English interface's default ---------------------------------------------------------
  for (const [label, options] of [["setting off", { examZhSymbols: false }], ["English interface", { language: "en" }]]) {
    await open({ rows: ["", STORED], ...options });
    for (const input of ["\\cong ", "\\sim ", "\\backcong "]) {
      const result = await typed(input);
      assert(result === input.trim(), `${label}: typed ${JSON.stringify(input)} gives ${JSON.stringify(result)}`);
    }
    await focusEmpty(0);
    const plain = await paste(0, { "text/plain": "A\\cong B\\sim C" });
    assert(plain === "A\\cong B\\sim C", `${label}: pasted text is changed: ${JSON.stringify(plain)}`);
    const text = (await copyWhole(1)).text;
    assert(text.includes(STORED), `${label}: the copied text is not the stored formula: ${JSON.stringify(text)}`);
    const shown = await sourceText();
    assert(shown && shown.includes(STORED), `${label}: the source panel is not the stored formula: ${JSON.stringify(shown)}`);
    const tiles = await tileCommands();
    const expectOff = { cong: "\\cong", backcong: "\\backcong", sim: "\\sim", backsim: "\\backsim" };
    for (const [id, command] of Object.entries(expectOff)) {
      assert(tiles[id] === command, `${label}: the ${id} tile's tooltip names ${JSON.stringify(tiles[id])}, expected ${JSON.stringify(command)}`);
    }
  }
  console.log("exam-zh spelling off (setting off, English interface): LaTeX read and written as is");
  console.log("exam-zh signs: passed");
}, { windowSize: "1280,1000" });
