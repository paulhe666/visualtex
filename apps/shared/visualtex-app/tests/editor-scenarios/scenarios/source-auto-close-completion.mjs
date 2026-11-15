// Environments in the LaTeX source panel (universal source format): completing \begin{name} inside a display block,
// Enter after a typed \begin{name}, an invalid environment that stays a draft, and a half-typed command that does
// not reach the document. Typed with real key presses, so the editor's key bindings and brace pairing take part.
import assert from "node:assert/strict";

export const names = ["source-auto-close-completion"];

// state the page starts with (merged into the editor's persisted state)
export const storage = () => ({
  editorLayout: "classic",
  sourceOpen: true,
  zoom: 0.45,
  lines: [{ id: "source-environments", latex: "", mode: "display" }],
  activeLineId: "source-environments",
});
export const localStorageItems = () => ({ "visualtex-desktop-editor-source-open": "true" });

// US-layout keys of the characters typed below: [code, key code, modifiers]
const KEYS = {
  "\\": ["Backslash", 220, 0],
  "{": ["BracketLeft", 219, 8],
  "}": ["BracketRight", 221, 8],
  "*": ["Digit8", 56, 8],
  "$": ["Digit4", 52, 8],
  "&": ["Digit7", 55, 8],
  "=": ["Equal", 187, 0],
};

// the environments the completion list offers, typed up to their last letter; a row that is a root aligned block
// with rows is the editor's own multi-line row, which the source writes with the format profile's multi-line
// environment (gather* by default)
const COMPLETION_CASES = [
  ["matrix", "matr"],
  ["pmatrix", "pmatri"],
  ["bmatrix", "bmatri"],
  ["Bmatrix", "Bmatri"],
  ["vmatrix", "vmatri"],
  ["Vmatrix", "Vmatri"],
  ["smallmatrix", "smallmatri"],
  ["cases", "case"],
  ["aligned", "aligne", "\\end{gather*}"],
  ["gathered", "gathere"],
  ["array", "arra"],
];

export async function run({ client, evaluate, key, sleep, waitForEvaluation }) {
  const press = async (text) => {
    for (const character of text) {
      const [code, keyCode, modifiers] = KEYS[character] ?? [
        `Key${character.toUpperCase()}`,
        character.toUpperCase().charCodeAt(0),
        /[A-Z]/.test(character) ? 8 : 0,
      ];
      await key(character, code, keyCode, modifiers);
    }
  };
  const click = async ({ x, y }) => {
    await client.send("Input.dispatchMouseEvent", { type: "mousePressed", x, y, button: "left", buttons: 1, clickCount: 1 });
    await client.send("Input.dispatchMouseEvent", { type: "mouseReleased", x, y, button: "left", buttons: 0, clickCount: 1 });
  };
  const source = () => evaluate(`(() => {
        const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
        const lines = [...document.querySelectorAll(".classic-source-pane-slot .cm-line")]
          .map((line) => (line.textContent ?? "").replace(/\\u200b/g, ""));
        return {
          lines,
          rows: (persisted.state?.lines ?? []).map((line) => ({ mode: line.mode, latex: line.latex })),
          listOpen: Boolean(document.querySelector(".cm-tooltip-autocomplete")),
          error: Boolean(document.querySelector(".source-error-chip")),
          focused: document.querySelector(".classic-source-pane-slot .cm-editor")?.classList.contains("cm-focused") ?? false,
          field: document.querySelector("math-field")?.value ?? null,
        };
      })()`);
  const waitForSource = async (description, check, timeoutMs = 8000) => {
    const started = Date.now();
    let state;
    while (Date.now() - started < timeoutMs) {
      state = await source();
      if (check(state)) return state;
      await sleep(60);
    }
    throw new Error(`Timed out waiting for ${description}: ${JSON.stringify(state)}`);
  };
  const emptyDisplayRow = [{ mode: "display", latex: "" }];

  const reload = async (description, ready = "true") => {
    await evaluate(`window.__sourceScenarioOldPage = true`);
    await client.send("Page.reload", { ignoreCache: true });
    await waitForEvaluation(`(() => ({
          ready:
            !window.__sourceScenarioOldPage &&
            document.querySelector(".workspace")?.dataset.editorLayout === "classic" &&
            Boolean(document.querySelector(".classic-source-pane-slot .cm-content")) &&
            Boolean(document.querySelector("math-field")) &&
            (${ready}),
        }))()`, description);
    await sleep(300);
  };

  // a document of one empty display row, the caret in the (empty) source
  const emptyDocument = async (label) => {
    await evaluate(`(() => {
          const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
          const line = { id: crypto.randomUUID(), latex: "", mode: "display" };
          persisted.state = { ...(persisted.state || {}), lines: [line], activeLineId: line.id };
          localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
        })()`);
    await reload(`${label}: empty document`);
    const point = await evaluate(`(() => {
          const rect = document.querySelector(".classic-source-pane-slot .cm-line")?.getBoundingClientRect();
          return rect ? { x: rect.left + 10, y: rect.top + rect.height / 2 } : null;
        })()`);
    assert.ok(point, `${label}: source line`);
    await click(point);
    await waitForSource(`${label}: source focus`, (state) => state.focused);
  };

  // $$, Enter, $$, then back into the block: the caret on its indented inner line
  const displayBlock = async (label) => {
    await press("$$");
    await key("Enter", "Enter", 13);
    await press("$$");
    await key("ArrowUp", "ArrowUp", 38);
    await key("End", "End", 35);
    await key("Enter", "Enter", 13);
    const block = await waitForSource(`${label}: display block`, (state) => state.lines.join("\n") === "$$\n  \n$$");
    assert.deepEqual(block.rows, emptyDisplayRow, `${label}: an empty block leaves the row as it was`);
    assert.equal(block.error, false, `${label}: ${JSON.stringify(block)}`);
  };

  const fieldCenter = () => evaluate(`(() => {
        const rect = document.querySelector("math-field")?.getBoundingClientRect();
        return rect ? { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 } : null;
      })()`);

  // 1. completing \begin{name} in a display block applies the whole environment, and only then
  const completed = [];
  for (const [environment, typedName, sourceEnd = `\\end{${environment}}`] of COMPLETION_CASES) {
    const begin = `\\begin{${environment}}`;
    const end = `\\end{${environment}}`;
    await emptyDocument(environment);
    await displayBlock(environment);
    await press(`\\begin{${typedName}`);
    await waitForEvaluation(`(() => {
          const selected = document
            .querySelector('.cm-tooltip-autocomplete [aria-selected="true"] .cm-completionLabel')
            ?.textContent?.trim() ?? "";
          return { ready: selected === ${JSON.stringify(begin)}, selected };
        })()`, `${environment}: completion candidate`);
    const typed = await source();
    assert.deepEqual(typed.rows, emptyDisplayRow, `${environment}: the half-typed environment is not applied: ${JSON.stringify(typed)}`);
    await sleep(120);
    await key("Enter", "Enter", 13);
    const applied = await waitForSource(`${environment}: completed environment applied`, (state) =>
      state.rows.length === 1 &&
      state.rows[0].mode === "display" &&
      state.rows[0].latex.startsWith(begin) &&
      state.rows[0].latex.endsWith(end) &&
      !state.error);
    assert.ok(!applied.lines.join("\n").includes(`${end}}`), `${environment}: a brace left after the environment: ${JSON.stringify(applied.lines)}`);

    const point = await fieldCenter();
    assert.ok(point, `${environment}: formula field`);
    await click(point);
    await sleep(180);
    const afterClick = await source();
    assert.deepEqual(afterClick.rows, applied.rows, `${environment}: a click into the formula keeps it: ${JSON.stringify(afterClick)}`);

    await reload(`${environment}: reloaded`, `(document.querySelector("math-field")?.value ?? "").includes(${JSON.stringify(end)})`);
    const reloaded = await source();
    assert.deepEqual(reloaded.rows, applied.rows, `${environment}: kept after a reload: ${JSON.stringify(reloaded)}`);
    assert.ok(reloaded.lines.some((line) => line.includes(sourceEnd)), `${environment}: source after a reload: ${JSON.stringify(reloaded.lines)}`);
    assert.equal(reloaded.error, false, `${environment}: ${JSON.stringify(reloaded)}`);
    completed.push({ environment, latex: applied.rows[0].latex });
  }

  // 2. outside math: a half-typed \begin is not a text row; the typed } steps over the paired brace and Enter closes
  // the environment around the caret
  await emptyDocument("align*");
  await press("\\begin{align*");
  await sleep(150);
  const halfTyped = await source();
  assert.deepEqual(halfTyped.rows, emptyDisplayRow, `a half-typed \\begin is not applied: ${JSON.stringify(halfTyped)}`);
  await press("}");
  await waitForSource("align*: the list closes once the name is closed", (state) => !state.listOpen);
  const named = await source();
  assert.deepEqual(named.lines, ["\\begin{align*}"], `the typed } steps over the paired one: ${JSON.stringify(named.lines)}`);
  await key("Enter", "Enter", 13);
  const topLevel = await waitForSource("align*: Enter closes the environment", (state) =>
    state.lines.join("\n") === "\\begin{align*}\n  \n\\end{align*}" &&
    state.rows.length === 1 &&
    state.rows[0].mode === "display" &&
    !state.error);
  await press("a&=b");
  // the row is the editor's multi-line row; its alignment point is a marker in the formula
  const topLevelBody = await waitForSource("align*: typed inside the environment", (state) =>
    state.lines[1] === "  a&=b" &&
    /^\\begin\{aligned\}\s*a&=b\s*\\end\{aligned\}$/.test(
      (state.rows[0]?.latex ?? "").replace(/\\class\{visualtex-align-marker\}\{\\kern0pt\}/g, "&"),
    ) &&
    !state.error);

  // 3. inside a display block: Enter after a typed \begin{name} closes it one level deeper
  const enterCases = [];
  for (const environment of ["gather*", "multline*", "equation*"]) {
    const begin = `\\begin{${environment}}`;
    const end = `\\end{${environment}}`;
    await emptyDocument(environment);
    await displayBlock(environment);
    await press(begin);
    await waitForSource(`${environment}: the list closes once the name is closed`, (state) => !state.listOpen);
    await sleep(150);
    await key("Enter", "Enter", 13);
    const opened = await waitForSource(`${environment}: Enter closes the environment`, (state) =>
      state.lines.join("\n") === ["$$", `  ${begin}`, "    ", `  ${end}`, "$$"].join("\n"));
    await press("x");
    const body = await waitForSource(`${environment}: typed inside the environment`, (state) =>
      state.rows.length === 1 &&
      state.rows[0].mode === "display" &&
      state.rows[0].latex.startsWith(begin) &&
      state.rows[0].latex.endsWith(end) &&
      /^x$/.test(state.rows[0].latex.slice(begin.length, -end.length).trim()) &&
      !state.error);
    assert.equal(opened.lines.length, 5);
    await reload(`${environment}: reloaded`, `(document.querySelector("math-field")?.value ?? "").includes(${JSON.stringify(end)})`);
    const reloaded = await source();
    assert.deepEqual(reloaded.rows, body.rows, `${environment}: kept after a reload: ${JSON.stringify(reloaded)}`);
    enterCases.push({ environment, latex: body.rows[0].latex });
  }

  // 4. an environment the formula cannot take stays a draft: marked, not applied, and the source keeps the caret
  await emptyDocument("array without columns");
  await displayBlock("array without columns");
  await press("\\begin{array}");
  await waitForSource("array: the list closes once the name is closed", (state) => !state.listOpen);
  await sleep(150);
  await key("Enter", "Enter", 13);
  const invalid = await waitForSource("array without columns: a marked draft", (state) =>
    state.lines.join("\n") === "$$\n  \\begin{array}\n    \n  \\end{array}\n$$" && state.error);
  assert.deepEqual(invalid.rows, emptyDisplayRow, `an invalid environment is not applied: ${JSON.stringify(invalid)}`);
  const invalidPoint = await fieldCenter();
  assert.ok(invalidPoint, "array without columns: formula field");
  await click(invalidPoint);
  await sleep(180);
  const afterInvalidClick = await source();
  assert.deepEqual(afterInvalidClick.rows, emptyDisplayRow, JSON.stringify(afterInvalidClick));
  assert.ok(afterInvalidClick.lines.includes("  \\end{array}"), JSON.stringify(afterInvalidClick));
  assert.equal(afterInvalidClick.focused, true, `the source keeps the caret while its draft is invalid: ${JSON.stringify(afterInvalidClick)}`);

  console.log(JSON.stringify({ completed, topLevel: topLevelBody.rows[0].latex, topLevelOpened: topLevel.lines, enterCases }, null, 2));
  console.log("Source panel environment completion and auto-close regression passed");
}
