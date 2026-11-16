// Every argument of a structure is reachable with the arrow keys, and what is typed there is saved in that argument.
// They were not: \mel's operator (its slot sits in a \smash of the expansion: \mel{\psi}{\hat{a}}{\psi} let the caret
// visit only the two psi), \expval's and \vev's operator, \va's and \vu's vector (inside the expansion's \vec), the
// cells of \pmqty, \bmqty, \vmqty and the small ones and \qq's text (plain macros: their slot marker fell apart inside
// the environment, or was read as text), and the body of accents (\hat, \vec, \bar...), \underarc, \underparen and
// \mathop once something was typed in it.
import assert from "node:assert/strict";
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const compact = (latex) => latex.replace(/\s+/g, "");

await withAppBrowser(async ({ baseUrl, client }) => {
  const evaluate = async (expression) => {
    const result = await client.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
    return result.result.value;
  };
  for (let tries = 0; tries < 300 && !String(await evaluate("location.href")).startsWith(baseUrl); tries += 1) await sleep(50);
  await evaluate(`(() => {
    for (const k of ["visualtex.onboarding.v3.completed", "visualtex.office.macos.first-run.v1.completed", "visualtex.onboarding.macos.desktop.v1.2.0.completed", "visualtex.office.macos.native-first-run.v1.2.0.completed"]) localStorage.setItem(k, "true");
    const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
    persisted.state = { ...(persisted.state || {}), lines: [{ id: "l1", latex: "" }], activeLineId: "l1" };
    localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
    localStorage.setItem("visualtex.workspace-content.v1", JSON.stringify({ state: { content: "formula" }, version: 0 }));
  })()`);
  await client.send("Page.reload", { ignoreCache: true });
  await sleep(1500);
  for (let t = 0; t < 200 && !(await evaluate(`Boolean(document.querySelector("math-field.visual-mathfield")?._mathfield)`)); t += 1) await sleep(50);
  const key = async (value, code, vk) => {
    const common = { key: value, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk };
    const text = value.length === 1 ? { text: value, unmodifiedText: value } : {};
    await client.send("Input.dispatchKeyEvent", { type: value.length === 1 ? "keyDown" : "rawKeyDown", ...common, ...text });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
    await sleep(25);
  };
  const set = (latex) => evaluate(`(() => {
    const f = document.querySelector("math-field.visual-mathfield");
    f.setValue(${JSON.stringify(latex)}, { mode: "math", format: "latex", insertionMode: "replaceAll", silenceNotifications: true });
    f.focus(); f.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
    f.position = 0;
    return f._mathfield.model.lastOffset;
  })()`);
  const position = () => evaluate(`document.querySelector("math-field.visual-mathfield")._mathfield.model.position`);
  // the stops after `letter` inside the formula's slots (not copies in a phantom), and after the structure holding it
  const offsets = (letter) => evaluate(`(() => {
    const m = document.querySelector("math-field.visual-mathfield")._mathfield.model;
    const visible = (a) => { for (let p = a.parent; p; p = p.parent) if (p.type === "phantom" && p.command !== "\\\\smash") return false; return true; };
    return m.atoms.filter((a) => a.value === ${JSON.stringify(letter)} && a.type !== "first" && visible(a)).map((a) => m.offsetOf(a));
  })()`);
  // Right from the start until the caret stands at `target` (or the end); the stops visited
  const walkTo = async (target, last) => {
    const seen = [0];
    for (let i = 0; i < 80; i += 1) {
      const now = await position();
      if (now === target) return seen;
      if (i > 0 && now === last) break;
      await key("ArrowRight", "ArrowRight", 39);
      seen.push(await position());
    }
    return seen;
  };

  // 1. the user's formula: the caret reaches the operator's slot, inside and after \hat{a}; typing goes there
  const user = "\\mel{\\psi}{\\hat{a}}{\\psi}";
  let last = await set(user);
  const [insideHat] = await offsets("a");
  assert.ok(insideHat !== undefined, "the operator's a is in the formula");
  const visited = await walkTo(last, last);
  assert.ok(visited.includes(insideHat), `Right from the start reaches the a under the hat in \\mel's middle argument (visited ${visited.join(",")}, a at ${insideHat})`);
  await set(user);
  await walkTo(insideHat, last);
  await key("x", "KeyX", 88);
  assert.equal(compact(await evaluate(`document.querySelector("math-field.visual-mathfield").getValue("latex")`)), compact("\\mel{\\psi}{\\hat{ax}}{\\psi}"), "typing under the hat edits \\mel's operator");
  // after the hat, still in the operator's slot
  await set(user);
  await walkTo(insideHat, last);
  await key("ArrowRight", "ArrowRight", 39);
  await key("y", "KeyY", 89);
  assert.equal(compact(await evaluate(`document.querySelector("math-field.visual-mathfield").getValue("latex")`)), compact("\\mel{\\psi}{\\hat{a}y}{\\psi}"), "after the hat the caret is still in \\mel's operator");

  // 2. one argument of each kind: reach the letter, type after it, and the saved LaTeX has it in that argument
  const cases = [
    ["\\expval{A}{\\psi}", "A", "\\expval{Ax}{\\psi}"],
    ["\\vev{A}", "A", "\\vev{Ax}"],
    ["\\mel*{a}{B}{c}", "B", "\\mel*{a}{Bx}{c}"],
    ["\\pmqty{a&b\\\\c&d}", "b", "\\pmqty{a&bx\\\\c&d}"],
    ["\\bmqty{a&b\\\\c&d}", "c", "\\bmqty{a&b\\\\cx&d}"],
    ["\\vmqty{a&b}", "b", "\\vmqty{a&bx}"],
    ["\\spmqty{a&b}", "a", "\\spmqty{ax&b}"],
    ["\\qq{word}", "d", "\\qq{wordx}"],
    ["\\va{q}", "q", "\\va{qx}"],
    ["\\vu{q}", "q", "\\vu{qx}"],
    ["\\hat{q}", "q", "\\hat{qx}"],
    ["\\vec{v}", "v", "\\vec{vx}"],
    ["\\bar{z}", "z", "\\bar{zx}"],
    ["\\underarc{q}", "q", "\\underarc{qx}"],
    ["\\mathop{q}", "q", "\\mathop{qx}"],
  ];
  for (const [latex, letter, expected] of cases) {
    last = await set(latex);
    const [target] = await offsets(letter);
    assert.ok(target !== undefined, `${latex}: the ${letter} is in the formula`);
    const seen = await walkTo(target, last);
    assert.ok(seen.includes(target), `${latex}: Right reaches the ${letter} (visited ${seen.join(",")}, target ${target})`);
    await key("x", "KeyX", 88);
    const saved = compact(await evaluate(`document.querySelector("math-field.visual-mathfield").getValue("latex")`));
    assert.equal(saved, compact(expected), `${latex}: what is typed after ${letter} is saved in that argument`);
  }
  console.log("every argument reachable and edited in place: passed");
}, { windowSize: "1100,800" });
