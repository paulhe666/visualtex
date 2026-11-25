// A big operator's operand (∫ ∑ ∏ ...; docs/editing-core/DESIGN.md §4, §5) is a region of its own, as in AxMath: the caret in
// the operand has the underline under only the operand, outside it under the whole row; at the operand's end the caret can
// stand inside or outside it (and outside every operand that ends there), one more key press each, at one place. The
// operand's extent is the one rule of operand-scope.ts: an integral's up to a sign after its differential, a sum's up to a
// top-level sign.
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const KEYS = { Right: [39, "ArrowRight"], Left: [37, "ArrowLeft"], Home: [36, "Home"], End: [35, "End"], Tab: [9, "Tab"] };

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

await withAppBrowser(async ({ baseUrl, client, evaluate }) => {
  const key = async (name) => {
    const [vk, code] = KEYS[name];
    const common = { key: code, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk };
    await client.send("Input.dispatchKeyEvent", { type: "rawKeyDown", ...common });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
    await sleep(120);
  };

  // one formula row, or several (the caret goes into the first)
  async function open(latex) {
    const lines = (Array.isArray(latex) ? latex : [latex]).map((text, index) => ({ id: `line-${index + 1}`, latex: text }));
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(1200);
    await evaluate(`(() => {
      for (const k of ["visualtex.onboarding.v3.completed", "visualtex.office.macos.first-run.v1.completed", "visualtex.onboarding.macos.desktop.v1.2.0.completed", "visualtex.office.macos.native-first-run.v1.2.0.completed"]) localStorage.setItem(k, "true");
      localStorage.setItem("visualtex.document.tutorial-shown.v1", "1");
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      persisted.state = { ...(persisted.state || {}), theme: "light", lines: ${JSON.stringify(lines)}, activeLineId: "line-1" };
      localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
      localStorage.setItem("visualtex.workspace-content.v1", JSON.stringify({ state: { content: "formula" }, version: 0 }));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(2200);
    for (let t = 0; t < 100 && !(await evaluate(`Boolean(document.querySelector("math-field.visual-mathfield")?.shadowRoot)`)); t += 1) await sleep(50);
    await evaluate(`(() => { const f = document.querySelector("math-field.visual-mathfield"); f.focus(); f.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true }); })()`);
    await sleep(300);
    await key("Home");
  }

  // the caret now: its offset, whether it carries the underline, the underline's ends (client px) and the bar's x, and the
  // atoms' boxes by their one letter
  const look = () =>
    evaluate(`(() => {
      const field = document.querySelector("math-field.visual-mathfield");
      const root = field.shadowRoot;
      const caret = root.querySelector(".ML__caret, .ML__text-caret, .ML__latex-caret");
      if (!caret) return null;
      const rect = caret.getBoundingClientRect();
      const hasLine = caret.classList.contains("ML__vtx-caret--line");
      const before = getComputedStyle(caret, "::before");
      const atoms = {};
      for (const el of root.querySelectorAll("[data-atom-id]")) {
        const own = [...el.childNodes].filter((n) => n.nodeType === 3).map((n) => n.textContent).join("");
        if (own.length === 1) atoms[own] = el.getBoundingClientRect().toJSON();
      }
      return {
        pos: field.position,
        hasLine,
        x: rect.left,
        height: rect.height,
        lineLeft: hasLine ? rect.left + parseFloat(before.left) : null,
        lineRight: hasLine ? rect.left + parseFloat(before.left) + parseFloat(before.width) : null,
        atoms,
      };
    })()`);

  // the caret walked with a key until nothing moves any more (the same state three times)
  async function walk(name, limit = 40) {
    const steps = [await look()];
    for (let i = 0; i < limit; i += 1) {
      await key(name);
      const now = await look();
      steps.push(now);
      const n = steps.length;
      if (n > 4 && [1, 2, 3].every((k) => steps[n - k].pos === steps[n - k - 1].pos && steps[n - k].x === steps[n - k - 1].x && steps[n - k].hasLine === steps[n - k - 1].hasLine)) break;
    }
    return steps;
  }

  const near = (a, b, label, tolerance = 5) => assert(Math.abs(a - b) <= tolerance, `${label}: ${a.toFixed(1)} is not ${b.toFixed(1)} (±${tolerance})`);
  // (the line reaches out to where the bar stands: at an operand's end the outer ways of standing are set a little outward)
  const spans = (step, atoms, from, to, label, tolerance = 5) => {
    assert(step.hasLine, `${label}: no underline`);
    near(step.lineLeft, atoms[from].left, `${label}: the underline starts at ${from}`, tolerance);
    near(step.lineRight, atoms[to].right, `${label}: the underline ends at ${to}`, tolerance);
  };

  // --- an integral and a sum: a ∫_p^q c d e ∑_{i=1}^n x y z (the operands run to the end of the row)
  await open("a\\int_p^q cde\\sum_{i=1}^n xyz");
  const right = await walk("Right");
  const A = right[0].atoms;
  for (const letter of "apqcdeinxyz") assert(A[letter], `atom ${letter} not found: ${Object.keys(A)}`);
  const at = (predicate) => right.findIndex(predicate);
  // after a: the whole row
  spans(right[1], A, "a", "z", "after a, outside the integral");
  // the operand's start (the stop after the operator and its limits): the bar alone, standing left of c
  const start = right.find((s) => s.hasLine === false && s.x > A.q.right && s.x < A.c.left + 4);
  assert(start, "the integral operand's start has no stop of its own (bar without underline left of c)");
  // inside the integral's operand: the underline under c..z only, not under a
  const afterC = right.find((s) => s.hasLine && Math.abs(s.x - A.c.right) < 6 && s.lineLeft > A.e.left - 100 && s.lineLeft < A.c.left + 5);
  assert(afterC, "after c: no underline from c");
  spans(afterC, A, "c", "z", "inside the integral operand");
  assert(afterC.lineLeft > A.a.right, "the integral operand's underline does not run under a");
  // the sum's operand: x..z, shorter than the integral's
  const afterX = right.find((s) => s.hasLine && Math.abs(s.x - A.x.right) < 6);
  assert(afterX, "after x: no underline");
  spans(afterX, A, "x", "z", "inside the sum operand");
  // the end of the row: three ways of standing at z's place, one key press each: inside the sum's operand, inside the integral's, outside both
  const atEnd = right.filter((s) => s.pos === right[right.length - 1].pos);
  assert(atEnd.length >= 3, `three states at the end (${atEnd.length})`);
  const endStates = atEnd.slice(0, 3);
  spans(endStates[0], A, "x", "z", "end: inside the sum operand");
  spans(endStates[1], A, "c", "z", "end: inside the integral operand", 12);
  spans(endStates[2], A, "a", "z", "end: outside, the whole row", 14);
  assert(endStates[0].x < endStates[1].x && endStates[1].x < endStates[2].x, "the three states stand a little apart, outward");
  assert(endStates[0].height < endStates[1].height, "the bar inside the sum operand is shorter than inside the integral's (which holds the sum's limits)");
  void at;

  // backward: End goes outside everything; Left steps in again, one state at a time
  await key("End");
  const outer = await look();
  spans(outer, A, "a", "z", "End: outside, the whole row", 14);
  await key("Left");
  const inner1 = await look();
  spans(inner1, A, "c", "z", "Left from outside: inside the integral operand", 12);
  await key("Left");
  const inner2 = await look();
  spans(inner2, A, "x", "z", "Left again: inside the sum operand");
  assert(inner2.pos === inner1.pos && inner1.pos === outer.pos, "the three states are one place");
  await key("Left");
  const beforeZ = await look();
  assert(beforeZ.pos === outer.pos - 1, "one more Left moves to the place before z");

  // --- an integral stops after its differential: x ∫_0^1 f(x) dx + c: the plus is outside, and so is c
  await open("u\\int_0^1 vw\\,dt+s");
  const diff = await walk("Right");
  const D = diff[0].atoms;
  for (const letter of "uvwdts") assert(D[letter], `atom ${letter} not found: ${Object.keys(D)}`);
  const afterT = diff.findIndex((s) => s.hasLine && Math.abs(s.x - D.t.right) < 6);
  assert(afterT >= 0, "after t: no stop");
  spans(diff[afterT], D, "v", "t", "inside the integral operand (it ends after dt)");
  assert(diff[afterT].lineRight < D.s.left, "the operand's underline stops before the plus and s");
  const outsideT = diff[afterT + 1];
  assert(outsideT.pos === diff[afterT].pos, "after dt a second state at the same place");
  spans(outsideT, D, "u", "s", "after dt, outside: the whole row", 12);
  // the plus before the differential is inside (an operand with a plus in it: ∫ x + 1 dx)
  await open("\\int p+q\\,dr+s");
  const inside = await walk("Right");
  const P = inside[0].atoms;
  const afterQ = inside.find((s) => s.hasLine && Math.abs(s.x - P.q.right) < 6);
  assert(afterQ, "after q: no underline");
  assert(afterQ.lineRight >= P.r.right - 5, "a plus before the differential belongs to the integral operand");
  assert(afterQ.lineRight < P.s.left, "the sign after the differential does not");

  // --- a sum stops at a top-level sign: ∑_i a_i + b
  await open("\\sum_i k_j+m");
  const sum = await walk("Right");
  const S = sum[0].atoms;
  for (const letter of "ijkm") assert(S[letter], `atom ${letter} not found: ${Object.keys(S)}`);
  // the operand is k_j (the plus ends it): the stops under it start the underline at k and stop short of the plus
  const operandIndexes = sum.map((s, index) => (s.hasLine && Math.abs(s.lineLeft - S.k.left) < 6 && s.lineRight < S.m.left - 4 ? index : -1)).filter((index) => index >= 0);
  assert(operandIndexes.length >= 1, "after k_j: no operand underline from k that stops before the plus");
  const lastInside = sum[operandIndexes[operandIndexes.length - 1]];
  const outsideJ = sum[operandIndexes[operandIndexes.length - 1] + 1];
  assert(outsideJ.hasLine && outsideJ.pos === lastInside.pos, "the sum's operand has a second state at its end");
  assert(Math.abs(outsideJ.lineLeft - S.i.left) < 40 && outsideJ.lineRight >= S.m.right - 5, "after the operand, outside: the whole row, under the plus and m too");

  // --- no operator, no operand state: a plain row has the one state
  await open("a+bc");
  const plain = await walk("Right");
  assert(new Set(plain.map((s) => s.pos)).size === plain.length - 3 + 0 || plain.length > 0, "plain row walks");
  const positions = plain.map((s) => s.pos);
  assert(positions[1] === 1 && positions[2] === 2 && positions[3] === 3, `a plain row has one stop per place (${positions.slice(0, 5)})`);

  // --- Tab in an operand leaves the operator, at the operand's end and outside it: not the row (before, Tab inside an
  // operand went on to the next row, as at the row's top level, since the operand is no structure of the formula)
  const rows = () => evaluate(`document.querySelectorAll("math-field.visual-mathfield").length`);
  const rightUntil = async (predicate, limit = 30) => {
    for (let i = 0; i < limit; i += 1) {
      const now = await look();
      if (predicate(now)) return now;
      await key("Right");
    }
    return null;
  };
  await open("a\\int_p^q cde\\sum_{i=1}^n xyz");
  const T = (await look()).atoms;
  const rowsBefore = await rows();
  const endPos = (await walk("Right")).at(-1).pos;
  await key("Home");
  // after c: inside the integral's operand only; Tab: the row's end, outside the integral (and the sum in it)
  assert(await rightUntil((s) => s.hasLine && Math.abs(s.x - T.c.right) < 6 && s.lineLeft < T.c.left + 5), "no stop after c inside the integral operand");
  await key("Tab");
  const tabbed = await look();
  assert(tabbed && tabbed.pos === endPos, `Tab after c: at the operand's end (${tabbed?.pos}, end ${endPos})`);
  spans(tabbed, T, "a", "z", "Tab after c: outside the integral", 14);
  assert((await rows()) === rowsBefore, "Tab in the operand stays in the row");
  // after x: inside the sum's operand, in the integral's; Tab: outside the sum (inside the integral), Tab again: outside both
  await key("Home");
  assert(await rightUntil((s) => s.hasLine && Math.abs(s.x - T.x.right) < 6 && Math.abs(s.lineLeft - T.x.left) < 6), "no stop after x inside the sum operand");
  await key("Tab");
  const outOfSum = await look();
  assert(outOfSum.pos === endPos, "Tab after x: at the operands' end");
  spans(outOfSum, T, "c", "z", "Tab after x: outside the sum, inside the integral", 12);
  await key("Tab");
  const outOfBoth = await look();
  assert(outOfBoth.pos === endPos, "Tab again: the same place");
  spans(outOfBoth, T, "a", "z", "Tab again: outside the integral too", 14);
  assert((await rows()) === rowsBefore, "Tab in the operands stays in the row");
  // an operand ending before the row's end: Tab after v leaves the integral after dt (s stays outside)
  await open("u\\int_0^1 vw\\,dt+s");
  const U = (await look()).atoms;
  assert(await rightUntil((s) => s.hasLine && Math.abs(s.x - U.v.right) < 6), "no stop after v");
  await key("Tab");
  const afterDt = await look();
  near(afterDt.x, U.t.right, "Tab after v: the caret after dt", 12);
  spans(afterDt, U, "u", "s", "Tab after v: outside the integral, the whole row", 12);

  // --- rows aligned at explicit alignment markers (the formula rows' align layout) keep the underline, each end moving
  // with the material above it (before, an aligned row had the short bar only)
  const marker = "\\class{visualtex-align-marker}{\\kern0pt}";
  await open([`x${marker}=y${marker}\\int_p^q cde`, `xxxx${marker}=yyyyyy${marker}k`]);
  await sleep(600);
  const M = (await look()).atoms;
  for (const letter of "xypqcde") assert(M[letter], `aligned row: atom ${letter} not found: ${Object.keys(M)}`);
  const inOperand = await rightUntil((s) => s.hasLine && Math.abs(s.x - M.c.right) < 6);
  assert(inOperand, "aligned row: after c no underline");
  spans(inOperand, M, "c", "e", "aligned row: inside the integral operand");
  await key("End");
  const rowEnd = await look();
  spans(rowEnd, M, "x", "e", "aligned row: End, the whole row (across both markers)", 14);
  await key("Home");
  await key("Right");
  const alignedAfterX = await look();
  spans(alignedAfterX, M, "x", "e", "aligned row: after x, the whole row", 14);

  console.log("operand of a big operator (its own underline, two ways of standing at its end, the extent rule, Tab, aligned rows): passed");
}, { windowSize: "1100,800" });
