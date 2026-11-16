// The L-shaped caret (AxMath's, docs/editing-core/DESIGN.md §5): in the formula field the caret carries a line under the whole
// region (slot) it is in: from the first item's left edge to the last item's right edge. Inside an accent it covers only the
// accent's content, in a numerator or denominator only that; at the top level the whole row; no line at the start of a region.
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const KEYS = { Right: [39, "ArrowRight"], Left: [37, "ArrowLeft"], Home: [36, "Home"] };

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

  async function open(latex) {
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(1200);
    await evaluate(`(() => {
      for (const k of ["visualtex.onboarding.v3.completed", "visualtex.office.macos.first-run.v1.completed", "visualtex.onboarding.macos.desktop.v1.2.0.completed", "visualtex.office.macos.native-first-run.v1.2.0.completed"]) localStorage.setItem(k, "true");
      localStorage.setItem("visualtex.document.tutorial-shown.v1", "1");
      const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
      persisted.state = { ...(persisted.state || {}), theme: "light", lines: [{ id: "line-1", latex: ${JSON.stringify(latex)} }], activeLineId: "line-1" };
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

  // the caret now: whether it has the line, the line's ends, the bar's height and its x (client px), and the atoms' boxes by their letter
  const look = () =>
    evaluate(`(() => {
      const root = document.querySelector("math-field.visual-mathfield").shadowRoot;
      const caret = root.querySelector(".ML__caret, .ML__text-caret, .ML__latex-caret");
      if (!caret) return null;
      const rect = caret.getBoundingClientRect();
      const hasLine = caret.classList.contains("ML__vtx-caret--line");
      const before = getComputedStyle(caret, "::before");
      const after = getComputedStyle(caret, "::after");
      const barWidth = parseFloat(after.width) || 0;
      const atoms = {};
      for (const el of root.querySelectorAll("[data-atom-id]")) {
        const own = [...el.childNodes].filter((n) => n.nodeType === 3).map((n) => n.textContent).join("");
        if (own.length === 1) atoms[own] = el.getBoundingClientRect().toJSON();
      }
      return {
        hasLine,
        x: rect.left,
        height: rect.height,
        barWidth,
        lineLeft: hasLine ? rect.left + parseFloat(before.left) : null,
        lineRight: hasLine ? rect.left + parseFloat(before.left) + parseFloat(before.width) : null,
        atoms,
      };
    })()`);

  const near = (a, b, label, tolerance = 4) => assert(Math.abs(a - b) <= tolerance, `${label}: ${a.toFixed(1)} is not ${b.toFixed(1)} (±${tolerance})`);

  // --- an accent: a, the hat over BC, d
  await open("a\\hat{BC}d");
  const steps = [];
  for (let i = 0; i < 7; i += 1) {
    steps.push(await look());
    if (i < 6) await key("Right");
  }
  const atoms = steps[0].atoms;
  assert(atoms.a && atoms.B && atoms.C && atoms.d, `atoms not found: ${Object.keys(atoms)}`);
  // 0: the start of the row: the bar alone
  assert(!steps[0].hasLine, "no line at the start of the row");
  // 1: after a (the top level): the line under the whole row
  assert(steps[1].hasLine, "a line after a");
  near(steps[1].lineLeft, atoms.a.left, "the top-level line starts at a", 5);
  near(steps[1].lineRight, atoms.d.right, "the top-level line ends at d", 5);
  // 2: inside the accent, at its start: the bar alone, shorter than the top-level one (the hat is not in the accent's content)
  assert(!steps[2].hasLine, "no line at the start of the accent's content");
  // 3 and 4: inside the accent, after B and after C: the line under BC only
  for (const i of [3, 4]) {
    assert(steps[i].hasLine, `a line inside the accent (step ${i})`);
    near(steps[i].lineLeft, atoms.B.left, `step ${i}: the line starts at B`, 5);
    near(steps[i].lineRight, atoms.C.right, `step ${i}: the line ends at C`, 5);
  }
  assert(steps[3].lineRight < steps[1].lineRight - 5, "the line inside the accent ends before the top-level line does (it does not run under d)");
  assert(steps[3].lineLeft > steps[1].lineLeft + 5, "the line inside the accent starts after the top-level line does (it does not run under a)");
  assert(steps[3].height < steps[1].height, `the bar inside the accent (${steps[3].height.toFixed(1)}) is shorter than the top-level bar (${steps[1].height.toFixed(1)}), which holds the hat`);
  // 5 and 6: after the accent and after d: the top level again
  for (const i of [5, 6]) {
    assert(steps[i].hasLine, `a line at the top level (step ${i})`);
    near(steps[i].lineLeft, atoms.a.left, `step ${i}: the line starts at a`, 5);
    near(steps[i].lineRight, atoms.d.right, `step ${i}: the line ends at d`, 5);
  }
  // the bar stands on the end of the line when the caret is at the end of its region
  near(steps[4].x, steps[4].lineRight - steps[4].barWidth / 2, "inside the accent at its end the bar stands on the line's end", 2);
  near(steps[6].x, steps[6].lineRight - steps[6].barWidth / 2, "at the end of the row the bar stands on the line's end", 2);

  // --- a fraction: a, bc over de, f
  await open("a\\frac{bc}{de}f");
  const f = [];
  for (let i = 0; i < 10; i += 1) {
    f.push(await look());
    if (i < 9) await key("Right");
  }
  const fa = f[0].atoms;
  assert(fa.b && fa.c && fa.d && fa.e && fa.f, `fraction atoms not found: ${Object.keys(fa)}`);
  assert(!f[2].hasLine, "no line at the start of the numerator");
  for (const i of [3, 4]) {
    assert(f[i].hasLine, `a line in the numerator (step ${i})`);
    near(f[i].lineLeft, fa.b.left, `numerator step ${i}: the line starts at b`, 5);
    near(f[i].lineRight, fa.c.right, `numerator step ${i}: the line ends at c`, 5);
  }
  assert(!f[5].hasLine, "no line at the start of the denominator");
  for (const i of [6, 7]) {
    assert(f[i].hasLine, `a line in the denominator (step ${i})`);
    near(f[i].lineLeft, fa.d.left, `denominator step ${i}: the line starts at d`, 5);
    near(f[i].lineRight, fa.e.right, `denominator step ${i}: the line ends at e`, 5);
  }
  assert(f[1].hasLine && f[1].lineRight > fa.f.right - 6, "the top-level line runs under the fraction to f");
  // an empty slot has no line (its marker is its caret)
  await open("a\\frac{}{}b");
  await key("Right");
  await key("Right");
  const empty = await look();
  assert(empty && !empty.hasLine, "no line in an empty slot");

  console.log("L-shaped caret (a line under the region the caret is in, only its content inside an accent or a fraction, none at a region's start or in an empty slot): passed");
}, { windowSize: "1100,800" });
