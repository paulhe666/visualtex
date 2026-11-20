// align: rows typed without '&' line up at their first relation; a typed '&' still sets the point

export const names = ["multiline-auto-align"];

export async function run({ clearField, evaluate, key, sleep, typeText, waitForEvaluation }) {
  const charKeys = {
    "(": ["Digit9", 57, 8], ")": ["Digit0", 48, 8], "^": ["Digit6", 54, 8],
    "&": ["Digit7", 55, 8], "+": ["Equal", 187, 8], "=": ["Equal", 187, 0],
    "<": ["Comma", 188, 8], " ": ["Space", 32, 0],
  };
  const typeChars = async (text) => {
    for (const character of text) {
      if (/[0-9]/.test(character)) await key(character, `Digit${character}`, 48 + Number(character));
      else if (charKeys[character]) {
        const [code, vk, modifiers] = charKeys[character];
        await key(character, code, vk, modifiers);
      } else await typeText(character);
    }
  };
  const rowBreak = () => key("Enter", "Enter", 13, 8);
  // the value and, per row of the aligned environment, the left edge of its first relation on screen
  const read = () => evaluate(`(() => {
    const field = document.querySelector('math-field');
    const mf = field?._mathfield;
    const root = mf?.model?.root;
    const array = (root?.body ?? []).find((atom) => atom.type === 'array');
    const lefts = [];
    if (array) {
      for (const row of array.rows) {
        const rel = row.flatMap((cell) => cell ?? []).find((atom) => atom.type === 'mrel');
        const element = rel?.id ? field.shadowRoot.querySelector('[data-atom-id="' + rel.id + '"]') : null;
        lefts.push(element ? Math.round(element.getBoundingClientRect().left * 10) / 10 : null);
      }
    }
    return { value: field?.value ?? '', lefts };
  })()`);
  const compact = (latex) => latex.replace(/\s+/g, "");

  await evaluate(`(() => {
    const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
    persisted.state = persisted.state || {};
    persisted.state.latexFormatProfile = { ...(persisted.state.latexFormatProfile || {}), multilineEnvironment: 'align', numbered: false };
    localStorage.setItem('visualtex-editor', JSON.stringify(persisted));
    location.reload();
    return true;
  })()`).catch(() => {});
  await sleep(900);
  await waitForEvaluation(`(() => ({ ready: Boolean(document.querySelector('math-field')?.isConnected) }))()`, 'reloaded editor');

  // 1. no '&' typed: every row lines up at its first relation, and is written with '&' there
  await clearField();
  await typeChars("ab=(x+1)^2");
  await key("ArrowRight", "ArrowRight", 39);
  await rowBreak();
  await typeChars("=x^2+2x+1");
  await rowBreak();
  await typeChars("abcdef<2x");
  await sleep(400);
  const auto = await read();
  const expected = String.raw`\begin{aligned}ab&=\left(x+1\right)^2\\&=x^2+2x+1\\abcdef&<2x\end{aligned}`;
  if (compact(auto.value) !== expected) throw new Error(`rows must be written with '&' before their first relation: ${auto.value}`);
  if (auto.lefts.length !== 3 || auto.lefts.some((x) => x === null) || Math.max(...auto.lefts) - Math.min(...auto.lefts) > 0.6) {
    throw new Error(`the first relations must line up: ${JSON.stringify(auto.lefts)}`);
  }

  // 2. a typed '&' sets the point where it is typed
  await clearField();
  await typeChars("a=b");
  await rowBreak();
  await typeChars("c+d&=e");
  await sleep(400);
  const manual = await read();
  const expectedManual = String.raw`\begin{aligned}a&=b\\c+d&=e\end{aligned}`;
  if (compact(manual.value) !== expectedManual) throw new Error(`a typed '&' must stay the row's point: ${manual.value}`);

  // 3. a typed '&' away from the relation wins over the relation
  await clearField();
  await typeChars("x=1");
  await rowBreak();
  await typeChars("&y=22");
  await sleep(400);
  const left = await read();
  const expectedLeft = String.raw`\begin{aligned}x&=1\\&y=22\end{aligned}`;
  if (compact(left.value) !== expectedLeft) throw new Error(`'&' at a row's start must keep the row left-aligned: ${left.value}`);

  // 4. a row inserted between rows, and an empty row removed again: the rows after them still line up
  await clearField();
  await typeChars("a=b");
  await rowBreak();
  await typeChars("cc=d");
  await key("ArrowUp", "ArrowUp", 38);
  await rowBreak();
  await typeChars("eee=f");
  await rowBreak();
  await key("Backspace", "Backspace", 8);
  await sleep(400);
  const inserted = await read();
  const expectedInserted = String.raw`\begin{aligned}a&=b\\eee&=f\\cc&=d\end{aligned}`;
  if (compact(inserted.value) !== expectedInserted) throw new Error(`rows after an inserted or removed row must keep lining up: ${inserted.value}`);
  if (inserted.lefts.length !== 3 || Math.max(...inserted.lefts) - Math.min(...inserted.lefts) > 0.6) {
    throw new Error(`the first relations must line up after rows were added and removed: ${JSON.stringify(inserted.lefts)}`);
  }

  console.log(JSON.stringify({ auto, manual, left, inserted }, null, 2));
  console.log("Multiline auto alignment regression passed");
}
