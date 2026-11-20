// multiline formulas and their alignment

export const names = ["multiline-align"];

export async function run({ clearField, evaluate, key, sleep, typeText, waitForEvaluation }) {
  const charKeys = {
    "(": ["Digit9", 57, 8], ")": ["Digit0", 48, 8], "^": ["Digit6", 54, 8],
    "&": ["Digit7", 55, 8], "+": ["Equal", 187, 8], "=": ["Equal", 187, 0],
    "'": ["Quote", 222, 0], " ": ["Space", 32, 0],
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
  const readMultiline = () => evaluate(`(() => {
        const fields = [...document.querySelectorAll('math-field')];
        const source = document.querySelector('.cm-content')?.innerText ?? null;
        return { values: fields.map((field) => field.value), source };
      })()`);

  const useMultilineEnvironment = async (environment) => {
    await evaluate(`(() => {
          const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}');
          persisted.state = persisted.state || {};
          persisted.state.latexFormatProfile = { ...(persisted.state.latexFormatProfile || {}), multilineEnvironment: ${JSON.stringify(environment)}, numbered: false };
          localStorage.setItem('visualtex-editor', JSON.stringify(persisted));
          location.reload();
          return true;
        })()`).catch(() => {});
    await sleep(900);
    await waitForEvaluation(`(() => ({ ready: Boolean(document.querySelector('math-field')?.isConnected) }))()`, 'reloaded editor');
  };

  // F03: '&' typed before the first Shift+Enter becomes the row's align point.
  await useMultilineEnvironment('align');
  await clearField();
  await typeChars("f(x)&=(x+1)^2");
  await key("ArrowRight", "ArrowRight", 39);
  await key("Enter", "Enter", 13, 8);
  await typeChars("&=x^2");
  await typeChars("+2x+1");
  await key("Enter", "Enter", 13, 8);
  await typeChars("f'(x)&=2x+2");
  await sleep(300);
  const shiftEnterAlign = (await readMultiline()).values[0].replace(/\s+/g, "");
  const expectedAlign = String.raw`\begin{aligned}f\left(x\right)&=\left(x+1\right)^2\\&=x^2+2x+1\\f^{\prime}\left(x\right)&=2x+2\end{aligned}`;
  if (shiftEnterAlign !== expectedAlign) {
    throw new Error(`First Shift+Enter row must split at its typed '&': ${shiftEnterAlign}`);
  }
  // Backspace on the empty row a Shift+Enter created removes that row and,
  // with one row left, the aligned/gathered wrapper too.
  const backspaceDowngrade = {};
  for (const environment of ["gather", "align"]) {
    await useMultilineEnvironment(environment);
    await clearField();
    await typeChars("x");
    await key("Enter", "Enter", 13, 8);
    await sleep(150);
    const upgraded = (await readMultiline()).values[0];
    if (!/\\begin\{(?:aligned|gathered)\}/.test(upgraded)) {
      throw new Error(`${environment}: Shift+Enter did not create a multiline formula: ${upgraded}`);
    }
    await key("Backspace", "Backspace", 8);
    await sleep(200);
    const state = await readMultiline();
    backspaceDowngrade[environment] = { upgraded, after: state.values };
    if (state.values.length !== 1 || state.values[0].replace(/\s+/g, "") !== "x") {
      throw new Error(`${environment}: Backspace must leave the ordinary formula x: ${JSON.stringify(state.values)}`);
    }
  }
  console.log(JSON.stringify({ shiftEnterAlign, backspaceDowngrade }, null, 2));
  console.log("Targeted desktop multiline alignment regression passed");
  return;
}
