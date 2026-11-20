// A typed accent whose content is deleted goes back to the source shell it was typed as (\hat{} with the caret between
// the braces, the LaTeX box), exactly as right after typing \hat: it used to stay an empty accent with a placeholder.
// Every accent command of the source-first list, emptied by Backspace, Delete, deleting a selection and cutting; then
// typing in the shell and leaving it gives the accent again. Typing undone by Backspace step by step; Delete in an
// empty shell.
import assert from "node:assert/strict";
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const compact = (latex) => latex.replace(/\s+/g, "");
const ACCENTS = ["\\acute", "\\grave", "\\dot", "\\ddot", "\\dddot", "\\ddddot", "\\tilde", "\\bar", "\\breve", "\\check",
  "\\hat", "\\vec", "\\widehat", "\\widetilde", "\\overline", "\\overrightarrow", "\\overleftarrow", "\\overleftrightarrow", "\\mathring"];

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
  // the app focuses the active formula once it has started, the caret at its end: on a slow machine that came after
  // the first caret this test set (the caret jumped to the end), so wait for it
  for (let t = 0; t < 100 && !(await evaluate(`document.activeElement === document.querySelector("math-field.visual-mathfield")`)); t += 1) await sleep(50);
  await sleep(300);
  const key = async (value, code, vk, modifiers = 0) => {
    const common = { key: value, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers };
    const text = value.length === 1 && !modifiers ? { text: value, unmodifiedText: value } : {};
    await client.send("Input.dispatchKeyEvent", { type: value.length === 1 && !modifiers ? "keyDown" : "rawKeyDown", ...common, ...text });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
    await sleep(40);
  };
  const field = `document.querySelector("math-field.visual-mathfield")`;
  const state = () => evaluate(`(() => { const f = ${field}; const m = f._mathfield.model;
    return { value: f.getValue("latex"), mode: m.mode, wrapper: f.dataset.visualtexSourceWrapperCommand ?? null, content: f.dataset.visualtexSourceWrapperContent ?? null,
      placeholder: m.atoms.some((a) => a.type === "placeholder"), position: m.position, last: m.lastOffset }; })()`);
  // the formula x\cmd{a}y with the caret after a (inside the accent) or before it
  const start = async (command, where) => {
    const target = await evaluate(`(() => { const f = ${field};
      f.setValue(${JSON.stringify("x" + command + "{a}y")}, { mode: "math", format: "latex", insertionMode: "replaceAll", silenceNotifications: true });
      f.focus(); f.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      const m = f._mathfield.model; f.position = 0;
      const a = m.atoms.find((atom) => atom.value === "a");
      return a ? m.offsetOf(a) - (${where === "before" ? 1 : 0}) : -1; })()`);
    assert.ok(target >= 0, `${command}{a}: the a is in the formula`);
    for (let i = 0; i < 12 && (await state()).position !== target; i += 1) await key("ArrowRight", "ArrowRight", 39);
    assert.equal((await state()).position, target, `${command}{a}: the caret reaches ${where} the a inside the accent`);
  };
  const expectShell = async (command, how) => {
    const now = await state();
    assert.equal(now.wrapper, command, `${command}, ${how}: the emptied accent is the source shell again (${JSON.stringify(now)})`);
    assert.equal(now.content, "", `${command}, ${how}: the shell is empty`);
    assert.equal(now.mode, "latex", `${command}, ${how}: the caret is in the LaTeX box`);
    assert.ok(!now.placeholder, `${command}, ${how}: no empty accent with a placeholder is left`);
    assert.equal(compact(now.value), compact(`x${command}{}y`), `${command}, ${how}: the formula shows ${command}{} between x and y`);
  };
  // typing in the shell and leaving it: the accent again, with the new content
  const retype = async (command, how) => {
    await key("b", "KeyB", 66);
    await key("ArrowRight", "ArrowRight", 39);
    await key("ArrowRight", "ArrowRight", 39);
    const now = await state();
    assert.equal(compact(now.value), compact(`x${command}{b}y`), `${command}, ${how}: typed in the shell and left, it is the accent again (${now.value})`);
    assert.equal(now.wrapper, null, `${command}, ${how}: the shell is closed`);
  };

  const rendered = [];
  for (const command of ACCENTS) {
    // (an accent the kernel does not draw as a structure with a body is not this case)
    await start(command, "after");
    await key("Backspace", "Backspace", 8);
    await expectShell(command, "Backspace");
    await retype(command, "Backspace");

    await start(command, "before");
    await key("Delete", "Delete", 46);
    await expectShell(command, "Delete");
    await retype(command, "Delete");

    await start(command, "after");
    await key("ArrowLeft", "ArrowLeft", 37, 8);
    await key("Backspace", "Backspace", 8);
    await expectShell(command, "selection deleted");

    await start(command, "after");
    await key("ArrowLeft", "ArrowLeft", 37, 8);
    await evaluate(`(() => { const sink = ${field}.shadowRoot.querySelector('[part="keyboard-sink"]');
      sink.dispatchEvent(new ClipboardEvent("cut", { clipboardData: new DataTransfer(), bubbles: true, cancelable: true, composed: true })); })()`);
    await sleep(80);
    await expectShell(command, "cut");
    rendered.push(command);
  }
  // typing undone key by key: \hat, a (the accent drawn), Backspace = \hat{} again, Backspace = gone
  const typeCommand = async (command) => {
    for (const ch of command) await key(ch, ch === "\\" ? "Backslash" : `Key${ch.toUpperCase()}`, ch === "\\" ? 220 : ch.toUpperCase().charCodeAt(0));
    await sleep(250);
    await key("Tab", "Tab", 9);
    await sleep(250);
  };
  for (const command of ["\\hat", "\\vec", "\\overline", "\\tilde"]) {
    await evaluate(`(() => { const f = ${field}; f.setValue("x", { mode: "math", format: "latex", insertionMode: "replaceAll", silenceNotifications: true });
      f.focus(); f.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true }); f.position = f._mathfield.model.lastOffset; })()`);
    await typeCommand(command);
    const opened = await state();
    assert.equal(opened.wrapper, command, `typing ${command} opens its shell (${JSON.stringify(opened)})`);
    await key("a", "KeyA", 65);
    await sleep(150);
    const typed = await state();
    // (with the accent auto-exit on, one character draws the accent; off, the shell stays and is left with Right)
    if (typed.wrapper) await key("ArrowRight", "ArrowRight", 39);
    assert.equal(compact((await state()).value), compact(`x${command}{a}`), `${command} then a: the accent`);
    await key("Backspace", "Backspace", 8);
    const back = await state();
    assert.equal(back.wrapper, command, `${command}{a}, Backspace after it: its shell again, as before the a was typed (${JSON.stringify(back)})`);
    assert.equal(compact(back.value), compact(`x${command}{}`), `${command}{a}, Backspace: ${command}{} with the caret inside`);
    await key("Backspace", "Backspace", 8);
    assert.equal(compact((await state()).value), "x", `${command}{}, Backspace: gone, as before ${command} was typed`);
  }
  // Delete in an empty shell takes the shell whole (it used to delete the closing brace: x\hat{ was left)
  await evaluate(`(() => { const f = ${field}; f.setValue("x", { mode: "math", format: "latex", insertionMode: "replaceAll", silenceNotifications: true });
    f.focus(); f.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true }); f.position = f._mathfield.model.lastOffset; })()`);
  await typeCommand("\\hat");
  await key("Delete", "Delete", 46);
  const deletedShell = await state();
  assert.equal(compact(deletedShell.value), "x", `Delete in an empty \\hat{} shell takes it whole (${JSON.stringify(deletedShell)})`);
  assert.equal(deletedShell.mode, "math", "the caret is back in math after the shell went");
  // Backspace after an accent with more content takes back its last character; Delete before one its first
  await start("\\hat", "after");
  await evaluate(`(() => { const f = ${field}; f.setValue("x\\\\hat{ab}y", { mode: "math", format: "latex", insertionMode: "replaceAll", silenceNotifications: true });
    f.focus(); f.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
    const m = f._mathfield.model; m.position = m.offsetOf(m.atoms.find((a) => a.type === "accent")); })()`);
  await key("Backspace", "Backspace", 8);
  assert.equal(compact((await state()).value), compact("x\\hat{a}y"), "Backspace after \\hat{ab} takes back the b");
  await key("Backspace", "Backspace", 8);
  assert.equal((await state()).wrapper, "\\hat", "Backspace after \\hat{a} opens its shell");
  await evaluate(`(() => { const f = ${field}; f.setValue("x\\\\vec{a}y", { mode: "math", format: "latex", insertionMode: "replaceAll", silenceNotifications: true });
    f.focus(); f.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
    const m = f._mathfield.model; m.position = m.offsetOf(m.atoms.find((a) => a.value === "x")); })()`);
  await key("Delete", "Delete", 46);
  const forward = await state();
  assert.equal(forward.wrapper, "\\vec", `Delete right before \\vec{a} opens its shell (${JSON.stringify(forward)})`);

  // leaving an emptied shell without typing removes the accent, as leaving a fresh \hat{} does
  await start("\\hat", "after");
  await key("Backspace", "Backspace", 8);
  await key("ArrowRight", "ArrowRight", 39);
  await key("ArrowRight", "ArrowRight", 39);
  assert.equal(compact((await state()).value), "xy", "an emptied shell left without typing goes, as a fresh one does");
  console.log(`accents back to their source shell when emptied (${rendered.length} commands, 4 ways): passed`);
}, { windowSize: "1100,800" });
