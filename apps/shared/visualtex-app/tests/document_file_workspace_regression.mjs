// Document mode on files, in the built app in a browser (its library is the browser's copy): the documents of the
// earlier library are converted at the first start and the one that cannot be is reported once; typing saves into the
// document's Markdown file and changes only its own line; a file changed by another program is read again when
// nothing was edited here, and when both changed the dialog's three answers do what they say (New waits for the
// answer); a document another program put into the documents folder is listed when the window comes back; a new
// document is a Markdown file without section numbers.
import assert from "node:assert/strict";
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const LINES = {
  format: "visualtex-document",
  version: 6,
  doc: {
    type: "doc",
    content: [
      { type: "heading", attrs: { level: 1 }, content: [{ type: "text", text: "# Lines" }] },
      { type: "paragraph", content: [{ type: "text", text: "First line with " }, { type: "math_inline", attrs: { latex: "x^2" } }, { type: "text", text: "." }] },
      { type: "paragraph", attrs: { align: "center" }, content: [{ type: "text", text: "centred line" }] },
      { type: "paragraph", content: [{ type: "text", text: "last line" }] },
    ],
  },
};
const BROKEN = '{"format":"visualtex-document","version":99,"doc":{"type":"doc","content":[]}}';

await withAppBrowser(async ({ baseUrl, client }) => {
  const evaluate = async (expression) => {
    const result = await client.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
    return result.result.value;
  };
  for (let tries = 0; tries < 300 && !String(await evaluate("location.href")).startsWith(baseUrl); tries += 1) await sleep(50);
  const until = async (expression, what) => {
    for (let tries = 0; tries < 160; tries += 1) {
      if (await evaluate(expression)) return;
      await sleep(50);
    }
    assert.fail(`timed out: ${what}`);
  };
  // a file of the browser's library: its text, and the same with other text (as another program writes it)
  const fileText = (id) => evaluate(`(() => { const raw = localStorage.getItem("visualtex.document.file." + ${JSON.stringify(id)}); return raw === null ? null : new TextDecoder().decode(Uint8Array.from(atob(JSON.parse(raw).base64), (c) => c.charCodeAt(0))); })()`);
  const fileMemory = (id) => evaluate(`JSON.parse(localStorage.getItem("visualtex.document.file." + ${JSON.stringify(id)})).memory`);
  const writeElsewhere = (id, text) =>
    evaluate(`(() => { const key = "visualtex.document.file." + ${JSON.stringify(id)}; const stored = JSON.parse(localStorage.getItem(key)); const bytes = new TextEncoder().encode(${JSON.stringify(text)}); stored.base64 = btoa(String.fromCharCode(...bytes)); localStorage.setItem(key, JSON.stringify(stored)); })()`);
  const lines = () => evaluate(`[...document.querySelectorAll(".vtx-document > p, .vtx-document > h1, .vtx-document > h2")].map((n) => n.textContent.replace(/\\u200b/g, ""))`);
  const names = () => evaluate(`[...document.querySelectorAll(".document-library-item:not(.is-trashed) .document-library-item-name")].map((n) => n.textContent)`);
  const currentName = () => evaluate(`document.querySelector(".document-library-item.is-current .document-library-item-name")?.textContent ?? null`);
  const dialog = () => evaluate(`(() => { const d = document.querySelector(".vtx-file-dialog"); return d ? { title: d.querySelector(".vtx-file-dialog-title").textContent, lines: [...d.querySelectorAll(".vtx-file-dialog-lines li")].map((n) => n.textContent), buttons: [...d.querySelectorAll("button")].map((n) => n.textContent) } : null; })()`);
  const press = (index) => evaluate(`document.querySelectorAll(".vtx-file-dialog button")[${index}].click()`);
  // the caret at the end of the document's last line, then text typed
  const typeAtEnd = async (text) => {
    const at = await evaluate(`(() => { const all = document.querySelectorAll(".vtx-document > p"); const r = all[all.length - 1].getBoundingClientRect(); return { x: r.left + 8, y: r.top + r.height / 2 }; })()`);
    for (const type of ["mousePressed", "mouseReleased"]) await client.send("Input.dispatchMouseEvent", { type, x: at.x, y: at.y, button: "left", clickCount: 1 });
    await sleep(200);
    const end = { key: "End", code: "End", windowsVirtualKeyCode: 35, nativeVirtualKeyCode: 35 };
    await client.send("Input.dispatchKeyEvent", { type: "rawKeyDown", ...end });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...end });
    await client.send("Input.insertText", { text });
  };
  const open = (name) => evaluate(`[...document.querySelectorAll(".document-library-item")].find((n) => n.querySelector(".document-library-item-name")?.textContent === ${JSON.stringify(name)}).click()`);

  // ---- the first start with the earlier library: one document converts, one cannot
  await evaluate(`(() => {
    localStorage.clear();
    for (const k of ["visualtex.onboarding.v3.completed", "visualtex.office.macos.first-run.v1.completed", "visualtex.onboarding.macos.desktop.v1.2.0.completed", "visualtex.office.macos.native-first-run.v1.2.0.completed"]) localStorage.setItem(k, "true");
    localStorage.setItem("visualtex.document.tutorial-shown.v1", "1");
    localStorage.setItem("visualtex.workspace-content.v1", JSON.stringify({ state: { content: "document" }, version: 0 }));
    localStorage.setItem("visualtex.document-library.v1", JSON.stringify({ state: { open: true, lastId: null, collapsed: [] }, version: 0 }));
    localStorage.setItem("visualtex.document.library.v1", JSON.stringify({ folders: [], documents: [
      { id: "lines", title: "Lines", titleSet: true, folder: null, created: 1, modified: 20, trashed: null },
      { id: "broken", title: "Broken", titleSet: true, folder: null, created: 1, modified: 10, trashed: null },
    ] }));
    localStorage.setItem("visualtex.document.library.doc.lines", ${JSON.stringify(JSON.stringify(LINES))});
    localStorage.setItem("visualtex.document.library.doc.broken", ${JSON.stringify(BROKEN)});
  })()`);
  await client.send("Page.reload", { ignoreCache: true });
  await until(`Boolean(document.querySelector(".vtx-file-dialog"))`, "the report of the document kept as .vtxdoc");
  const report = await dialog();
  assert.equal(report.buttons.length, 1);
  assert.equal(report.lines.length, 1);
  assert.match(report.lines[0], /^Broken — /);
  assert.deepEqual((await names()).sort(), ["Broken.vtxdoc", "Lines"]);
  await press(0);
  await until(`!document.querySelector(".vtx-file-dialog")`, "the report closed");
  await until(`JSON.parse(localStorage.getItem("visualtex.document.files.migration")).reported === true`, "the report remembered as shown");
  // the newest document is open, as a Markdown file with its formatting in the memory
  await until(`document.querySelector(".vtx-document h1")?.textContent.includes("Lines")`, "the converted document open");
  assert.equal(await currentName(), "Lines");
  const converted = "# Lines\nFirst line with $x^2$.\ncentred line\nlast line\n";
  assert.equal(await fileText("Lines.md"), converted);
  assert.match(await fileMemory("Lines.md"), /"align":\s*"center"/);
  assert.equal(await evaluate(`getComputedStyle(document.querySelectorAll(".vtx-document > p")[1]).textAlign`), "center");
  // (a document of an earlier version keeps its section numbers; they are the document's own setting now)
  assert.match(await fileMemory("Lines.md"), /"headingNumbering":\s*"1"/);
  assert.equal(await fileText("Broken.vtxdoc"), BROKEN);
  // the earlier library is still there
  assert.ok(await evaluate(`localStorage.getItem("visualtex.document.library.doc.lines") !== null`));

  // ---- typing is saved into the file; one line changes
  await typeAtEnd(" edited");
  await until(`${JSON.stringify(converted.replace("last line", "last line edited"))} === (() => { const raw = localStorage.getItem("visualtex.document.file.Lines.md"); return new TextDecoder().decode(Uint8Array.from(atob(JSON.parse(raw).base64), (c) => c.charCodeAt(0))); })()`, "the edit saved to the file");

  // ---- changed by another program, nothing edited here: read again when the window comes back
  await writeElsewhere("Lines.md", "# Lines\nFirst line with $x^2$.\ncentred line\nlast line edited\nadded elsewhere\n");
  await evaluate(`window.dispatchEvent(new Event("focus"))`);
  await until(`[...document.querySelectorAll(".vtx-document > p")].some((p) => p.textContent.includes("added elsewhere"))`, "the file read again");
  assert.equal(await dialog(), null);
  // (the lines that were not rewritten keep what the memory says of them)
  assert.equal(await evaluate(`getComputedStyle(document.querySelectorAll(".vtx-document > p")[1]).textAlign`), "center");

  // ---- changed by another program and here: the dialog; "save a copy" keeps both
  const theirs = "# Lines\nFirst line with $x^2$.\ncentred line\nlast line edited\nrewritten elsewhere\n";
  await writeElsewhere("Lines.md", theirs);
  await typeAtEnd(" mine");
  await until(`Boolean(document.querySelector(".vtx-file-dialog"))`, "the question about the changed file");
  const question = await dialog();
  assert.equal(question.buttons.length, 3);
  assert.equal(await fileText("Lines.md"), theirs, "nothing is written while the question is open");
  await press(2);
  await until(`!document.querySelector(".vtx-file-dialog")`, "the question answered");
  await until(`(document.querySelector(".document-library-item.is-current .document-library-item-name")?.textContent ?? "").startsWith("Lines ")`, "the copy open");
  const copy = await currentName();
  assert.notEqual(copy, "Lines");
  assert.equal(await fileText(`${copy}.md`), "# Lines\nFirst line with $x^2$.\ncentred line\nlast line edited\nadded elsewhere mine\n");
  assert.equal(await fileText("Lines.md"), theirs);

  // ---- "overwrite" writes this side; "reload" takes the other
  await open("Lines");
  await until(`[...document.querySelectorAll(".vtx-document > p")].some((p) => p.textContent.includes("rewritten elsewhere"))`, "the original open again");
  await writeElsewhere("Lines.md", "# Lines\nother text\n");
  await typeAtEnd(" kept");
  await until(`Boolean(document.querySelector(".vtx-file-dialog"))`, "the question again");
  // (New waits while the question is asked: the edits here would be dropped)
  const listed = (await names()).length;
  await evaluate(`window.dispatchEvent(new Event("visualtex-document-new-request"))`);
  await sleep(500);
  assert.equal((await names()).length, listed, "nothing made while the question is asked");
  assert.equal(await currentName(), "Lines");
  assert.ok(await dialog());
  await press(1);
  await until(`!document.querySelector(".vtx-file-dialog")`, "overwritten");
  await until(`(() => { const raw = localStorage.getItem("visualtex.document.file.Lines.md"); return new TextDecoder().decode(Uint8Array.from(atob(JSON.parse(raw).base64), (c) => c.charCodeAt(0))).endsWith("rewritten elsewhere kept\\n"); })()`, "this side written");
  await writeElsewhere("Lines.md", "# Lines\nfrom the other program\n");
  await typeAtEnd(" lost");
  await until(`Boolean(document.querySelector(".vtx-file-dialog"))`, "the question a third time");
  await press(0);
  await until(`!document.querySelector(".vtx-file-dialog")`, "reloaded");
  await until(`[...document.querySelectorAll(".vtx-document > p")].some((p) => p.textContent.includes("from the other program"))`, "the other side shown");
  assert.ok(!(await lines()).join("|").includes("lost"));
  assert.equal(await fileText("Lines.md"), "# Lines\nfrom the other program\n");

  // ---- a document another program put into the documents folder is listed when the window comes back
  await evaluate(`(() => {
    const index = JSON.parse(localStorage.getItem("visualtex.document.files.v1"));
    index.documents.push({ id: "From elsewhere.md", title: "From elsewhere", folder: null, created: 5, modified: 5, trashed: null, path: "From elsewhere.md", kind: "markdown", outside: false });
    localStorage.setItem("visualtex.document.files.v1", JSON.stringify(index));
    localStorage.setItem("visualtex.document.file.From elsewhere.md", JSON.stringify({ base64: btoa("# Elsewhere\\n"), memory: null, assets: {} }));
  })()`);
  assert.ok(!(await names()).includes("From elsewhere"));
  await evaluate(`window.dispatchEvent(new Event("focus"))`);
  await until(`[...document.querySelectorAll(".document-library-item-name")].some((n) => n.textContent === "From elsewhere")`, "the file put there listed");

  // ---- a new document: a Markdown file, its sections not numbered
  const before = await names();
  await evaluate(`window.dispatchEvent(new Event("visualtex-document-new-request"))`);
  await until(`document.querySelectorAll(".document-library-item:not(.is-trashed)").length === ${before.length + 1}`, "the new document listed");
  const created = (await names()).find((name) => !before.includes(name));
  assert.equal(await currentName(), created);
  assert.equal(await fileText(`${created}.md`), "");
  await until(`Boolean(document.querySelector(".vtx-sheet.is-headings-none"))`, "sections unnumbered");
  await typeAtEnd("# Heading");
  await until(`(() => { const raw = localStorage.getItem("visualtex.document.file." + ${JSON.stringify(`${created}.md`)}); return new TextDecoder().decode(Uint8Array.from(atob(JSON.parse(raw).base64), (c) => c.charCodeAt(0))) === "# Heading\\n"; })()`, "the new document's text saved");
  assert.equal(await fileMemory(`${created}.md`), null);
}, { windowSize: "1400,900" });

console.log("document files in the workspace: passed");
