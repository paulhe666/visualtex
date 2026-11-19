// Documents as files, in the app's code above the library (document mode): a document opened from the library, edited
// and saved through its file and memory; a save refused when the file changed underneath; the browser's copy of
// the library; and the conversion of the earlier library's .vtxdoc documents (file versions 1 to 6) to Markdown
// files, each verified by reading it back. Run: tsx --import ../shared/visualtex-app/testing/node-loader.mjs <this file>
import { test } from "node:test";
import assert from "node:assert/strict";
import type { Node } from "prosemirror-model";
import { documentSchema as S, imageEntry, parseDocument, registerImageBytes, sameDocument } from "../../visualtex-document/src/index.ts";
import { createLocalLibrary, type DocumentLibraryApi } from "../src/document/documentLibraryApi.ts";
import { documentContent, loadDocumentFile, newDocumentContent, readsBackTheSame, renameNewPictures } from "../src/document/documentFiles.ts";
import { keptDocuments, migrateLegacyLibrary } from "../src/document/documentMigration.ts";
import { memoryStorage } from "./memory_storage.ts";

const n = S.nodes;
const utf8 = (text: string) => new TextEncoder().encode(text);
const text = (bytes: Uint8Array) => new TextDecoder().decode(bytes);
const empty = { memory: null, assets: [] };

function png(shade: number): Uint8Array {
  const base = Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="), (c) => c.charCodeAt(0));
  const bytes = new Uint8Array(base.length + 1);
  bytes.set(base);
  bytes[base.length] = shade;
  return bytes;
}

const base64 = (bytes: Uint8Array) => btoa(String.fromCharCode(...bytes));

/** The document with one block changed, as an edit in the editor leaves it. */
function withBlock(doc: Node, index: number, block: Node): Node {
  const children: Node[] = [];
  doc.forEach((child, _offset, i) => children.push(i === index ? block : child));
  return doc.type.create(doc.attrs, children);
}

// ------------------------------------------------------------------------------ the library --

test("the browser's library keeps documents as files are kept", async () => {
  const api = createLocalLibrary(memoryStorage());
  const a = await api.create(null, "Notes", "markdown", { bytes: utf8("one\n"), ...empty });
  const b = await api.create(null, "Notes", "markdown", { bytes: utf8("two\n"), ...empty });
  assert.deepEqual([a.id, b.id, a.title, a.kind, a.folder], ["Notes.md", "Notes 2.md", "Notes", "markdown", null]);
  await assert.rejects(api.create(null, " . ", "markdown", { bytes: utf8(""), ...empty }));
  const { id: folder } = await api.createFolder("课程");
  const moved = await api.move(b.id, folder);
  assert.equal(moved.id, "课程/Notes 2.md");
  const renamed = await api.rename(moved.id, "Notes");
  assert.equal(renamed.id, "课程/Notes.md");
  assert.equal(text((await api.read(renamed.id)).bytes), "two\n");
  // a save is checked against what was read
  const opened = await api.read(a.id);
  const saved = await api.save(a.id, { bytes: utf8("one more\n"), memory: "{}", assets: [{ path: "assets/p.png", bytes: png(1) }] }, opened.fingerprint);
  assert.equal(saved.status, "saved");
  const stale = await api.save(a.id, { bytes: utf8("stale\n"), ...empty }, opened.fingerprint);
  assert.equal(stale.status, "changed");
  assert.equal(text((await api.read(a.id)).bytes), "one more\n");
  assert.equal((await api.save(a.id, { bytes: utf8("forced\n"), ...empty }, opened.fingerprint, true)).status, "saved");
  const again = await api.read(a.id);
  assert.deepEqual([again.assets, [...again.pictures!.keys()]], [["p.png"], ["assets/p.png"]]);
  assert.equal((await api.check(a.id)).fingerprint, again.fingerprint);
  // trash and back, search, folders
  const trashed = (await api.trash(renamed.id)).documents.find((doc) => doc.trashed !== null)!;
  assert.ok(trashed.id.startsWith(".trash/"));
  assert.deepEqual(await api.search("two"), []);
  const restored = await api.restore(trashed.id);
  assert.ok(restored.documents.some((doc) => doc.id === "课程/Notes.md" && doc.trashed === null));
  assert.deepEqual(await api.search("two"), ["课程/Notes.md"]);
  const copy = await api.duplicate("课程/Notes.md", "Notes 副本");
  assert.equal(copy.id, "课程/Notes 副本.md");
  await api.renameFolder(folder, "Courses");
  assert.equal(text((await api.read("Courses/Notes.md")).bytes), "two\n");
  const flat = await api.deleteFolder("Courses");
  assert.ok(flat.documents.every((doc) => doc.folder === null));
  assert.equal(flat.documents.length, 3);
});

test("a document opened from its file, edited and saved: text, memory and pictures", async () => {
  const api = createLocalLibrary(memoryStorage());
  const entry = await api.create(null, "Paper", "markdown", { bytes: utf8("# Paper\r\n\r\nText with $x$.\r\n"), ...empty });
  const file = await loadDocumentFile(await api.read(entry.id));
  assert.equal(file.doc.childCount, 3);
  // nothing changed: the same bytes
  assert.equal(text(documentContent(file, file.doc).bytes), "# Paper\r\n\r\nText with $x$.\r\n");
  // a centred line and a pasted picture
  const key = registerImageBytes(png(2), "Diagram.PNG")!;
  const edited = file.doc.type.create({ ...file.doc.attrs, paper: "a5" }, [
    file.doc.child(0),
    file.doc.child(1),
    file.doc.child(2),
    n.paragraph.create({ align: "center" }, S.text("centred")),
    n.figure.create(null, [n.paragraph.create(null, n.image.create({ src: key, width: 0.3 })), n.caption.create(null, S.text("A diagram"))]),
  ]);
  const content = documentContent(file, edited);
  assert.equal(text(content.bytes), "# Paper\r\n\r\nText with $x$.\r\ncentred\r\n\r\n![](assets/Diagram.png)\r\nA diagram\r\n");
  assert.deepEqual(content.assets.map((asset) => asset.path), ["assets/Diagram.png"]);
  assert.ok(content.memory && JSON.parse(content.memory).settings.paper === "a5");
  const saved = await api.save(entry.id, content, file.fingerprint);
  assert.equal(saved.status, "saved");
  // read again: the same document; saved again without a change: the same file, the picture not sent again
  const reopened = await loadDocumentFile(await api.read(entry.id));
  assert.ok(sameDocument(reopened.doc, edited));
  assert.equal(reopened.dropped, 0);
  const unchanged = documentContent(reopened, reopened.doc);
  assert.equal(text(unchanged.bytes), text(content.bytes));
  assert.deepEqual(unchanged.assets, []);
  assert.equal(reopened.imagePaths.get(key), "assets/Diagram.png");
  // the file changed underneath by another program: the save is refused, and the document read again has the
  // other program's line without the formatting of the line it replaced
  const elsewhere = text(content.bytes).replace("centred", "centered elsewhere");
  await api.save(entry.id, { bytes: utf8(elsewhere), memory: content.memory, assets: [] }, null, true);
  const mine = withBlock(reopened.doc, 0, n.heading.create({ level: 1 }, S.text("# Paper, edited here")));
  assert.equal((await api.save(entry.id, documentContent(reopened, mine), reopened.fingerprint)).status, "changed");
  const theirs = await loadDocumentFile(await api.read(entry.id));
  assert.equal(theirs.dropped, 1);
  assert.deepEqual([theirs.doc.child(3).textContent, theirs.doc.child(3).attrs.align], ["centered elsewhere", null]);
  assert.equal(theirs.doc.child(4).type.name, "figure");
  assert.equal(theirs.doc.attrs.paper, "a5");
  // saved as a new file: every picture goes along
  const made = newDocumentContent(mine, {}, reopened.imagePaths);
  assert.deepEqual(made.content.assets.map((asset) => asset.path), ["assets/Diagram.png"]);
  assert.ok(readsBackTheSame(mine, made.content, made.imagePaths));
  assert.ok(!readsBackTheSame(mine, { ...made.content, memory: null }, made.imagePaths));
});

test("a new picture's name taken beside the document meanwhile: the picture takes another", async () => {
  const api = createLocalLibrary(memoryStorage());
  const entry = await api.create(null, "Notes", "markdown", { bytes: utf8("Text.\n"), ...empty });
  const file = await loadDocumentFile(await api.read(entry.id));
  const key = registerImageBytes(png(5), "Figure.png")!;
  const edited = file.doc.type.create(file.doc.attrs, [file.doc.child(0), n.figure.create(null, [n.paragraph.create(null, n.image.create({ src: key, width: 0.3 })), n.caption.create()])]);
  assert.deepEqual(documentContent(file, edited).assets.map((asset) => asset.path), ["assets/Figure.png"]);
  // the library refused the save: another program put an assets/Figure.png beside the document
  assert.ok(renameNewPictures(file, "assets/Figure.png"));
  const again = documentContent(file, edited);
  assert.deepEqual(again.assets.map((asset) => asset.path), ["assets/Figure-2.png"]);
  assert.equal(text(again.bytes), "Text.\n\n![](assets/Figure-2.png)\n");
  assert.equal((await api.save(entry.id, again, file.fingerprint)).status, "saved");
  // saved: the picture has its file, nothing is named again
  file.unwritten.clear();
  assert.equal(renameNewPictures(file, "assets/Figure-2.png"), false);
  assert.equal(file.imagePaths.get(key), "assets/Figure-2.png");
});

test("a .vtxdoc of an earlier version is read and saved in its own format", async () => {
  const api = createLocalLibrary(memoryStorage());
  const old = JSON.stringify({ format: "visualtex-document", version: 4, doc: { type: "doc", content: [{ type: "heading", attrs: { level: 1 }, content: [{ type: "text", text: "Old" }] }] } });
  const entry = await api.create(null, "Old", "vtxdoc", { bytes: utf8(old), ...empty });
  assert.equal(entry.id, "Old.vtxdoc");
  const file = await loadDocumentFile(await api.read(entry.id));
  assert.equal(file.kind, "vtxdoc");
  assert.equal(file.doc.child(0).textContent, "# Old");
  // (a file from before version 7 keeps its section numbers)
  assert.equal(file.doc.attrs.headingNumbering, "1");
  const content = documentContent(file, file.doc);
  assert.equal(content.memory, null);
  assert.equal(JSON.parse(text(content.bytes)).format, "visualtex-document");
  await assert.rejects(loadDocumentFile({ ...(await api.read(entry.id)), bytes: utf8('{"format":"visualtex-latex-project"}') }), /LaTeX/);
});

// ------------------------------------------------------- the earlier library's documents --

const picture = png(40);
const pictureKey = registerImageBytes(picture)!;
const paragraph = (...content: unknown[]) => ({ type: "paragraph", content });
const plain = (value: string, ...marks: unknown[]) => ({ type: "text", text: value, ...(marks.length ? { marks } : {}) });
const item = (...content: unknown[]) => ({ type: "list_item", content });
const cell = (...content: unknown[]) => ({ type: "table_cell", content });
const table = { type: "table", content: [{ type: "table_row", content: [cell(paragraph(plain("a"))), cell(paragraph(plain("b")))] }, { type: "table_row", content: [cell(paragraph(plain("1"))), cell(paragraph(plain("2", { type: "strong" })))] }] };

/** What documents of each file version held (fileFormat.ts): paragraphs with marks, then images, boards, captioned
 *  figures and tables, Markdown source in the text, and lines of Markdown. */
const body = [
  { type: "heading", attrs: { level: 1 }, content: [plain("Fourier series")] },
  paragraph(plain("Bold", { type: "strong" }), plain(" and "), plain("italic", { type: "em" }), plain(" with "), { type: "math_inline", attrs: { latex: "a_n" } }, plain(" and a "), plain("link", { type: "link", attrs: { href: "https://example.com" } }), plain(".")),
  { type: "paragraph", attrs: { align: "center" }, content: [plain("centred, red", { type: "textColor", attrs: { color: "#ff0000" } })] },
  { type: "math_display", attrs: { latex: "f(x)=\\sum_n a_n e^{inx}", numbered: true } },
  { type: "heading", attrs: { level: 2, unnumbered: true }, content: [plain("Lists")] },
  { type: "bullet_list", content: [item(paragraph(plain("one"))), item(paragraph(plain("two")), { type: "bullet_list", content: [item(paragraph(plain("nested")))] })] },
  { type: "ordered_list", attrs: { start: 3 }, content: [item(paragraph(plain("three")))] },
  { type: "blockquote", content: [paragraph(plain("quoted"))] },
  { type: "code_block", attrs: { language: "js" }, content: [plain("let a = 1;")] },
  { type: "horizontal_rule" },
];
const FIXTURES: Record<string, unknown> = {
  v1: { format: "visualtex-document", version: 1, doc: { type: "doc", attrs: { paper: "a5" }, content: [...body, table] } },
  v2: { format: "visualtex-document", version: 2, doc: { type: "doc", content: [...body, paragraph({ type: "image", attrs: { src: pictureKey, alt: "p", width: 0.4 } }), paragraph(plain("Text "), { type: "image", attrs: { src: pictureKey, width: 0.1 } })] }, images: { [pictureKey]: base64(picture) } },
  v3: {
    format: "visualtex-document",
    version: 3,
    doc: { type: "doc", content: [...body, { type: "image_board", content: [{ type: "board_image", attrs: { src: pictureKey, width: 0.3, x: 0.1, top: 0 } }, { type: "board_image", attrs: { src: pictureKey, width: 0.2, x: 0.6, top: 0.05 } }] }] },
    images: { [pictureKey]: base64(picture) },
  },
  v4: {
    format: "visualtex-document",
    version: 4,
    doc: {
      type: "doc",
      attrs: { equationNumbering: "1.1", headingNumbering: "1." },
      content: [
        { type: "title", content: [plain("A title")] },
        { type: "author", content: [plain("An author")] },
        { type: "date", content: [plain("2026")] },
        { type: "abstract", content: [paragraph(plain("The abstract."))] },
        ...body,
        { type: "figure", content: [paragraph({ type: "image", attrs: { src: pictureKey, width: 0.5 } }), { type: "caption", content: [plain("A figure")] }] },
        { type: "table_figure", content: [{ type: "caption", content: [plain("A table")] }, table] },
        paragraph(plain("A note"), { type: "footnote", attrs: { text: "The note." } }, plain(" and a citation "), { type: "cite", attrs: { targets: ["ref1"] } }, plain(".")),
        { type: "page_command", attrs: { command: "newpage", value: "" } },
        { type: "bibliography", content: [{ type: "bib_entry", attrs: { id: "ref1" }, content: [plain("Fourier, 1822")] }] },
      ],
    },
    images: { [pictureKey]: base64(picture) },
  },
  v5: {
    format: "visualtex-document",
    version: 5,
    doc: {
      type: "doc",
      content: [
        { type: "heading", attrs: { level: 1 }, content: [plain("# Source text")] },
        paragraph(plain("**Bold** and *italic* with "), { type: "math_inline", attrs: { latex: "x" } }, plain(".")),
        { type: "bullet_list", content: [item(paragraph(plain("one"))), item(paragraph(plain("two")))] },
        { type: "math_display", attrs: { latex: "a=b", numbered: false } },
        { type: "table_figure", content: [{ type: "caption" }, table] },
      ],
    },
  },
  v6: {
    format: "visualtex-document",
    version: 6,
    doc: {
      type: "doc",
      content: [
        { type: "heading", attrs: { level: 1 }, content: [plain("# Lines")] },
        paragraph(plain("- a list line with "), { type: "math_inline", attrs: { latex: "x^2" } }),
        paragraph(),
        paragraph(plain("> a quote")),
        { type: "math_display", attrs: { latex: "E=mc^2", numbered: true, id: "abcde12345" } },
        paragraph(plain("See "), { type: "eqref", attrs: { target: "abcde12345" } }, plain(".")),
        { type: "paragraph", attrs: { align: "right", indent: 1 }, content: [plain("right", { type: "fontSize", attrs: { size: "Large" } })] },
      ],
    },
  },
};

/** The earlier library as a browser kept it (the app's kept the same documents as files). */
function earlierLibrary(documents: Array<{ id: string; title: string; text: string; folder?: string; trashed?: number; modified?: number }>): Storage {
  const storage = memoryStorage();
  storage.setItem(
    "visualtex.document.library.v1",
    JSON.stringify({
      folders: [{ id: "f1", name: "课程", created: 1 }],
      documents: documents.map((doc, i) => ({ id: doc.id, title: doc.title, titleSet: true, folder: doc.folder ?? null, created: 1000 + i, modified: doc.modified ?? 2000 + i, trashed: doc.trashed ?? null })),
    }),
  );
  for (const doc of documents) storage.setItem("visualtex.document.library.doc." + doc.id, doc.text);
  return storage;
}

const names = { untitled: "未命名文档", differs: "does not convert exactly" };

test("file versions 1 to 6 convert to Markdown files that read back as the same documents", async () => {
  const storage = earlierLibrary(Object.entries(FIXTURES).map(([id, file]) => ({ id, title: `Document ${id}`, text: JSON.stringify(file), folder: id === "v4" ? "f1" : undefined })));
  const api = createLocalLibrary(storage);
  const state = await migrateLegacyLibrary(api, names);
  assert.deepEqual(keptDocuments(state), []);
  assert.equal(state.reported, true);
  assert.equal(Object.keys(state.documents).length, 6);
  const library = await api.list();
  assert.deepEqual(library.documents.map((doc) => doc.id).sort(), ["Document v1.md", "Document v2.md", "Document v3.md", "Document v5.md", "Document v6.md", "课程/Document v4.md"]);
  for (const [id, fixture] of Object.entries(FIXTURES)) {
    const original = parseDocument(S, JSON.stringify(fixture)).doc;
    const converted = library.documents.find((doc) => doc.title === `Document ${id}`)!;
    const opened = await api.read(converted.id);
    const file = await loadDocumentFile(opened);
    assert.ok(sameDocument(file.doc, original), `${id}\n${text(opened.bytes)}`);
    assert.equal(file.dropped, 0);
    // the time of the document it came from, and Markdown a person reads
    assert.equal(converted.modified, 2000 + Object.keys(FIXTURES).indexOf(id));
    assert.doesNotMatch(text(opened.bytes), /<p |<img |<span |<div /);
    // (documents written before sections were unnumbered by default keep their numbers)
    assert.ok(["1", "1."].includes(file.doc.attrs.headingNumbering), id);
  }
  // the pictures became files named by their content, shared by the documents that show them
  const v2 = await api.read("Document v2.md");
  assert.match(text(v2.bytes), /!\[p\]\(assets\/image-[0-9a-f]{12}\.png\)/);
  assert.equal(v2.pictures!.size, 1);
  assert.deepEqual([...v2.pictures!.values()][0], imageEntry(pictureKey)!.bytes);
  assert.match(text((await api.read("Document v1.md")).bytes), /^# Fourier series\n\*\*Bold\*\* and \*italic\* with \$a_n\$ and a \[link\]\(https:\/\/example\.com\)\.\n/);
  // the earlier library stays; a second start converts nothing again
  assert.ok(storage.getItem("visualtex.document.library.doc.v1"));
  await migrateLegacyLibrary(api, names);
  assert.equal((await api.list()).documents.length, 6);
});

test("a document that does not convert stays a .vtxdoc and is reported; deleted and LaTeX ones are left", async () => {
  const good = JSON.stringify(FIXTURES.v6);
  const storage = earlierLibrary([
    { id: "good", title: "", text: good },
    { id: "newer", title: "From a newer version", text: JSON.stringify({ format: "visualtex-document", version: 99, doc: { type: "doc", content: [] } }) },
    { id: "broken", title: "Broken", text: '{"format":"visualtex-document","version":6,"doc":{"type":"doc","content":[{"type":"no-such-node"}]}}' },
    { id: "latex", title: "TeX project", text: '{"format":"visualtex-latex-project","path":"x"}' },
    { id: "deleted", title: "Deleted", text: good, trashed: 5 },
  ]);
  const api = createLocalLibrary(storage);
  const state = await migrateLegacyLibrary(api, names);
  assert.equal(state.reported, false);
  const kept = keptDocuments(state);
  assert.deepEqual(kept.map((doc) => doc.title).sort(), ["Broken", "From a newer version"]);
  assert.match(kept.find((doc) => doc.title === "From a newer version")!.reason, /newer/);
  const library = await api.list();
  // (a document without a name of its own is named by its first heading)
  assert.deepEqual(library.documents.map((doc) => doc.id).sort(), ["Broken.vtxdoc", "From a newer version.vtxdoc", "Lines.md"]);
  assert.equal(library.documents.find((doc) => doc.id === "Broken.vtxdoc")!.kind, "vtxdoc");
  assert.equal(text((await api.read("Broken.vtxdoc")).bytes), storage.getItem("visualtex.document.library.doc.broken"));
});

test("a document kept as .vtxdoc in a later run is reported, though the report was shown before", async () => {
  const broken = '{"format":"visualtex-document","version":6,"doc":{"type":"doc","content":[{"type":"no-such-node"}]}}';
  const storage = earlierLibrary([
    { id: "first", title: "First", text: broken },
    { id: "later", title: "Later", text: broken },
  ]);
  const api = createLocalLibrary(storage);
  // the second one cannot be read at the first start
  const unreadable: DocumentLibraryApi = { ...api, legacy: { ...api.legacy, read: (id) => (id === "later" ? Promise.reject(new Error("busy")) : api.legacy.read(id)) } };
  const first = await migrateLegacyLibrary(unreadable, names);
  assert.deepEqual([first.finished, first.reported, keptDocuments(first).map((doc) => doc.title)], [null, false, ["First"]]);
  // the report is shown and answered
  await api.legacy.setState({ ...first, reported: true });
  const second = await migrateLegacyLibrary(api, names);
  assert.notEqual(second.finished, null);
  assert.equal(second.reported, false);
  assert.deepEqual(keptDocuments(second).map((doc) => doc.title).sort(), ["First", "Later"]);
  // answered again: nothing more to report at the next start
  await api.legacy.setState({ ...second, reported: true });
  assert.equal((await migrateLegacyLibrary(api, names)).reported, true);
});

test("a conversion cut short goes on where it stopped, without doing a document twice", async () => {
  const storage = earlierLibrary(["v1", "v5", "v6"].map((id, i) => ({ id, title: `Doc ${id}`, text: JSON.stringify(FIXTURES[id]), modified: 10 + i })));
  const api = createLocalLibrary(storage);
  // the second document's creation fails once (the app closed, the disk full)
  let creations = 0;
  const failing: DocumentLibraryApi = { ...api, create: (...args) => (++creations === 2 ? Promise.reject(new Error("disk full")) : api.create(...args)) };
  // (its copy as a .vtxdoc fails too: the third call)
  const interrupted: DocumentLibraryApi = { ...failing, create: (...args) => (creations === 2 && args[2] === "vtxdoc" ? Promise.reject(new Error("disk full")) : failing.create(...args)) };
  const first = await migrateLegacyLibrary(interrupted, names);
  assert.deepEqual(Object.keys(first.documents).sort(), ["v1", "v6"]);
  assert.equal(first.finished, null);
  // the next start does the one that is left
  const second = await migrateLegacyLibrary(api, names);
  assert.notEqual(second.finished, null);
  assert.deepEqual(Object.keys(second.documents).sort(), ["v1", "v5", "v6"]);
  assert.deepEqual((await api.list()).documents.map((doc) => doc.id).sort(), ["Doc v1.md", "Doc v5.md", "Doc v6.md"]);
});
