// A document as a Markdown file and its memory (shared/visualtex-document markdownFile.ts): every node, mark and
// setting comes back from the two; a file made elsewhere is the same file when saved unchanged, and one edit changes
// its own line; the formulas the text cannot say plainly read back; a line another program rewrote loses its own
// formatting only. Run: tsx --import ../shared/visualtex-app/testing/node-loader.mjs <this file>
import { test } from "node:test";
import assert from "node:assert/strict";
import { EditorState } from "prosemirror-state";
import type { Node } from "prosemirror-model";
import {
  assignIds,
  boardFixes,
  decodeFileText,
  documentSchema as S,
  encodeFileText,
  figureFixes,
  imageEntry,
  newPicturePaths,
  parseDocumentMemory,
  readDocumentFile,
  registerImageBytes,
  sameDocument,
  serializeDocumentMemory,
  withImageFiles,
  writeDocumentFile,
  type DocumentMemory,
} from "../../visualtex-document/src/index.ts";

const n = S.nodes;
const t = (text: string, ...marks: Array<[string, Record<string, unknown>?]>) => S.text(text, marks.map(([name, attrs]) => S.marks[name].create(attrs)));
const inline = (content: Array<Node | string>) => content.filter((c) => c !== "").map((c) => (typeof c === "string" ? S.text(c) : c));
const p = (...content: Array<Node | string>) => n.paragraph.create(null, inline(content));
const line = (attrs: Record<string, unknown> | null, ...content: Array<Node | string>) => n.paragraph.create(attrs, inline(content));
const mi = (latex: string) => n.math_inline.create({ latex });
const md = (latex: string, attrs: Record<string, unknown> = {}) => n.math_display.create({ latex, ...attrs });
const doc = (blocks: Node[], attrs: Record<string, unknown> | null = null) => n.doc.create(attrs, blocks);

/** A PNG of one pixel (its colour makes the bytes, so each is a picture of its own). */
function png(shade: number): Uint8Array {
  const base = Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="), (c) => c.charCodeAt(0));
  const bytes = new Uint8Array(base.length + 1);
  bytes.set(base);
  bytes[base.length] = shade;
  return bytes;
}

/** The document as the editor holds it after opening: boards, figures and ids as its plugins leave them. */
function asOpened(document: Node): Node {
  let state = EditorState.create({ doc: document });
  for (const fix of [boardFixes, figureFixes]) {
    const tr = fix(state);
    if (tr) state = state.apply(tr);
  }
  const tr = assignIds(state.doc, null, () => state.tr);
  return tr ? state.apply(tr).doc : state.doc;
}

/** Ids made by the editor are given anew when a document is read: left out of a comparison unless referred to; a
 *  footnote without a label takes a number in the file. */
function comparable(document: Node): string {
  const referenced = new Set<string>();
  document.descendants((node) => {
    if (node.type.name === "eqref" && node.attrs.target) referenced.add(node.attrs.target);
    if (node.type.name === "cite") for (const id of node.attrs.targets as string[]) referenced.add(id);
    return true;
  });
  return JSON.stringify(document.toJSON(), (key, value) =>
    (key === "id" && typeof value === "string" && /^[0-9a-z]{10}$/.test(value) && !referenced.has(value)) || (key === "label" && /^\d+$/.test(String(value))) ? null : value,
  );
}

/** Written and read back (the memory through its JSON text, as it is stored); the pictures found by their paths. */
function roundTrip(document: Node, options: Parameters<typeof writeDocumentFile>[1] = {}) {
  const written = writeDocumentFile(document, options);
  const memory = parseDocumentMemory(serializeDocumentMemory(written.memory));
  const byPath = new Map([...written.images].map(([key, path]) => [path, key]));
  const read = withImageFiles((path) => byPath.get(path) ?? null, () => readDocumentFile(S, written.text, memory));
  return { written, read, memory: memory as DocumentMemory };
}

function assertRoundTrip(document: Node, message: string) {
  const { written, read } = roundTrip(document);
  assert.equal(comparable(read.doc), comparable(document), `${message}\n--- text ---\n${written.text}`);
  assert.equal(read.dropped, 0, message);
  return written;
}

// ------------------------------------------------------------------ the file's text --

test("a plain document is plain Markdown and needs no memory", () => {
  const document = doc([
    n.heading.create({ level: 1 }, t("# Heading")),
    p("Text with ", mi("x^2"), " and **bold** `code`."),
    md("E = mc^2"),
    p("- item"),
    p("\t- nested"),
    p(""),
    p("> quote"),
    p("```js"),
    p("let $a$ = 1;"),
    p("```"),
  ]);
  const { written, read } = roundTrip(document);
  assert.equal(written.text, "# Heading\nText with $x^2$ and **bold** `code`.\n\n$$\nE = mc^2\n$$\n\n- item\n\t- nested\n\n> quote\n```js\nlet $a$ = 1;\n```\n");
  assert.deepEqual(written.memory.blocks, []);
  assert.deepEqual(written.memory.settings, {});
  assert.equal(read.doc.child(0).type.name, "heading");
  assert.equal(read.doc.child(8).textContent, "let $a$ = 1;");
});

test("an object stands with one blank line on each side, which is not a line of the document", () => {
  const cases: Array<[Node[], string]> = [
    [[p("a"), md("x"), p("b")], "a\n\n$$\nx\n$$\n\nb\n"],
    [[p("a"), p(""), md("x"), p(""), p("b")], "a\n\n\n$$\nx\n$$\n\n\nb\n"],
    [[md("x"), md("y")], "$$\nx\n$$\n\n$$\ny\n$$\n"],
    [[md("x"), p(""), md("y")], "$$\nx\n$$\n\n\n\n$$\ny\n$$\n"],
    [[p(""), md("x"), p("")], "\n\n$$\nx\n$$\n\n\n"],
    [[md("x")], "$$\nx\n$$\n"],
  ];
  for (const [blocks, text] of cases) {
    const written = assertRoundTrip(doc(blocks), text);
    assert.equal(written.text, text);
    assert.deepEqual(written.memory.blocks, []);
  }
});

test("nothing only VisualTeX reads is written into the file", () => {
  const key = registerImageBytes(png(1))!;
  const document = asOpened(
    doc(
      [
        n.title.create(null, t("My Title")),
        n.author.create(null, t("An Author")),
        n.date.create(null, t("2026")),
        n.abstract.create(null, [p("Abstract text.")]),
        line({ align: "center" }, t("red", ["textColor", { color: "#ff0000" }]), t(" big", ["fontSize", { size: "Large" }]), t(" marked", ["highlight", { color: "#00ff00" }])),
        n.page_command.create({ command: "newpage" }),
        n.page_command.create({ command: "pagenumbering", value: "roman" }),
        p(n.image.create({ src: key, alt: "pic", width: 0.31 })),
        md("a=b", { numbered: true }),
      ],
      { paper: "a5", headingNumbering: "1.", lineSpacing: 1.5 },
    ),
  );
  const { written, read } = roundTrip(document);
  assert.doesNotMatch(written.text, /<[a-zA-Z!\/]/, written.text);
  assert.doesNotMatch(written.text, /^---$/m);
  assert.match(written.text, /^# My Title\nAn Author\n2026\n/);
  // page commands are the memory's (decision 2): no line in the file, and they come back
  assert.doesNotMatch(written.text, /\\newpage|\\pagenumbering/);
  const commands: string[] = [];
  read.doc.forEach((node) => {
    if (node.type.name === "page_command") commands.push(`${node.attrs.command}:${node.attrs.value}`);
  });
  assert.deepEqual(commands, ["newpage:", "pagenumbering:roman"]);
  assert.match(written.text, /!\[pic\]\(assets\/image-[0-9a-f]{12}\.png\)/);
  assert.match(written.text, /a=b \\tag\{1\}/);
  assert.deepEqual(written.memory.settings, { paper: "a5", lineSpacing: 1.5, headingNumbering: "1." });
});

// ------------------------------------------------------- every node, mark and setting --

test("every node, mark and setting comes back from the text and the memory", () => {
  const a = registerImageBytes(png(2))!;
  const b = registerImageBytes(png(3))!;
  const c = registerImageBytes(png(4))!;
  const cell = (content: Node[], attrs: Record<string, unknown> = {}) => n.table_cell.create(attrs, content);
  const table = (rows: Node[][]) => n.table.create(null, rows.map((cells) => n.table_row.create(null, cells)));
  const caption = (...content: Array<Node | string>) => n.caption.create(null, content.map((x) => (typeof x === "string" ? S.text(x) : x)));
  const blocks: Node[] = [
    n.title.create(null, [t("A "), mi("\\LaTeX"), t(" title")]),
    n.author.create(null, [t("First Author"), n.hard_break.create(), t("Institute")]),
    n.author.create(null, t("Second Author")),
    n.date.create(null, t("2026-11-18")),
    n.abstract.create(null, [p("First paragraph of the abstract."), p("Second, with ", mi("a+b"), "."), md("x=1")]),
    n.toc.create(),
    n.heading.create({ level: 1 }, t("# One")),
    n.heading.create({ level: 2, unnumbered: true }, t("## Two {-}")),
    n.heading.create({ level: 3, align: "center" }, t("### Three")),
    line({ align: "center" }, "centred"),
    line({ align: "right", indent: 2 }, "right, indented"),
    line({ indent: 1 }, "indented"),
    p(t("red", ["textColor", { color: "#c00000" }]), t(" and "), t("large", ["fontSize", { size: "Large" }]), t(" and "), t("green ground", ["highlight", { color: "#b7e1cd" }]), t(" and "), t("both", ["textColor", { color: "#0000ff" }], ["fontSize", { size: "small" }])),
    p(t("bold", ["strong"]), t(" "), t("em", ["em"]), t(" "), t("u", ["underline"]), t(" "), t("s", ["strike"]), t(" "), t("code", ["code"]), t(" "), t("k", ["kbd"]), t(" x"), t("2", ["superscript"]), t(" y"), t("i", ["subscript"]), t(" "), t("link", ["link", { href: "https://example.com", title: "T" }]), t(" "), t("mark", ["highlight", { color: "#fff59d" }])),
    p("Inline ", mi("x^2"), ", a note", n.footnote.create({ text: "The note with $a^2$." }), ", a named note", n.footnote.create({ text: "Named.", label: "why" }), " and a cite ", n.cite.create({ targets: ["refa", "refb"] }), "."),
    md("E = mc^2", { id: "eqenergy01", numbered: true }),
    p("See ", n.eqref.create({ target: "eqenergy01" }), " and ", n.eqref.create({ target: "named" }), " and ", n.eqref.create({ target: "gone" }), "."),
    md("a^2+b^2=c^2 \\tag{P}"),
    md("\\int f", { id: "named", numbered: true }),
    md("\\begin{aligned}\n  a &= b \\\\\n  c &= d\n\\end{aligned}"),
    md(""),
    n.table_figure.create(null, [caption("Above ", mi("t")), table([[cell([p("h1")]), cell([p("h2")], { align: "right" })], [cell([p("a | b")]), cell([p("**b**")])]])]),
    n.table_figure.create({ rules: "booktabs" }, [table([[cell([p("x")], { colspan: 2 })], [cell([p("1")], { colwidth: [120] }), cell([md("\\frac12"), p("two"), p("paras")])]]), caption("Below")]),
    n.table_figure.create(null, [table([[cell([p("only")])]])]),
    n.table_figure.create(null, [caption(), table([[cell([p(t("coloured", ["textColor", { color: "#123456" }]))])]])]),
    n.figure.create(null, [line({ align: "center" }, n.image.create({ src: a, alt: "one", width: 0.4 })), caption("A caption with ", mi("y"))]),
    n.figure.create(null, [p(n.image.create({ src: b, alt: "", width: 0.3, wrap: "right" })), caption()]),
    p("Text the picture wraps, long enough to be a paragraph."),
    n.figure.create(null, [n.image_board.create(null, [n.board_image.create({ src: a, width: 0.3, x: 0.05, top: 0 }), n.board_image.create({ src: c, alt: "c", width: 0.2, x: 0.5, top: 0.1 })]), caption("Board")]),
    p("In the text ", n.image.create({ src: c, alt: "inline", width: 0.1 }), " a picture."),
    n.page_command.create({ command: "newpage" }),
    n.page_command.create({ command: "setcounter", value: "5" }),
    n.page_command.create({ command: "thispagestyle", value: "empty" }),
    n.html_block.create({ html: "<details>\n<summary>More</summary>\nBody\n</details>" }),
    p("kept ", n.html_inline.create({ html: '<img src="missing.png">' }), " html"),
    n.bullet_list.create(null, [n.list_item.create(null, [p("old list item")]), n.list_item.create({ checked: true }, [p("task")])]),
    n.ordered_list.create({ start: 3 }, [n.list_item.create(null, [p("three")])]),
    n.blockquote.create({ alert: "note" }, [p("old quote")]),
    n.code_block.create({ language: "js" }, t("let a = 1;\n\nlet b = 2;")),
    n.horizontal_rule.create(),
    p("```"),
    p("| not | a table |"),
    p("```"),
    p(""),
    p(""),
    n.bibliography.create(null, [n.bib_entry.create({ id: "refa", key: "knuth84" }, [t("Knuth, "), mi("\\TeX")]), n.bib_entry.create({ id: "refb" }, t("Lamport"))]),
  ];
  const document = asOpened(
    doc(blocks, { paper: "b5", marginTop: 20, marginBottom: 21, marginLeft: 22, marginRight: 23, fontSize: 12, cjkFont: "KaiTi", latinFont: "Georgia", lineSpacing: 1.25, equationNumbering: "1-1", equationBrackets: "[]", headingNumbering: "1" }),
  );
  const written = assertRoundTrip(document, "the whole document");
  // written again from what was read: the same file and the same memory
  const { read } = roundTrip(document);
  const again = writeDocumentFile(read.doc, { source: read.source });
  assert.equal(again.text, written.text);
  assert.deepEqual(again.memory, written.memory);
  // each block alone, and between plain lines
  document.forEach((block, _offset, index) => {
    if (["bibliography"].includes(block.type.name)) return;
    const alone = doc([block], document.attrs);
    const between = doc([p("before"), block, p("after")]);
    for (const candidate of [alone, between]) {
      // (a reference needs its target in the document)
      let dangling = false;
      candidate.descendants((node) => {
        if (node.type.name === "eqref" || node.type.name === "cite") dangling = true;
        return true;
      });
      if (!dangling) assertRoundTrip(candidate, `block ${index} (${block.type.name})`);
    }
  });
});

test("ids of any form are given anew; only a reference or a formula's name keeps one", () => {
  const document = doc([
    n.paragraph.create({ id: "p-1" }, t("an older document's id")),
    n.heading.create({ id: "sec:intro", level: 1 }, t("# Intro")),
    md("a=b", { id: "block-7" }),
    md("c=d", { id: "eq:named" }),
    md("e=f", { id: "target-9", numbered: true }),
    p("See ", n.eqref.create({ target: "target-9" }), "."),
  ]);
  const { written, read } = roundTrip(document);
  // (a formula's id that is not one made here is its name: it goes into the text)
  assert.equal(written.text, ["an older document's id", "# Intro", "", "$$", "a=b \\label{block-7}", "$$", "", "$$", "c=d \\label{eq:named}", "$$", "", "$$", "e=f \\tag{1} \\label{target-9}", "$$", "", "See $\\eqref{target-9}$.", ""].join("\n"));
  assert.ok(sameDocument(read.doc, document));
  assert.equal(read.doc.child(0).attrs.id, null);
  assert.equal(read.doc.child(1).attrs.id, null);
  assert.deepEqual([read.doc.child(2).attrs.id, read.doc.child(3).attrs.id, read.doc.child(4).attrs.id], ["block-7", "eq:named", "target-9"]);
  // the lines whose ids are simply dropped need no memory; the formulas do (a name alone reads as a numbered
  // formula, a tag as the formula's own)
  assert.equal(written.memory.blocks.length, 3);
});

test("each setting alone", () => {
  const settings: Record<string, unknown> = { paper: "letter", marginTop: 10, marginBottom: 11, marginLeft: 12, marginRight: 13, fontSize: 9, cjkFont: "SimHei", latinFont: "Arial", lineSpacing: 2, equationNumbering: "1.1", equationBrackets: "[]", headingNumbering: "none" };
  for (const [name, value] of Object.entries(settings)) {
    const { written, read } = roundTrip(doc([p("text")], { [name]: value }));
    assert.equal(written.text, "text\n");
    assert.deepEqual(written.memory.settings, { [name]: value });
    assert.equal(read.doc.attrs[name], value, name);
  }
});

test("a page command made here has no line in the file; one a file has stays its text", () => {
  const pc = (command: string, value = "") => n.page_command.create({ command, value });
  const notes = p("Noted", n.footnote.create({ text: "The note." }));
  const cases: Array<[Node[], string]> = [
    [[p("a"), pc("newpage"), p("b")], "a\nb\n"],
    [[pc("pagenumbering", "roman"), p("first")], "first\n"],
    [[p("last"), pc("clearpage")], "last\n"],
    [[p("a"), pc("newpage"), md("x"), pc("thispagestyle", "empty"), pc("setcounter", "3"), p("b")], "a\n\n$$\nx\n$$\n\nb\n"],
    [[md("x"), pc("newpage"), md("y")], "$$\nx\n$$\n\n$$\ny\n$$\n"],
    [[notes, pc("newpage")], "Noted[^1]\n\n[^1]: The note.\n"],
    [[pc("newpage")], "\n"],
  ];
  for (const [blocks, text] of cases) {
    const written = assertRoundTrip(doc(blocks), text);
    assert.equal(written.text, text);
  }
  // the line a page command goes with rewritten elsewhere: that page command is dropped, the others stay
  const document = doc([p("one"), pc("newpage"), p("two"), pc("clearpage"), p("three")]);
  const { written, memory } = roundTrip(document);
  const lines = (document: Node) => {
    const all: string[] = [];
    document.forEach((node) => all.push(node.type.name === "page_command" ? `\\${node.attrs.command}` : node.textContent));
    return all;
  };
  assert.deepEqual(lines(readDocumentFile(S, written.text.replace("two", "2"), memory).doc), ["one", "2", "\\clearpage", "three"]);
  // a file's own page command lines are its text, kept as they are; one made here beside them has none
  const file = "One\n\n\\newpage\n\nTwo \\pagenumbering{roman} in a line.\n";
  const own = readDocumentFile(S, file);
  assert.deepEqual(lines(own.doc), ["One", "", "\\newpage", "", "Two \\pagenumbering{roman} in a line."]);
  assert.equal(writeDocumentFile(own.doc, { source: own.source }).text, file);
  const added: Node[] = [];
  own.doc.forEach((node, _offset, index) => added.push(...(index === 0 ? [node, pc("clearpage")] : [node])));
  const again = writeDocumentFile(doc(added), { source: own.source });
  assert.equal(again.text, file);
  assert.deepEqual(lines(readDocumentFile(S, again.text, again.memory).doc), ["One", "\\clearpage", "", "\\newpage", "", "Two \\pagenumbering{roman} in a line."]);
});

// -------------------------------------------------------------------- formulas --

test("formulas the dollars cannot say plainly read back", () => {
  const cases: Array<{ name: string; block: Node; text: string; plain: boolean }> = [
    { name: "a digit after", block: p(mi("x"), "2 apples"), text: "$x$\u20602 apples\n", plain: true },
    { name: "two together", block: p(mi("a"), mi("b"), mi("c")), text: "$a$\u2060$b$\u2060$c$\n", plain: true },
    { name: "a dollar before", block: p("cost $", mi("x"), " b"), text: "cost $\u2060$x$ b\n", plain: true },
    { name: "a dollar after", block: p(mi("x"), "$ each"), text: "$x$\u2060$ each\n", plain: true },
    { name: "an escaped dollar inside", block: p("a ", mi("\\$5"), " b"), text: "a $\\$5$ b\n", plain: true },
    { name: "spaces inside", block: p("a ", mi(" x + y "), " b"), text: "a $x + y$ b\n", plain: false },
    { name: "a closing backslash", block: p("a ", mi("x\\ "), "b"), text: "a $x\\ {}$b\n", plain: false },
    { name: "several lines", block: p("a ", mi("\\begin{cases}a\\\\\nb\\end{cases}"), " b"), text: "a $\\begin{cases}a\\\\ b\\end{cases}$ b\n", plain: false },
    // an empty formula keeps its dollars, a joiner before them (it was left out of the text, 2026-11-24)
    { name: "empty", block: p("a ", mi(""), " b"), text: "a \u2060$$ b\n", plain: true },
    { name: "empty, alone in its line", block: p(mi("")), text: "\u2060$$\n", plain: true },
    { name: "empty, first in its line", block: p("  ", mi(""), " b"), text: "  \u2060$$ b\n", plain: true },
    { name: "empty, between formulas", block: p(mi("a"), mi(""), mi(""), mi("b")), text: "$a$\u2060$$\u2060$$\u2060$b$\n", plain: true },
    { name: "empty, between dollars and a digit", block: p("cost $", mi(""), "$5 ", mi(""), "2"), text: "cost $\u2060$$\u2060$5 \u2060$$\u20602\n", plain: true },
    { name: "$$ in a line's text", block: p("a $$ b $$x^2$$ c"), text: "a $$ b $$x^2$$ c\n", plain: true },
    { name: "in a heading", block: n.heading.create({ level: 2 }, [t("## H "), mi("x^2")]), text: "## H $x^2$\n", plain: true },
    { name: "after an open code span", block: p("`code then ", mi("x")), text: "`code then $x$\n", plain: true },
    { name: "between code spans", block: p("`a` then ", mi("x"), " and `b`"), text: "`a` then $x$ and `b`\n", plain: true },
    { name: "inside what reads as a code span", block: p("`a ", mi("x"), " b`"), text: "`a $x$ b`\n", plain: false },
    { name: "prices", block: p("$5 and $10, \\$3"), text: "$5 and $10, \\$3\n", plain: true },
    { name: "a display formula with a blank line", block: md("a\n\nb"), text: "$$\na\n\nb\n$$\n", plain: true },
    { name: "an empty display formula", block: md(""), text: "$$\n$$\n", plain: true },
    { name: "a display formula's own tag", block: md("a=b \\tag{5}"), text: "$$\na=b \\tag{5}\n$$\n", plain: true },
    { name: "a numbered display formula", block: md("a=b", { numbered: true }), text: "$$\na=b \\tag{1}\n$$\n", plain: false },
    { name: "a literal $$ line", block: p("$$ not math"), text: "$$ not math\n", plain: true },
  ];
  for (const { name, block, text, plain } of cases) {
    const document = doc([block]);
    const { written, read } = roundTrip(document);
    assert.equal(written.text, text, name);
    assert.equal(comparable(read.doc), comparable(document), name);
    // the text alone says it (no memory) where it can
    assert.equal(written.memory.blocks.length === 0, plain, `${name}: memory ${JSON.stringify(written.memory.blocks)}`);
  }
});

test("\\( \\) and \\[ \\] are formulas; a file's \\tag stays the formula's own", () => {
  const text = "Inline \\(a+b\\) and $c$.\n\n\\[\nx = 1\n\\]\n\n\\[ y = 2 \\]\n\n$$z = 3 \\tag{7}$$\n\n\\[1\\] an escaped bracket\n";
  const read = readDocumentFile(S, text);
  const kinds: string[] = [];
  read.doc.forEach((node) => kinds.push(node.type.name));
  assert.deepEqual(kinds, ["paragraph", "math_display", "math_display", "math_display", "paragraph"]);
  assert.equal(read.doc.child(0).child(1).attrs.latex, "a+b");
  assert.equal(read.doc.child(1).attrs.latex, "x = 1");
  assert.equal(read.doc.child(2).attrs.latex, "y = 2");
  assert.deepEqual([read.doc.child(3).attrs.latex, read.doc.child(3).attrs.numbered], ["z = 3 \\tag{7}", false]);
  assert.equal(writeDocumentFile(read.doc, { source: read.source }).text, text);
});

test("a named formula and its references", () => {
  const text = "$$\nE = mc^2 \\label{eq:energy}\n$$\n\nBy $\\eqref{eq:energy}$ and $\\ref{eq:energy}$, not $\\eqref{other}$.\n";
  const read = readDocumentFile(S, text);
  assert.deepEqual([read.doc.child(0).attrs.id, read.doc.child(0).attrs.latex, read.doc.child(0).attrs.numbered], ["eq:energy", "E = mc^2", true]);
  const kinds: string[] = [];
  read.doc.child(1).forEach((node) => kinds.push(node.type.name));
  assert.deepEqual(kinds, ["text", "eqref", "text", "eqref", "text", "math_inline", "text"]);
  assert.equal(writeDocumentFile(read.doc, { source: read.source }).text, text);
  // written anew, the name stays with the formula and the references
  const fresh = writeDocumentFile(read.doc);
  assert.equal(fresh.text, "$$\nE = mc^2 \\tag{1} \\label{eq:energy}\n$$\n\nBy $\\eqref{eq:energy}$ and $\\eqref{eq:energy}$, not $\\eqref{other}$.\n");
  assert.equal(comparable(readDocumentFile(S, fresh.text, fresh.memory).doc), comparable(read.doc));
});

test("a name a file gave a formula stays in the file, even one shaped like the editor's ids", () => {
  // ("definition": ten lowercase letters, as the ids the editor gives)
  const text = "$$\nx = 1 \\label{definition}\n$$\n\nBy $\\eqref{definition}$.\n";
  const read = readDocumentFile(S, text);
  assert.deepEqual([read.doc.child(0).attrs.id, read.doc.child(0).attrs.numbered, read.doc.child(1).child(1).type.name], ["definition", true, "eqref"]);
  assert.equal(writeDocumentFile(asOpened(read.doc), { source: read.source }).text, text);
  // the formula and the reference's line edited: the name stays with both
  const edited = doc([md("x = 2", { id: "definition", numbered: true }), p("Now by ", n.eqref.create({ target: "definition" }), ".")]);
  const written = writeDocumentFile(edited, { source: read.source });
  assert.equal(written.text, "$$\nx = 2 \\tag{1} \\label{definition}\n$$\n\nNow by $\\eqref{definition}$.\n");
  const back = readDocumentFile(S, written.text, parseDocumentMemory(serializeDocumentMemory(written.memory)));
  assert.ok(sameDocument(back.doc, edited));
  assert.equal(back.doc.child(0).attrs.id, "definition");
  // written anew from it (Save as), the name goes along
  assert.match(writeDocumentFile(edited, { labels: read.source.labels }).text, /^\$\$\nx = 2 \\tag\{1\} \\label\{definition\}\n\$\$\n\nNow by \$\\eqref\{definition\}\$\.\n$/);
});

// ------------------------------------------------------------------- footnotes --

test("footnotes are Markdown footnotes", () => {
  const document = doc([p("One", n.footnote.create({ text: "First." }), " two", n.footnote.create({ text: "Second\nwith a second line.", label: "b" }), "."), p("Three", n.footnote.create({ text: "" }))]);
  const { written, read } = roundTrip(document);
  assert.equal(written.text, "One[^1] two[^b].\nThree[^2]\n\n[^1]: First.\n[^b]: Second\n    with a second line.\n[^2]:\n");
  assert.deepEqual(written.memory.blocks, []);
  assert.equal(read.doc.child(0).child(1).attrs.text, "First.");
  assert.equal(read.doc.child(0).child(3).attrs.label, "b");
});

test("a file's footnotes: read when they end the file, left as text otherwise", () => {
  const sound = "Text[^note] and more[^2].\n\n[^note]: The note.\n\n[^2]: Another,\n    on two lines.\n";
  const read = readDocumentFile(S, sound);
  assert.equal(read.doc.childCount, 1);
  assert.deepEqual([read.doc.child(0).child(1).type.name, read.doc.child(0).child(1).attrs.text, read.doc.child(0).child(1).attrs.label], ["footnote", "The note.", "note"]);
  assert.equal(read.doc.child(0).child(3).attrs.text, "Another,\non two lines.");
  assert.equal(writeDocumentFile(read.doc, { source: read.source }).text, sound);
  for (const text of [
    "Text[^a].\n\n[^a]: Defined in the middle.\n\nMore text.\n",
    "Twice[^a] and again[^a].\n\n[^a]: One note.\n",
    "No reference.\n\n[^a]: Unused.\n",
    "In code `[^a]`.\n\n[^a]: Note.\n",
  ]) {
    const plain = readDocumentFile(S, text);
    plain.doc.descendants((node) => {
      assert.notEqual(node.type.name, "footnote", text);
      return true;
    });
    assert.equal(writeDocumentFile(plain.doc, { source: plain.source }).text, text);
  }
});

// ------------------------------------------------------------ files made elsewhere --

const CORPUS: Record<string, string> = {
  plain: "# Title\n\nA paragraph of text.\nA second line.\n\n## Section\n\n- one\n- two\n  - nested\n\n1. first\n2. second\n\n> a quote\n> more\n\n---\n\nThe end.\n",
  frontMatter: "---\ntitle: Notes\ntags: [a, b]\nprice: $5$\n---\n\n# Notes\n\nBody text.\n",
  dollars: "It costs $5 and $10.\nA formula $a^2+b^2$ here, and $x$, and \\(y\\).\nEscaped \\$5 and `$code$`.\n",
  displayLoose: "Before.\n\n$$\n\\int_0^1 x\\,dx\n$$\n\nAfter.\n",
  displayTight: "Before.\n$$\n\\int_0^1 x\\,dx\n$$\nAfter.\n",
  displayOneLine: "Before.\n\n$$a^2 + b^2 = c^2$$\n\nAfter.\n\n$$ spaced $$\n",
  displaySpread: "Before.\n\n\n$$\nx\n$$\n\n\n\nAfter.\n\n$$\ny\n$$\n\n\n$$\nz\n$$\n",
  displayInList: "- item\n\n  $$\n  a = b\n  $$\n\n- next\n",
  bracketMath: "Text.\n\n\\[\nx = 1\n\\]\n\n\\[ y = 2 \\]\n\nInline \\(z\\).\n",
  tagged: "$$\nE = mc^2 \\tag{5}\n$$\n\n$$\nF = ma \\label{eq:f}\n$$\n\nSee $\\eqref{eq:f}$.\n",
  table: "Text.\n\n| a | b |\n|---|:-:|\n| 1 | 2 |\n| $x$ | **y** |\n\nAfter.\n",
  tableUnspaced: "|a|b|\n|-|-|\n|1|2|\n",
  tableNoPipes: "a | b\n--- | ---\n1 | 2\n3 | 4\n\nAfter.\n",
  tableCaptions: "Table: The first\n\n| a |\n|---|\n| 1 |\n\n| b |\n|---|\n| 2 |\n\nTable: Below the second\n\n| c |\n|---|\n| 3 |\n<!-- no caption -->\n",
  tableTight: "Text\n| a | b |\n|---|---|\n| 1 | 2 |\nMore text\n",
  code: "```python\nprice = \"$5 and $10\"\n$$\nnot math\n$$\n| not | a table |\n|---|---|\n```\n\n~~~\n[^1]: not a note\n~~~\n",
  html: "<div align=\"center\">\n<img src=\"logo.png\" width=\"100\">\n</div>\n\n<details>\n<summary>More</summary>\n\nHidden $x$.\n\n</details>\n\nText with <span style=\"color:red\">colour</span> and <br> a break.\n\n<!-- a comment -->\n",
  links: "A [link][ref] and ![picture](missing.png) and <https://example.com>.\n\n[ref]: https://example.com \"Title\"\n",
  pictures: "![alone](not-found.png)\n\nText ![inline](not-found.png) text.\n\n![a](x.png) ![b](y.png)\n",
  footnotes: "A note[^1] and another[^long].\n\n[^1]: Short.\n[^long]: Long\n    on two lines.\n",
  footnotesMiddle: "A note[^1].\n\n[^1]: In the middle.\n\nMore.\n",
  setext: "Title\n=====\n\nSection\n-------\n\nText.\n",
  whitespace: "  indented text  \n\ttab\n\n\n\ntrailing spaces   \n   \n",
  chinese: "# 标题\n\n正文，公式 $a^2$，价格 $5。\n\n$$\n\\sum_{i=1}^n i\n$$\n\n表：说明\n\n| 列 |\n|---|\n| 值 |\n",
  commands: "[toc]\n\n# One\n\n\\newpage\n\n\\pagenumbering{roman}\n\nText \\newpage in a line.\n",
  noFinalNewline: "# Title\n\nLast line without a newline",
  blankEdges: "\n\n# Title\n\n$$\nx\n$$\n\n\n",
  onlyFormula: "$$\nx\n$$",
  empty: "",
  oneNewline: "\n",
};

test("a file made elsewhere, opened and saved without edits, is the same file", () => {
  for (const [name, source] of Object.entries(CORPUS)) {
    for (const [variant, text] of [["lf", source], ["crlf", source.replace(/\n/g, "\r\n")]] as const) {
      const read = readDocumentFile(S, text);
      // unchanged: the file itself
      assert.equal(writeDocumentFile(read.doc, { source: read.source }).text, text, `${name} ${variant}`);
      // as the editor holds it (ids, boards, figures): still unchanged
      const opened = asOpened(read.doc);
      assert.equal(comparable(opened), comparable(read.doc), `${name} ${variant}: the editor changes what was read`);
      const written = writeDocumentFile(opened, { source: read.source });
      assert.equal(written.text, text, `${name} ${variant} (opened)`);
      // and nothing of it needs the memory
      assert.deepEqual(written.memory.blocks, [], `${name} ${variant}`);
    }
  }
});

test("one edit changes only its own line", () => {
  let edited = 0;
  for (const [name, source] of Object.entries(CORPUS)) {
    for (const [variant, text, eol] of [["lf", source, "\n"], ["crlf", source.replace(/\n/g, "\r\n"), "\r\n"]] as const) {
      const read = readDocumentFile(S, text);
      // every line of plain text in turn gets a word more
      for (let index = 0; index < read.doc.childCount; index++) {
        const block = read.doc.child(index);
        if (!block.isTextblock || !block.childCount || block.childCount !== 1 || !block.firstChild!.isText) continue;
        const from = text.split(eol);
        if (from.filter((row) => row === block.textContent).length !== 1) continue;
        const children: Node[] = [];
        read.doc.forEach((child, _o, i) => children.push(i === index ? child.type.create(child.attrs, S.text(child.textContent + " edited")) : child));
        const written = writeDocumentFile(read.doc.type.create(read.doc.attrs, children), { source: read.source });
        const to = written.text.split(eol);
        const expected = from.map((row) => (row === block.textContent ? row + " edited" : row));
        assert.deepEqual(to, expected, `${name} ${variant}: line ${index}`);
        edited++;
      }
    }
  }
  assert.ok(edited > 100, `${edited} edits tried`);
});

test("a file whose lines end in different ways keeps every line's own end", () => {
  const cases = [
    "# Title\r\nfirst line\nsecond line\r\n\r\n$$\r\nx\n$$\n\nlast line\r\n",
    "old Mac\rfile\r\rlast\r",
    "no final newline\r\nmixed\nend",
  ];
  for (const text of cases) {
    const read = readDocumentFile(S, text);
    assert.equal(writeDocumentFile(read.doc, { source: read.source }).text, text, JSON.stringify(text));
    // the rows with their ends, as the file has them
    const rows = text.match(/[^\r\n]*(?:\r\n|\n|\r|$)/g)!.filter((row, i, all) => row !== "" || i < all.length - 1);
    read.doc.forEach((block, _offset, index) => {
      if (!block.isTextblock || block.childCount !== 1 || !block.firstChild!.isText) return;
      const at = rows.findIndex((row) => row.replace(/\r?\n|\r/, "") === block.textContent);
      // each line edited in turn: only its own text changes, its line end stays
      const children: Node[] = [];
      read.doc.forEach((child, _o, i) => children.push(i === index ? p(block.textContent + " edited") : child));
      const expected = rows.map((row, i) => (i === at ? row.replace(block.textContent, block.textContent + " edited") : row)).join("");
      assert.equal(writeDocumentFile(doc(children), { source: read.source }).text, expected, `${JSON.stringify(text)}: ${block.textContent}`);
    });
    // a line added after the first one: it takes the line end most lines have; the others keep theirs
    const children: Node[] = [];
    read.doc.forEach((child, _o, i) => {
      children.push(child);
      if (i === 0) children.push(p("added"));
    });
    const most = read.source.eol;
    assert.equal(writeDocumentFile(doc(children), { source: read.source }).text, rows[0] + "added" + most + rows.slice(1).join(""), JSON.stringify(text));
  }
});

test("text written into an empty file ends with a newline; a file without one keeps that", () => {
  const empty = readDocumentFile(S, "");
  assert.equal(writeDocumentFile(empty.doc, { source: empty.source }).text, "");
  assert.equal(writeDocumentFile(doc([p("first words")]), { source: empty.source }).text, "first words\n");
  const bare = readDocumentFile(S, "no newline");
  assert.equal(writeDocumentFile(doc([p("no newline"), p("more")]), { source: bare.source }).text, "no newline\nmore");
});

test("an edited object is written anew, the others stay", () => {
  const text = "Text.\n\n$$a^2$$\n\n|a|b|\n|-|-|\n|1|2|\n\n$$b^2$$\n";
  const read = readDocumentFile(S, text);
  const children: Node[] = [];
  read.doc.forEach((child, _o, i) => children.push(i === 1 ? md("a^3") : child));
  assert.equal(writeDocumentFile(doc(children), { source: read.source }).text, "Text.\n\n$$\na^3\n$$\n\n|a|b|\n|-|-|\n|1|2|\n\n$$b^2$$\n");
});

// ------------------------------------------------------------------ the memory --

test("a line another program rewrote loses its own formatting; the others keep theirs", () => {
  const document = doc(
    [
      n.title.create(null, t("The Title")),
      line({ align: "center" }, "centred line"),
      p(t("red", ["textColor", { color: "#ff0000" }]), t(" word")),
      md("a=b", { numbered: true }),
      line({ align: "right" }, "right line"),
      p("same"),
      line({ align: "center" }, "same"),
    ],
    { paper: "a5" },
  );
  const { written } = roundTrip(document);
  assert.equal(written.text, "# The Title\ncentred line\nred word\n\n$$\na=b \\tag{1}\n$$\n\nright line\nsame\nsame\n");
  assert.equal(written.memory.blocks.length, 6);
  // rewritten elsewhere: one line changed, lines added before and between
  const elsewhere = written.text.replace("red word", "blue word").replace("# The Title", "New first line\n\n# The Title").replace("right line", "added\n\nright line");
  const read = readDocumentFile(S, elsewhere, written.memory);
  assert.equal(read.dropped, 1);
  const kinds: string[] = [];
  read.doc.forEach((node) => kinds.push(`${node.type.name}:${node.attrs.align ?? ""}:${node.textContent}`));
  assert.deepEqual(kinds, ["paragraph::New first line", "paragraph::", "title::The Title", "paragraph:center:centred line", "paragraph::blue word", "math_display::", "paragraph::added", "paragraph::", "paragraph:right:right line", "paragraph::same", "paragraph:center:same"]);
  assert.equal(read.doc.child(4).firstChild!.marks.length, 0);
  assert.equal(read.doc.child(5).attrs.numbered, true);
  assert.equal(read.doc.attrs.paper, "a5");
  // a formula's text changed elsewhere: it is the file's formula, with the tag the file gives it
  const retagged = readDocumentFile(S, written.text.replace("a=b", "a=c"), written.memory);
  assert.deepEqual([retagged.doc.child(3).attrs.latex, retagged.doc.child(3).attrs.numbered], ["a=c \\tag{1}", false]);
  // a memory that is not one (another version, damaged): the text alone
  for (const broken of ["", "{", '{"format":"other"}', JSON.stringify({ ...written.memory, version: 99 })]) assert.equal(parseDocumentMemory(broken), null);
  const damaged = { ...written.memory, blocks: [{ hash: written.memory.blocks[0].hash, occurrence: 0, lines: 1, nodes: [{ type: "no-such-node" }] }, ...written.memory.blocks.slice(1)] };
  const partly = readDocumentFile(S, written.text, damaged);
  assert.equal(partly.dropped, 1);
  assert.equal(partly.doc.child(0).type.name, "heading");
  assert.equal(partly.doc.child(1).attrs.align, "center");
});

test("pictures: the file's paths stay, a new picture gets one in assets", () => {
  const a = registerImageBytes(png(10))!;
  const b = registerImageBytes(png(11))!;
  const files = new Map([["img/photo one.png", a], ["../shared/b.png", b]]);
  const text = "![first](<img/photo one.png>)\n\nText ![second](../shared/b.png) inline.\n\n![a](<img/photo one.png>) ![b](../shared/b.png)\n";
  const read = withImageFiles((path) => files.get(path) ?? null, () => readDocumentFile(S, text));
  const kinds: string[] = [];
  read.doc.forEach((node) => kinds.push(node.type.name));
  assert.deepEqual(kinds, ["figure", "paragraph", "figure"]);
  assert.equal(read.doc.child(1).child(1).type.name, "image");
  assert.equal(read.doc.child(2).firstChild!.type.name, "image_board");
  assert.deepEqual([...read.source.imagePaths], [[a, "img/photo one.png"], [b, "../shared/b.png"]]);
  const opened = asOpened(read.doc);
  assert.equal(writeDocumentFile(opened, { source: read.source }).text, text);
  // resized and captioned: the paths as the file had them; a new picture beside them
  const c = registerImageBytes(png(12))!;
  const figure = opened.child(0);
  const resized = n.figure.create(null, [p(n.image.create({ ...figure.firstChild!.firstChild!.attrs, width: 0.25 })), n.caption.create(null, t("Caption"))]);
  const added = n.figure.create(null, [p(n.image.create({ src: c, alt: "new", width: 0.5 })), n.caption.create()]);
  const changed = doc([resized, opened.child(1), opened.child(2), added]);
  const written = writeDocumentFile(changed, { source: read.source, imagePaths: new Map([[c, "assets/diagram.png"]]) });
  assert.equal(written.text, "![first](<img/photo one.png>)\nCaption\n\nText ![second](../shared/b.png) inline.\n\n![a](<img/photo one.png>) ![b](../shared/b.png)\n\n![new](assets/diagram.png)\n");
  assert.deepEqual([...written.images], [[a, "img/photo one.png"], [b, "../shared/b.png"], [c, "assets/diagram.png"]]);
  const back = withImageFiles((path) => new Map([...files, ["assets/diagram.png", c]]).get(path) ?? null, () => readDocumentFile(S, written.text, written.memory));
  assert.equal(comparable(back.doc), comparable(changed));
  // the memory's pictures are found by their paths when their bytes changed on the disk
  const d = registerImageBytes(png(13))!;
  const replaced = withImageFiles((path) => (path === "img/photo one.png" ? d : (new Map([...files, ["assets/diagram.png", c]]).get(path) ?? null)), () => readDocumentFile(S, written.text, written.memory));
  assert.equal(replaced.doc.child(0).firstChild!.firstChild!.attrs.src, d);
  assert.equal(replaced.doc.child(0).firstChild!.firstChild!.attrs.width, 0.25);
});

test("a picture in a table is the file its path names when the document is read", () => {
  const first = registerImageBytes(png(30))!;
  const second = registerImageBytes(png(31))!;
  const text = "| x |\n|---|\n| ![p](fig.png) |\n";
  const pictures = (document: Node) => {
    const keys: string[] = [];
    document.descendants((node) => {
      if (node.type.name === "image") keys.push(node.attrs.src);
      return true;
    });
    return keys;
  };
  // the same text in two folders (or read again after the picture's file was replaced): each its own picture
  assert.deepEqual(pictures(withImageFiles((path) => (path === "fig.png" ? first : null), () => readDocumentFile(S, text)).doc), [first]);
  assert.deepEqual(pictures(withImageFiles((path) => (path === "fig.png" ? second : null), () => readDocumentFile(S, text)).doc), [second]);
});

test("a new picture's file: its own name in assets, a pasted one named by its time, never renumbered", () => {
  const photo = registerImageBytes(png(20), "C:\\Users\\me\\Pictures\\Photo 1.PNG")!;
  const same = registerImageBytes(png(21), "photo 1.png")!;
  const drawn = registerImageBytes(png(22), "diagram.svg")!;
  const odd = registerImageBytes(png(23), 'a:b*c?".png')!;
  const pasted = registerImageBytes(png(24), "image.png")!;
  const pastedToo = registerImageBytes(png(25))!;
  for (const key of [pasted, pastedToo]) imageEntry(key)!.added = new Date(2026, 10, 18, 9, 5, 7).getTime();
  const old = registerImageBytes(png(26), "old.png")!;
  const picture = (key: string) => n.figure.create(null, [p(n.image.create({ src: key })), n.caption.create()]);
  const document = doc([photo, same, drawn, odd, pasted, pastedToo, old].map(picture));
  const known = new Map([[old, "img/old.png"]]);
  const paths = newPicturePaths(document, known, ["assets/diagram.png"]);
  assert.deepEqual([...paths], [
    [photo, "assets/Photo 1.png"],
    [same, "assets/photo 1-2.png"],
    [drawn, "assets/diagram-2.png"],
    [odd, "assets/a-b-c.png"],
    [pasted, "assets/image-20261118-090507.png"],
    [pastedToo, "assets/image-20261118-090507-2.png"],
  ]);
  // given once, a picture keeps its path whatever is added before it
  const later = doc([picture(registerImageBytes(png(27), "photo 1.png")!), ...[photo, same].map(picture)]);
  const all = new Map([...known, ...paths]);
  assert.deepEqual([...newPicturePaths(later, all)].map(([, path]) => path), ["assets/photo 1-3.png"]);
  const written = writeDocumentFile(document, { imagePaths: all });
  assert.match(written.text, /^!\[\]\(<assets\/Photo 1\.png>\)\n\n!\[\]\(<assets\/photo 1-2\.png>\)\n/);
  assert.match(written.text, /!\[\]\(img\/old\.png\)\n$/);
});

// ------------------------------------------------------------- bytes and text --

test("the file's encoding, byte order mark and line ends are kept", () => {
  const text = "# 标题\r\n\r\n正文 $x$。\r\n";
  const utf8 = new TextEncoder().encode(text);
  const withMark = new Uint8Array([0xef, 0xbb, 0xbf, ...utf8]);
  const utf16le = new Uint8Array([0xff, 0xfe, ...[...text].flatMap((ch) => [ch.charCodeAt(0) & 0xff, ch.charCodeAt(0) >> 8])]);
  const utf16be = new Uint8Array([0xfe, 0xff, ...[...text].flatMap((ch) => [ch.charCodeAt(0) >> 8, ch.charCodeAt(0) & 0xff])]);
  // "标题" and "正文" in GBK
  const gbk = new Uint8Array([0x23, 0x20, 0xb1, 0xea, 0xcc, 0xe2, 0x0d, 0x0a, 0x0d, 0x0a, 0xd5, 0xfd, 0xce, 0xc4, 0x20, 0x24, 0x78, 0x24, 0xa1, 0xa3, 0x0d, 0x0a]);
  for (const [name, bytes] of Object.entries({ "utf-8": utf8, "utf-8-bom": withMark, "utf-16le": utf16le, "utf-16be": utf16be, gb18030: gbk })) {
    const decoded = decodeFileText(bytes);
    assert.equal(decoded.encoding, name);
    assert.equal(decoded.text, text, name);
    const read = readDocumentFile(S, decoded.text);
    // unchanged
    const same = encodeFileText(writeDocumentFile(read.doc, { source: read.source }).text, decoded.encoding);
    assert.deepEqual([...same.bytes], [...bytes], name);
    // edited: the same encoding and line ends
    const edited = doc([read.doc.child(0), read.doc.child(1), p("正文改了")]);
    const out = encodeFileText(writeDocumentFile(edited, { source: read.source }).text, decoded.encoding);
    assert.equal(out.encoding, name);
    assert.equal(decodeFileText(out.bytes).text, "# 标题\r\n\r\n正文改了\r\n", name);
  }
});

test("GB 18030 and Windows-1252 files keep every byte of the lines that were not edited", () => {
  const hex = (bytes: Uint8Array) => [...bytes].map((byte) => byte.toString(16).padStart(2, "0")).join(" ");
  /** The file read, its second line edited to `edit`, written back. */
  const editSecond = (bytes: Uint8Array, edit: string) => {
    const decoded = decodeFileText(bytes);
    const read = readDocumentFile(S, decoded.text);
    assert.equal(hex(encodeFileText(writeDocumentFile(read.doc, { source: read.source }).text, decoded.encoding).bytes), hex(bytes), "unchanged");
    const children: Node[] = [];
    read.doc.forEach((child, _offset, index) => children.push(index === 1 ? p(edit) : child));
    return { decoded, out: encodeFileText(writeDocumentFile(doc(children), { source: read.source }).text, decoded.encoding) };
  };
  // GB 18030: an ideographic space (A1 A1, which A3 A0 also reads as), a bullet and an emoji (four bytes each)
  const first = [0xa1, 0xa1, 0xa1, 0xa1, 0xd5, 0xfd, 0xce, 0xc4, 0x81, 0x36, 0xa6, 0x31, 0x94, 0x39, 0xfc, 0x36, 0x0d, 0x0a];
  const gb = new Uint8Array([...first, 0x61, 0x0d, 0x0a, 0xa1, 0xa1, 0x62, 0x0d, 0x0a]);
  const { decoded, out } = editSecond(gb, "改 • 😀");
  assert.equal(decoded.encoding, "gb18030");
  assert.equal(decoded.text, "　　正文•😀\r\na\r\n　b\r\n");
  assert.equal(out.encoding, "gb18030");
  assert.equal(hex(out.bytes), hex(new Uint8Array([...first, 0xb8, 0xc4, 0x20, 0x81, 0x36, 0xa6, 0x31, 0x20, 0x94, 0x39, 0xfc, 0x36, 0x0d, 0x0a, 0xa1, 0xa1, 0x62, 0x0d, 0x0a])));
  // bytes that are neither UTF-8 nor GB 18030 (Latin-1): read as Windows-1252, written back as they were
  const latin = new Uint8Array([0x63, 0x61, 0x66, 0xe9, 0x0a, 0x78, 0x0a, 0x9c, 0x75, 0x76, 0x72, 0x65, 0x20, 0x80, 0x0a]);
  const kept = editSecond(latin, "y");
  assert.equal(kept.decoded.encoding, "windows-1252");
  assert.equal(kept.decoded.text, "café\nx\nœuvre €\n");
  assert.equal(kept.out.encoding, "windows-1252");
  assert.equal(hex(kept.out.bytes), hex(new Uint8Array([0x63, 0x61, 0x66, 0xe9, 0x0a, 0x79, 0x0a, 0x9c, 0x75, 0x76, 0x72, 0x65, 0x20, 0x80, 0x0a])));
  // a character the encoding has not: the file becomes UTF-8, its text the same
  const widened = editSecond(latin, "中文");
  assert.equal(widened.out.encoding, "utf-8");
  assert.equal(new TextDecoder().decode(widened.out.bytes), "café\n中文\nœuvre €\n");
  assert.equal(encodeFileText("\ud800", "gb18030").encoding, "utf-8");
});

test("a long document is written in good time", () => {
  const blocks: Node[] = [];
  for (let i = 0; i < 4000; i++) {
    if (i % 40 === 7) blocks.push(md(`x_{${i}} = ${i}`, { numbered: i % 80 === 7 }));
    else if (i % 55 === 3) blocks.push(line({ align: "center" }, `centred ${i}`));
    else blocks.push(p(`Line ${i} with `, mi(`a_{${i}}`), " and text that goes on for a while to be a real line."));
  }
  const document = asOpened(doc(blocks));
  const started = performance.now();
  const written = writeDocumentFile(document);
  const first = performance.now() - started;
  const read = readDocumentFile(S, written.text, written.memory);
  assert.equal(comparable(read.doc), comparable(document));
  const again = performance.now();
  writeDocumentFile(asOpened(read.doc), { source: read.source });
  const second = performance.now() - again;
  assert.ok(first < 3000 && second < 3000, `first ${first.toFixed(0)} ms, again ${second.toFixed(0)} ms`);
});
