/** Run against the DOCX saved by real Word after importing the supplied source. */
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { DOMParser } from "@xmldom/xmldom";
import { parseLatexMarkdownDocument } from "../src/office/documentImport/documentImportParser.ts";
import { normalizeFormulaEditorDocument } from "../src/office/shared/formulaEditorDocument.ts";
import { renderOfficeFormulaArtifacts } from "../src/office/shared/formulaRenderArtifacts.ts";

globalThis.DOMParser ??= DOMParser;
const xml = new DOMParser().parseFromString("<root/>", "application/xml");
Object.getPrototypeOf(xml).querySelector ??= function(name) { return this.getElementsByTagName(name)?.item(0) ?? null; };
if (!("children" in Object.getPrototypeOf(xml.documentElement))) Object.defineProperty(Object.getPrototypeOf(xml.documentElement), "children", { get() { return Array.from(this.childNodes ?? []).filter(n => n.nodeType === 1); } });

const [sourcePath, resultPath, kind] = process.argv.slice(2);
assert(sourcePath && resultPath && ["image", "omml"].includes(kind));
const expected = parseLatexMarkdownDocument(readFileSync(sourcePath, "utf8"), "auto", 11)
  .filter(block => block.kind === "formula")
  .map(block => {
    const document = normalizeFormulaEditorDocument([{ id: "fixture", latex: block.latex }], "raw");
    const artifact = renderOfficeFormulaArtifacts({ ...document, host: "word", displayMode: block.displayMode, numbered: block.numbered });
    return { latex: artifact.canonicalLatex, numbered: block.numbered };
  });
const result = JSON.parse(execFileSync("python3", ["-c", String.raw`
import json, sys, zipfile
from pathlib import Path
import xml.etree.ElementTree as E
sys.path.insert(0, str(Path(sys.argv[2]).resolve()))
from inspect_word_conversion_docx import inspect, NS, W
result = inspect(sys.argv[1])
with zipfile.ZipFile(sys.argv[1]) as archive:
    root = E.fromstring(archive.read('word/document.xml'))
order = []
for node in root.iter():
    if node.tag == W+'bookmarkStart' and node.get(W+'name','').startswith('VT_F_'):
        compact = node.get(W+'name')[5:]
        formula = next(f for f in result['formulas'] if f['formulaId'].replace('-','') == compact)
        order.append(formula['formulaId'])
    elif node.tag == '{'+NS['wp']+'}docPr':
        title = node.get('title','')
        if title.startswith('visualtex:formula-ref:v1:'): order.append(title.split(':')[3])
result['ordered'] = [next(f for f in result['formulas'] if f['formulaId']==i) for i in order]
print(json.dumps(result))
`, resultPath, new URL(".", import.meta.url).pathname], { encoding: "utf8" }));
assert.equal(result.ordered.length, expected.length, "Missing, duplicated or unrendered formulas");
assert.equal(result.formulas.length, expected.length, "Dangling formula metadata");
for (const [index, actual] of result.ordered.entries()) {
  assert.equal(actual.kind, kind, `Formula ${index + 1} output kind`);
  assert.equal(actual.latex, expected[index].latex, `Formula ${index + 1} source/order changed`);
  assert.equal(actual.numbered, expected[index].numbered);
  if (kind === "image") {
    const scale = actual.fontSizePt / 14;
    assert.ok(Math.abs(actual.width - actual.referenceWidthPt * scale) < 0.12, `Formula ${index + 1} width changed`);
    assert.ok(Math.abs(actual.height - actual.referenceHeightPt * scale) < 0.12, `Formula ${index + 1} height changed`);
    assert.ok(actual.width > 0 && actual.height > 0);
  } else {
    assert.ok(actual.mathText?.length, `Formula ${index + 1} has no native content`);
    assert.ok(!/\\[A-Za-z]+/.test(actual.mathText), `Formula ${index + 1} contains literal commands`);
  }
}
const numbered = expected.filter(formula => formula.numbered);
assert.equal(result.sequences.length, numbered.length);
assert.ok(result.sequences.every(field => field.owner?.startsWith("VT_N_")));
assert.deepEqual(result.sequences.map(field => Number(field.result)), numbered.map((_, i) => i + 1));
console.log(`PASS: real Word import ${expected.length} ${kind} formulas; source/order, ${numbered.length} numbers and ${kind === "image" ? "image geometry" : "native content"} verified`);
