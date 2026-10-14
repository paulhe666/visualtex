/** Validate a DOCX saved by real Word after either public redraw command.
 * Usage: tsx scripts/word_redraw_text_order_regression.mts source.txt result.docx
 * Reproduces smart-spacing damage with: 中文：甲 $x$ 乙；Before $y$ after.
 * Formula counts alone cannot detect the old result: 中文：甲乙[formula]；...
 */
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { findWindowsWordLatexRedrawSpans } from "../src/office/redraw/wordLatexRedrawParser.ts";

const [sourcePath, resultPath] = process.argv.slice(2);
assert(sourcePath && resultPath, "Provide original source and a DOCX saved by real Word");
const source = readFileSync(sourcePath, "utf8").replace(/\r\n?/g, "\n");
const spans = findWindowsWordLatexRedrawSpans(source);
assert(spans.length > 0);
let expected = "";
let cursor = 0;
for (const span of spans) {
  expected += source.slice(cursor, span.start) + "[FORMULA]";
  cursor = span.end;
}
expected += source.slice(cursor);

const result = JSON.parse(execFileSync("python3", ["-c", String.raw`
import json, sys, zipfile
import xml.etree.ElementTree as E
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
with zipfile.ZipFile(sys.argv[1]) as archive:
    root = E.fromstring(archive.read('word/document.xml'))
counts = {'image': 0, 'omml': 0}
def walk(node):
    if node.tag == W+'drawing':
        extent = node.find('.//'+WP+'extent')
        assert extent is not None and int(extent.get('cx')) > 0 and int(extent.get('cy')) > 0
        counts['image'] += 1
        return '[FORMULA]'
    if node.tag == M+'oMath':
        counts['omml'] += 1
        assert node.find('.//'+M+'t') is not None, 'Empty native formula'
        return '[FORMULA]'
    if node.tag == W+'t': return node.text or ''
    if node.tag == W+'tab': return '\t'
    if node.tag in (W+'br', W+'cr'): return '\n'
    result = ''.join(walk(child) for child in node)
    return result + ('\n' if node.tag == W+'p' else '')
print(json.dumps({'text': walk(root.find(W+'body')), 'counts': counts}))
`, resultPath], { encoding: "utf8" }));

// Word owns a final paragraph mark; do not normalize ANY intervening whitespace
// or punctuation, which would hide the original smart-spacing regression.
// VTPlaceCaretAfterInlineNativeEquation inserts one zero-width U+2060 after
// native inline formulas to exit Word's math zone. Recognize only that boundary;
// retain every source space, punctuation mark and any other joiner.
const visibleText = result.counts.omml > 0
  ? result.text.replace(/\[FORMULA\]\u2060/g, "[FORMULA]")
  : result.text;
assert.equal(visibleText.replace(/\n+$/, ""), expected.replace(/\n+$/, ""));
assert.equal(result.counts.image + result.counts.omml, spans.length);
console.log(`PASS: real Word ${spans.length} formulas; exact surrounding text, spaces, punctuation and order preserved (${JSON.stringify(result.counts)})`);
