"""Verify actual Word snapshots for adjacent OMML copy/paste and independent edit."""
import json
import runpy
import sys
import unicodedata
from pathlib import Path

helpers = runpy.run_path(str(Path(__file__).with_name('validate-copy-inline-evidence.py')))
load, require, cfb_hashes = helpers['load'], helpers['require'], helpers['cfb_hashes']
root = Path(__file__).parent / 'evidence'

def signature(math):
    return tuple(math[key] for key in ('text', 'font', 'size', 'type'))

def ids(data):
    return {b['name'] for b in data['bookmarks'] if b['name'].startswith('VTOMML_')}

def main():
    if len(sys.argv) not in (4, 5):
        raise SystemExit('usage: script before-label two-copies-label edited-label [reopened-label]')
    before, bx = load(sys.argv[1])
    copies, cx = load(sys.argv[2])
    edited, ex = load(sys.argv[3])
    require(len(before['maths']) == 3 and len(copies['maths']) == len(edited['maths']) == 5,
            'Two adjacent pastes must add precisely two separate OMath objects')
    require(len(before['shapes']) == len(copies['shapes']) == len(edited['shapes']) == 6,
            'Adjacent OMML repair changed unrelated OLE inventory')
    require(cfb_hashes(bx) == cfb_hashes(cx) == cfb_hashes(ex), 'An unrelated OLE payload changed')
    require(before['paragraphs'] == copies['paragraphs'] == edited['paragraphs'],
            'Adjacent inline paste inserted a paragraph break')
    require(len(ids(copies)) == 5 and ids(before) < ids(copies) == ids(edited),
            'New OMML identities must be unique; previous identities must survive edits')
    for data in (copies, edited):
        require(signature(data['maths'][0]) == signature(before['maths'][0]), 'Original source formula changed')
        require(signature(data['maths'][3]) == signature(before['maths'][1]), 'Existing adjacent equation changed')
        require(signature(data['maths'][4]) == signature(before['maths'][2]), 'Other original OMML changed')
    for index in (1, 2):
        require(signature(copies['maths'][index]) == signature(before['maths'][0]),
                'A copied formula lost source content or inline typography')
    require('s+t=10' in unicodedata.normalize('NFKC', edited['maths'][1]['text']), 'Selected adjacent copy edit did not apply')
    require(signature(edited['maths'][2]) == signature(copies['maths'][2]), 'Editing one adjacent copy changed the other')
    if len(sys.argv) == 5:
        reopened, _ = load(sys.argv[4])
        require(ids(reopened) == ids(edited), 'Save/reopen lost an independent OMML identity')
        require([signature(m) for m in reopened['maths']] == [signature(m) for m in edited['maths']],
                'Save/reopen changed equation content or layout')
        require(reopened['paragraphs'] == edited['paragraphs'], 'Save/reopen changed paragraph structure')
    report = {'status': 'passed', 'evidence': sys.argv[1:], 'actualOMathCount': 5,
              'originalsPreserved': True, 'independentAdjacentCopies': True,
              'inlineTypographyPreserved': True, 'noNewParagraphs': True,
              'independentEdit': True, 'saveReopen': len(sys.argv) == 5}
    (root / 'copy-adjacent-omml-validation.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
    print(json.dumps(report, indent=2))

if __name__ == '__main__':
    main()
