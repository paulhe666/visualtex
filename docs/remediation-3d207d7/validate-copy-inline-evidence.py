"""Validate evidence captured from real Word UI operations; never create fixtures/XML."""
import base64
import hashlib
import json
import posixpath
import sys
import zlib
import unicodedata
from pathlib import Path
import xml.etree.ElementTree as ET

ROOT = Path(__file__).parent / 'evidence'
NS = {
    'pkg': 'http://schemas.microsoft.com/office/2006/xmlPackage',
    'o': 'urn:schemas-microsoft-com:office:office',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}

def require(condition, message):
    if not condition:
        raise AssertionError(message)

def load(label):
    summary = json.loads((ROOT / (label + '-inventory.json')).read_text(encoding='utf-8-sig'))
    require(len(summary['documents']) == 1, 'Evidence must identify exactly one owned test document')
    entry = summary['documents'][0]
    data = json.loads((ROOT / (entry['key'] + '.json')).read_text(encoding='utf-8-sig'))
    xml = ET.parse(ROOT / (entry['key'] + '.xml'))
    return data, xml

def decode(value):
    require(value.startswith('visualtex:v1:deflate:'), 'Missing VisualTeX cache payload')
    value = value.split(':')[-1]
    return json.loads(zlib.decompress(base64.urlsafe_b64decode(value + '=' * (-len(value) % 4)), -15))

def cfb_hashes(xml):
    parts = {part.attrib['{' + NS['pkg'] + '}name']: part for part in xml.findall('pkg:part', NS)}
    rels = parts['/word/_rels/document.xml.rels']
    targets = {r.attrib['Id']: posixpath.normpath('/word/' + r.attrib['Target'])
               for r in rels.findall('.//rel:Relationship', NS)}
    result = []
    for ole in parts['/word/document.xml'].findall('.//o:OLEObject', NS):
        target = targets[ole.attrib['{' + NS['r'] + '}id']]
        data = base64.b64decode(parts[target].find('pkg:binaryData', NS).text)
        result.append(hashlib.sha256(data).hexdigest())
    return result

def native_metadata(data):
    return [decode(s['alternative']) for s in data['shapes'] if s['prog'] == 'VisualTeX.Formula.1']

def main():
    if len(sys.argv) < 4:
        raise SystemExit('usage: script sources-label copies-label edited-label [reopened-label]')
    source, source_xml = load(sys.argv[1])
    copies, copies_xml = load(sys.argv[2])
    edited, edited_xml = load(sys.argv[3])
    require(len(source['shapes']) == 2 and len(source['maths']) == 1, 'Invalid three-format source fixture')
    for data in (copies, edited):
        require(len(data['shapes']) == 6 and len(data['maths']) == 3, 'Copy/edit changed formula count')
        require(len(data['tables']) == 0, 'An inline copy unexpectedly created a table')
        require(len([b for b in data['bookmarks'] if b['name'].startswith('VTOMML_')]) == 3,
                'Each OMML must retain its own identity')
    originals = cfb_hashes(source_xml)
    for label, data, xml in [('copies', copies, copies_xml), ('edited', edited, edited_xml)]:
        require(cfb_hashes(xml)[:2] == originals, label + ': an original OLE payload was modified')
        require(data['maths'][0] == source['maths'][0], label + ': original OMML changed')
        require(native_metadata(data)[0] == native_metadata(source)[0], label + ': original metadata changed')
    metadata = native_metadata(copies)
    require(len({m['formulaId'] for m in metadata}) == 3, 'Duplicate native FormulaId')
    require(len({line['id'] for m in metadata for line in m['lines']}) == 3, 'Duplicate formula-line identity')
    for m in metadata:
        for key in ('latex', 'displayMode', 'numbered', 'fontSizePt', 'formulaLetterFont', 'formulaChineseFont'):
            require(m.get(key) == metadata[0].get(key), 'Lost copied format: ' + key)
    books = {b['name']: b for b in copies['bookmarks']}
    for shape, m in zip([s for s in copies['shapes'] if s['prog'] == 'VisualTeX.Formula.1'], metadata):
        b = books['VTO_' + m['formulaId'].replace('-', '')]
        require((b['start'], b['end']) == (shape['start'], shape['end']), 'VTO spans adjacent content')
    for prog in ('VisualTeX.Formula.1', 'Equation.DSMT4'):
        shapes = [s for s in copies['shapes'] if s['prog'] == prog]
        require(len(shapes) == 3, 'Wrong object type after paste')
        require(all(s['width'] == shapes[0]['width'] and s['height'] == shapes[0]['height'] for s in shapes),
                'Copied OLE geometry differs from source')
    require(cfb_hashes(copies_xml)[4:] == [originals[1]] * 2, 'MathType copy lost native OLE content')
    for m in copies['maths']:
        require((m['font'], m['size'], m['type'], m['text']) ==
                tuple(source['maths'][0][k] for k in ('font', 'size', 'type', 'text')),
                'OMML copy lost content, font, size or inline layout')
    changed = native_metadata(edited)
    require(changed[1]['formulaId'] == metadata[1]['formulaId'] and changed[1]['latex'] == 'x+y=7',
            'Native copy edit did not preserve identity and new source')
    require(changed[2] == metadata[2], 'Editing one native copy changed the other')
    require('u+v=8' in unicodedata.normalize('NFKC', edited['maths'][1]['text']), 'OMML copy edit did not apply')
    require(edited['maths'][2]['text'] == copies['maths'][2]['text'], 'Editing OMML copy affected its sibling')
    after_hashes = cfb_hashes(edited_xml)
    before_hashes = cfb_hashes(copies_xml)
    require(after_hashes[4] != before_hashes[4] and after_hashes[5] == before_hashes[5],
            'MathType edit did not affect precisely the chosen copied OLE')
    if len(sys.argv) > 4:
        reopened, reopened_xml = load(sys.argv[4])
        require(len(reopened['shapes']) == 6 and len(reopened['maths']) == 3, 'Reopen changed formula counts')
        require(native_metadata(reopened) == native_metadata(edited), 'Save/reopen lost native identity or content')
        # Word commits live VisualTeX IPersistStorage metadata on Save. The outer
        # CFB bytes legitimately change; check the actual equation/preview streams
        # and persisted IDs instead of mistaking serialization for content damage.
        saved_storage = json.loads((ROOT / (sys.argv[4] + '-embedded-storage.json')).read_text(encoding='utf-8-sig'))
        edit_storage = json.loads((ROOT / (sys.argv[3] + '-embedded-storage.json')).read_text(encoding='utf-8-sig'))
        require(len(saved_storage) == len(edit_storage) == 6, 'Incomplete embedded-storage evidence')
        native_ids = []
        cached_native = iter(native_metadata(reopened))
        for before_stream, saved_stream in zip(edit_storage, saved_storage):
            require(before_stream['kind'] == saved_stream['kind'], 'Reopen changed OLE class')
            if saved_stream['kind'] == 'mathTypeOle':
                require(before_stream['nativeHash'] == saved_stream['nativeHash'], 'Save/reopen changed MathType MTEF')
            else:
                cached = next(cached_native)
                native_ids.append(saved_stream['formulaId'])
                require(saved_stream['previewHash'] == before_stream['previewHash'], 'Save/reopen changed VisualTeX preview')
                for field in ('formulaId', 'latex', 'lines', 'fontSizePt', 'formulaLetterFont', 'formulaChineseFont', 'displayMode', 'numbered'):
                    require(saved_stream['metadata'].get(field) == cached.get(field), 'Persisted OLE metadata mismatch: ' + field)
        require(len(set(native_ids)) == 3, 'Saved embedded FormulaIds are not unique')
        require([m['text'] for m in reopened['maths']] == [m['text'] for m in edited['maths']],
                'Save/reopen changed OMML content')
    report = {'status': 'passed', 'source': sys.argv[1], 'copies': sys.argv[2], 'edited': sys.argv[3],
              'reopened': sys.argv[4] if len(sys.argv) > 4 else None,
              'nativeFormulaIds': [m['formulaId'] for m in metadata],
              'originalOlePayloadsUnchanged': True, 'allNineFormulasRetained': True,
              'copyGeometryAndFormatsPreserved': True, 'selectedCopiesIndependentlyEdited': True}
    (ROOT / 'copy-inline-ui-validation.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
    print(json.dumps(report, indent=2))

if __name__ == '__main__':
    main()
