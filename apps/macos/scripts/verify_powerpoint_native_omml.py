#!/usr/bin/env python3
"""Validate actual saved Office output, ignoring only schema fallback branches."""
import argparse
import json
import pathlib
import xml.etree.ElementTree as ET
import zipfile

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
}

def effective(node):
    if node.tag == '{' + NS['mc'] + '}AlternateContent':
        selected = node.find('mc:Choice', NS)
        if selected is None:
            selected = node.find('mc:Fallback', NS)
        if selected is None:
            raise AssertionError('Empty OOXML AlternateContent')
        result = []
        for child in selected:
            result.extend(effective(child))
        return result
    result = ET.Element(node.tag, node.attrib)
    result.text, result.tail = node.text, node.tail
    for child in node:
        result.extend(effective(child))
    return [result]

EXPECTED = {
    'inline_mid': ['before  after'], 'inline_caret': ['before after'],
    'inline_start': ['after'], 'inline_end': ['before'], 'inline_empty': [''],
    'unicode': ['前文  后文'], 'multiline_replace': ['before  after'],
    'font_size': ['before  after'], 'format_preservation': ['before  after'],
    'display_middle': ['before ', '', ' after'], 'display_existing': ['before', '', 'after'],
    'display_start': ['', ' after'], 'display_end': ['before ', ''], 'display_empty': [''],
    'display_bullets': ['before ', '', ' after'],
    'standalone_inline': [''], 'standalone_display': [''],
    'stale': ['before X after changed'], 'cancel': ['before X after'],
}

def verify(deck, results):
    checked = []
    with zipfile.ZipFile(deck) as archive:
        for case in results:
            name = case['name']
            raw = ET.fromstring(archive.read(f"ppt/slides/slide{case['slideIndex']}.xml"))
            root = effective(raw)[0]
            tree = root.find('p:cSld/p:spTree', NS)
            shapes = [n for n in tree if n.tag not in {
                '{' + NS['p'] + '}nvGrpSpPr', '{' + NS['p'] + '}grpSpPr', '{' + NS['p'] + '}extLst'
            }]
            assert len(shapes) == 1 and shapes[0].tag == '{' + NS['p'] + '}sp', (name, 'extra/nontext shape')
            shape = shapes[0]
            body = shape.find('p:txBody', NS)
            assert body is not None, (name, 'missing native text body')
            paragraphs = body.findall('a:p', NS)
            plain = [''.join(t.text or '' for t in p.findall('.//a:t', NS)) for p in paragraphs]
            assert plain == EXPECTED[name], (name, plain, EXPECTED[name])
            math = body.findall('.//m:oMath', NS)
            count = 0 if name in ('stale', 'cancel') else 1
            assert len(math) == count, (name, 'wrong native math count', len(math))
            assert not shape.findall('.//p:pic', NS), (name, 'picture instead of native math')
            assert not shape.findall('.//p:grpSp', NS), (name, 'group container')
            assert all('visualtex' not in v.lower() for n in shape.iter() for v in n.attrib.values()), (name, 'VisualTeX object metadata')
            if math:
                assert len(math[0].findall('.//m:f', NS)) == 1, (name, 'missing fraction')
                den = math[0].find('.//m:f/m:den', NS)
                assert ''.join(t.text or '' for t in den.findall('.//m:t', NS)) == '2'
                expected_size = '3600' if name == 'font_size' else '2400'
                sizes = {n.get('sz') for n in math[0].findall('.//a:rPr', NS) if n.get('sz')}
                assert sizes == {expected_size}, (name, 'wrong mathematical font sizes', sizes)
            if name.startswith('display_') or name == 'standalone_display':
                formula_paragraphs = [p for p in paragraphs if p.findall('.//m:oMath', NS)]
                assert len(formula_paragraphs) == 1
                assert formula_paragraphs[0].find('a:pPr', NS).get('algn') == 'ctr', (name, 'not centered')
            if name in ('format_preservation', 'display_bullets'):
                for paragraph in ([paragraphs[0]] if name == 'format_preservation' else [paragraphs[0], paragraphs[-1]]):
                    props = paragraph.find('a:pPr', NS)
                    assert props.get('algn') == 'r', (name, 'changed neighbor alignment')
                    assert props.find('a:spcBef/a:spcPts', NS).get('val') == '1100', (name, 'changed paragraph spacing')
                    assert props.find('a:spcAft/a:spcPts', NS).get('val') == '1300', (name, 'changed paragraph spacing')
                    if name == 'display_bullets':
                        assert props.find('a:buAutoNum', NS) is not None, (name, 'lost neighbor numbering')
                prefix = next(r for r in paragraphs[0].findall('a:r', NS) if r.find('a:t', NS).text == 'before ')
                suffix = next(r for r in paragraphs[-1].findall('a:r', NS) if r.find('a:t', NS).text == 'after')
                assert prefix.find('a:rPr', NS).get('b') == '1'
                assert suffix.find('a:rPr', NS).get('i') == '1'
            checked.append({'case': name, 'native_math': count, 'paragraphs': len(paragraphs), 'shapes': 1})
    return checked

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    base = pathlib.Path(__file__).resolve().parents[1] / 'build-logs/powerpoint-native-omml'
    parser.add_argument('--deck', type=pathlib.Path, default=base / 'NativeOmmlAcceptance.pptm')
    parser.add_argument('--results', type=pathlib.Path, default=base / 'native-integration-results.json')
    args = parser.parse_args()
    result = verify(args.deck, json.loads(args.results.read_text()))
    (base / 'native-structure-verification.json').write_text(json.dumps(result, ensure_ascii=False, indent=2))
    print(f'PASS: {len(result)} saved PowerPoint cases; native OMML, exact surrounding text, layout, font size, no extra shapes/metadata, preserved neighbor paragraph formatting.')
