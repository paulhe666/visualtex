"""Read already captured real Word evidence; this is not a UI acceptance runner."""
import json
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

directory = Path(__file__).parent
evidence = directory / 'evidence'
stage = sys.argv[1]
source = (directory / 'redraw-source.tex').read_text(encoding='utf-8-sig')
prose = [part.strip() for part in re.split(r'\$\$.*?\$\$', source, flags=re.S) if part.strip()]
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}

def tree(element):
    if element is None:
        return None
    return [element.tag, dict(sorted(element.attrib.items())), element.text or '', [tree(child) for child in element]]

def properties(root, text):
    found = []
    for paragraph in root.findall('.//w:body/w:p', ns):
        for run in paragraph.findall('w:r', ns):
            if text in ''.join(node.text or '' for node in run.findall('w:t', ns)):
                found.append({'run': tree(run.find('w:rPr', ns)), 'paragraph': tree(paragraph.find('w:pPr', ns))})
    if len(found) != 1:
        raise ValueError(f'Expected one actual body run containing {text!r}, got {len(found)}')
    return found[0]

rows = []
for mode in ['omml', 'vt', 'mt']:
    for breaks in ['cr', 'soft']:
        label = f'{stage}-{mode}-{breaks}'
        before_files = list(evidence.glob(label+'-source-__*.json'))
        after_files = list(evidence.glob(label+'-result-__*.json'))
        if len(before_files) != 1 or len(after_files) != 1:
            raise ValueError(f'Missing or ambiguous real evidence: {label}')
        before_path, after_path = before_files[0], after_files[0]
        before = json.loads(before_path.read_text(encoding='utf-8-sig'))
        after = json.loads(after_path.read_text(encoding='utf-8-sig'))
        bx, ax = ET.parse(before_path.with_suffix('.xml')).getroot(), ET.parse(after_path.with_suffix('.xml')).getroot()
        body = []
        for text in prose:
            bp, ap = properties(bx, text), properties(ax, text)
            matches = [p for p in after['paras'] if text in p['text']]
            body.append({'text': text, 'runFormattingEqual': bp['run'] == ap['run'],
                         'paragraphFormattingEqual': bp['paragraph'] == ap['paragraph'],
                         'beforeXml': bp, 'afterXml': ap, 'actualComParagraphs': matches})
        numbers = [f for f in after['fields'] if f['type'] == 12 and f['result']]
        formulas = sorted(after['maths'] + after['shapes'], key=lambda x:x['start'])
        positions = [b['actualComParagraphs'][0]['start'] for b in body]
        expected_order = len(formulas) == 2 and positions[0] < formulas[0]['start'] < positions[1] < formulas[1]['start']
        rows.append({'label': label, 'document': after['name'], 'before': str(before_path), 'after': str(after_path),
                     'sourceParagraphCount': before['paragraphs'], 'sourceManualBreaks': len(bx.findall('.//w:body//w:br', ns)),
                     'ommlCount':len(after['maths']), 'oleCount':len(after['shapes']),
                     'xmlOmmlCount':len(ax.findall('.//w:body//m:oMath', ns)),
                     'xmlOleCount':len(ax.findall('.//w:body//w:object', ns)),
                     'numberResults':[f['result'] for f in numbers], 'numberFields':numbers,
                     'sourceFormulaOrderPreserved':expected_order, 'body':body,
                     'screenshot':str(evidence/(label+'-visible.png')), 'visualReviewRequired':True})
report = {'scope':'Recorded COM/XML cross-check of the original two-formula redraw source; screenshots require separate real review', 'rows':rows}
output = evidence/(stage+'-redraw-com-xml-review.json')
output.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
for row in rows:
    print(json.dumps({'label':row['label'],'document':row['document'],'omml':row['ommlCount'],'ole':row['oleCount'],
                      'numbers':row['numberResults'],'order':row['sourceFormulaOrderPreserved'],
                      'sourceParagraphs':row['sourceParagraphCount'],'sourceManualBreaks':row['sourceManualBreaks'],
                      'bodyFormat':all(p['runFormattingEqual'] and p['paragraphFormattingEqual'] for p in row['body'])},ensure_ascii=False))
