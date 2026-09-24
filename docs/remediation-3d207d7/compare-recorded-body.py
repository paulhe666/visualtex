"""Supplementary comparison of recorded plain body paragraphs, never UI acceptance.

Reports its scope explicitly. Mixed formula/text paragraphs and fields are excluded
and must be inspected separately; no Word connection or document mutation.
"""
import json
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}
w = '{' + ns['w'] + '}'

def tree(element):
    if element is None:
        return None
    return [element.tag, dict(sorted(element.attrib.items())), element.text or '',
            [tree(child) for child in element]]

def read(path):
    root = ET.parse(path).getroot()
    result = []
    for p in root.findall('.//w:body//w:p', ns):
        if any(p.find(query, ns) is not None for query in
               ['.//m:oMath', './/w:object', './/w:fldChar', './/w:fldSimple']):
            continue
        text = ''.join(node.text or '' for node in p.findall('.//w:t', ns))
        if not text:
            continue
        runs = [[r.find('w:t', ns).text or '', tree(r.find('w:rPr', ns))]
                for r in p.findall('w:r', ns) if r.find('w:t', ns) is not None]
        result.append(dict(text=text, paragraphProperties=tree(p.find('w:pPr', ns)), runs=runs))
    return result

before, after, output = map(Path, sys.argv[1:4])
a, b = read(before), read(after)
report = dict(before=str(before), after=str(after), scope='plain nonempty body paragraphs without formula or field',
              equal=a == b, beforeCount=len(a), afterCount=len(b),
              beforeParagraphs=a, afterParagraphs=b)
output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps({k: v for k, v in report.items() if not k.endswith('Paragraphs')}))
