"""Compare all recorded prose, including paragraphs containing inline formulas.

This reads evidence only. It neither drives Word nor declares UI acceptance.
Formula payloads and fields are excluded; ordinary runs, breaks, tabs and paragraph
formatting stay in scope. Empty paragraphs are reported by the COM evidence.
"""
import json
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}
W = '{' + NS['w'] + '}'

def tree(element):
    if element is None:
        return None
    return [element.tag, dict(sorted(element.attrib.items())), element.text or '',
            [tree(child) for child in element]]

def read(path):
    root = ET.parse(path).getroot()
    result = []
    for paragraph in root.findall('.//w:body/w:p', NS):
        # Field result text is not prose; keep its distinct validation in the
        # actual Fields/Bookmarks evidence rather than treating it as body text.
        if paragraph.find('.//w:fldChar', NS) is not None:
            continue
        runs = []
        for run in paragraph.findall('w:r', NS):
            if run.find('w:object', NS) is not None:
                continue
            content = [tree(child) for child in run
                       if child.tag in {W+'t', W+'tab', W+'br', W+'cr'}]
            if content:
                runs.append([tree(run.find('w:rPr', NS)), content])
        text = ''.join(node.text or '' for run in paragraph.findall('w:r', NS)
                       for node in run.findall('w:t', NS))
        if not text:
            continue
        result.append({'text': text, 'paragraphProperties': tree(paragraph.find('w:pPr', NS)),
                       'runs': runs})
    return result

before, after, output = map(Path, sys.argv[1:4])
a, b = read(before), read(after)
report = {'before': str(before), 'after': str(after),
          'scope': 'ordinary text/runs and paragraph properties, including mixed inline formula paragraphs; field paragraphs excluded',
          'equal': a == b, 'beforeCount': len(a), 'afterCount': len(b),
          'beforeParagraphs': a, 'afterParagraphs': b}
output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps({key: value for key, value in report.items() if not key.endswith('Paragraphs')}))
