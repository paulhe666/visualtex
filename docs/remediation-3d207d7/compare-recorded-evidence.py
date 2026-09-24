"""Read saved evidence only. This is diagnostic comparison, never Word acceptance."""
import json
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

before, after, output = map(Path, sys.argv[1:4])
differences = []

def compare(a, b, path):
    if type(a) is not type(b):
        differences.append(dict(path=path, before=a, after=b))
    elif isinstance(a, dict):
        for key in sorted(a.keys() | b.keys()):
            compare(a.get(key), b.get(key), path + '/' + key)
    elif isinstance(a, list):
        if len(a) != len(b):
            differences.append(dict(path=path+'/length', before=len(a), after=len(b)))
        for index, (left, right) in enumerate(zip(a, b)):
            compare(left, right, path + '/' + str(index))
    elif a != b:
        differences.append(dict(path=path, before=a, after=b))

def node(element):
    return dict(tag=element.tag, attrs=element.attrib, text=element.text or '',
                tail=element.tail or '', children=[node(child) for child in element])

if before.suffix == '.json':
    read = lambda path: json.loads(path.read_text(encoding='utf-8-sig'))
else:
    read = lambda path: node(ET.parse(path).getroot())
compare(read(before), read(after), '')
report = dict(before=str(before), after=str(after), count=len(differences), differences=differences)
output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(dict(count=len(differences), first=differences[:12]), ensure_ascii=False, indent=2))
