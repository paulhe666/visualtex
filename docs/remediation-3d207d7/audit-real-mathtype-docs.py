"""Read-only OOXML inventory; preserve exact originals as evidence copies."""
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent
OUT = ROOT / 'evidence' / 'real-mathtype-20260926'
OUT.mkdir(parents=True, exist_ok=True)
DOWNLOADS = Path.home() / 'Downloads'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
O = '{urn:schemas-microsoft-com:office:office}'
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def audit(path, label):
    dest = OUT / (label + '-disk-original.docx')
    original_hash = sha(path)
    if dest.exists():
        if sha(dest) != original_hash:
            raise RuntimeError('Original changed since first snapshot; refusing overwrite: ' + str(dest))
    else:
        shutil.copy2(path, dest)
    with zipfile.ZipFile(dest) as z:
        xml = z.read('word/document.xml')
        root = ET.fromstring(xml)
        styles = ET.fromstring(z.read('word/styles.xml'))
        paras = []
        changed = []
        for i, p in enumerate(root.iter(W + 'p'), 1):
            runs = []
            for r in p.iter(W + 'r'):
                text = ''.join(t.text or '' for t in r.iter(W + 't'))
                rp = r.find(W + 'rPr')
                props = {} if rp is None else {c.tag.split('}')[-1]: {k.split('}')[-1]:v for k,v in c.attrib.items()} for c in rp}
                objects = [e.attrib for e in r.iter(O+'OLEObject')]
                row = {'text':text,'props':props,'ole':objects}
                runs.append(row)
                if ('position' in props or 'vertAlign' in props) and (text or objects):
                    changed.append({'paragraph':i, 'context':''.join(t.text or '' for t in p.iter(W+'t')), **row})
            paras.append({'index':i,'text':''.join(x['text'] for x in runs),'properties':ET.tostring(p.find(W+'pPr'),encoding='unicode') if p.find(W+'pPr') is not None else '', 'runs':runs})
        embeddings = [n for n in z.namelist() if n.startswith('word/embeddings/')]
        style_positions = [ET.tostring(s,encoding='unicode') for s in styles.findall(W+'style') if s.find('.//'+W+'position') is not None or s.find('.//'+W+'vertAlign') is not None]
        summary = {'path':str(path),'copy':str(dest),'sha256':original_hash,'paragraphs':len(paras),'omaths':len(list(root.iter(M+'oMath'))),'oleCount':len(list(root.iter(O+'OLEObject'))),'embeddingCount':len(embeddings),'dollars':sum(p['text'].count('$') for p in paras),'shiftedRuns':len(changed)}
        detail = {**summary,'shifted':changed,'stylesWithVerticalPosition':style_positions,'paras':paras}
        (OUT/(label+'-disk-inventory.json')).write_text(json.dumps(detail,ensure_ascii=False,indent=2),encoding='utf-8')
        (OUT/(label+'-disk-document.xml')).write_bytes(xml)
        (OUT/(label+'-disk-text.txt')).write_text('\n'.join(p['text'] for p in paras),encoding='utf-8')
        print(json.dumps(summary,ensure_ascii=False))
        print('SHIFTED RUN SAMPLES:',json.dumps(changed[:5],ensure_ascii=False))

files = [('physics',DOWNLOADS/'11.核物理、相对论4.docx')]
for p in DOWNLOADS.rglob('高一数学第一次月考卷01*.docx'):
    if p.name.endswith('（考试版）.docx'):
        files.append(('math-exam',p))
    elif p.name.endswith('（考试版A3）.docx'):
        files.append(('math-exam-a3',p))
if len(files) != 3:
    raise RuntimeError('Expected exactly the three named source files, got '+str(files))
for label,path in files:
    audit(path,label)
