#!/usr/bin/env python3
"""Inspect the saved, real PowerPoint equations produced by the edit tests."""
import json
import pathlib
import re
import zipfile
import xml.etree.ElementTree as ET
from verify_powerpoint_native_omml import NS, effective

base=pathlib.Path(__file__).resolve().parents[1]/'build-logs/powerpoint-native-omml'
records=json.loads((base/'native-edit-integration.json').read_text())
checked=[]
with zipfile.ZipFile(base/'NativeOmmlAcceptance.pptm') as archive:
    for record in records:
        name=record['name']
        slide=int(re.search(r'slide=(\d+)',record['report']).group(1))
        root=effective(ET.fromstring(archive.read(f'ppt/slides/slide{slide}.xml')))[0]
        tree=root.find('p:cSld/p:spTree',NS)
        shapes=tree.findall('p:sp',NS)
        assert len(shapes)==1 and not tree.findall('p:grpSp',NS) and not tree.findall('p:pic',NS),(name,'not one original native text box')
        shape=shapes[0];body=shape.find('p:txBody',NS)
        ordinary=''.join(x.text or '' for x in body.findall('.//a:t',NS))
        expected=''.join(record['beforeSnapshot']['ordinary'])
        current_paragraphs=[''.join(t.text or '' for t in p.findall('.//a:t',NS)) for p in body.findall('a:p',NS)]
        assert current_paragraphs==record['beforeSnapshot']['ordinary'],(name,'ordinary text, whitespace or paragraphs changed',current_paragraphs,record['beforeSnapshot']['ordinary'])
        math=body.findall('.//m:oMath',NS)
        native_text=''.join(x.text or '' for x in body.findall('.//m:t',NS))
        if name not in ('ordinary','ambiguous','miss','style_changed'):
            assert '5' in native_text and any('𝑦' in (x.text or '') for x in body.findall('.//m:t',NS)),(name,'edit not applied')
            assert len(body.findall('.//m:f',NS))==(2 if name in ('multiple','merged') else 1),(name,'wrong replaced expression count')
        else:
            assert '5' not in native_text,(name,'rejected operation modified formula')
        if name=='display':
            assert len(body.findall('a:p',NS))==3,(name,'display paragraph changed')
        else:
            assert len(body.findall('a:p',NS))==1,(name,'inline paragraph split')
        assert all('visualtex' not in value.lower() for element in shape.iter() for value in element.attrib.values()),(name,'managed formula metadata')
        assert all(element.tag not in ('{'+NS['p']+'}tags','{'+NS['p']+'}oleObj') for element in shape.iter()),(name,'unexpected formula management object')
        checked.append({'case':name,'slide':slide,'nativeMathZones':len(math),'ordinaryText':ordinary,'passed':True})
(base/'native-edit-structure-verification.json').write_text(json.dumps(checked,ensure_ascii=False,indent=2))
print(f'PASS: {len(checked)} real saved native edits; exact neighboring text/whitespace, native mathematical content, no new objects or VisualTeX metadata.')
