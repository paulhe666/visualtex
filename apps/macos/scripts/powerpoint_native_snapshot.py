#!/usr/bin/env python3
import json,pathlib,sys,zipfile,xml.etree.ElementTree as E
from verify_powerpoint_native_omml import effective,NS
base=pathlib.Path(__file__).resolve().parents[1]/'build-logs/powerpoint-native-omml'
slide=int(sys.argv[1])
with zipfile.ZipFile(base/'NativeOmmlAcceptance.pptm') as z:
 root=effective(E.fromstring(z.read(f'ppt/slides/slide{slide}.xml')))[0]
 body=root.find('.//p:txBody',NS)
 print(json.dumps({'ordinary':[''.join(t.text or '' for t in p.findall('.//a:t',NS)) for p in body.findall('a:p',NS)],'mathCount':len(body.findall('.//m:oMath',NS))},ensure_ascii=False))
