"""Verify actual Word-saved image baselines against embedded SVG math geometry."""
import sys, json, zipfile, xml.etree.ElementTree as E
from inspect_word_conversion_docx import decode_metadata, NS, W
P='{'+NS['wp']+'}'
rows=[]
with zipfile.ZipFile(sys.argv[1]) as z:
 root=E.fromstring(z.read('word/document.xml'))
 parents={c:p for p in root.iter() for c in p}
 for drawing in root.iter(W+'drawing'):
  props=drawing.find('.//'+P+'docPr'); ext=drawing.find('.//'+P+'extent')
  if props is None or not props.get('title','').startswith('visualtex:formula-ref:v1:'):continue
  meta=decode_metadata(props.get('descr',''))
  if meta['displayMode']!='inline':continue
  run=drawing
  while run.tag!=W+'r': run=parents[run]
  pos=run.find('./'+W+'rPr/'+W+'position')
  position=0 if pos is None else float(pos.get(W+'val'))/2
  height=int(ext.get('cy'))/12700
  descent=-meta['referenceBaselinePt']*height/meta['referenceHeightPt']
  error=position+descent
  assert abs(error)<.06, (meta['lines'],error,position,height,meta['referenceBaselinePt'])
  rows.append({'latex':'\n'.join(x['latex'] for x in meta['lines']),'fontSizePt':meta['fontSizePt'],'positionPt':position,'descentPt':descent,'errorPt':error})
assert rows, 'No real Word inline formulas found'
report={'count':len(rows),'maxBaselineErrorPt':max(abs(x['errorPt']) for x in rows),'formulas':rows}
print(json.dumps(report,ensure_ascii=False,indent=2))
