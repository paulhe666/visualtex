"""Measure a production Word document at its unchanged 160% native zoom.
ROIs refer to the preserved five-image wrapped-paragraph acceptance document.
Its rendered font/size/style must be retained; normalize by painted CJK height.
"""
from pathlib import Path
import sys,json
from PIL import Image,ImageDraw,ImageFont
import numpy as np
root=Path(sys.argv[1])
references={'text':(40,0),'sum':(39,1),'display':(35,.5)}
def bounds(ink,roi):
 x,y,X,Y=roi;ys,xs=np.where(ink[y:Y,x:X]);assert len(ys),f'Empty ROI {roi}'
 return [int(x+xs.min()),int(y+ys.min()),int(x+xs.max()+1),int(y+ys.max()+1)]
def centre(b):return (b[1]+b[3])/2
rows=[];sheet=Image.new('RGB',(1370,5*300),'white');draw=ImageDraw.Draw(sheet)
font=ImageFont.truetype('/System/Library/Fonts/Supplemental/Arial.ttf',22)
for i,mode in enumerate(['bottom','baseline','center','top','auto']):
 image=Image.open(root/f'{mode}-native.png').convert('RGB');ink=(np.asarray(image)<130).all(axis=2)
 ys=np.where(ink[350:950,655:791].any(axis=1))[0]+350
 bands=[g for g in np.split(ys,np.where(np.diff(ys)>4)[0]+1) if len(g)>=15]
 assert len(bands)==2, f'{mode}: unexpected body rows {[g.tolist() for g in bands]}'
 bodies=[bounds(ink,(655,int(g[0]),791,int(g[-1])+1)) for g in bands]
 for idx,kind,line,x,X in [(1,'text',0,865,893),(2,'text',0,1364,1393),(3,'text',1,552,582),(4,'sum',1,948,978),(5,'display',1,1306,1340)]:
  body=bodies[line];cy=centre(body);span=body[3]-body[1]
  anchor=bounds(ink,(x,int(cy-span),X,int(cy+span)))
  h,delta=references[kind];error=(centre(anchor)-cy)*h/span-delta
  rows.append(dict(mode=mode,image=idx,kind=kind,body=body,anchor=anchor,errorPx=round(error,3)))
  assert abs(error)<=3,f'{mode}/image {idx}: {error:.3f} reference pixels'
 top=min(b[1] for b in bodies)-45;bottom=max(b[3] for b in bodies)+45
 draw.text((10,i*300+8),mode,font=font,fill='black')
 sheet.paste(image.crop((325,top,1660,bottom)),(10,i*300+35))
report={'result':'PASS','method':'Native Word, unchanged production theme font and image styles, black pixels <130, normalize to reference CJK height','count':len(rows),'maxReferenceErrorPx':max(abs(r['errorPx']) for r in rows),'rows':rows}
(root/'actual-document-pixel-acceptance.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
sheet.save(root/'actual-document-five-modes.png')
print(json.dumps({k:report[k] for k in ['result','count','maxReferenceErrorPx']}))
