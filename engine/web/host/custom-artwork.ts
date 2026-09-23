import type {PathOp,Rect,RgbColor,LayoutSnapshot} from './layout';
export interface SymbolMetrics {widthEm:number;ascentEm:number;descentEm:number}
export interface ArtworkPath {
  path:PathOp[];transform:[number,number,number,number,number,number];innerTransform:[number,number,number,number,number,number];ink:Rect;
  clip?:Rect;outlineFontId?:string;operation:'paint'|'erase';fill:boolean;strokeWidth:number;lineCap:'butt'|'round'|'square';lineJoin:'miter'|'round'|'bevel';
}
export interface CustomSymbolAsset {
  sha256:string;command:string;role:string;limitsBehavior:'auto'|'limits'|'nolimits';metrics:SymbolMetrics;
  definition:Record<string,unknown>;geometry:ArtworkPath[];
  diagnostics:{shapeIndex:number;code:string;message:string}[];
}
export interface ArtworkInstance {
  nodeId:number;assetSha256:string;pathIndex:number;shape:ArtworkPath;x:number;y:number;scale:number;color?:RgbColor;
}
export type PaintItem={kind:'glyph'|'rule'|'artwork';index:number};
const invalid=():never=>{throw new Error('Invalid custom artwork geometry or paint order');};
const rect=(r:Rect)=>r&&[r.xMin,r.yMin,r.xMax,r.yMax].every(Number.isFinite)&&r.xMax>=r.xMin&&r.yMax>=r.yMin;
const contains=(a:Rect,b:Rect)=>b.xMin>=a.xMin-1e-8&&b.yMin>=a.yMin-1e-8&&b.xMax<=a.xMax+1e-8&&b.yMax<=a.yMax+1e-8;
const transformed=(m:ArtworkPath['transform'],x:number,y:number):[number,number]=>[m[0]*x+m[2]*y+m[4],m[1]*x+m[3]*y+m[5]];
const boundedPoint=(p:[number,number])=>p.every(v=>Number.isFinite(v)&&Math.abs(v)<=10000);
const transformedRect=(m:ArtworkPath['transform'],r:Rect):Rect|null=>{
  const points=[transformed(m,r.xMin,r.yMin),transformed(m,r.xMax,r.yMin),transformed(m,r.xMax,r.yMax),transformed(m,r.xMin,r.yMax)];
  if(points.some(p=>!boundedPoint(p)))return null;
  const xs=points.map(p=>p[0]),ys=points.map(p=>p[1]);return{xMin:Math.min(...xs),yMin:Math.min(...ys),xMax:Math.max(...xs),yMax:Math.max(...ys)};
};
const validMatrix=(m:unknown):m is ArtworkPath['transform']=>Array.isArray(m)&&m.length===6&&m.every(v=>Number.isFinite(v)&&Math.abs(v)<=1e8)&&Math.abs(m[0]*m[3]-m[1]*m[2])>=1e-12;
const sameColor=(a:RgbColor|undefined,b:RgbColor|undefined)=>a===undefined?b===undefined:b!==undefined&&a.r===b.r&&a.g===b.g&&a.b===b.b;
/** Independent host boundary validation. Nothing from the resource's authored
 * JSON is installed as innerHTML/CSS or put into an SVG attribute directly. */
export function validateArtworkGeometry(scene:LayoutSnapshot):void {
  const artworks=scene.artworks??[],order=scene.paintOrder??[];
  if(!Array.isArray(artworks)||!Array.isArray(order))invalid();
  if(!artworks.length){if(order.length)invalid();return;}
  if(artworks.length>4096||order.length!==scene.glyphs.length+scene.rules.length+artworks.length)invalid();
  const used=new Set<string>(),closed=new Set<string>();let current:{key:string;next:number;x:number;y:number;scale:number;color?:RgbColor}|null=null;
  for(const item of order){
    const length=item?.kind==='glyph'?scene.glyphs.length:item?.kind==='rule'?scene.rules.length:item?.kind==='artwork'?artworks.length:-1;
    if(!Number.isSafeInteger(item.index)||item.index<0||item.index>=length||used.has(`${item.kind}:${item.index}`))invalid();used.add(`${item.kind}:${item.index}`);
    if(item.kind!=='artwork'){if(current){closed.add(current.key);current=null;}continue;}
    const a=artworks[item.index],key=`${a?.nodeId}:${a?.assetSha256}`;
    if(current!==null){if(current.key===key){if(a.pathIndex!==current.next||!Object.is(a.x,current.x)||!Object.is(a.y,current.y)||!Object.is(a.scale,current.scale)||!sameColor(a.color,current.color))invalid();current.next++;continue;}closed.add(current.key);current=null;}
    if(closed.has(key)||a.pathIndex!==0)invalid();current={key,next:1,x:a.x,y:a.y,scale:a.scale,color:a.color};
  }
  const frame={xMin:0,yMin:0,xMax:scene.widthEm,yMax:scene.heightEm};let total=0;
  for(const a of artworks){
    const rawShape=a?.shape as (ArtworkPath&{textSource?:unknown})|undefined;
    if(!a||!/^[0-9a-f]{64}$/u.test(a.assetSha256)||!Number.isSafeInteger(a.pathIndex)||a.pathIndex<0||a.pathIndex>=8192
      ||!Number.isSafeInteger(a.nodeId)||a.nodeId<=0||![a.x,a.y,a.scale].every(Number.isFinite)||a.scale<=0||!rawShape||rawShape.textSource!==undefined||!rect(a.shape.ink)||!Array.isArray(a.shape.path)||a.shape.path.length===0
      ||(a.shape.outlineFontId!==undefined&&(!scene.fonts.some(font=>font.id===a.shape.outlineFontId)||typeof a.shape.outlineFontId!=='string'))
      ||!validMatrix(a.shape.transform)||!validMatrix(a.shape.innerTransform)||(a.shape.clip!==undefined&&(!rect(a.shape.clip)||[a.shape.clip.xMin,a.shape.clip.yMin,a.shape.clip.xMax,a.shape.clip.yMax].some(v=>Math.abs(v)>10000)))
      ||!['paint','erase'].includes(a.shape.operation)||typeof a.shape.fill!=='boolean'||!Number.isFinite(a.shape.strokeWidth)||a.shape.strokeWidth<0||a.shape.strokeWidth>100
      ||!['butt','round','square'].includes(a.shape.lineCap)||!['miter','round','bevel'].includes(a.shape.lineJoin))invalid();
    total+=a.shape.path.length;if(total>131072)invalid();
    const owner=scene.regions.find(r=>r.nodeId===a.nodeId);if(!owner)throw new Error('Invalid custom artwork geometry or paint order');
    if(owner.role!=='customSymbol')invalid();
    const ownerInk=owner.ink;if(ownerInk!==null&&!rect(ownerInk))invalid();
    const r=a.shape.ink,world={xMin:r.xMin*a.scale+a.x,yMin:r.yMin*a.scale+a.y,xMax:r.xMax*a.scale+a.x,yMax:r.yMax*a.scale+a.y};
    if(a.shape.operation==='paint'&&(!ownerInk||!contains(frame,world)||!contains(ownerInk,world)))invalid();
    if(a.shape.clip!==undefined){const clipped=transformedRect(a.shape.transform,a.shape.clip);if(!clipped||!rect(clipped))invalid();}
    let moved=false;for(const op of a.shape.path){
      if(op.op==='move')moved=true;if(!moved)invalid();
      if(op.op==='close')continue;
      const length=op.op==='move'||op.op==='line'?2:op.op==='quad'?4:op.op==='cubic'?6:-1;
      if(!Array.isArray(op.points)||op.points.length!==length||!op.points.every(v=>Number.isFinite(v)&&Math.abs(v)<=10000))invalid();
      for(let i=0;i<op.points.length;i+=2){const q=transformed(a.shape.innerTransform,op.points[i],op.points[i+1]);if(!boundedPoint(q))invalid();const p=transformed(a.shape.transform,q[0],q[1]);if(!boundedPoint(p)||(a.shape.clip===undefined&&!contains(r,{xMin:p[0],yMin:p[1],xMax:p[0],yMax:p[1]})))invalid();}
    }
  }
}
