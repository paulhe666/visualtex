import type { LayoutSnapshot, PathOp, Rect, CaretStop } from './layout';
import type { Position } from './editing';
import { rgbHex } from './paint.js';
import { validateGridGeometry } from './grid-geometry.js';
import { validateArtworkGeometry } from './custom-artwork.js';

const NS = 'http://www.w3.org/2000/svg';
let nextView = 1;
const finite = (value: number): number => {
  if (!Number.isFinite(value)) throw new Error('SVG view received non-finite geometry');
  return value;
};
const number = (value: number) => String(finite(value));
const samePosition = (a: Position, b: Position) => a.kind === b.kind && a.nodeId === b.nodeId
  && (a.kind === 'row' && b.kind === 'row' ? a.offset === b.offset
    : a.kind === 'text' && b.kind === 'text' && a.grapheme === b.grapheme);
function element<K extends keyof SVGElementTagNameMap>(name: K): SVGElementTagNameMap[K] {
  return document.createElementNS(NS, name);
}
function rectPath(rect: Rect): SVGRectElement {
  const {xMin,yMin,xMax,yMax}=rect;
  [xMin,yMin,xMax,yMax].forEach(finite);
  if(xMax<xMin || yMax<yMin) throw new Error('Invalid SVG rectangle');
  const result=element('rect');
  result.setAttribute('x',number(xMin));result.setAttribute('y',number(yMin));
  result.setAttribute('width',number(xMax-xMin));result.setAttribute('height',number(yMax-yMin));
  return result;
}
function pathData(path: PathOp[]): string {
  return path.map(op=>{
    if(op.op==='close') return 'Z';
    const command={move:'M',line:'L',quad:'Q',cubic:'C'}[op.op];
    return command+op.points.map(number).join(' ');
  }).join('');
}
export interface SceneIdentity {documentId: string; revision: number; sourceId: number}
export interface ViewOverlay extends SceneIdentity {
  caret: Position | null;
  rectangles: Rect[];
}
export interface ViewPointer extends SceneIdentity {
  xEm: number;
  yEm: number;
  shiftKey: boolean;
  pointerId: number;
  /** 'double' follows a completed double click (pointer events carry no click count). */
  phase: 'down'|'move'|'up'|'double';
}
export interface SvgViewOptions {
  onPointer?: (event: ViewPointer) => void;
  accessibleLabel?: string;
  /** Editor-only structural affordance. Exported core SVG remains undecorated. */
  showEmptySlots?: boolean;
}

/** Thin renderer of core geometry. No LaTeX parsing, math layout, tree editing,
 * hit-testing heuristics or Office operations live in this view. */
export class SvgSceneView {
  readonly container: HTMLDivElement;
  readonly svg: SVGSVGElement;
  private scene: LayoutSnapshot | null = null;
  private overlayLayer: SVGGElement;
  private caretElement: SVGRectElement | null = null;
  private zoom = 1;
  private disposed = false;
  private readonly abort = new AbortController();
  private readonly namespace = `vt-core-view-${nextView++}`;
  private sceneBuilds = 0;
  private overlayBuilds = 0;
  private pointerActive: number | null = null;

  constructor(parent: HTMLElement, private readonly options: SvgViewOptions = {}) {
    this.container=document.createElement('div');this.container.className='vt-core-scene';
    this.container.setAttribute('role','group');
    this.container.setAttribute('aria-label',options.accessibleLabel ?? '数学公式编辑区域');
    this.svg=element('svg');this.svg.classList.add('vt-core-svg');
    this.svg.setAttribute('aria-hidden','true');
    this.svg.style.display='block';this.svg.style.overflow='visible';this.svg.style.touchAction='none';
    this.overlayLayer=element('g');this.overlayLayer.classList.add('vt-core-overlay');
    this.container.append(this.svg);parent.append(this.container);
    const listen=(phase:ViewPointer['phase'])=>(event:PointerEvent)=>{
      if(!this.scene || this.disposed) return;
      if(phase==='down') {
        if(event.button!==0) return;
        this.pointerActive=event.pointerId;this.svg.setPointerCapture(event.pointerId);
      } else if(this.pointerActive!==event.pointerId) return;
      const point=this.clientToLayout(event.clientX,event.clientY);
      this.options.onPointer?.({...this.identity(),...point,shiftKey:event.shiftKey,pointerId:event.pointerId,phase});
      if(phase==='up') {this.pointerActive=null;this.svg.releasePointerCapture(event.pointerId);}
    };
    this.svg.addEventListener('pointerdown',listen('down'),{signal:this.abort.signal});
    this.svg.addEventListener('pointermove',listen('move'),{signal:this.abort.signal});
    this.svg.addEventListener('pointerup',listen('up'),{signal:this.abort.signal});
    this.svg.addEventListener('pointercancel',()=>{this.pointerActive=null;},{signal:this.abort.signal});
    this.svg.addEventListener('dblclick',event=>{
      if(!this.scene||this.disposed||event.button!==0)return;
      const point=this.clientToLayout(event.clientX,event.clientY);
      this.options.onPointer?.({...this.identity(),...point,shiftKey:event.shiftKey,pointerId:-1,phase:'double'});
    },{signal:this.abort.signal});
  }
  get stats() {return {sceneBuilds:this.sceneBuilds,overlayBuilds:this.overlayBuilds,zoom:this.zoom};}
  get snapshotIdentity(): SceneIdentity | null {return this.scene?this.identity():null;}
  private identity(): SceneIdentity {
    if(!this.scene) throw new Error('View has no current layout');
    return {documentId:this.scene.documentId,revision:this.scene.revision,sourceId:this.scene.sourceId};
  }
  private active() {if(this.disposed) throw new Error('SVG scene view has been disposed');}
  setSnapshot(snapshot: LayoutSnapshot, accessibleLabel?: string): void {
    this.active();
    if(!snapshot.documentId || !Number.isInteger(snapshot.revision) || !Number.isInteger(snapshot.sourceId)
      || ![snapshot.width,snapshot.height,snapshot.widthEm,snapshot.heightEm,snapshot.fontSizePx].every(v=>Number.isFinite(v)&&v>0))
      throw new Error('Invalid layout snapshot');
    if(this.scene?.documentId===snapshot.documentId && snapshot.revision<this.scene.revision)
      throw new Error('Stale layout revision');
    // Build separately; invalid resource data cannot replace the current DOM.
    const next=structuredClone(snapshot);
    validateGridGeometry(next);
    validateArtworkGeometry(next);
    const glyphNodes:SVGElement[]=[],ruleNodes:SVGElement[]=[],artworkNodes:SVGElement[]=[];
    const defs=element('defs'), glyphs=element('g'), decorations=element('g'), overlay=element('g');
    glyphs.classList.add('vt-core-glyphs');glyphs.setAttribute('fill','currentColor');
    decorations.classList.add('vt-core-decorations');decorations.setAttribute('pointer-events','none');
    overlay.classList.add('vt-core-overlay');overlay.setAttribute('pointer-events','none');
    for(const [index,resource] of next.resources.entries()) {
      const path=element('path');path.id=`${this.namespace}-glyph-${index}`;
      path.setAttribute('d',pathData(resource.shape.path));defs.append(path);
    }
    for(const glyph of next.glyphs) {
      if(!next.resources[glyph.resource] || !Number.isInteger(glyph.resource) || glyph.scale<=0)
        throw new Error('Invalid glyph resource or scale');
      const use=element('use');use.setAttribute('href',`#${this.namespace}-glyph-${glyph.resource}`);
      use.setAttribute('transform',`translate(${number(glyph.x)} ${number(glyph.y)}) scale(${number(glyph.scale)})`);
      if(glyph.color!==undefined)use.setAttribute('fill',rgbHex(glyph.color));
      use.dataset.nodeId=String(glyph.nodeId);glyphs.append(use);glyphNodes.push(use);
    }
    for(const rule of next.rules) {
      const rect=rectPath(rule.rect);if(rule.color!==undefined)rect.setAttribute('fill',rgbHex(rule.color));
      rect.dataset.nodeId=String(rule.nodeId);glyphs.append(rect);ruleNodes.push(rect);
    }
    const artworks=next.artworks??[];
    const buildArtworkShape=(artwork:typeof artworks[number],artworkIndex:number,paint:string,scope:'paint'|'mask',identify:boolean):SVGGElement=>{
      const outer=element('g'),path=element('path');path.setAttribute('d',pathData(artwork.shape.path));
      outer.setAttribute('transform',`matrix(${artwork.shape.transform.map(number).join(' ')})`);path.setAttribute('transform',`matrix(${artwork.shape.innerTransform.map(number).join(' ')})`);
      path.setAttribute('fill',artwork.shape.fill?paint:'none');
      if(artwork.shape.strokeWidth>0){path.setAttribute('stroke',paint);path.setAttribute('stroke-width',number(artwork.shape.strokeWidth));path.setAttribute('stroke-linecap',artwork.shape.lineCap);path.setAttribute('stroke-linejoin',artwork.shape.lineJoin);}
      else path.setAttribute('stroke','none');
      if(identify){path.dataset.nodeId=String(artwork.nodeId);path.dataset.assetSha256=artwork.assetSha256;}
      if(artwork.shape.clip!==undefined){const clip=element('clipPath');clip.id=`${this.namespace}-artwork-${scope}-clip-${artworkIndex}`;clip.setAttribute('clipPathUnits','userSpaceOnUse');clip.append(rectPath(artwork.shape.clip));defs.append(clip);const clipped=element('g');clipped.setAttribute('clip-path',`url(#${clip.id})`);clipped.append(path);outer.append(clipped);}else outer.append(path);
      return outer;
    };
    const union=(a:Rect,b:Rect):Rect=>({xMin:Math.min(a.xMin,b.xMin),yMin:Math.min(a.yMin,b.yMin),xMax:Math.max(a.xMax,b.xMax),yMax:Math.max(a.yMax,b.yMax)});
    const groupedArtworkNodes:Array<SVGElement|null|undefined>=new Array(artworks.length);let groupIndex=0;
    for(let at=0;at<(next.paintOrder?.length??0);){const item=next.paintOrder![at];if(item.kind!=='artwork'){at++;continue;}const first=artworks[item.index],indices:number[]=[];while(at<next.paintOrder!.length){const part=next.paintOrder![at];if(part.kind!=='artwork')break;const candidate=artworks[part.index];if(candidate.nodeId!==first.nodeId||candidate.assetSha256!==first.assetSha256)break;indices.push(part.index);at++;}
      const group=element('g');group.setAttribute('transform',`translate(${number(first.x)} ${number(first.y)}) scale(${number(first.scale)})`);const visible=indices.filter(i=>artworks[i].shape.operation==='paint');const hasErase=indices.some(i=>artworks[i].shape.operation==='erase');
      if(visible.length){const paint=first.color!==undefined?rgbHex(first.color):'currentColor';if(hasErase){const bounds=visible.map(i=>artworks[i].shape.ink).reduce(union);if(bounds.xMax>bounds.xMin&&bounds.yMax>bounds.yMin){const mask=element('mask');mask.id=`${this.namespace}-artwork-mask-${groupIndex}`;mask.setAttribute('maskUnits','userSpaceOnUse');mask.setAttribute('maskContentUnits','userSpaceOnUse');mask.setAttribute('mask-type','luminance');mask.setAttribute('x',number(bounds.xMin));mask.setAttribute('y',number(bounds.yMin));mask.setAttribute('width',number(bounds.xMax-bounds.xMin));mask.setAttribute('height',number(bounds.yMax-bounds.yMin));for(const i of indices)mask.append(buildArtworkShape(artworks[i],i,artworks[i].shape.operation==='paint'?'#ffffff':'#000000','mask',false));defs.append(mask);const masked=element('g');masked.setAttribute('mask',`url(#${mask.id})`);for(const i of visible)masked.append(buildArtworkShape(artworks[i],i,paint,'paint',true));group.append(masked);}}
        else for(const i of visible)group.append(buildArtworkShape(artworks[i],i,paint,'paint',true));}
      groupedArtworkNodes[indices[0]]=group;for(const i of indices.slice(1))groupedArtworkNodes[i]=null;groupIndex++;
    }
    if(artworks.length)glyphs.replaceChildren(...next.paintOrder!.flatMap(item=>item.kind==='glyph'?[glyphNodes[item.index]]:item.kind==='rule'?[ruleNodes[item.index]]:groupedArtworkNodes[item.index]?[groupedArtworkNodes[item.index]!]:[]));
    if(this.options.showEmptySlots){
      for(const region of next.regions){
        if(region.role!=='emptySlot'&&region.role!=='emptyRow')continue;
        const logical=region.logical;
        const logicalWidth=Math.max(0.34,logical.xMax-logical.xMin);
        const logicalHeight=Math.max(0.58,logical.yMax-logical.yMin);
        const width=Math.max(0.28,logicalWidth-0.10);
        const height=Math.max(0.46,Math.min(0.82,logicalHeight*0.72));
        const x=logical.xMin+(logicalWidth-width)/2;
        const centerY=(logical.yMin+logical.yMax)/2;
        const slot=element('rect');
        slot.classList.add('vt-core-empty-slot');slot.dataset.nodeId=String(region.nodeId);
        slot.setAttribute('x',number(x));slot.setAttribute('y',number(centerY-height/2));
        slot.setAttribute('width',number(width));slot.setAttribute('height',number(height));slot.setAttribute('rx','0.055');
        decorations.append(slot);
      }
    }
    this.svg.setAttribute('viewBox',`0 0 ${number(next.widthEm)} ${number(next.heightEm)}`);
    this.svg.replaceChildren(defs,glyphs,decorations,overlay);this.scene=next;this.overlayLayer=overlay;this.caretElement=null;
    this.resize();this.sceneBuilds++;
    if(accessibleLabel!==undefined) this.container.setAttribute('aria-label',accessibleLabel);
  }
  setOverlay(overlay: ViewOverlay): void {
    this.active();if(!this.scene) throw new Error('Layout required before selection');
    const id=this.identity();
    if(id.documentId!==overlay.documentId || id.revision!==overlay.revision || id.sourceId!==overlay.sourceId)
      throw new Error('Selection belongs to a different layout revision');
    const fragment=document.createDocumentFragment();
    for(const rect of overlay.rectangles) {
      const selection=rectPath(rect);selection.classList.add('vt-core-selection');
      selection.setAttribute('fill','Highlight');selection.setAttribute('fill-opacity','0.25');fragment.append(selection);
    }
    let caret: SVGRectElement | null=null;
    if(overlay.caret) {
      const stop=this.scene.carets.find(c=>samePosition(c.position,overlay.caret!));
      if(!stop) throw new Error('Selection position is absent from the current layout');
      caret=this.drawCaret(stop);fragment.append(caret);
    }
    this.overlayLayer.replaceChildren(fragment);this.caretElement=caret;this.overlayBuilds++;
  }
  private drawCaret(stop: CaretStop): SVGRectElement {
    const caret=rectPath({xMin:stop.x,yMin:stop.top,xMax:stop.x+1/(this.scene!.fontSizePx*this.zoom),yMax:stop.bottom});
    caret.classList.add('vt-core-caret');caret.setAttribute('fill','currentColor');return caret;
  }
  setCaretVisible(visible: boolean): void {this.active();if(this.caretElement)this.caretElement.style.visibility=visible?'visible':'hidden';}
  setZoom(zoom: number): void {
    this.active();if(!Number.isFinite(zoom)||zoom<0.1||zoom>8)throw new Error('Invalid display zoom');
    this.zoom=zoom;if(this.scene)this.resize();
    if(this.caretElement)this.caretElement.setAttribute('width',number(1/(this.scene!.fontSizePx*zoom)));
  }
  private resize(): void {
    this.svg.setAttribute('width',number(this.scene!.width*this.zoom));
    this.svg.setAttribute('height',number(this.scene!.height*this.zoom));
  }
  clientToLayout(clientX: number,clientY: number): {xEm:number;yEm:number} {
    this.active();finite(clientX);finite(clientY);
    const matrix=this.svg.getScreenCTM();if(!matrix)throw new Error('SVG view is not attached to a viewport');
    const p=new DOMPoint(clientX,clientY).matrixTransform(matrix.inverse());
    return {xEm:finite(p.x),yEm:finite(p.y)};
  }
  layoutToClient(xEm:number,yEm:number): {x:number;y:number} {
    this.active();finite(xEm);finite(yEm);
    const matrix=this.svg.getScreenCTM();if(!matrix)throw new Error('SVG view is not attached to a viewport');
    const p=new DOMPoint(xEm,yEm).matrixTransform(matrix);return {x:finite(p.x),y:finite(p.y)};
  }
  dispose(): void {
    if(this.disposed)return;
    this.disposed=true;this.abort.abort();this.scene=null;this.container.remove();this.pointerActive=null;
  }
}
