import {EngineEditorHost,type EditorHostOptions} from './editor-host.js';
import {EngineEditorSessionOwner,type EngineRuntimePort} from './editor-session.js';
import type {EditorSnapshot} from './editing.js';
import type {LayoutOptions} from './layout.js';

export interface VirtualFormulaRecord {
  documentId:string;
  source:string;
}
export interface VirtualFormulaViewportOptions {
  layout:LayoutOptions;
  overscanPx?:number;
  minHeightPx?:number;
  maxMounted?:number;
  onChange?:(index:number,snapshot:EditorSnapshot)=>void;
  onError?:(index:number,error:Error)=>void;
  onEnter?:(index:number)=>void;
}
interface FormulaEntry {
  index:number;
  record:VirtualFormulaRecord;
  shell:HTMLDivElement;
  owner:EngineEditorSessionOwner|null;
  host:EngineEditorHost|null;
  intersecting:boolean;
  lastHeight:number;
}

/** Viewport owner for a long formula document. Shells are passive layout
 * placeholders; only nearby formulas receive SVG/input DOM. Once materialized,
 * a formula keeps one EngineEditorSessionOwner and the same core handle across
 * repeated park/resume cycles. */
export class VirtualFormulaViewport {
  readonly element:HTMLDivElement;
  private readonly entries:FormulaEntry[]=[];
  private readonly observer:IntersectionObserver;
  private readonly abort=new AbortController();
  private readonly overscanPx:number;
  private readonly minHeightPx:number;
  private readonly maxMounted:number;
  private disposed=false;
  private reconciling=false;

  constructor(private readonly runtime:EngineRuntimePort,private readonly scrollRoot:HTMLElement,
    records:readonly VirtualFormulaRecord[],private readonly options:VirtualFormulaViewportOptions){
    if(records.length===0)throw new Error('Virtual formula viewport requires at least one formula');
    this.overscanPx=options.overscanPx??160;
    this.minHeightPx=options.minHeightPx??72;
    this.maxMounted=options.maxMounted??32;
    if(!Number.isFinite(this.overscanPx)||this.overscanPx<0||this.overscanPx>5000)throw new Error('Invalid virtual viewport overscan');
    if(!Number.isFinite(this.minHeightPx)||this.minHeightPx<16||this.minHeightPx>4096)throw new Error('Invalid virtual formula minimum height');
    if(!Number.isInteger(this.maxMounted)||this.maxMounted<1||this.maxMounted>48)throw new Error('Virtual viewport maxMounted must be between 1 and 48');
    const ids=new Set<string>();
    for(const record of records){
      if(!record.documentId||ids.has(record.documentId))throw new Error('Virtual formula documentId must be unique and nonempty');
      ids.add(record.documentId);
    }
    const doc=scrollRoot.ownerDocument;
    this.element=doc.createElement('div');this.element.className='vt-core-virtual-list';
    for(const [index,record] of records.entries()){
      const shell=doc.createElement('div');shell.className='vt-core-virtual-formula';shell.dataset.formulaIndex=String(index);
      shell.style.minHeight=`${this.minHeightPx}px`;shell.style.position='relative';shell.style.contain='layout paint';
      shell.setAttribute('aria-label',`公式 ${index+1}`);this.element.append(shell);
      this.entries.push({index,record:{...record},shell,owner:null,host:null,intersecting:false,lastHeight:this.minHeightPx});
    }
    scrollRoot.append(this.element);
    const win=doc.defaultView;if(!win?.IntersectionObserver)throw new Error('IntersectionObserver is required for formula viewport virtualization');
    this.observer=new win.IntersectionObserver(changes=>{
      for(const change of changes){
        const index=Number((change.target as HTMLElement).dataset.formulaIndex);
        const entry=this.entries[index];if(entry?.shell===change.target)entry.intersecting=change.isIntersecting;
      }
      this.reconcile();
    },{root:scrollRoot,rootMargin:`${this.overscanPx}px 0px`,threshold:0});
    for(const entry of this.entries)this.observer.observe(entry.shell);
    doc.addEventListener('focusin',()=>queueMicrotask(()=>this.reconcile()),{signal:this.abort.signal});
  }

  private report(entry:FormulaEntry,reason:unknown):void{
    const error=reason instanceof Error?reason:new Error(String(reason));this.options.onError?.(entry.index,error);
  }
  private pinned(entry:FormulaEntry):boolean{
    const host=entry.host;if(!host)return false;
    return host.input.hasUncommittedInput||entry.shell.contains(entry.shell.ownerDocument.activeElement);
  }
  private mount(entry:FormulaEntry):boolean{
    if(entry.host||this.disposed)return Boolean(entry.host);
    let created=false;
    try{
      if(!entry.owner){entry.owner=new EngineEditorSessionOwner(this.runtime,entry.record.documentId,entry.record.source);created=true;}
      const hostOptions:EditorHostOptions={documentId:entry.record.documentId,layout:this.options.layout,
        onChange:snapshot=>{entry.record.source=snapshot.document.source;this.options.onChange?.(entry.index,snapshot);},
        onError:error=>this.options.onError?.(entry.index,error),onEnter:()=>this.options.onEnter?.(entry.index)};
      entry.host=new EngineEditorHost(this.runtime,entry.shell,hostOptions,entry.owner);
      entry.shell.dataset.mounted='true';entry.shell.removeAttribute('aria-busy');return true;
    }catch(error){
      this.report(entry,error);entry.shell.setAttribute('aria-busy','true');
      if(created&&entry.owner&&!entry.owner.isParked&&!entry.owner.isDestroyed){try{entry.owner.park();}catch{/* retained active owner is visible in stats and can be retried */}}
      return false;
    }
  }
  private unmount(entry:FormulaEntry):boolean{
    if(!entry.host||!entry.owner||this.pinned(entry))return false;
    try{
      const measured=entry.shell.getBoundingClientRect().height;
      if(Number.isFinite(measured)&&measured>0)entry.lastHeight=Math.max(this.minHeightPx,measured);
      // Park first. If the retention budget rejects the move, keep the existing
      // input/view alive rather than leaving an active session with no owner UI.
      entry.owner.park();
      entry.host.dispose();entry.host=null;entry.shell.dataset.mounted='false';
      entry.shell.style.minHeight=`${entry.lastHeight}px`;return true;
    }catch(error){this.report(entry,error);return false;}
  }
  private reconcile():void{
    if(this.disposed||this.reconciling)return;this.reconciling=true;
    try{
      for(const entry of this.entries)if(!entry.intersecting)this.unmount(entry);
      let mounted=this.entries.reduce((count,entry)=>count+Number(entry.host!==null),0);
      for(const entry of this.entries){
        if(!entry.intersecting||entry.host||mounted>=this.maxMounted)continue;
        if(this.mount(entry))mounted++;
      }
    }finally{this.reconciling=false;}
  }
  /** Deterministic test/app hook after scroll/layout changes. Normal operation
   * is still driven by IntersectionObserver. */
  refresh():void{this.reconcile();}
  inspect(index:number):EditorSnapshot|null{
    const entry=this.entries[index];if(!entry)throw new Error('Formula index outside virtual document');
    return entry.owner?.inspect()??null;
  }
  source(index:number):string{
    const entry=this.entries[index];if(!entry)throw new Error('Formula index outside virtual document');
    return entry.owner?.inspect().document.source??entry.record.source;
  }
  sessionIdentity(index:number):{handle:number;parked:boolean;destroyed:boolean}|null{
    const entry=this.entries[index];if(!entry)throw new Error('Formula index outside virtual document');
    return entry.owner?{handle:entry.owner.handle,parked:entry.owner.isParked,destroyed:entry.owner.isDestroyed}:null;
  }
  shell(index:number):HTMLElement{
    const entry=this.entries[index];if(!entry)throw new Error('Formula index outside virtual document');return entry.shell;
  }
  focus(index:number):boolean{
    const entry=this.entries[index];if(!entry)throw new Error('Formula index outside virtual document');
    if(!entry.host)return false;entry.host.focus();return true;
  }
  get stats(){
    return {formulas:this.entries.length,mounted:this.entries.filter(e=>e.host).length,
      materialized:this.entries.filter(e=>e.owner).length,parked:this.entries.filter(e=>e.owner?.isParked).length,
      inputDom:this.element.querySelectorAll('.vt-core-input').length,svgDom:this.element.querySelectorAll('.vt-core-svg').length};
  }
  dispose():void{
    if(this.disposed)return;this.disposed=true;this.observer.disconnect();this.abort.abort();
    for(const entry of this.entries){
      entry.host?.dispose();entry.host=null;
      try{entry.owner?.destroy();}catch(error){this.report(entry,error);}
      entry.owner=null;entry.shell.remove();
    }
    this.element.remove();
  }
}
