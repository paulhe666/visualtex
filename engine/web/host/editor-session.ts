import type {ClipboardFragment,EngineCommand,EngineReply,EngineResponse} from './protocol';
import type {EditorAction,EditorActionResult,EditorSnapshot,OcrInsertionTarget} from './editing';
import type {LayoutOptions,LayoutSnapshot,SceneQuery,SceneQueryResult} from './layout';
import type {SourceEdit,SourceChangeMap} from './source-edit';

export interface EngineRuntimePort {dispatch(request:string):string}
declare const ocrTicketBrand:unique symbol;
/** Opaque object identity. Not a serializable session handle or OCR payload. */
export interface OcrInsertionTicket {readonly [ocrTicketBrand]:true}
export type OcrInsertionResult=
  | {status:'applied';latex:string;result:EditorActionResult}
  | {status:'rejected';latex:string;reason:string};

/** Persistent core owner for one formula. It contains no DOM. A visible
 * EngineEditorHost may attach/detach repeatedly while this owner preserves the
 * same handle, event sequence, document identities, selection and history. */
export class EngineEditorSessionOwner {
  readonly handle:number;
  private state:EditorSnapshot;
  private sequence=0;
  private parked=false;
  private destroyed=false;
  private pendingOcr:{ticket:OcrInsertionTicket;target:OcrInsertionTarget}|null=null;
  constructor(private readonly runtime:EngineRuntimePort,documentId:string,source='',symbols:readonly unknown[]=[]){
    const created=this.send(symbols.length?{op:'createWithSymbols',documentId,source,symbols:structuredClone([...symbols])}:{op:'create',documentId,source},'session');
    this.handle=created.handle;
    this.state=this.send({op:'inspectEditor',handle:this.handle},'editor').snapshot;
  }
  private send<K extends EngineReply['kind']>(command:EngineCommand,kind:K):Extract<EngineReply,{kind:K}>{
    const response=JSON.parse(this.runtime.dispatch(JSON.stringify({protocol:1,command}))) as EngineResponse;
    if(response.status==='error')throw new Error(`${response.error.code}: ${response.error.message}`);
    if(response.status!=='ok'||response.result.kind!==kind)throw new Error('Unexpected engine response');
    return response.result as Extract<EngineReply,{kind:K}>;
  }
  private active():void{
    if(this.destroyed)throw new Error('Editor session owner disposed');
    if(this.parked)throw new Error('Editor session owner is parked');
  }
  belongsTo(runtime:EngineRuntimePort):boolean{return runtime===this.runtime;}
  get documentId():string{return this.state.document.documentId;}
  get revision():number{return this.state.document.revision;}
  get isParked():boolean{return this.parked;}
  get isDestroyed():boolean{return this.destroyed;}
  inspect():EditorSnapshot{return structuredClone(this.state);}
  current():EditorSnapshot{return this.state;}
  action(action:EditorAction):EditorActionResult{
    this.active();
    const result=this.send({op:'editorAction',handle:this.handle,expectedRevision:this.state.document.revision,eventSequence:++this.sequence,action},'editorAction').result;
    this.state=result.editor;return result;
  }
  replaceSource(source:string):EditorSnapshot{
    this.active();
    this.cancelPendingOcr();
    this.send({op:'replaceSource',handle:this.handle,expectedRevision:this.state.document.revision,source},'document');
    this.state=this.send({op:'inspectEditor',handle:this.handle},'editor').snapshot;
    return this.inspect();
  }
  patchSource(edit:SourceEdit):{editor:EditorSnapshot;mapping:SourceChangeMap}{
    this.active();this.cancelPendingOcr();
    const reply=this.send({op:'patchSource',handle:this.handle,expectedRevision:this.state.document.revision,edit},'sourcePatched');
    this.state=this.send({op:'inspectEditor',handle:this.handle},'editor').snapshot;
    return {editor:this.inspect(),mapping:structuredClone(reply.mapping)};
  }
  captureOcrInsertion():OcrInsertionTicket{
    this.active();
    const reply=this.send({op:'captureOcrTarget',handle:this.handle,expectedRevision:this.state.document.revision},'ocrTarget');
    const ticket=Object.freeze({}) as OcrInsertionTicket;
    this.pendingOcr={ticket,target:structuredClone(reply.target)};return ticket;
  }
  cancelOcrInsertion(ticket:OcrInsertionTicket):boolean{
    if(!this.pendingOcr||this.pendingOcr.ticket!==ticket)return false;
    const {target}=this.pendingOcr;this.pendingOcr=null;
    if(this.parked||this.destroyed)return true;
    return this.send({op:'cancelOcrTarget',handle:this.handle,requestId:target.requestId},'ocrTargetCancelled').cancelled;
  }
  cancelPendingOcr():void{if(this.pendingOcr)this.cancelOcrInsertion(this.pendingOcr.ticket);}
  completeOcrInsertion(ticket:OcrInsertionTicket,latex:string):OcrInsertionResult{
    const pending=this.pendingOcr;
    if(!pending||pending.ticket!==ticket)return {status:'rejected',latex,reason:'OCR request expired, was cancelled, or belongs to another owner'};
    this.pendingOcr=null;
    try{
      this.active();
      const result=this.send({op:'editorAction',handle:this.handle,expectedRevision:pending.target.revision,eventSequence:++this.sequence,action:{kind:'applyOcrLatex',target:pending.target,latex}},'editorAction').result;
      this.state=result.editor;return {status:'applied',latex,result};
    }catch(error){return {status:'rejected',latex,reason:error instanceof Error?error.message:String(error)};}
  }
  copySelection():ClipboardFragment{
    this.active();
    return this.send({op:'copySelection',handle:this.handle,expectedRevision:this.state.document.revision},'clipboard').fragment;
  }
  layout(options:LayoutOptions,includeSvg:boolean):{sceneToken:number;snapshot:LayoutSnapshot;svg:string|null}{
    this.active();
    const result=this.send({op:'layout',handle:this.handle,expectedRevision:this.state.document.revision,options,includeSvg},'layout');
    return {sceneToken:result.sceneToken,snapshot:result.snapshot,svg:result.svg};
  }
  presentationMathMl(displayMode:boolean):{mathMl:string;alignedGridIds:number[]}{
    this.active();
    return this.send({op:'exportPresentationMathMl',handle:this.handle,expectedRevision:this.state.document.revision,displayMode},'presentationMathMl').result;
  }
  query(sceneToken:number,query:SceneQuery):SceneQueryResult{
    this.active();
    return this.send({op:'queryScene',handle:this.handle,expectedRevision:this.state.document.revision,sceneToken,query},'sceneQuery').result;
  }
  park():void{
    this.active();
    if(this.state.composition)throw new Error('Finish or cancel pending input before parking the editor session');
    this.send({op:'parkSession',handle:this.handle,expectedRevision:this.state.document.revision},'sessionParked');
    this.pendingOcr=null;this.parked=true;
  }
  resume():void{
    if(this.destroyed)throw new Error('Editor session owner disposed');
    if(!this.parked)return;
    this.send({op:'resumeSession',handle:this.handle,expectedRevision:this.state.document.revision},'sessionResumed');
    this.parked=false;
    this.state=this.send({op:'inspectEditor',handle:this.handle},'editor').snapshot;
  }
  destroy():void{
    if(this.destroyed)return;
    this.send({op:'destroy',handle:this.handle,expectedRevision:this.state.document.revision},'destroyed');
    this.pendingOcr=null;this.destroyed=true;this.parked=false;
  }
}
