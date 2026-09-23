import type {EditorSnapshot} from './editing.js';
import type {SourceEdit,SourceChangeMap} from './source-edit.js';
import {EngineEditorSessionOwner} from './editor-session.js';

declare const sourceSyncBrand:unique symbol;
export interface SourceSyncOrigin {readonly [sourceSyncBrand]:true}
export interface SourceViewFrame {
  documentId:string;
  revision:number;
  sourceId:number;
  source:string;
  origin:SourceSyncOrigin;
}
export interface SourceViewEditResult {
  frame:SourceViewFrame;
  mapping:SourceChangeMap;
  editor:EditorSnapshot;
}

/** CodeMirror-facing source bridge. The core session remains the only content
 * and undo/redo owner. View transactions carry a base revision/source identity;
 * core-to-view synchronization carries an opaque origin object so update
 * listeners can suppress echo without comparing source strings heuristically. */
export class SourceEditorBridge {
  private readonly coreOrigins=new WeakSet<object>();
  private readonly viewTransactions=new WeakSet<object>();
  constructor(private readonly owner:EngineEditorSessionOwner){}
  private origin():SourceSyncOrigin{
    const value=Object.freeze({}) as SourceSyncOrigin;this.coreOrigins.add(value as object);return value;
  }
  frame():SourceViewFrame{
    const editor=this.owner.inspect(),document=editor.document;
    return {documentId:document.documentId,revision:document.revision,sourceId:document.sourceId,source:document.source,origin:this.origin()};
  }
  isCoreSyncOrigin(value:unknown):value is SourceSyncOrigin{return Boolean(value)&&typeof value==='object'&&this.coreOrigins.has(value as object);}
  /** A CodeMirror update annotated with a core origin is display synchronization,
   * not a new source transaction. All other edits must use applyViewEdit. */
  shouldDispatchViewChange(origin:unknown):boolean{return !this.isCoreSyncOrigin(origin);}
  applyViewEdit(base:Pick<SourceViewFrame,'documentId'|'revision'|'sourceId'>,edit:SourceEdit,transactionIdentity:object):SourceViewEditResult{
    if(!transactionIdentity||typeof transactionIdentity!=='object')throw new Error('Source view transaction identity must be an object');
    if(this.viewTransactions.has(transactionIdentity))throw new Error('Source view transaction was already consumed');
    const current=this.owner.inspect().document;
    if(base.documentId!==current.documentId||base.revision!==current.revision||base.sourceId!==current.sourceId)throw new Error('Source view transaction is stale');
    this.viewTransactions.add(transactionIdentity);
    const result=this.owner.patchSource(edit);const document=result.editor.document;
    return {editor:result.editor,mapping:result.mapping,frame:{documentId:document.documentId,revision:document.revision,sourceId:document.sourceId,source:document.source,origin:this.origin()}};
  }
  /** Source-pane undo/redo is delegated to the same core history as visual edits.
   * There is intentionally no second CodeMirror document-history owner here. */
  history(redo:boolean):SourceViewFrame{
    this.owner.action({kind:redo?'redo':'undo'});return this.frame();
  }
}
