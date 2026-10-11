import {SvgSceneView} from './svg-view.js';
import {rgbHex} from './paint.js';
import {nearestGrid} from './grid-context.js';
import {BrowserInputBridge} from './input-bridge.js';
import {CommandInputState,type CommandDescriptor,type CommandInputSnapshot,type CommandTarget,type CommandUsage} from './command-input.js';
import {EngineEditorSessionOwner,type EngineRuntimePort,type OcrInsertionTicket,type OcrInsertionResult} from './editor-session.js';
import type {ViewPointer} from './svg-view';
import type {EditCommand,EditorAction,EditorActionResult,EditorSnapshot,Position,Selection} from './editing';
import type {LayoutOptions,LayoutSnapshot,SceneQuery,SceneQueryResult,MathStyle} from './layout';
import type {LimitMode} from './syntax';

export type {EngineRuntimePort} from './editor-session.js';
export interface CoreEditorSelectionToken {
  schema:'visualtex-core-selection-v1';
  ownerId:string;
  documentId:string;
  revision:number;
  sourceId:number;
  selection:Selection;
}
export interface EditorInputBehavior {
  autoPairDelimiters:boolean;
  autoExitSuperscript:boolean;
  autoExitSubscript:boolean;
  autoExitAccent:boolean;
  autoExitWrapperCommand:boolean;
}
export interface EditorHostOptions {
  documentId:string;
  initialSource?:string;
  symbols?:readonly unknown[];
  layout:LayoutOptions;
  onChange?:(snapshot:EditorSnapshot)=>void;
  /** Read-only visual geometry notification. Product shells may use this to
   * preserve their historical row sizing without taking ownership of layout. */
  onLayout?:(snapshot:LayoutSnapshot)=>void;
  onError?:(error:Error)=>void;
  onEnter?:()=>void;
  commands?:readonly CommandDescriptor[];
  commandUsage?:Readonly<Record<string,CommandUsage>>;
  commandLimit?:number;
  /** Product hosts may render their own rich candidate surface while this host
   * remains the sole owner of transient command state and confirmation. */
  externalCommandPopup?:boolean;
  onCommandState?:(snapshot:CommandInputSnapshot)=>void;
  onCommandAccepted?:(id:string,query:string)=>void;
  inputBehavior?:EditorInputBehavior;
  inlineShortcuts?:Readonly<Record<string,string>>;
}
let nextSelectionOwner=1;
const caret=(p:Position):Selection=>({kind:'range',anchor:p,focus:p});
const samePosition=(a:Position,b:Position)=>a.kind===b.kind&&a.nodeId===b.nodeId&&(a.kind==='row'&&b.kind==='row'?a.offset===b.offset:a.kind==='text'&&b.kind==='text'&&a.grapheme===b.grapheme);
const roles:Record<string,string>={root:'公式',fraction:'分式',numerator:'分子',denominator:'分母',radical:'根式',degree:'根指数',radicand:'被开方项',fence:'定界符',fenceBody:'括号内部',
  accent:'重音或装饰',accentBody:'重音正文',annotation:'上下标注',annotationLabel:'标注内容',annotationBase:'标注正文',
  largeOperator:'大运算符',generalOperator:'自定义算子',customSymbol:'自定义符号',operatorBody:'算子本体',macroCall:'兼容宏',
  'macroArgument:1':'宏参数 1','macroArgument:2':'宏参数 2','macroArgument:3':'宏参数 3',lowerLimit:'下限',upperLimit:'上限',
  space:'数学空白',phantom:'隐藏占位',phantomBody:'隐藏内容',boxed:'方框',boxedBody:'方框内容',
  color:'着色公式',colorBody:'着色内容',colorChange:'颜色声明',styleChange:'样式声明',style:'数学样式',styleBody:'样式内容',
  grid:'框内多行公式',gridRow:'公式行',middleDelimiter:'中间定界符',
  scripts:'上下标',base:'底数',subscript:'下标',superscript:'上标',text:'文本',symbol:'符号',group:'分组'};

/** A per-document owner bridge, not a second editable state store. All mutations
 * and history are core transactions. The cached snapshot is a read projection. */
export class EngineEditorHost {
  readonly view:SvgSceneView;
  readonly input:BrowserInputBridge;
  readonly status:HTMLDivElement;
  readonly commandPopup:HTMLDivElement;
  private scene:LayoutSnapshot|null=null;
  private sceneToken=0;
  private typingGroup=1;
  private readonly owner:EngineEditorSessionOwner;
  private readonly ownsOwner:boolean;
  private commandInput:CommandInputState|null;
  private dragAnchor:Position|null=null;
  private pendingMove:ViewPointer|null=null;
  private moveFrame:number|null=null;
  private desiredX:number|null=null;
  private extendVertical=false;
  private zoom:number|null=null;
  private focused=false;
  private disposed=false;
  private readonly abort=new AbortController();
  private readonly selectionOwnerId=`vt-core-owner-${nextSelectionOwner++}`;
  private errors:string[]=[];
  private recentMathInput='';
  constructor(runtime:EngineRuntimePort,parent:HTMLElement,private readonly options:EditorHostOptions,owner?:EngineEditorSessionOwner){
    if(owner&&!owner.belongsTo(runtime))throw new Error('Editor session owner belongs to another runtime');
    if(owner&&options.symbols?.length)throw new Error('An attached owner already has an immutable custom-symbol catalog');
    this.owner=owner??new EngineEditorSessionOwner(runtime,options.documentId,options.initialSource??'',options.symbols??[]);
    this.ownsOwner=!owner;
    this.commandInput=options.commands?.length?new CommandInputState(options.commands,options.commandUsage??{},options.commandLimit??8):null;
    if(this.owner.documentId!==options.documentId)throw new Error('Editor session owner documentId mismatch');
    this.owner.resume();
    this.view=new SvgSceneView(parent,{onPointer:event=>this.pointer(event),showEmptySlots:true});
    this.status=document.createElement('div');this.status.className='vt-core-position';
    this.status.setAttribute('role','status');this.status.setAttribute('aria-live','polite');parent.append(this.status);
    this.commandPopup=document.createElement('div');this.commandPopup.className='vt-core-command-popup';
    this.commandPopup.setAttribute('role','listbox');this.commandPopup.setAttribute('aria-label','公式命令候选');this.commandPopup.hidden=true;parent.append(this.commandPopup);
    this.input=new BrowserInputBridge(parent,{
      key:event=>this.key(event),
      insert:text=>this.guardKeystroke(text,()=>{
        const mode=this.inputMode();
        if(this.commandInput?.active){
          if(this.commandInput.append(text,this.commandTarget())){this.renderCommandInput();return true;}
          this.cancelCommandInput();
        }
        if(mode==='math'&&text==='\\'&&this.commandInput&&this.state.selection.kind==='range'&&samePosition(this.state.selection.anchor,this.state.selection.focus)){
          this.owner.cancelPendingOcr();this.commandInput.begin(this.commandTarget());this.renderCommandInput();return true;
        }
        if(mode==='math'&&(text==='^'||text==='_')) return this.commitAndRender({kind:'edit',command:{kind:'insertStructure',structure:text==='^'?'superscript':'subscript'}});
        if(mode==='math'&&['(',')','[',']','{','}'].includes(text)) {
          this.typingGroup++;
          if(this.options.inputBehavior?.autoPairDelimiters===false){
            const path=this.selectionPath();
            const committed=this.commitAndRender({kind:'edit',command:{kind:'insertText',text,mode:'math'}});
            this.autoExitAfterAtomicInput(path);return committed;
          }
          return this.commitAndRender({kind:'edit',command:{kind:'typeDelimiter',text}});
        }
        const path=this.selectionPath();
        const groupId=`typing-${this.typingGroup}-${mode}`;
        const committed=this.commitAndRender({kind:'edit',command:{kind:'insertText',text,mode},groupId});
        if(mode==='math'){
          this.applyInlineShortcut(text,groupId);
          this.autoExitAfterAtomicInput(path);
        }else this.recentMathInput='';
        return committed;
      }),
      remove:backward=>{this.typingGroup++;this.recentMathInput='';return this.commitAndRender({kind:'edit',command:{kind:backward?'deleteBackward':'deleteForward'}});},
      history:redo=>{this.typingGroup++;this.recentMathInput='';return this.commitAndRender({kind:redo?'redo':'undo'});},
      beginComposition:()=>{this.cancelCommandInput();this.typingGroup++;this.recentMathInput='';const result=this.dispatch({kind:'beginComposition',mode:this.inputMode()});this.paintResult(result);return result.editor.composition?.id??null;},
      updateComposition:(id,preedit)=>this.commitAndRender({kind:'updateComposition',id,preedit}),
      commitComposition:(id,text)=>{this.typingGroup++;return this.commitAndRender({kind:'commitComposition',id,text});},
      cancelComposition:id=>{this.typingGroup++;return this.commitAndRender({kind:'cancelComposition',id});},
      copy:()=>this.owner.copySelection(),
      cut:fragment=>{
        if(fragment.revision!==this.state.document.revision||fragment.selectionEpoch!==this.state.selectionEpoch)throw new Error('Clipboard target is stale');
        this.typingGroup++;return this.commitAndRender({kind:'edit',command:{kind:'deleteBackward'}});
      },
      paste:(text,latex,structure)=>{
        this.typingGroup++;this.recentMathInput='';
        if(this.inputMode()==='text')return this.commitAndRender({kind:'edit',command:{kind:'insertText',text,mode:'text'}});
        const source=latex??text;
        return this.commitAndRender({kind:'edit',command:structure?{kind:'pasteStructured',latex:source,fragment:structure}:{kind:'pasteLatex',latex:source}});
      },
      lineBreak:()=>{this.typingGroup++;this.recentMathInput='';return this.commitAndRender({kind:'edit',command:{kind:'splitLine'}});},
      paragraph:()=>this.paragraph(),
      focus:focused=>{this.focused=focused;if(!focused){this.typingGroup++;this.recentMathInput='';this.cancelCommandInput();}try{this.view.setCaretVisible(focused);}catch{/* layout may not exist yet */}},error:error=>this.reportError(error),
    });
    const position=()=>{try{this.positionInput();}catch(error){this.reportError(error);}};
    window.addEventListener('resize',position,{signal:this.abort.signal});
    window.addEventListener('scroll',position,{signal:this.abort.signal,capture:true});
    // Unsupported or unfinished source remains a live, recoverable session.
    // Failure to draw must not destroy the user's source or its history.
    try{this.refresh(true);}catch(error){this.reportError(error);}
  }
  private get state():EditorSnapshot{return this.owner.current();}
  private commandTarget():CommandTarget{return {documentId:this.state.document.documentId,revision:this.state.document.revision,selectionEpoch:this.state.selectionEpoch,selection:structuredClone(this.state.selection)};}
  inspectCommandInput():CommandInputSnapshot|null{return this.commandInput?.snapshot()??null;}
  setInputBehavior(inputBehavior:EditorInputBehavior):void{this.options.inputBehavior={...inputBehavior};}
  setInlineShortcuts(inlineShortcuts:Readonly<Record<string,string>>):void{this.options.inlineShortcuts={...inlineShortcuts};this.recentMathInput='';}
  setCommandConfiguration(commands:readonly CommandDescriptor[],usage:Readonly<Record<string,CommandUsage>>={},limit=8):void{
    if(this.commandInput?.active)this.commandInput.cancel();
    this.commandInput=commands.length?new CommandInputState(commands,usage,limit):null;
    this.renderCommandInput();
  }
  selectCommandCandidate(index:number):boolean {
    if(!this.commandInput?.active)return false;
    const selected=this.commandInput.select(index,this.commandTarget());
    if(selected)this.renderCommandInput();
    return selected;
  }
  confirmCommandCandidate(index?:number):boolean {
    if(!this.commandInput?.active)return false;
    if(index!==undefined&&!this.commandInput.select(index,this.commandTarget()))return false;
    return this.confirmCommandInput();
  }
  private cancelCommandInput():void {
    if(!this.commandInput?.active)return;
    this.commandInput.cancel();this.renderCommandInput();
  }
  private renderCommandInput():void {
    if(!this.commandInput){this.commandPopup.hidden=true;return;}
    const snapshot=this.commandInput.snapshot();this.options.onCommandState?.(structuredClone(snapshot));
    this.commandPopup.replaceChildren();
    if(!snapshot.active||this.options.externalCommandPopup){this.commandPopup.hidden=true;return;}
    const title=document.createElement('div');title.className='vt-core-command-query';title.textContent=`\\${snapshot.query}`;this.commandPopup.append(title);
    const list=document.createElement('div');list.className='vt-core-command-candidates';
    for(const [index,candidate] of snapshot.candidates.entries()){
      const item=document.createElement('div');item.className='vt-core-command-candidate';item.setAttribute('role','option');item.setAttribute('aria-selected',String(index===snapshot.selectedIndex));item.dataset.commandId=candidate.id;item.textContent=candidate.label;list.append(item);
    }
    if(!snapshot.candidates.length){const empty=document.createElement('div');empty.className='vt-core-command-empty';empty.textContent='No matching command';list.append(empty);}
    this.commandPopup.append(list);this.commandPopup.hidden=false;
    const anchor=this.input.textarea.getBoundingClientRect();Object.assign(this.commandPopup.style,{position:'fixed',left:`${Math.max(8,anchor.left)}px`,top:`${Math.min(window.innerHeight-40,anchor.bottom+4)}px`,zIndex:'8'});
  }
  private confirmCommandInput():boolean {
    if(!this.commandInput?.active)return false;
    const confirmation=this.commandInput.confirm(this.commandTarget());this.renderCommandInput();
    if(!confirmation)return true;
    this.typingGroup++;const result=this.dispatch({kind:'edit',command:confirmation.action});this.paintResult(result);this.options.onCommandAccepted?.(confirmation.id,confirmation.query);return true;
  }
  private commandKey(event:KeyboardEvent):boolean|null {
    if(!this.commandInput?.active)return null;
    if(event.key==='Escape'){this.cancelCommandInput();return true;}
    if(event.key==='Backspace'){this.commandInput.backspace(this.commandTarget());this.renderCommandInput();return true;}
    if(event.key==='ArrowUp'||event.key==='ArrowDown'){this.commandInput.move(event.key==='ArrowUp'?-1:1,this.commandTarget());this.renderCommandInput();return true;}
    if(event.key==='Enter'||event.key==='Tab'||event.key===' '){return this.confirmCommandInput();}
    if(event.key.length===1&&/^[A-Za-z0-9-]$/u.test(event.key))return false;
    if(!event.isComposing)this.cancelCommandInput();
    return null;
  }
  private reportError(reason:unknown){
    const error=reason instanceof Error?reason:new Error(String(reason));this.errors.push(error.message);
    if(this.errors.length>128)this.errors.shift();
    this.status.textContent=error.message;this.options.onError?.(error);
  }
  inspect():EditorSnapshot{return structuredClone(this.state);}
  /** The current source without cloning the whole editor snapshot. */
  get source():string{return this.state.document.source;}
  captureSelectionToken():CoreEditorSelectionToken{return {schema:'visualtex-core-selection-v1',ownerId:this.selectionOwnerId,documentId:this.state.document.documentId,revision:this.state.document.revision,sourceId:this.state.document.sourceId,selection:structuredClone(this.state.selection)};}
  restoreSelectionToken(token:CoreEditorSelectionToken):boolean{
    if(token.schema!=='visualtex-core-selection-v1'||token.ownerId!==this.selectionOwnerId||token.documentId!==this.state.document.documentId||token.revision!==this.state.document.revision||token.sourceId!==this.state.document.sourceId)return false;
    const result=this.dispatch({kind:'setSelection',selection:structuredClone(token.selection)});this.refresh(result.contentChanged);this.input.focus();return true;
  }
  get layoutSnapshot():LayoutSnapshot|null{return this.scene?structuredClone(this.scene):null;}
  get diagnostics(){return this.errors.slice();}
  private query(query:SceneQuery):SceneQueryResult{
    if(!this.scene||this.scene.revision!==this.state.document.revision)throw new Error('Current layout is unavailable');
    return this.owner.query(this.sceneToken,query);
  }
  private rawAction(action:EditorAction):EditorActionResult{
    if(this.disposed)throw new Error('Editor host disposed');
    return this.owner.action(action);
  }
  private dispatch(action:EditorAction):EditorActionResult{
    if(this.commandInput?.active)this.cancelCommandInput();
    const result=this.rawAction(action);
    if(result.verticalMove){
      const extend=action.kind==='moveVertical'&&this.extendVertical;
      const down=result.verticalMove.direction==='down';
      let selection=result.verticalMove.selection;
      let from:Position|null=null;
      if(extend&&selection.kind==='range')from=selection.focus;
      else{
        const queried=this.query({kind:'selection',selection});
        if(queried.kind!=='selection')return result;
        from=queried.caret??(selection.kind==='gridCells'?{kind:'row' as const,nodeId:selection.focusCell,offset:0,affinity:'after' as const}:null);
        if(!from&&selection.kind==='range'){
          // A highlighted range first collapses to its upper/lower edge,
          // then the caret moves vertically from there.
          selection=this.rawAction({kind:'moveHorizontal',direction:down?'right':'left',extend:false}).editor.selection;
          if(selection.kind==='range')from=selection.focus;
        }
      }
      if(!from)return result;
      const target=this.query({kind:'vertical',from,down,desiredX:this.desiredX});
      if(target.kind!=='vertical')throw new Error('Invalid vertical navigation response');
      this.desiredX=target.desiredX;
      const anchor=extend&&selection.kind==='range'?selection.anchor:target.position;
      return this.rawAction({kind:'setSelection',selection:{kind:'range',anchor,focus:target.position}});
    }
    if(action.kind!=='updateComposition')this.desiredX=null;
    return result;
  }
  private inputMode():'math'|'text'{return this.state.selection.kind==='range'&&this.state.selection.focus.kind==='text'?'text':'math';}
  private selectionPath():string[]{
    try{const q=this.query({kind:'selection',selection:this.state.selection});return q.kind==='selection'?q.description.path:[];}catch{return [];}
  }
  private applyInlineShortcut(text:string,groupId:string):void{
    const shortcuts=this.options.inlineShortcuts??{};
    if(!Object.keys(shortcuts).length){this.recentMathInput='';return;}
    if(!/^[A-Za-z0-9<>=+\-*/]+$/u.test(text)){this.recentMathInput='';return;}
    this.recentMathInput=(this.recentMathInput+text).slice(-64);
    let matchedKey='';let replacement='';
    for(const [key,value] of Object.entries(shortcuts)){
      if(!key||key.length>64||!this.recentMathInput.endsWith(key)||key.length<=matchedKey.length)continue;
      const before=this.recentMathInput.slice(0,-key.length).at(-1)??'';
      if(/^[A-Za-z]/u.test(key)&&/[A-Za-z]/u.test(before))continue;
      matchedKey=key;replacement=value;
    }
    if(!matchedKey)return;
    try{
      const result=this.dispatch({kind:'edit',command:{kind:'replaceRecentMath',expected:matchedKey,latex:replacement},groupId});
      this.refresh(result.contentChanged);this.recentMathInput='';
    }catch{this.recentMathInput='';}
  }
  private autoExitAfterAtomicInput(path:string[]):void{
    const behavior=this.options.inputBehavior;if(!behavior)return;
    const should=(path.includes('superscript')&&behavior.autoExitSuperscript)
      ||(path.includes('subscript')&&behavior.autoExitSubscript)
      ||(path.includes('accentBody')&&behavior.autoExitAccent)
      ||(path.includes('styleBody')&&behavior.autoExitWrapperCommand);
    if(!should)return;
    try{const result=this.dispatch({kind:'edit',command:{kind:'exitStructure'}});this.refresh(result.contentChanged);}catch(error){this.reportError(error);}
  }
  private paintResult(result:EditorActionResult):void{
    // A rendering exception cannot uncommit an already accepted core action.
    try{this.refresh(result.contentChanged);}catch(error){this.reportError(error);}
  }
  private commitAndRender(action:EditorAction):boolean{const result=this.dispatch(action);this.paintResult(result);return true;}
  private key(event:KeyboardEvent):boolean{
    const commandResult=this.commandKey(event);if(commandResult!==null)return commandResult;
    // Word/line deletion (Ctrl/Alt/Cmd+Backspace/Delete) has no formula
    // meaning beyond one atom. Owning it here keeps the browser from sending
    // deleteWord*/deleteSoftLine* input the bridge cannot map.
    if((event.key==='Backspace'||event.key==='Delete')&&(event.ctrlKey||event.metaKey||event.altKey))return this.deleteKey(event.key==='Backspace');
    if((event.ctrlKey||event.metaKey)&&!event.altKey){
      const directions={ArrowLeft:'left',ArrowRight:'right',ArrowUp:'up',ArrowDown:'down'} as const;
      if(event.shiftKey&&event.key in directions){
        const selection=this.state.selection;
        const id=selection.kind==='gridCells'?selection.focusCell:selection.kind==='range'?selection.focus.nodeId:null;
        if(id===null||!nearestGrid(this.state.document.analysis.tree,id))return false;
        this.typingGroup++;
        try{this.commitAndRender({kind:'extendGridSelection',direction:directions[event.key as keyof typeof directions]});}
        catch(error){this.reportError(error);}
        return true;
      }
      if(event.key==='Home'||event.key==='End')return this.moveBoundary(event.key==='End',true,event.shiftKey);
      const key=event.key.toLowerCase();
      if(key==='z'||key==='y'){this.typingGroup++;return this.commitAndRender({kind:key==='y'||event.shiftKey?'redo':'undo'});}
      if(key==='a'){
        const root=this.state.document.analysis.tree;if(root.kind!=='row')return false;this.typingGroup++;
        return this.commitAndRender({kind:'setSelection',selection:{kind:'range',anchor:{kind:'row',nodeId:root.id,offset:0,affinity:'after'},focus:{kind:'row',nodeId:root.id,offset:root.entries.length,affinity:'after'}}});
      }
      return false;
    }
    if(event.ctrlKey||event.metaKey||event.altKey)return false;
    if(event.key==='Escape'&&this.state.selection.kind==='gridCells'){
      return this.commitAndRender({kind:'setSelection',selection:caret({kind:'row',nodeId:this.state.selection.focusCell,offset:0,affinity:'after'})});
    }
    if(event.key==='Tab'){
      if(this.state.selection.kind==='gridCells')this.dispatch({kind:'setSelection',selection:caret({kind:'row',nodeId:this.state.selection.focusCell,offset:0,affinity:'after'})});
      const selection=this.state.selection;
      if(selection.kind!=='range'||selection.anchor.nodeId!==selection.focus.nodeId)return false;
      const environment=nearestGrid(this.state.document.analysis.tree,selection.focus.nodeId)?.environment;
      if(!environment||environment==='gathered')return false;
      this.typingGroup++;
      try{this.commitAndRender({kind:'edit',command:{kind:'moveGridCell',direction:event.shiftKey?'left':'right'}});}
      catch(error){this.reportError(error);}
      return true;
    }
    if(event.key==='Home'||event.key==='End')return this.moveBoundary(event.key==='End',false,event.shiftKey);
    if(event.key==='ArrowLeft'||event.key==='ArrowRight'){this.typingGroup++;return this.commitAndRender({kind:'moveHorizontal',direction:event.key==='ArrowLeft'?'left':'right',extend:event.shiftKey});}
    if(event.key==='ArrowUp'||event.key==='ArrowDown'){
      this.typingGroup++;this.extendVertical=event.shiftKey;
      try{return this.commitAndRender({kind:'moveVertical',direction:event.key==='ArrowUp'?'up':'down'});}
      finally{this.extendVertical=false;}
    }
    if(event.key==='Backspace'||event.key==='Delete')return this.deleteKey(event.key==='Backspace');
    if(event.key==='Enter'){
      this.typingGroup++;
      if(event.shiftKey){
        // A contained row break never reaches the outer Enter/Apply callback.
        // Consume even a refused interior split: native textarea newline input
        // must not mutate the formula after an atomic core rejection.
        try{this.commitAndRender({kind:'edit',command:{kind:'splitLine'}});}
        catch(error){this.reportError(error);}
        return true;
      }
      try{this.paragraph();}catch(error){this.reportError(error);}
      return true;
    }
    return false;
  }
  private guardKeystroke(text:string,insert:()=>boolean):boolean{
    try{return insert();}
    catch(error){
      // A single refused keystroke has nothing worth preserving: report it and
      // stay ready for the next key instead of holding a pending edit that
      // swallows all further typing until Enter/Escape. Longer text (IME or
      // dictation commits) is still held for retry.
      if([...text].length>1)throw error;
      this.recentMathInput='';this.reportError(error);return true;
    }
  }
  private deleteKey(backward:boolean):boolean{
    this.typingGroup++;this.recentMathInput='';
    // This key is owned even when the core rejects a destructive range.
    // Do not fall through to beforeinput, retry the deletion and hold an
    // empty native pending edit after the already-reported refusal.
    try{this.commitAndRender({kind:'edit',command:{kind:backward?'deleteBackward':'deleteForward'}});}
    catch(error){this.reportError(error);}
    return true;
  }
  private paragraph():boolean{
    this.typingGroup++;
    const selection=this.state.selection;
    const id=selection.kind==='gridCells'?selection.focusCell:selection.kind==='range'?selection.focus.nodeId:null;
    if(id!==null&&nearestGrid(this.state.document.analysis.tree,id)?.environment==='cases'){
      return this.commitAndRender({kind:'edit',command:{kind:'splitLine'}});
    }
    this.options.onEnter?.();return true;
  }
  private moveBoundary(toEnd:boolean,document:boolean,extend:boolean):boolean{
    const query=this.query({kind:'boundary',selection:this.state.selection,toEnd,document});
    if(query.kind!=='boundary')throw new Error('Invalid boundary navigation response');
    this.typingGroup++;
    const anchor=extend&&this.state.selection.kind==='range'?this.state.selection.anchor:query.position;
    return this.commitAndRender({kind:'setSelection',selection:{kind:'range',anchor,focus:query.position}});
  }
  private gridEdit(command:EditCommand):void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm or cancel pending input before a matrix operation');
    this.typingGroup++;
    const result=this.dispatch({kind:'edit',command});this.refresh(result.contentChanged);this.input.focus();
  }
  applyMatrix(rows:number,columns:number):void{this.gridEdit({kind:'insertMatrix',rows,columns});}
  applyCases(rows=2):void{this.gridEdit({kind:'insertCases',rows});}
  applyAligned(rows=2,pairs=1):void{this.gridEdit({kind:'insertAligned',rows,pairs});}
  insertGridRow(after=true):void{this.gridEdit({kind:'insertGridRow',after});}
  deleteGridRow():void{this.gridEdit({kind:'deleteGridRow'});}
  insertGridColumn(after=true):void{this.gridEdit({kind:'insertGridColumn',after});}
  deleteGridColumn():void{this.gridEdit({kind:'deleteGridColumn'});}
  clearGridCells():void{this.gridEdit({kind:'clearGridCells'});}
  pasteGridCells(latex:string):void{this.gridEdit({kind:'pasteGridCells',latex});}
  splitLine():void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm or cancel pending input before a contained line break');
    this.typingGroup++;
    const result=this.dispatch({kind:'edit',command:{kind:'splitLine'}});this.refresh(result.contentChanged);this.input.focus();
  }
  private refresh(contentChanged:boolean){
    // Publish accepted core state even when its new source cannot be laid out.
    // The source pane and undo stack must never remain on a previous revision.
    this.options.onChange?.(this.inspect());
    try{
      if(contentChanged||!this.scene||this.scene.revision!==this.state.document.revision){
        const layout=this.owner.layout(this.options.layout,false);
        this.view.setSnapshot(layout.snapshot);this.sceneToken=layout.sceneToken;this.scene=layout.snapshot;
        this.options.onLayout?.(structuredClone(layout.snapshot));
      }
      const overlay=this.query({kind:'selection',selection:this.state.selection});
      if(overlay.kind!=='selection')throw new Error('Invalid selection geometry response');
      this.view.container.hidden=false;this.view.container.removeAttribute('aria-disabled');
      this.input.setEnabled(true);
      this.view.setOverlay({...this.view.snapshotIdentity!,caret:overlay.caret,caretGeometry:overlay.caretGeometry,rectangles:overlay.rectangles});
      this.view.setCaretVisible(this.focused);
      const d=overlay.description;
      const label=`${d.path.map(p=>/^gridRow:\d+$/u.test(p)?`第 ${p.split(':')[1]} 行`:/^gridCell:\d+$/u.test(p)?`第 ${p.split(':')[1]} 列`:roles[p]??p).join('，')}，位置 ${d.offset}/${d.length}${d.text?`，${d.text}`:''}`;
      this.status.textContent=label;this.input.textarea.setAttribute('aria-label',label);
      this.input.textarea.setAttribute('aria-multiline',String(this.state.document.analysis.facts.grids>0));
      this.input.setPreedit(this.state.composition?.preedit??'');
      this.positionInput(overlay.caret);
    }catch(error){
      this.scene=null;this.sceneToken=0;this.dragAnchor=null;
      this.view.container.hidden=true;this.view.container.setAttribute('aria-disabled','true');
      const blankEditable=this.state.document.source.trim()===''&&this.state.document.analysis.complete&&this.state.document.analysis.supported;
      this.input.setEnabled(blankEditable);
      this.input.setPreedit(this.state.composition?.preedit??'');
      this.status.textContent=blankEditable
        ?'空白公式可直接输入；首个数学内容出现后建立矢量布局。'
        :'当前源码尚不能排版；源码和撤销历史已保留，可在源码区修复。';
      if(!blankEditable)throw error;
    }
  }
  private positionInput(position?:Position|null){
    if(!this.scene)return;
    if(position===undefined){const q=this.query({kind:'selection',selection:this.state.selection});position=q.kind==='selection'?q.caret:null;}
    const c=position?this.scene.carets.find(c=>{
      const a=c.position,b=position!;return a.kind===b.kind&&a.nodeId===b.nodeId&&(a.kind==='row'&&b.kind==='row'?a.offset===b.offset:a.kind==='text'&&b.kind==='text'&&a.grapheme===b.grapheme);
    }):this.scene.carets[0];
    if(!c)return;
    const top=this.view.layoutToClient(c.x,c.top),bottom=this.view.layoutToClient(c.x,c.bottom);
    this.input.setAnchor(top.x,top.y,Math.max(1,bottom.y-top.y));
  }
  private pointer(event:ViewPointer){
    // A drag produces far more pointermove events than frames; hit-test and
    // re-select once per frame with the latest point. down/up stay immediate
    // and supersede any queued move.
    if(event.phase==='move'){
      const first=this.pendingMove===null;this.pendingMove=event;
      if(first)this.moveFrame=requestAnimationFrame(()=>{this.moveFrame=null;const move=this.pendingMove;this.pendingMove=null;if(move)this.applyPointer(move);});
      return;
    }
    this.cancelPendingMove();this.applyPointer(event);
  }
  private cancelPendingMove(){
    if(this.moveFrame!==null)cancelAnimationFrame(this.moveFrame);
    this.moveFrame=null;this.pendingMove=null;
  }
  private applyPointer(event:ViewPointer){
    if(this.disposed||this.input.hasUncommittedInput)return;
    if(!this.scene||event.documentId!==this.scene.documentId||event.revision!==this.scene.revision||event.sourceId!==this.scene.sourceId)return;
    try{
      const query=this.query({kind:'hitTest',xEm:event.xEm,yEm:event.yEm});if(query.kind!=='hitTest')return;
      if(event.phase==='double'){
        if(event.shiftKey)return;
        // Double click selects the slot (or text run) under the pointer.
        const at=caret(query.position);
        const start=this.query({kind:'boundary',selection:at,toEnd:false,document:false});
        const end=this.query({kind:'boundary',selection:at,toEnd:true,document:false});
        if(start.kind!=='boundary'||end.kind!=='boundary')return;
        this.dragAnchor=null;this.typingGroup++;
        this.dispatch({kind:'setSelection',selection:{kind:'range',anchor:start.position,focus:end.position}});this.refresh(false);this.input.focus();return;
      }
      if(event.phase==='down')this.dragAnchor=event.shiftKey&&this.state.selection.kind==='range'?this.state.selection.anchor:
        event.shiftKey&&this.state.selection.kind==='gridCells'?{kind:'row',nodeId:this.state.selection.anchorCell,offset:0,affinity:'after'}:query.position;
      if(!this.dragAnchor)return;
      this.dispatch({kind:'pointerSelection',anchor:this.dragAnchor,focus:query.position});this.refresh(false);
      if(event.phase==='up')this.dragAnchor=null;
      this.input.focus();
    }catch(error){this.reportError(error);}
  }
  beginOcrInsertion():OcrInsertionTicket{
    if(this.disposed)throw new Error('Editor host disposed');
    if(this.input.hasUncommittedInput||this.commandInput?.active)throw new Error('Finish pending input before capturing an OCR target');
    return this.owner.captureOcrInsertion();
  }
  cancelOcrInsertion(ticket:OcrInsertionTicket):boolean{return this.owner.cancelOcrInsertion(ticket);}
  applyOcrResult(ticket:OcrInsertionTicket,latex:string):OcrInsertionResult{
    if(this.disposed||this.input.hasUncommittedInput||this.commandInput?.active){
      this.owner.cancelOcrInsertion(ticket);
      return {status:'rejected',latex,reason:'OCR target view is disposed or has uncommitted input'};
    }
    const outcome=this.owner.completeOcrInsertion(ticket,latex);
    if(outcome.status==='applied'){this.typingGroup++;this.desiredX=null;this.paintResult(outcome.result);}
    // No focus(), current-editor lookup, source normalization or retry here.
    return outcome;
  }
  applyEditCommand(command:EditCommand):void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before applying an edit command');
    this.typingGroup++;
    const result=this.dispatch({kind:'edit',command:structuredClone(command)});this.refresh(result.contentChanged);this.input.focus();
  }
  applyStructure(structure:'fraction'|'binomial'|'skewedFraction'|'superscript'|'subscript'|'text'|'squareRoot'|'indexedRoot'):void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before inserting a structure');
    const result=this.dispatch({kind:'edit',command:{kind:'insertStructure',structure}});this.refresh(result.contentChanged);this.input.focus();
  }
  applyCustomSymbol(command:string):void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before inserting a custom symbol');
    this.typingGroup++;
    const result=this.dispatch({kind:'edit',command:{kind:'insertCustomSymbol',command}});this.refresh(result.contentChanged);this.input.focus();
  }
  applySymbol(command:string):void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before inserting a symbol');
    this.typingGroup++;
    const result=this.dispatch({kind:'edit',command:{kind:'insertSymbol',command}});
    this.refresh(result.contentChanged);this.input.focus();
  }
  applyFence(left:string,right:string):void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before inserting a fence');
    this.typingGroup++;
    const result=this.dispatch({kind:'edit',command:{kind:'insertFence',left,right}});this.refresh(result.contentChanged);this.input.focus();
  }
  applyOperator(command:string,withLower=true,withUpper=true):void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before inserting an operator');
    this.typingGroup++;
    const result=this.dispatch({kind:'edit',command:{kind:'insertOperator',command,withLower,withUpper}});this.refresh(result.contentChanged);this.input.focus();
  }
  applyGeneralOperator(withLower=false,withUpper=false):void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before inserting an operator body');
    this.typingGroup++;
    const result=this.dispatch({kind:'edit',command:{kind:'insertGeneralOperator',withLower,withUpper}});this.refresh(result.contentChanged);this.input.focus();
  }
  setOperatorLimits(mode:LimitMode):void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before changing limits');
    this.typingGroup++;
    const result=this.dispatch({kind:'edit',command:{kind:'setOperatorLimits',mode}});this.refresh(result.contentChanged);this.input.focus();
  }
  setLayoutOptions(layout:LayoutOptions):void{
    // Identical options are a no-op: a full relayout per call is the costly
    // part, and callers may re-apply the same resolved fonts on every render.
    if(JSON.stringify(layout)===JSON.stringify(this.options.layout))return;
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before changing layout resources');
    const previous=this.options.layout;this.options.layout=structuredClone(layout);
    try{this.refresh(true);}catch(error){this.options.layout=previous;this.refresh(true);throw error;}
  }
  setMathStyle(style:MathStyle):void{
    if(!['display','text','script','scriptScript'].includes(style))throw new Error('Invalid mathematical style');
    if(this.options.layout.style===style)return;
    this.setLayoutOptions({...this.options.layout,style});
  }
  applyAccent(command:string):void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before inserting an accent');
    this.typingGroup++;
    const result=this.dispatch({kind:'edit',command:{kind:'insertAccent',command}});this.refresh(result.contentChanged);this.input.focus();
  }
  applyAnnotation(command:string):void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before inserting an annotation');
    this.typingGroup++;
    const result=this.dispatch({kind:'edit',command:{kind:'insertAnnotation',command}});this.refresh(result.contentChanged);this.input.focus();
  }
  applyStyle(command:string):void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before applying a mathematical style');
    this.typingGroup++;
    const result=this.dispatch({kind:'edit',command:{kind:'setStyle',style:command}});this.refresh(result.contentChanged);this.input.focus();
  }
  applySpace(command:string):void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before inserting math space');
    this.typingGroup++;const result=this.dispatch({kind:'edit',command:{kind:'insertSpace',command}});
    this.refresh(result.contentChanged);this.input.focus();
  }
  applyPhantom(command:string):void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before inserting phantom');
    this.typingGroup++;const result=this.dispatch({kind:'edit',command:{kind:'insertPhantom',command}});
    this.refresh(result.contentChanged);this.input.focus();
  }
  applyColor(color:import('./layout').RgbColor):void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before applying color');
    rgbHex(color);this.typingGroup++;
    const result=this.dispatch({kind:'edit',command:{kind:'setColor',color}});this.refresh(result.contentChanged);this.input.focus();
  }
  applyBoxed():void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before inserting a box');
    this.typingGroup++;const result=this.dispatch({kind:'edit',command:{kind:'insertBoxed'}});
    this.refresh(result.contentChanged);this.input.focus();
  }
  exitStructure():void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before leaving a structure');
    this.typingGroup++;
    const result=this.dispatch({kind:'edit',command:{kind:'exitStructure'}});this.refresh(result.contentChanged);this.input.focus();
  }
  replaceSource(source:string):void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before replacing source');
    this.cancelCommandInput();this.owner.replaceSource(source);this.refresh(true);
  }
  history(redo:boolean):void{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before changing history');
    this.typingGroup++;const result=this.dispatch({kind:redo?'redo':'undo'});this.refresh(result.contentChanged);
  }
  setZoom(zoom:number):void{
    if(zoom===this.zoom)return;
    this.view.setZoom(zoom);this.zoom=zoom;this.positionInput();
    if(this.scene)this.options.onLayout?.(structuredClone(this.scene));
  }
  focus():void{this.input.focus();}
  exportSvg():string{
    if(this.input.hasUncommittedInput)throw new Error('Confirm, retry or cancel pending input before export');
    const result=this.owner.layout(this.options.layout,true);
    // Export also replaces the runtime's scene token; do not keep a stale hit map.
    this.sceneToken=result.sceneToken;this.scene=result.snapshot;
    if(result.svg===null)throw new Error('Missing requested vector export');return result.svg;
  }
  dispose():void{
    if(this.disposed)return;
    this.cancelPendingMove();this.input?.dispose();this.abort.abort();
    try{this.owner.cancelPendingOcr();if(this.ownsOwner)this.owner.destroy();}
    finally{this.disposed=true;this.view?.dispose();this.status?.remove();this.commandPopup?.remove();}
  }
}
