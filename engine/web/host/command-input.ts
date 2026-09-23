import type {EditCommand,Selection} from './editing.js';

export interface CommandTarget {
  documentId:string;
  revision:number;
  selectionEpoch:number;
  selection:Selection;
}
export interface CommandUsage {
  useCount?:number;
  lastUsedAt?:number;
  acceptedPrefixes?:Record<string,number>;
  pinned?:boolean;
}
export interface CommandDescriptor {
  id:string;
  trigger:string;
  aliases?:readonly string[];
  label?:string;
  /** Product-facing command spelling retained only for candidate presentation. */
  displayCommand?:string;
  /** Renderer-independent LaTeX used by the product candidate preview. */
  previewLatex?:string;
  priority?:number;
  action:EditCommand;
}
export interface CommandCandidate {
  id:string;
  trigger:string;
  label:string;
  displayCommand:string;
  previewLatex:string|null;
  score:number;
  action:EditCommand;
}
export interface CommandInputSnapshot {
  active:boolean;
  query:string;
  selectedIndex:number;
  candidates:CommandCandidate[];
  target:CommandTarget|null;
}
export interface CommandConfirmation {
  id:string;
  query:string;
  action:EditCommand;
}

const clone=<T>(value:T):T=>structuredClone(value);
const normalize=(value:string)=>value.trim().replace(/^\\/u,'').toLocaleLowerCase();
function sameTarget(a:CommandTarget,b:CommandTarget):boolean {
  return a.documentId===b.documentId&&a.revision===b.revision&&a.selectionEpoch===b.selectionEpoch
    &&JSON.stringify(a.selection)===JSON.stringify(b.selection);
}
function validateDescriptor(descriptor:CommandDescriptor):void {
  if(!descriptor.id||descriptor.id.length>128||descriptor.id.trim()!==descriptor.id||/[\u0000-\u001f]/u.test(descriptor.id))throw new Error('Invalid command descriptor id');
  const trigger=normalize(descriptor.trigger);
  if(!/^[a-z][a-z0-9-]{0,63}$/u.test(trigger))throw new Error(`Invalid command trigger: ${descriptor.trigger}`);
  for(const alias of descriptor.aliases??[])if(!normalize(alias)||normalize(alias).length>64)throw new Error(`Invalid command alias: ${alias}`);
  if(descriptor.priority!==undefined&&(!Number.isFinite(descriptor.priority)||Math.abs(descriptor.priority)>10000))throw new Error('Invalid command priority');
}
function score(query:string,descriptor:CommandDescriptor,usage?:CommandUsage):number {
  const trigger=normalize(descriptor.trigger),aliases=(descriptor.aliases??[]).map(normalize);
  let text=-Infinity;
  for(const candidate of [trigger,...aliases]){
    if(candidate===query)text=Math.max(text,10000);
    else if(candidate.startsWith(query))text=Math.max(text,7000-(candidate.length-query.length)*8);
    else if(query.length>=2&&candidate.includes(query))text=Math.max(text,3500-candidate.indexOf(query)*16);
  }
  if(!Number.isFinite(text))return text;
  let personalized=0;
  if(usage){
    personalized+=Math.min(500,Math.log2((usage.useCount??0)+1)*70);
    personalized+=Math.min(700,(usage.acceptedPrefixes?.[query]??0)*90);
    if(usage.pinned)personalized+=900;
    if((usage.lastUsedAt??0)>0){const days=Math.max(0,(Date.now()-(usage.lastUsedAt??0))/86_400_000);personalized+=250*Math.exp(-days/21);}
  }
  return text+(descriptor.priority??0)+personalized;
}
function rank(query:string,descriptors:readonly CommandDescriptor[],usage:Readonly<Record<string,CommandUsage>>,limit:number):CommandCandidate[]{
  const normalized=normalize(query);if(!normalized)return [];
  return descriptors.map((descriptor,index)=>({descriptor,index,value:score(normalized,descriptor,usage[descriptor.id])}))
    .filter(item=>Number.isFinite(item.value))
    .sort((a,b)=>b.value-a.value||a.index-b.index)
    .slice(0,limit)
    .map(({descriptor,value})=>({id:descriptor.id,trigger:normalize(descriptor.trigger),label:descriptor.label??`\\${normalize(descriptor.trigger)}`,displayCommand:descriptor.displayCommand??`\\${normalize(descriptor.trigger)}`,previewLatex:descriptor.previewLatex??null,score:value,action:clone(descriptor.action)}));
}

/** Host-owned transient command draft. It never writes the mathematical document.
 * The caller must execute the returned command as one normal core transaction. */
export class CommandInputState {
  private target:CommandTarget|null=null;
  private query='';
  private selectedIndex=0;
  constructor(private readonly descriptors:readonly CommandDescriptor[],private readonly usage:Readonly<Record<string,CommandUsage>>={},private readonly limit=8){
    if(!Number.isInteger(limit)||limit<1||limit>32)throw new Error('Invalid command candidate limit');
    const ids=new Set<string>();
    for(const descriptor of descriptors){validateDescriptor(descriptor);if(ids.has(descriptor.id))throw new Error(`Duplicate command descriptor id: ${descriptor.id}`);ids.add(descriptor.id);}
  }
  get active():boolean{return this.target!==null;}
  begin(target:CommandTarget):void {
    if(this.active)throw new Error('Command input is already active');
    this.target=clone(target);this.query='';this.selectedIndex=0;
  }
  cancel():void {this.target=null;this.query='';this.selectedIndex=0;}
  ensureFresh(current:CommandTarget):boolean {
    if(!this.target)return false;
    if(!sameTarget(this.target,current)){this.cancel();return false;}
    return true;
  }
  append(text:string,current:CommandTarget):boolean {
    if(!this.ensureFresh(current))return false;
    if(!/^[A-Za-z0-9-]+$/u.test(text))return false;
    if(this.query.length+text.length>64)return false;
    this.query+=text.toLocaleLowerCase();this.selectedIndex=0;return true;
  }
  backspace(current:CommandTarget):'updated'|'cancelled'|'stale' {
    if(!this.ensureFresh(current))return 'stale';
    if(!this.query){this.cancel();return 'cancelled';}
    this.query=this.query.slice(0,-1);this.selectedIndex=0;return 'updated';
  }
  move(delta:-1|1,current:CommandTarget):boolean {
    if(!this.ensureFresh(current))return false;
    const candidates=this.candidates();if(!candidates.length)return true;
    this.selectedIndex=(this.selectedIndex+delta+candidates.length)%candidates.length;return true;
  }
  select(index:number,current:CommandTarget):boolean {
    if(!this.ensureFresh(current))return false;
    const candidates=this.candidates();
    if(!Number.isInteger(index)||index<0||index>=candidates.length)return false;
    this.selectedIndex=index;return true;
  }
  confirm(current:CommandTarget):CommandConfirmation|null {
    if(!this.ensureFresh(current))return null;
    const candidate=this.candidates()[this.selectedIndex]??null;
    if(!candidate)return null;
    const confirmation={id:candidate.id,query:this.query,action:clone(candidate.action)};this.cancel();return confirmation;
  }
  snapshot():CommandInputSnapshot {
    const candidates=this.candidates();
    if(this.selectedIndex>=candidates.length)this.selectedIndex=Math.max(0,candidates.length-1);
    return {active:this.active,query:this.query,selectedIndex:this.selectedIndex,candidates,target:this.target?clone(this.target):null};
  }
  private candidates():CommandCandidate[]{return rank(this.query,this.descriptors,this.usage,this.limit);}
}

/** Minimal self-owned capability seed. Production P10 will adapt the existing
 * VisualTeX registry to descriptors after P09 capability classification. */
export const CORE_COMMAND_DESCRIPTORS:readonly CommandDescriptor[]=[
  {id:'symbol-partial',trigger:'partial',aliases:['partial-symbol'],label:'\\partial',priority:500,action:{kind:'insertSymbol',command:'\\partial'}},
  {id:'structure-frac',trigger:'frac',aliases:['fraction'],label:'\\frac',priority:480,action:{kind:'insertStructure',structure:'fraction'}},
  {id:'structure-sqrt',trigger:'sqrt',aliases:['root'],label:'\\sqrt',priority:470,action:{kind:'insertStructure',structure:'squareRoot'}},
  {id:'template-partial-derivative',trigger:'partialderivative',aliases:['pd-template','partial-derivative'],label:'partial derivative template',priority:300,action:{kind:'pasteLatex',latex:'\\frac{\\partial \\placeholder{}}{\\partial \\placeholder{}}'}},
];
