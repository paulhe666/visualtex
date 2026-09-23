// Read compatibility only. No store imports, storage writes, DOM, Office,
// font fallback, macro execution or MathLive-offset -> NodeId conversion.
export const MAX_LEGACY_ARCHIVE_BYTES=96_000_000;
const MAX_FORMULAS=10_000,MAX_SOURCE=1_048_576;
type RecordValue=Record<string,unknown>;
export type LegacyArchiveKind='formula-document'|'editor-state'|'configuration'|'document-checkpoint';
export interface LegacyFormulaSource {id:string;latex:string;mode:'inline'|'display'}
export interface LegacyFontPreferences {formulaLetterFont:string;formulaChineseFont:string}
export interface LegacyHotkey {
  id:string;target:RecordValue;
  chord:{code:string;key:string;ctrlKey:boolean;altKey:boolean;shiftKey:boolean;metaKey:boolean};
  updatedAt:number;
}
const record=(v:unknown):v is RecordValue=>v!==null&&typeof v==='object'&&!Array.isArray(v);
function object(v:unknown,name:string):RecordValue{if(!record(v))throw new Error(`Invalid legacy ${name}`);return v;}
function string(v:unknown,name:string,max=MAX_SOURCE):string{if(typeof v!=='string'||v.length>max)throw new Error(`Invalid legacy ${name}`);return v;}
function id(v:unknown):string{const s=string(v,'formula id',256);if(!s||/[\u0000-\u001f\u007f]/u.test(s))throw new Error('Invalid legacy formula id');return s;}
function parse(raw:string):RecordValue{
  if(typeof raw!=='string'||raw.length>MAX_LEGACY_ARCHIVE_BYTES||new TextEncoder().encode(raw).length>MAX_LEGACY_ARCHIVE_BYTES)throw new Error('Legacy archive byte budget exceeded');
  const value=JSON.parse(raw.charCodeAt(0)===0xfeff?raw.slice(1):raw) as unknown;let count=0;
  const walk=(v:unknown,depth:number):void=>{if(++count>1_000_000||depth>48)throw new Error('Legacy archive nesting/item budget exceeded');if(typeof v==='number'&&!Number.isFinite(v))throw new Error('Non-finite legacy archive value');if(Array.isArray(v))for(const x of v)walk(x,depth+1);else if(record(v))for(const x of Object.values(v))walk(x,depth+1);};walk(value,0);return object(value,'archive');
}
function formulaSources(value:unknown,blocks:boolean):LegacyFormulaSource[]{
  if(!Array.isArray(value)||value.length>MAX_FORMULAS)throw new Error('Legacy formula count/type exceeds the supported boundary');
  const used=new Set<string>();return value.map(v=>{const row=object(v,'formula'),key=id(row.id);if(used.has(key))throw new Error('Duplicate legacy formula id; refusing to alias two owners');used.add(key);
    const mode=blocks?row.displayMode:row.mode;if(mode!==undefined&&!(blocks?['inline','block']:['inline','display']).includes(String(mode)))throw new Error('Invalid legacy formula display mode');
    return {id:key,latex:string(row.latex,'LaTeX source'),mode:mode==='inline'?'inline':'display'};
  });
}
/** Immutable original envelope. An unchanged save returns the exact input
 * string, including formatting/BOM. A changed save modifies only a CAS-matched
 * formula source and retains all unknown metadata and the original schema. */
export class LegacyArchive {
  readonly kind:LegacyArchiveKind;
  private readonly envelope:RecordValue;
  private readonly sources:LegacyFormulaSource[];
  constructor(private readonly raw:string){
    const v=parse(raw);this.envelope=v;
    if(v.schema==='visualtex-user-configuration'){
      if(v.version!==undefined&&v.version!==1)throw new Error('Unsupported legacy configuration version');this.kind='configuration';
      object(v.editorSettings??{},'editor settings');const storage=object(v.storage,'configuration storage');for(const value of Object.values(storage))string(value,'storage entry',MAX_LEGACY_ARCHIVE_BYTES);
      if(v.personalization!==undefined)object(v.personalization,'personalization');this.sources=[];
    }else if(Array.isArray(v.formulas)){
      this.kind='formula-document';string(v.title,'title',10000);if(v.version!==undefined&&(!Number.isSafeInteger(v.version)||Number(v.version)<1))throw new Error('Invalid legacy document version');object(v.settings??{},'document settings');if(v.macros!==undefined){for(const value of Object.values(object(v.macros,'macros')))string(value,'macro source');}this.sources=formulaSources(v.formulas,true);
    }else if(record(v.state)){
      this.kind='editor-state';const state=v.state;this.sources=state.lines!==undefined?formulaSources(state.lines,false):typeof state.latex==='string'?[{id:'legacy-single-formula',latex:string(state.latex,'legacy single source'),mode:'display'}]:[];
      if(!this.sources.length&&state.lines===undefined&&state.latex===undefined)throw new Error('Legacy editor state has no formula source');
    }else if(record(v.document)&&Array.isArray(v.document.lines)){
      this.kind='document-checkpoint';id(v.id);this.sources=formulaSources(v.document.lines,false);
      object(v.document.selectionByLineId??{},'legacy selection archive');
    }else throw new Error('Unrecognized legacy JSON envelope; no .vtx binary format is assumed');
  }
  originalText():string{return this.raw;}
  snapshot():RecordValue{return structuredClone(this.envelope);}
  formulas():LegacyFormulaSource[]{return structuredClone(this.sources);}
  settings():RecordValue{
    const v=this.envelope;return structuredClone(this.kind==='configuration'?object(v.editorSettings??{},'settings'):this.kind==='formula-document'?object(v.settings??{},'settings'):this.kind==='editor-state'?object(v.state,'state'):object(v.document,'checkpoint document'));
  }
  storage():Record<string,string>{return this.kind==='configuration'?structuredClone(this.envelope.storage) as Record<string,string>:{};}
  fontPreferences(externalStorage:Readonly<Record<string,string>>={}):LegacyFontPreferences{
    const settings=this.settings(),storage={...externalStorage,...this.storage()};
    // Explicit document settings win on file-open; separately stored fonts win
    // on workspace hydration. Imported configuration storage wins over current
    // storage. Unknown preferences are retained, never silently mapped to CSS.
    const value=(setting:string,key:string,fallback:string)=>this.kind==='formula-document'?(settings[setting]??storage[key]??fallback):(storage[key]??settings[setting]??fallback);
    return {formulaLetterFont:string(value('formulaLetterFont','visualtex.formula-letter-font','katex'),'letter font preference',128),formulaChineseFont:string(value('formulaChineseFont','visualtex.formula-chinese-font','system'),'Chinese font preference',128)};
  }
  historyItems():RecordValue[]{
    const v=this.envelope,container=this.kind==='configuration'?object(v.personalization??{},'personalization'):this.kind==='editor-state'?object(v.state,'state'):{};
    const items=container.history??[];if(!Array.isArray(items)||items.length>10_000)throw new Error('Invalid legacy snippet history');
    return items.map(item=>{const r=object(item,'history item');id(r.id);string(r.latex,'history LaTeX');if(r.lines!==undefined)formulaSources(r.lines,false);return structuredClone(r);});
  }
  hotkeys():LegacyHotkey[]{
    const source=this.storage()['visualtex-formula-hotkeys-v1'];if(source===undefined)return [];
    const envelope=parse(source),state=object(envelope.state??envelope,'shortcut state'),bindings=state.bindings??[];
    if(!Array.isArray(bindings)||bindings.length>10_000)throw new Error('Invalid legacy shortcut bindings');
    const used=new Set<string>();return bindings.map(value=>{const b=object(value,'shortcut'),key=id(b.id);if(used.has(key))throw new Error('Duplicate legacy shortcut id');used.add(key);const target=object(b.target,'shortcut target'),chord=object(b.chord,'shortcut chord');id(target.id);object(target.command,'shortcut command');
      string(chord.code,'shortcut code',64);string(chord.key,'shortcut key',64);for(const k of ['ctrlKey','altKey','shiftKey','metaKey'])if(typeof chord[k]!=='boolean')throw new Error('Invalid legacy shortcut modifier');if(typeof b.updatedAt!=='number'||!Number.isFinite(b.updatedAt))throw new Error('Invalid legacy shortcut timestamp');
      return structuredClone(b) as unknown as LegacyHotkey;
    });
  }
  customSymbolLibrary():RecordValue|null{
    const raw=this.storage()['visualtex.custom-symbols.v1'];if(raw===undefined)return null;
    const value=parse(raw);if(value.version!==1||!Array.isArray(value.symbols))throw new Error('Unsupported legacy custom-symbol library envelope');return structuredClone(value);
  }
  /** Old MathLive offsets are available only as archival data. They must not be
   * passed to core selection APIs; a new session owns its own initial caret. */
  legacySelectionArchive():RecordValue|null{
    return this.kind==='document-checkpoint'?structuredClone(object(object(this.envelope.document,'checkpoint').selectionByLineId??{},'selection archive')):null;
  }
  replaceFormulaSource(formulaId:string,expectedSource:string,replacement:string):LegacyArchive{
    string(replacement,'replacement source');const current=this.sources.find(f=>f.id===formulaId);if(!current||current.latex!==expectedSource)throw new Error('Legacy formula source conflict');if(current.latex===replacement)return this;
    const next=structuredClone(this.envelope);
    if(this.kind==='formula-document'){const row=(next.formulas as RecordValue[]).find(f=>f.id===formulaId)!;row.latex=replacement;}
    else if(this.kind==='editor-state'){const state=object(next.state,'state');if(Array.isArray(state.lines))(state.lines as RecordValue[]).find(f=>f.id===formulaId)!.latex=replacement;else state.latex=replacement;}
    else if(this.kind==='document-checkpoint')throw new Error('Historical checkpoints are read-only; restore into a fresh document before editing');
    else throw new Error('A configuration archive is not a formula document');
    return new LegacyArchive(JSON.stringify(next,null,2));
  }
}
