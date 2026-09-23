import type {EditCommand} from './editing.js';
import type {CommandDescriptor} from './command-input.js';
import type {FontIdentity,LayoutOptions,CustomSymbolTextFontBinding} from './layout.js';
import type {LegacyFontPreferences,LegacyHotkey} from './legacy-archive.js';

export interface VerifiedPreferenceBinding {
  preference:string;
  role:'math'|'text';
  fontId:string;
}
export interface ResolvedFontPreferences {
  requested:LegacyFontPreferences;
  layout:LayoutOptions;
  unresolved:string[];
}
export interface HotkeyEventLike {
  code:string;key:string;ctrlKey:boolean;altKey:boolean;shiftKey:boolean;metaKey:boolean;
}
export interface CompiledLegacyHotkey {
  id:string;
  targetId:string;
  chord:LegacyHotkey['chord'];
  action:EditCommand;
  source:'descriptor'|'legacy-template';
}
const clone=<T>(value:T):T=>structuredClone(value);
const record=(value:unknown):value is Record<string,unknown>=>Boolean(value)&&typeof value==='object'&&!Array.isArray(value);
const text=(value:unknown,name:string,max=65_536):string=>{if(typeof value!=='string'||!value||value.length>max||/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/u.test(value))throw new Error(`Invalid ${name}`);return value;};
function verifiedId(fonts:readonly FontIdentity[],fontId:string):string{
  if(!fonts.some(font=>font.id===fontId))throw new Error('Font preference points to an unregistered resource identity');
  return fontId;
}
/** Resolve saved preferences only through explicit host-provided verified resources.
 * Unknown preferences are retained in `requested` and reported unresolved; no CSS
 * family, OS lookup, or same-name substitution happens here. */
export function resolveVerifiedFontPreferences(base:LayoutOptions,prefs:LegacyFontPreferences,fonts:readonly FontIdentity[],bindings:readonly VerifiedPreferenceBinding[],customTextFonts:readonly CustomSymbolTextFontBinding[]=[]):ResolvedFontPreferences{
  const keys=new Set<string>();for(const binding of bindings){const key=`${binding.role}:${binding.preference}`;if(keys.has(key))throw new Error('Duplicate verified font preference binding');keys.add(key);}
  const lookup=(preference:string,role:VerifiedPreferenceBinding['role'])=>bindings.find(binding=>binding.preference===preference&&binding.role===role);
  const math=lookup(prefs.formulaLetterFont,'math'),textBinding=lookup(prefs.formulaChineseFont,'text'),unresolved:string[]=[];
  if(!math)unresolved.push(`math:${prefs.formulaLetterFont}`);if(!textBinding)unresolved.push(`text:${prefs.formulaChineseFont}`);
  const custom=customTextFonts.map(binding=>({...binding,fontId:verifiedId(fonts,binding.fontId)}));
  return {requested:clone(prefs),unresolved,layout:{...clone(base),...(math?{mathFont:verifiedId(fonts,math.fontId)}:{}),...(textBinding?{textFont:verifiedId(fonts,textBinding.fontId)}:{}),customSymbolTextFonts:custom}};
}
function fallbackAction(target:Record<string,unknown>):EditCommand{
  const command=record(target.command)?target.command:null;if(!command)throw new Error('Legacy hotkey target is missing a command record');
  const template=text(command.insertTemplate,'legacy hotkey insert template');
  // Older persisted builds used #? for an editable slot. Convert only this exact
  // inert marker into the core placeholder syntax; arbitrary macro text is not executed.
  const latex=template.replaceAll('#?','\\placeholder{}');
  return {kind:'pasteLatex',latex};
}
/** Adapt one existing product command without importing the production registry
 * into the engine host. Callers may provide structural descriptors for commands
 * whose cursor/slot behavior must remain native; all other bounded templates are
 * inserted through the same core parser transaction. */
export function compileLegacyCommand(commandValue:unknown,descriptors:readonly CommandDescriptor[]=[]):EditCommand{
  const command=record(commandValue)?commandValue:null;if(!command)throw new Error('Invalid legacy command record');
  const commandId=typeof command.id==='string'?command.id:'';const descriptor=commandId?descriptors.find(item=>item.id===commandId):undefined;
  return descriptor?clone(descriptor.action):fallbackAction({command});
}
/** Compile persisted hotkeys into core edit commands. Known command IDs use the
 * structural descriptor supplied by the P09/P10 registry adapter. Unknown custom
 * tiles remain data and may only become a bounded pasteLatex transaction. */
export function compileLegacyHotkeys(bindings:readonly LegacyHotkey[],descriptors:readonly CommandDescriptor[]=[]):CompiledLegacyHotkey[]{
  const byId=new Map(descriptors.map(descriptor=>[descriptor.id,descriptor]));const chordIds=new Set<string>(),bindingIds=new Set<string>();
  return bindings.map(binding=>{
    if(bindingIds.has(binding.id))throw new Error('Duplicate legacy hotkey binding id');bindingIds.add(binding.id);
    const target=record(binding.target)?binding.target:null;if(!target)throw new Error('Invalid legacy hotkey target');const command=record(target.command)?target.command:null;if(!command)throw new Error('Invalid legacy hotkey command');
    const commandId=typeof command.id==='string'?command.id:'';const descriptor=commandId?byId.get(commandId):undefined;
    const chordId=[binding.chord.ctrlKey,binding.chord.altKey,binding.chord.shiftKey,binding.chord.metaKey,binding.chord.code].join(':');if(chordIds.has(chordId))throw new Error('Conflicting legacy hotkey chord');chordIds.add(chordId);
    return {id:binding.id,targetId:text(target.id,'legacy hotkey target id',256),chord:clone(binding.chord),action:descriptor?clone(descriptor.action):fallbackAction(target),source:descriptor?'descriptor':'legacy-template'};
  });
}
export function matchLegacyHotkey(event:HotkeyEventLike,bindings:readonly CompiledLegacyHotkey[]):CompiledLegacyHotkey|null{
  return bindings.find(binding=>binding.chord.code===event.code&&binding.chord.ctrlKey===event.ctrlKey&&binding.chord.altKey===event.altKey&&binding.chord.shiftKey===event.shiftKey&&binding.chord.metaKey===event.metaKey)??null;
}
