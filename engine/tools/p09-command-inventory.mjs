// P09 read-only inventory of the existing Windows command registry against the
// current self-owned parser. This does not import the production editor or write
// application state; reports are emitted only below engine/target.
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {mkdirSync,readFileSync,writeFileSync} from 'node:fs';
import {dirname,resolve,relative} from 'node:path';
import {fileURLToPath,pathToFileURL} from 'node:url';
import {createRequire} from 'node:module';

const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const repo=resolve(root,'..');
const stamp=new Date().toISOString().replaceAll(':','-').replaceAll('.','-');
const out=resolve(root,'target/p09-command-inventory',stamp);mkdirSync(out,{recursive:true});
const registryModule=await import(pathToFileURL(resolve(repo,'apps/windows/src/autocomplete/commandRegistry.ts')).href);
const commands=registryModule.commandRegistry;
assert.ok(Array.isArray(commands)&&commands.length>0);
const {EngineRuntime}=createRequire(import.meta.url)(resolve(root,'web/pkg-node/visualtex_engine_wasm.js'));
const runtime=new EngineRuntime();
const dispatch=command=>{
 const response=JSON.parse(runtime.dispatch(JSON.stringify({protocol:1,command})));
 assert.equal(response.status,'ok',JSON.stringify(response));return response.result;
};
const placeholder='\\placeholder{}';
// A placeholder is a structural editor slot, not raw source text. Materialize
// it as a braced atom so an adjacent control word such as `\\partial` or
// delimiter name such as `\\langle` cannot be accidentally changed into the
// unrelated control word `\\partialx`/`\\langlex` by this audit itself.
const materialize=template=>template.split(placeholder).join('{x}');
const normalizeTrigger=value=>value.trim().replace(/^\\/u,'').toLocaleLowerCase();
const byTrigger=new Map();
for(const command of commands){
 const key=normalizeTrigger(command.command);const ids=byTrigger.get(key)??[];ids.push(command.id);byTrigger.set(key,ids);
}
const rows=[];let supportedTemplates=0,supportedContextualTemplates=0,coveredTemplates=0,directAtoms=0;
for(const command of commands){
 const source=materialize(command.insertTemplate);const analyzed=dispatch({op:'analyzeSource',source});
 const analysis=analyzed.analysis;const templateSupported=analysis.complete&&analysis.supported;
 const contextual=/^[\s]*[\^_]/u.test(source);const contextualSource=contextual?`x${source}`:null;
 const contextualAnalysis=contextual?dispatch({op:'analyzeSource',source:contextualSource}).analysis:null;
 const contextualSupported=Boolean(contextualAnalysis?.complete&&contextualAnalysis?.supported&&contextualAnalysis?.exportable);
 if(templateSupported)supportedTemplates++;if(contextualSupported)supportedContextualTemplates++;if(templateSupported||contextualSupported)coveredTemplates++;
 const exact=command.insertTemplate===command.command&&!command.insertTemplate.includes(placeholder);
 const exactResult=exact?dispatch({op:'analyzeSource',source:command.command}).analysis:null;
 const direct=Boolean(exactResult?.complete&&exactResult?.supported&&exactResult?.exportable);if(direct)directAtoms++;
 rows.push({id:command.id,command:command.command,insertTemplate:command.insertTemplate,previewLatex:command.previewLatex,category:command.category,defaultPriority:command.defaultPriority,aliases:command.aliases,keywords:command.keywords,
  trigger:normalizeTrigger(command.command),triggerCollisionIds:byTrigger.get(normalizeTrigger(command.command)),usesPlaceholder:command.insertTemplate.includes(placeholder),materializedTemplate:source,
  template:{complete:analysis.complete,supported:analysis.supported,exportable:analysis.exportable,diagnostics:analysis.diagnostics.map(d=>({code:d.code,severity:d.severity}))},
  contextualTemplate:contextual?{source:contextualSource,complete:contextualAnalysis.complete,supported:contextualAnalysis.supported,exportable:contextualAnalysis.exportable,diagnostics:contextualAnalysis.diagnostics.map(d=>({code:d.code,severity:d.severity}))}:null,
  capabilityCovered:templateSupported||contextualSupported,directAtom:direct});
}
runtime.free();
const collisions=[...byTrigger.entries()].filter(([,ids])=>ids.length>1).map(([trigger,ids])=>({trigger,ids})).sort((a,b)=>a.trigger.localeCompare(b.trigger));
const partial=rows.filter(row=>row.command==='\\partial'||row.command==='\\pdv'||row.id.includes('partial'));
const report={schema:'visualtex-p09-command-inventory-v2',status:coveredTemplates===commands.length?'PASS':'FAILED',registryCount:commands.length,supportedMaterializedTemplates:supportedTemplates,supportedContextualTemplates,coveredTemplates,directSupportedAtoms:directAtoms,triggerCollisions:collisions.length,partialRisk:partial,
 rows,source:{registry:'apps/windows/src/autocomplete/commandRegistry.ts',registrySha256:createHash('sha256').update(readFileSync(resolve(repo,'apps/windows/src/autocomplete/commandRegistry.ts'))).digest('hex')},
 scope:'Static built-in registry only. Runtime custom-symbol library is intentionally inventoried separately; no localStorage or production application state is read.'};
writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));
console.log(JSON.stringify({receipt:relative(root,resolve(out,'report.json')).replaceAll('\\','/'),status:report.status,registryCount:report.registryCount,supportedMaterializedTemplates:report.supportedMaterializedTemplates,supportedContextualTemplates:report.supportedContextualTemplates,coveredTemplates:report.coveredTemplates,directSupportedAtoms:report.directSupportedAtoms,triggerCollisions:report.triggerCollisions,partialRisk:report.partialRisk.map(({id,command,insertTemplate,directAtom,template,contextualTemplate})=>({id,command,insertTemplate,directAtom,templateSupported:template.supported,contextualSupported:Boolean(contextualTemplate?.supported)}))},null,2));
if(report.status!=='PASS')process.exitCode=1;
