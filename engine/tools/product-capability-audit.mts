// Read-only import of the product's static command data; only the NEW core is
// executed. Writes uniquely named audit evidence below engine/target only.
import {createHash} from 'node:crypto';
import {createRequire} from 'node:module';
import {mkdirSync,readFileSync,writeFileSync} from 'node:fs';
import {dirname,resolve,relative} from 'node:path';
import {fileURLToPath} from 'node:url';
import {commandRegistry} from '../../apps/windows/src/autocomplete/commandRegistry.ts';
import {EXTENDED_INTEGRAL_SYMBOLS} from '../../apps/windows/src/math/extendedIntegralCompatibility.ts';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const sha=(bytes:string|Buffer)=>createHash('sha256').update(bytes).digest('hex');
const wasmPath=resolve(root,'web/pkg-node/visualtex_engine_wasm_bg.wasm');
const {EngineRuntime}=createRequire(import.meta.url)(resolve(root,'web/pkg-node/visualtex_engine_wasm.js'));
const runtime=new EngineRuntime();
function send(command:unknown){return JSON.parse(runtime.dispatch(JSON.stringify({protocol:1,command})));}
const symbolFallback=process.argv.includes('--symbol-fallback');
const fontPaths:Record<string,string>={math:process.env.VISUALTEX_TEST_MATH_FONT??resolve(root,'../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf'),text:process.env.VISUALTEX_TEST_TEXT_FONT??'C:/Windows/Fonts/NotoSansSC-VF.ttf',
 ...(symbolFallback?{fallback:resolve(root,'target/test-resources/stix-two-2.13/STIXTwoMath-Regular.otf')}: {})};
const fonts=Object.fromEntries(Object.entries(fontPaths).map(([name,path])=>{
 const bytes=readFileSync(path);if(name==='fallback'&&sha(bytes)!=='f2076b9f1676438439dd41e23676f5ab99056e83d6b8f8c27841591ef2ccfa72')throw new Error('Unexpected explicit audit fallback font');
 const result=JSON.parse(runtime.registerFont(bytes,0));if(result.status!=='ok')throw new Error(JSON.stringify(result));return[name,result.result.font];
}));
let serial=0;
function inspect(source:string){
 const analysis=send({op:'analyzeSource',source});if(analysis.status!=='ok')return{parse:'ERROR',error:analysis.error};
 const p=analysis.result.analysis;const diagnostics=p.diagnostics.map((d:any)=>({code:d.code,source:source.slice(d.span.utf16Start,d.span.utf16End)}));
 if(!p.complete||!p.supported)return{parse:p.complete?'UNSUPPORTED':'INCOMPLETE',diagnostics,layout:'NOT_RUN',svg:'NOT_RUN'};
 const created=send({op:'create',documentId:`capability-audit-${++serial}`,source});if(created.status!=='ok')throw new Error(JSON.stringify(created));const handle=created.result.handle;
 try{
  const layout=send({op:'layout',handle,expectedRevision:0,includeSvg:p.exportable,options:{mathFont:fonts.math.id,textFont:fonts.text.id,mathFallbacks:symbolFallback?[fonts.fallback.id]:[],style:'display',fontSizePx:16,paddingEm:.125}});
  if(layout.status!=='ok')return{parse:'PASS',diagnostics,layout:'ERROR',error:layout.error,svg:'NOT_RUN'};
  return{parse:'PASS',diagnostics,layout:'PASS',svg:p.exportable?'PASS':'UNFILLED',nodes:p.facts.nodes};
 }finally{const result=send({op:'destroy',handle,expectedRevision:0});if(result.status!=='ok')throw new Error(JSON.stringify(result));}
}
let report:any;
try{
 const commands=commandRegistry.map(c=>({id:c.id,command:c.command,category:c.category,preview:c.previewLatex,template:c.insertTemplate,
  previewStatus:inspect(c.previewLatex),templateStatus:inspect(c.insertTemplate),edit:'NOT_RUN',mathML:'NOT_IMPLEMENTED',office:'NOT_RUN'}));
 const integrals=Object.entries(EXTENDED_INTEGRAL_SYMBOLS).map(([command,legacySemanticValue])=>({command,
  legacySemanticValue,authored:inspect(`\\${command}_{a}^{b}x`),
  unicodeGlyph:legacySemanticValue.length<=2&&!legacySemanticValue.includes('\\')?inspect(legacySemanticValue):{parse:'NOT_APPLICABLE'}}));
 const count=(key:'previewStatus'|'templateStatus')=>({total:commands.length,parse:commands.filter(c=>c[key].parse==='PASS').length,
  layout:commands.filter(c=>c[key].layout==='PASS').length,svg:commands.filter(c=>c[key].svg==='PASS').length});
 const missing=new Map<string,number>();for(const c of commands){for(const status of[c.previewStatus,c.templateStatus])for(const d of status.diagnostics??[]){
  const text=d.code==='unknownCommand'?(/^\\[A-Za-z]+/.exec(d.source)?.[0]??d.source):d.code;missing.set(text,(missing.get(text)??0)+1);
 }}
 report={schema:'visualtex-product-capability-audit-v1',capturedAt:new Date().toISOString(),wasmSha256:sha(readFileSync(wasmPath)),fonts,explicitSymbolFallback:symbolFallback,
  commandDataSha256:sha(JSON.stringify(commandRegistry)),uniqueIds:new Set(commands.map(c=>c.id)).size,
  preview:count('previewStatus'),template:count('templateStatus'),missing:[...missing].sort((a,b)=>b[1]-a[1]),commands,integrals,
  limitations:['Actual isolated new-core source/layout/SVG checks only','One fixed font selection and display style','Not actual command-toolbar editing coverage','No macro/custom symbol/Office actions','No legacy math engine execution or production writes']};
 if(runtime.activeSessions!==0)throw new Error('Leaked audit session');
}finally{runtime.free();}
const out=resolve(root,'target/capability-audit',report.capturedAt.replaceAll(':','-').replaceAll('.','-'));mkdirSync(out,{recursive:true});
writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));
console.log(JSON.stringify({receipt:relative(root,resolve(out,'report.json')).replaceAll('\\','/'),preview:report.preview,template:report.template,
 uniqueIds:report.uniqueIds,missing:report.missing.slice(0,35),integrals:report.integrals.map((i:any)=>({command:i.command,parse:i.authored.parse,glyph:i.unicodeGlyph.layout,error:i.unicodeGlyph.error?.message})),scope:report.limitations},null,2));
