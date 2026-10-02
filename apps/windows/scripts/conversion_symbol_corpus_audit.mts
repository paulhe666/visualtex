import fs from 'node:fs';
import path from 'node:path';
import { TeX } from 'mathjax-full/js/input/tex.js';
import { AllPackages } from 'mathjax-full/js/input/tex/AllPackages.js';
import { EXPLICIT_MATH_FONT_PACKAGE } from '../src/export/explicit-math-font.ts';
import { CharacterMap, DelimiterMap } from 'mathjax-full/js/input/tex/SymbolMap.js';
import { latexToMathMl } from '../src/export/runtime';
import { VISUALTEX_MATHML_MACROS, VISUALTEX_SVG_MACROS } from '../src/math/latexCompatibility';
import { getActiveCustomSymbols } from '../src/math/customSymbolRegistry';

const out = path.resolve(process.argv[2] ?? '../../artifacts/conversion-symbol-audit-20261002');
const files = ['manifest.json','focus.json','summary.json'].map(n=>path.join(out,n));
for (const file of files) if (fs.existsSync(file)) throw new Error('Refusing to replace evidence: '+file);
const packages = [...AllPackages.filter(n=>n!=='noundefined' && n!=='noerrors'), EXPLICIT_MATH_FONT_PACKAGE];
const tex = new TeX({packages, macros:VISUALTEX_MATHML_MACROS, formatError:(_jax:unknown,e:unknown)=>{throw e;}});
const handlers=(tex as any).configuration.handlers;
if (!handlers || typeof handlers.keys!=='function') throw new Error('MathJax handler API changed');
const inventory:any[]=[];
const regexHandlers:any[]=[];
const atoms=new Map<string,any>();
for (const handlerName of handlers.keys()) {
  const handler=handlers.get(handlerName);
  if (!handler._configuration?.[Symbol.iterator]) throw new Error('MathJax priority API changed');
  for (const {item:map} of handler._configuration) {
    if (!(map.map instanceof Map)) {regexHandlers.push({handler:handlerName,map:map.name,kind:map.constructor.name});continue;}
    for (const [key,value] of map.map) {
      const active=handler.applicable(key)===map;
      const character=map instanceof CharacterMap;
      const delimiter=map instanceof DelimiterMap;
      const entry={handler:handlerName,map:map.name,key,active,kind:map.constructor.name,
        scalarText:character?value.char:undefined,attributes:character?value.attributes:undefined,
        parser:character?map.parser?.name:value.func?.name,args:character?undefined:value.args,
        coverage:character&&active?'atom-candidate':active?'template-required':'shadowed'};
      inventory.push(entry);
      if (!active||!character||!['macro','delimiter'].includes(handlerName)||typeof value.char!=='string') continue;
      const command=delimiter ? key : key.startsWith('\\')?key:'\\'+key;
      // A command can be present under macro and delimiter handlers. Preserve both provenance entries.
      const identity=JSON.stringify([command,value.char,value.attributes??{}]);
      if (!atoms.has(identity)) atoms.set(identity,{command,scalarText:value.char,attributes:value.attributes??{},delimiter,provenance:[]});
      atoms.get(identity).provenance.push(entry);
    }
  }
}
const formulas:any[]=[];
const seen=new Map<string,any>();
function add(latex:string,category:string,source:unknown,displayMode:'inline'|'block'='inline',expected='parse') {
  const key=JSON.stringify([latex,displayMode,expected]);
  const prior=seen.get(key);
  if(prior){prior.provenance.push({category,source});return;}
  const item={index:formulas.length+1,id:'C'+String(formulas.length+1).padStart(5,'0'),latex,displayMode,expected,provenance:[{category,source}]};
  seen.set(key,item);formulas.push(item);
}
for(const atom of atoms.values()){
  // Delimiter lookup keys are lexical tokens, including raw punctuation and the null '.'.
  // Outside left/right some have different meanings (< is a relation, not an angle fence).
  if(atom.delimiter){
    const fenced='\\left'+atom.command+' x\\right.';
    add(fenced,'delimiter-context',atom,'block');
    add('x^{'+fenced+'}_{i}','delimiter-script-context',atom);
  }else{
    add(atom.command,'atomic-map',atom);
    add('x^{'+atom.command+' }_{i}','atomic-script-context',atom);
  }
}
const macroInventory=Object.entries(VISUALTEX_MATHML_MACROS).map(([name,def])=>{
  const argc=Array.isArray(def)?Number(def[1]):0;
  const latex='\\'+name+Array.from({length:argc},(_,i)=>'{'+['x','y','z'][i%3]+'}').join('');
  add(latex,'product-macro',{name,def});
  return{name,mathMlDefinition:def,svgDefinition:(VISUALTEX_SVG_MACROS as any)[name],argc};
});
const customSymbols=getActiveCustomSymbols().map((s:any)=>({id:s.id,command:s.command,ommlFallback:s.ommlFallback??null}));
for(const s of customSymbols)add(s.command.startsWith('\\')?s.command:'\\'+s.command,'custom-symbol-node-snapshot',s,'inline',s.ommlFallback?'parse':'unsupported-without-semantic-fallback');
const pairs=[['parallel','shortparallel'],['nparallel','nshortparallel'],['mid','shortmid'],['nmid','nshortmid'],['emptyset','varnothing'],['sim','thicksim'],['approx','thickapprox'],['propto','varpropto'],['subsetneq','varsubsetneq'],['subsetneqq','varsubsetneqq'],['supsetneq','varsupsetneq'],['supsetneqq','varsupsetneqq']];
for(const [ordinary,alternate]of pairs)for(const command of [ordinary,alternate]){
  add('a\\'+command+' b','alternate-glyph-contrast',{ordinary,alternate,command});
  add('x^{a\\'+command+' b}_{j}','alternate-script-contrast',{ordinary,alternate,command},'block');
}
const focused=[
'x^2','x_2','x_i^j','x^{a_b}','x_{a^b}','x^{a^{b_c}}',
'{}^{14}_{6}\\mathrm{C}','a+{}^{2}x','(x+y)^2','{\\frac{a}{b}}^{n+1}',
'\\sqrt{x_i^2+y_i^2}','\\sqrt[3]{x^{a_b}}',
'\\sum_{i=1}^{n}x_i^2','\\sum\\nolimits_{i=1}^{n}x_i','\\int_0^1x^2\\,dx',
'\\lim_{x\\to0}\\frac{\\sin x}{x}',"x'+x''+x^{\\prime\\prime}",
'\\hat{x}^{2}+\\vec{v}_{i}','\\overbrace{a+b}^{n}','\\underbrace{a+b}_{n}',
'\\left(\\begin{matrix}a&b\\\\c&d\\end{matrix}\\right)^2',
'\\begin{cases}x^2&x>0\\\\x_2&x\\le0\\end{cases}',
'\\begin{aligned}x_i^2&=a+b\\\\y_j&=\\frac{c}{d}\\end{aligned}',
'\\angle ABC+\\measuredangle DEF+\\sphericalangle GHI',
'\\boldsymbol{\\angle}+\\mathcal{\\angle}+\\mathbb{R}+\\mathfrak{g}',
'\\text{速度}\\;v_i=\\mathrm{d}x_i/\\mathrm{d}t',
'\\lvert x\\rvert_i^2+\\lVert y\\rVert_j^3',
'\\left.\\frac{df}{dx}\\right|_{x=0}',
'\\iint_D f(x,y)\\,dx\\,dy','\\oint_C \\mathbf{F}\\cdot d\\mathbf{r}'
];
for(const latex of focused)for(const mode of ['inline','block']as const)add(latex,'structural-focus',null,mode);
for(const command of ['varparalle','varparallel'])add('\\'+command,'unknown-command-control',null,'inline','reject-unless-custom-defined');
for(const item of formulas){
  try{item.mathMl=await latexToMathMl(item.latex,item.displayMode==='block');item.parseStatus='passed';}
  catch(error){item.mathMl='';item.parseStatus='rejected';item.exportError=String(error);}
  item.nativeExportStatus='not-tested';item.visualFidelityStatus='not-tested';
}
const focus=formulas.filter(x=>x.provenance.some((p:any)=>/^(alternate-|structural-|unknown-)/.test(p.category)));
const failures=formulas.filter(x=>x.parseStatus==='rejected').map(({id,latex,expected,exportError,provenance})=>({id,latex,expected,exportError,categories:provenance.map((p:any)=>p.category)}));
const counts={inventory:inventory.length,activeInventory:inventory.filter(x=>x.active).length,shadowedInventory:inventory.filter(x=>!x.active).length,atomIdentities:atoms.size,productMacros:macroInventory.length,customSymbolsInNodeSnapshot:customSymbols.length,cases:formulas.length,parsePassed:formulas.length-failures.length,parseRejected:failures.length,focusCases:focus.length,nativeExportTested:0};
const scope='Finite active command-map inventory plus explicitly generated contexts; not exhaustive formula combinations. Node custom-library snapshot is not proof of browser user-library contents. Parse success is not native-format or pixel-fidelity success.';
const sourceModes=['wordOmml','nativeOle','mathTypeOle'];
const directions=sourceModes.flatMap(source=>sourceModes.filter(target=>target!==source).map(target=>({source,target})));
fs.mkdirSync(out,{recursive:true});
fs.writeFileSync(files[0],JSON.stringify({schemaVersion:1,scope,packages,sourceModes,directions,counts,regexHandlers,inventory,atoms:[...atoms.values()],macroInventory,customSymbols,alternatePairs:pairs,formulas},null,2));
fs.writeFileSync(files[1],JSON.stringify({schemaVersion:1,scope,sourceModes,directions,formulas:focus},null,2));
fs.writeFileSync(files[2],JSON.stringify({schemaVersion:1,scope,counts,failures,files},null,2));
console.log(JSON.stringify({counts,failures,files},null,2));
