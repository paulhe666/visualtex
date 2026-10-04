import { readFileSync, writeFileSync } from 'node:fs';
import { DOMParser } from '@xmldom/xmldom';
import { MmlMi } from 'mathjax-full/js/core/MmlTree/MmlNodes/mi.js';
const base = '../../artifacts/conversion-symbol-audit-20261002/';
const read = (path:string) => JSON.parse(readFileSync(base + path, 'utf8').replace(/^\uFEFF/, ''));
const cases = read('cp251-preserve-structure/word-cases.json').cases;
const stages = read('cp251-word-preserved3/stages.json');
function inspect(xml:string) {
 const doc = new DOMParser().parseFromString(xml, 'application/xml');
 const elements = Array.from(doc.getElementsByTagName('*'));
 const tokens = elements.filter(e => ['mi','mo','mn','mtext'].includes(e.localName!)).map(e => {
  let style = '';
  for(let p:any=e; p?.nodeType===1; p=p.parentNode) if(p.hasAttribute('mathvariant')) { style=p.getAttribute('mathvariant'); break; }
  return {role:e.localName,text:e.textContent,style:style || (e.localName==='mi' && MmlMi.singleCharacter.test(e.textContent!) ? 'italic' : 'normal'),stretchy:e.getAttribute('stretchy') || null};
 });
 return {tokens, structures:['mfenced','mfrac','msup','msub'].map(name=>({name,count:elements.filter(e=>e.localName===name).length}))};
}
const rows = stages.map((stage:any) => {
 const expected=inspect(cases.find((c:any)=>c.name===stage.name).mathMl), actual=inspect(stage.mathMl);
 const contract=(tokens:any[])=>tokens.map(({stretchy,...token})=>token);
 const checks={tokens:JSON.stringify(contract(expected.tokens))===JSON.stringify(contract(actual.tokens)),structure:JSON.stringify(expected.structures)===JSON.stringify(actual.structures),fixed:expected.tokens.every((t,i)=>t.stretchy!=='false'||actual.tokens[i]?.stretchy==='false')};
 return {name:stage.name,stage:stage.stage,expected,actual,checks,passed:Object.values(checks).every(Boolean)};
});
const report={scope:'Strict ordered token role/scalar/effective-style, fixed flags and owned structure counts; not full pixels',rows,passed:rows.filter((r:any)=>r.passed).length,failed:rows.filter((r:any)=>!r.passed).length};
writeFileSync(base+'cp251-preserve-structure/word-stage-audit.json',JSON.stringify(report,null,2));
console.log(JSON.stringify({passed:report.passed,failed:report.failed,failures:rows.filter((r:any)=>!r.passed).map((r:any)=>({name:r.name,stage:r.stage,checks:r.checks}))},null,2));
