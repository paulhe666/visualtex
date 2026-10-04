import assert from 'node:assert/strict';
import {mkdir, readFile, writeFile} from 'node:fs/promises';
import {resolve, join} from 'node:path';
import {latexToMathMl, latexToSvg} from '../src/export/runtime.ts';
import {TeXFont} from 'mathjax-full/js/output/svg/fonts/tex.js';
import {materializeEffectiveOperatorVariants} from '../src/export/effectiveMathMlGlyphVariants.ts';

const root = resolve(process.argv[2] ?? (()=>{throw new Error('An owned output path is required.');})());
await mkdir(root,{recursive:true});
const source = JSON.parse(await readFile(resolve(process.argv[3]),'utf8'));
const updated=[];
for(const item of source) updated.push({...item,MathMl:await latexToMathMl(item.Latex,true)});
await writeFile(join(root,'inputs.json'),JSON.stringify(updated,null,2),'utf8');
const font=new TeXFont();
const observations=[];
for(const [name,command,expected] of [['cal','mathcal','normal'],['frak','mathfrak','normal'],['sf','mathsf','normal'],['tt','mathtt','normal'],['bb','mathbb','bold']]) {
  const latex=`\\${command}{-}`;
  const mathMl=await latexToMathMl(latex,true);
  assert.match(mathMl,new RegExp(`mathvariant="${expected}"`));
  const reference=expected==='bold' ? String.raw`\boldsymbol{-}` : '-';
  const options={displayMode:true,fontSizePt:12,paddingPx:0,background:'transparent'};
  const actual=await latexToSvg(latex,options);
  const control=await latexToSvg(reference,options);
  const paths=(s)=>[...s.matchAll(/<path\b[^>]*\bd="([^"]+)"/g)].map(x=>x[1]);
  assert.deepEqual(paths(actual.svg),paths(control.svg),`${name} changed actual rendered glyph`);
  assert.equal(actual.width,control.width); assert.equal(actual.height,control.height);
  observations.push({name,latex,expected,mathMl,ActualGlyphMatchesReference:true});
}
// Control actual alphabet glyphs as well as basic weight/slant: neither can
// be dropped merely because some other symbol has a renderer fallback.
assert.match(await latexToMathMl(String.raw`\mathcal{A}`,true),/mathvariant="script"/);
assert.match(await latexToMathMl(String.raw`\mathfrak{A}`,true),/mathvariant="fraktur"/);
assert.match(await latexToMathMl(String.raw`\boldsymbol{-}`,true),/mathvariant="bold"/);
assert.match(await latexToMathMl(String.raw`\mathit{A}`,true),/mathvariant="italic"/);
// Inspect every nonletter glyph exposed by this version's renderer. A change
// must be backed by IDENTICAL resolved glyph objects, not a hand-maintained
// list of selected characters or requested font names.
let checks=0,changes=0;
for(const variant of ['script','fraktur','sans-serif','monospace','double-struck','bold-script','bold-fraktur','bold-sans-serif','sans-serif-italic','sans-serif-bold-italic']) {
  for(const key in font.getVariant(variant).chars) {
    const scalar=Number(key); if(!Number.isInteger(scalar)||scalar<0||scalar>0xFFFF) continue;
    const glyph=String.fromCodePoint(scalar); if(/[\p{L}\p{N}]/u.test(glyph)) continue;
    const attrs={mathvariant:variant};
    const node={kind:'mo',getText:()=>glyph,attributes:{get:k=>attrs[k],set:(k,v)=>{attrs[k]=v;}}};
    materializeEffectiveOperatorVariants({walkTree:f=>f(node)},font);
    checks++;
    if(attrs.mathvariant!==variant) {
      changes++;
      assert.equal(font.getChar(variant,scalar),font.getChar(attrs.mathvariant,scalar));
    }
  }
}
await writeFile(join(root,'summary.json'),JSON.stringify({Inputs:updated.length,ChangedInputs:updated.filter((x,i)=>x.MathMl!==source[i].MathMl).map(x=>x.Id),GlyphChecks:checks,ResolvedFallbacks:changes,DirectSvgChecks:observations},null,2),'utf8');
console.log(JSON.stringify({Inputs:updated.length,ChangedInputs:updated.filter((x,i)=>x.MathMl!==source[i].MathMl).length,GlyphChecks:checks,ResolvedFallbacks:changes,DirectSvgChecks:5}));
