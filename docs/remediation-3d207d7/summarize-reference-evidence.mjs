// Read-only summary of evidence already captured from the real installed Ribbon flow.
// This does not operate Word or replace visual/content/navigation acceptance.
import {readFileSync} from 'node:fs';
import {fileURLToPath} from 'node:url';
const [label, number='103'] = process.argv.slice(2);
const read = suffix => JSON.parse(readFileSync(fileURLToPath(new URL(`./evidence/${label}${suffix}.json`, import.meta.url)), 'utf8').replace(/^\uFEFF/, ''));
const doc = read(`-__${number}`), format = read('-font');
const fonts = range => ({ascii:range?.NameAscii, eastAsia:range?.NameFarEast, size:range?.Size, bold:range?.Bold});
const references = format.fields.filter(field => field.type === 3).map(field => {
  const name = field.code.text.match(/^\s*REF\s+(\S+)/i)?.[1];
  const target = doc.bookmarks.find(bookmark => bookmark.name === name);
  return {name, result:field.result.text, target:target && {start:target.start,end:target.end,text:target.text}, matchesTarget:!!target && target.text === field.result.text, font:fonts(field.result), codeFont:fonts(field.code)};
});
console.log(JSON.stringify({label, document:doc.name, end:doc.end, maths:doc.maths.length, shapes:doc.shapes.length,
  tables:doc.tables.map(table=>[table.rows,table.cols]),
  sequenceFields:format.fields.filter(field=>field.type===12).map(field=>({code:field.code.text,result:field.result.text,font:fonts(field.result),hidden:field.result.Hidden,color:field.result.Color})),
  references,
  prose:doc.paras.filter(paragraph=>!paragraph.inTable && paragraph.text.includes('正文')).map(({text,font,size,lineRule,line,before,after,alignment})=>({text,font,size,lineRule,line,before,after,alignment}))
},null,2));
