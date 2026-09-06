// Read-only review of the original audit. This is not product acceptance.
import { readFileSync, readdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { inflateRawSync } from 'node:zlib';
function shape(s) {
  let metadata;
  if (s.alternative?.startsWith('visualtex:v1:deflate:')) {
    const m=JSON.parse(inflateRawSync(Buffer.from(s.alternative.slice(21), 'base64url')));
    metadata={id:m.formulaId,latex:m.latex,numbered:m.numbered,displayMode:m.displayMode};
  }
  return {...s, alternative:metadata || s.alternative};
}
const root = join(dirname(fileURLToPath(import.meta.url)), '../audit-3d207d7/evidence');
const filter = new RegExp(process.argv[2] || '.*');
for (const file of readdirSync(root).filter(f => f.endsWith('.json') && filter.test(f))) {
  const data = JSON.parse(readFileSync(join(root, file), 'utf8').replace(/^\uFEFF/, ''));
  const docs = data.documents || (data.maths ? [data] : null);
  console.log('\nFILE ' + file);
  if (!docs) { console.log(JSON.stringify(data, (key,value) => /^(svg|svgBase64|pngBase64|emfBase64)$/.test(key) ? `[${value?.length || 0} bytes]` : value)); continue; }
  for (const d of docs) console.log(JSON.stringify({
    name:d.name, saved:d.saved, end:d.end,
    tables:d.tables?.map(t=>({start:t.start,end:t.end,rows:t.rows,cols:t.cols,cells:t.cells?.map(c=>[c.row,c.col,c.start,c.end,c.math])})),
    maths:d.maths?.map(m=>[m.start,m.end,m.type,m.font,m.size,m.text]), shapes:d.shapes?.map(shape), controls:d.controls,
    fields:d.fields?.map(f=>[f.start,f.end,f.code,f.result]), bookmarks:d.bookmarks?.map(b=>[b.name,b.start,b.end,b.text]),
    paras:d.paras?.map(p=>[p.start,p.end,p.text,p.font,p.size,p.alignment,p.before,p.after]), variables:d.variables, xmlFonts:d.xmlFonts
  }));
}
