import {readFileSync, writeFileSync, existsSync} from 'node:fs';
import {dirname, join} from 'node:path';
import {fileURLToPath} from 'node:url';
const dir=dirname(fileURLToPath(import.meta.url));
const source=readFileSync(join(dir,'perf-20-formulas.tex'),'utf8');
const blocks=[...source.matchAll(/第 \d+ 段正文[^\n]*\n\$\$[^\n]*\$\$/g)].map(m=>m[0]);
if(blocks.length!==10)throw new Error('The canonical fixture must contain ten inline/display pairs.');
for(const count of [50,100]){
 const file=join(dir,`perf-${count}-formulas.tex`);
 const pairs=[];
 for(let i=0;i<count/2;i++){
  let block=blocks[i%blocks.length].replace(/第 \d+ 段正文/,`第 ${i+1} 段正文`);
  if(i>=10){
   // Additional copies have distinct contents as well as fresh Office IDs.
   // The first twenty equations remain identical across all document sizes.
   block=block.replace(/\$\$(.*?)\$\$/,(_,math)=>`$$${math}+0\\cdot${i+1}$$`);
   block=block.replace(/(?<!\$)\$([^$]+)\$(?!\$)/,(_,math)=>`$${math}+0\\cdot${i+1}$`);
  }
  pairs.push(block);
 }
 const text=`性能基准文档。包含 ${count/2} 个行内公式和 ${count/2} 个全部带编号的行间公式。\n\n${pairs.join('\n')}\n\n性能基准文档结束。\n`;
 if(existsSync(file)&&readFileSync(file,'utf8')!==text)throw new Error('Refusing to overwrite a different fixture: '+file);
 if(!existsSync(file))writeFileSync(file,text,'utf8');
 console.log(`FIXTURE|${count}|${file}`);
}
