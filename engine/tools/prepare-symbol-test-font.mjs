// Isolated, pinned test resource preparation. Never installs fonts, changes
// production assets, or includes font bytes in the source checkpoint package.
import {createHash} from 'node:crypto';
import {readFileSync,existsSync,mkdirSync,writeFileSync} from 'node:fs';
import {resolve,dirname} from 'node:path';
import {fileURLToPath} from 'node:url';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const directory=resolve(root,'target/test-resources/stix-two-2.13');
const url='https://raw.githubusercontent.com/stipub/stixfonts/v2.13/fonts/static_otf/STIXTwoMath-Regular.otf';
const expected='f2076b9f1676438439dd41e23676f5ab99056e83d6b8f8c27841591ef2ccfa72';
const path=resolve(directory,'STIXTwoMath-Regular.otf');
const hash=b=>createHash('sha256').update(b).digest('hex');
if(existsSync(path)){
  if(hash(readFileSync(path))!==expected)throw new Error('Existing test font has a different hash; do not overwrite or silently trust it.');
}else{
  const response=await fetch(url,{signal:AbortSignal.timeout(20000)});
  if(!response.ok)throw new Error(`Pinned test font request failed: ${response.status}`);
  const bytes=Buffer.from(await response.arrayBuffer());
  if(bytes.length!==838652||hash(bytes)!==expected)throw new Error('Pinned test font identity mismatch');
  mkdirSync(directory,{recursive:true});writeFileSync(path,bytes,{flag:'wx'});
}
const licensePath=resolve(directory,'OFL.txt');
if(!existsSync(licensePath)){
  const response=await fetch('https://raw.githubusercontent.com/stipub/stixfonts/v2.13/OFL.txt',{signal:AbortSignal.timeout(20000)});
  if(!response.ok)throw new Error(`Pinned font license request failed: ${response.status}`);
  const text=await response.text();
  if(text.length>64000||!text.includes('SIL OPEN FONT LICENSE'))throw new Error('Unexpected test font license');
  writeFileSync(licensePath,text,{flag:'wx'});
}
console.log(JSON.stringify({status:'PASS',fixture:path,sha256:expected,licenseSha256:hash(readFileSync(licensePath)),
  productionWrites:0,fontInstallation:false,sourceArchiveIncludesFonts:false},null,2));
