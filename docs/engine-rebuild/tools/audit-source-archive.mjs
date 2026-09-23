// Read-only archive audit. Never extracts files or mutates project/index/refs.
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {resolve,dirname} from 'node:path';
import {fileURLToPath} from 'node:url';
import {gunzipSync} from 'node:zlib';
import {createHash} from 'node:crypto';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'../../..');
const [receiptPath,archivePath]=process.argv.slice(2);
assert.ok(receiptPath&&archivePath,'Usage: audit-source-archive.mjs <receipt> <archive>');
const receipt=JSON.parse(readFileSync(resolve(root,receiptPath),'utf8'));
assert.equal(receipt.sourceUnchanged,true,'Verification did not use stable inputs');
const bytes=readFileSync(resolve(root,archivePath));
const tar=gunzipSync(bytes,{maxOutputLength:32*1024*1024});
const sha=buffer=>createHash('sha256').update(buffer).digest('hex');
const entries=new Map();
const field=(header,start,len)=>header.subarray(start,start+len).toString('utf8').replace(/\0.*$/su,'').trim();
for(let at=0;at+512<=tar.length;){
 const header=tar.subarray(at,at+512);if(header.every(b=>b===0))break;
 const type=field(header,156,1),prefix=field(header,345,155),name=(prefix?`${prefix}/`:'')+field(header,0,100);
 const sizeText=field(header,124,12);assert.match(sizeText,/^[0-7]+$/u,'Unsupported tar size');
 const size=parseInt(sizeText,8);assert.ok(Number.isSafeInteger(size)&&size>=0&&at+512+size<=tar.length,'Truncated tar entry');
 assert.ok(name.startsWith('package/')&&!name.split('/').includes('..'),'Unsafe archive path');
 assert.ok(type==='0'||type===''||type==='5',`Unsupported archive entry type: ${type} ${name}`);
 if(type!=='5'){
  assert.ok(!entries.has(name),`Duplicate archive entry ${name}`);
  entries.set(name,{size,sha256:sha(tar.subarray(at+512,at+512+size))});
 }
 at+=512+Math.ceil(size/512)*512;
}
const forbidden=[...entries.keys()].filter(path=>/(?:^|\/)(?:target|node_modules|\.tools|pkg|pkg-node)\//u.test(path)
 ||/\.(?:ttf|otf|ttc|woff2?|exe|dll|msi|dmg|wasm|pem|key)$/iu.test(path)||/(?:^|\/)\.env(?:\.|$)/u.test(path));
assert.deepEqual(forbidden,[],'Archive contains generated/binary/private resources');
for(const file of receipt.inputs){
 assert.equal(entries.get(`package/${file.path}`)?.sha256,file.sha256,`Archive differs from verified source: ${file.path}`);
 assert.equal(sha(readFileSync(resolve(root,'engine',file.path))),file.sha256,`Working input changed after validation: ${file.path}`);
}
console.log(JSON.stringify({schema:'visualtex-engine-archive-audit-v1',status:'PASS',archive:archivePath,
 archiveSha256:sha(bytes),archiveBytes:bytes.length,archiveFiles:entries.size,
 verifiedSourceFiles:receipt.inputs.length,sourceSha256:receipt.sourceSha256,
 archiveMatchedEveryVerifiedInput:true,currentSourceStillMatches:true,forbiddenArchiveFiles:forbidden,
 verificationReceipt:receiptPath,strictVerificationStatus:receipt.status,writes:0},null,2));
