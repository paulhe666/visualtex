import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {readFileSync} from 'node:fs';
import {resolve,dirname} from 'node:path';
import {fileURLToPath} from 'node:url';
import {gunzipSync} from 'node:zlib';

const repoRoot=resolve(dirname(fileURLToPath(import.meta.url)),'../../..');
const [receiptArg,archiveArg]=process.argv.slice(2);
assert.ok(receiptArg&&archiveArg,'Usage: audit-p10-source-archive.mjs <receipt> <archive>');
const receipt=JSON.parse(readFileSync(resolve(repoRoot,receiptArg),'utf8'));
assert.equal(receipt.status,'PASS');assert.equal(receipt.sourceUnchanged,true);assert.ok(Array.isArray(receipt.inputs));
const archiveBytes=readFileSync(resolve(repoRoot,archiveArg));const tar=gunzipSync(archiveBytes,{maxOutputLength:64*1024*1024});
const sha=bytes=>createHash('sha256').update(bytes).digest('hex');
const field=(header,start,len)=>header.subarray(start,start+len).toString('utf8').replace(/\0.*$/su,'').trim();
const entries=new Map();
for(let at=0;at+512<=tar.length;){
 const header=tar.subarray(at,at+512);if(header.every(byte=>byte===0))break;
 const nameField=field(header,0,100),prefix=field(header,345,155),name=(prefix?`${prefix}/`:'')+nameField,type=field(header,156,1),sizeText=field(header,124,12);
 assert.match(sizeText,/^[0-7]+$/u,`Bad tar size for ${name}`);const size=parseInt(sizeText,8);assert.ok(Number.isSafeInteger(size)&&size>=0&&at+512+size<=tar.length,`Truncated ${name}`);
 assert.equal(type,'0',`Unexpected tar entry type ${type} for ${name}`);assert.ok(name.startsWith('package/')&&!name.split('/').includes('..'),`Unsafe archive path ${name}`);assert.ok(!entries.has(name),`Duplicate ${name}`);
 entries.set(name,{size,sha256:sha(tar.subarray(at+512,at+512+size))});at+=512+Math.ceil(size/512)*512;
}
const expected=new Set(receipt.inputs.map(file=>`package/${file.path}`));assert.equal(entries.size,expected.size,'Archive file count differs from verified inputs');
for(const [name] of entries)assert.ok(expected.has(name),`Unexpected archive file ${name}`);
for(const file of receipt.inputs){
 const name=`package/${file.path}`;assert.equal(entries.get(name)?.sha256,file.sha256,`Archive differs from receipt: ${file.path}`);assert.equal(sha(readFileSync(resolve(repoRoot,file.path))),file.sha256,`Working source changed after verification: ${file.path}`);
}
const forbidden=[...entries.keys()].filter(path=>/(?:^|\/)(?:target|node_modules|dist|\.tools|pkg|pkg-node)\//u.test(path)||/\.(?:ttf|otf|ttc|woff2?|exe|dll|msi|dmg|wasm|pem|key)$/iu.test(path)||/(?:^|\/)\.env(?:\.|$)/u.test(path));assert.deepEqual(forbidden,[],'Archive contains generated/binary/private resources');
console.log(JSON.stringify({schema:'visualtex-p10-source-archive-audit-v1',status:'PASS',archive:archiveArg,archiveSha256:sha(archiveBytes),archiveBytes:archiveBytes.length,archiveFiles:entries.size,verifiedSourceFiles:receipt.inputs.length,sourceSha256:receipt.sourceSha256,archiveMatchedEveryVerifiedInput:true,currentSourceStillMatches:true,forbiddenArchiveFiles:forbidden,verificationReceipt:receiptArg,completion:receipt.completion,writes:0},null,2));
