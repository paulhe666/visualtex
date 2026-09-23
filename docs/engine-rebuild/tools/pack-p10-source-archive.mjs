import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {readFileSync,writeFileSync,mkdirSync} from 'node:fs';
import {resolve,dirname} from 'node:path';
import {fileURLToPath} from 'node:url';
import {gzipSync} from 'node:zlib';

const repoRoot=resolve(dirname(fileURLToPath(import.meta.url)),'../../..');
const [receiptArg,archiveArg]=process.argv.slice(2);
assert.ok(receiptArg&&archiveArg,'Usage: pack-p10-source-archive.mjs <receipt> <archive>');
const receiptPath=resolve(repoRoot,receiptArg),archivePath=resolve(repoRoot,archiveArg);
const receipt=JSON.parse(readFileSync(receiptPath,'utf8'));
assert.equal(receipt.status,'PASS');assert.equal(receipt.sourceUnchanged,true);assert.ok(Array.isArray(receipt.inputs)&&receipt.inputs.length>0);
const sha=bytes=>createHash('sha256').update(bytes).digest('hex');
const oct=(value,length)=>`${value.toString(8).padStart(length-1,'0')}\0`;
function nameFields(path){
 const bytes=Buffer.byteLength(path);if(bytes<=100)return{name:path,prefix:''};
 for(let index=path.length-1;index>0;index=path.lastIndexOf('/',index-1)){
  if(index<0)break;const prefix=path.slice(0,index),name=path.slice(index+1);
  if(Buffer.byteLength(prefix)<=155&&Buffer.byteLength(name)<=100)return{name,prefix};
 }
 throw new Error(`Path cannot fit ustar header: ${path}`);
}
function header(path,size){
 const {name,prefix}=nameFields(path);const h=Buffer.alloc(512,0);h.write(name,0,100,'utf8');h.write(oct(0o644,8),100,8,'ascii');h.write(oct(0,8),108,8,'ascii');h.write(oct(0,8),116,8,'ascii');h.write(oct(size,12),124,12,'ascii');h.write(oct(0,12),136,12,'ascii');h.fill(0x20,148,156);h.write('0',156,1,'ascii');h.write('ustar\0',257,6,'ascii');h.write('00',263,2,'ascii');h.write('VisualTeX',265,32,'utf8');h.write('VisualTeX',297,32,'utf8');if(prefix)h.write(prefix,345,155,'utf8');
 let sum=0;for(const byte of h)sum+=byte;const checksum=`${sum.toString(8).padStart(6,'0')}\0 `;h.write(checksum,148,8,'ascii');return h;
}
const chunks=[];let total=0;
for(const file of [...receipt.inputs].sort((a,b)=>a.path.localeCompare(b.path))){
 assert.match(file.path,/^(?:engine|apps\/windows)\//u);assert.ok(!file.path.split('/').includes('..'));
 const bytes=readFileSync(resolve(repoRoot,file.path));assert.equal(sha(bytes),file.sha256,`Working source differs: ${file.path}`);
 const archiveName=`package/${file.path}`;const h=header(archiveName,bytes.length);chunks.push(h,bytes);total+=h.length+bytes.length;
 const padding=(512-(bytes.length%512))%512;if(padding){const pad=Buffer.alloc(padding,0);chunks.push(pad);total+=padding;}
}
chunks.push(Buffer.alloc(1024,0));total+=1024;const tar=Buffer.concat(chunks,total);const gz=gzipSync(tar,{level:9,mtime:0});mkdirSync(dirname(archivePath),{recursive:true});writeFileSync(archivePath,gz);
console.log(JSON.stringify({status:'PASS',archive:archiveArg,archiveSha256:sha(gz),archiveBytes:gz.length,sourceSha256:receipt.sourceSha256,files:receipt.inputs.length},null,2));
