// Recovery utility scoped to this feature's owned test fixtures, never user docs.
import {execFileSync} from 'node:child_process';
const owned=new Set(['NativeOmmlTemplate.pptx','NativeOmmlProbe.pptx','NativeOmmlGuardProbe.pptx','NativeOmmlMinimalProbe [Repaired]','NativeOmmlProductionFixture.pptx','NativeOmmlAcceptance.pptm','演示文稿2']);
function as(command){return execFileSync('/usr/bin/osascript',['-e',`tell application "Microsoft PowerPoint" to ${command}`],{encoding:'utf8',timeout:12000}).trim();}
if(!process.argv.includes('--restart-owned-only'))throw Error('Explicit --restart-owned-only required');
const names=as('get name of every presentation').split(', ').filter(Boolean);
for(const n of names){
 if(!owned.has(n))throw Error('Unowned presentation: '+n);
 if(n==='演示文稿2'){
  if(as(`get count of slides of presentation ${JSON.stringify(n)}`)!=='1')throw Error('Unexpected unsaved test content');
  const content=as(`get content of text range of text frame of every shape of slide 1 of presentation ${JSON.stringify(n)}`);
  if(content.split(', ').some(t=>t!==''&&t!=='missing value'))throw Error('The unsaved test placeholder has text');
 }else if(n!=='NativeOmmlMinimalProbe [Repaired]'){
  if(!as(`get full name of presentation ${JSON.stringify(n)}`).includes('/visualtexmac/apps/macos/build-logs/powerpoint-native-omml/'))throw Error('Unowned fixture path');
 }
}
for(const n of names){as(`close presentation ${JSON.stringify(n)} saving no`);console.log('Closed owned fixture '+n);}
if(as('get count of presentations')!=='0')throw Error('Office did not close all owned fixtures');
as('quit');console.log('Closed empty test host normally.');
