import {cp254FontScopeControls} from './cp254_font_scope_candidate.ts';
const probeSources=[...cp254FontScopeControls.map(x=>x.source),...["\\boldsymbol{\\frac{\\mathsf{A}}{\\mathrm{B}}}","\\boldsymbol{\\hat{\\mathnormal{x}}}","\\boldsymbol{x_{\\mathrm{i}}^{\\mathsf{2}}}","\\boldsymbol{\\text{rate}+\\mathrm{A}}","\\boldsymbol{\\left(\\mathsf{A}\\right)}","\\mathbf{\\hat{\\mathnormal{x}}}"]];
import {loadVisualTexMathLiveKernel,visualTexMathLiveContourIntegralCompatibility} from '../vite.mathliveIntegralCompatibility.ts';
const plugin=visualTexMathLiveContourIntegralCompatibility();
const kernel=(await (plugin.transform as Function).call({},loadVisualTexMathLiveKernel(),'\u0000visualtex-mathlive-kernel')).code;
import {createServer} from 'node:http';
import {readFileSync,writeFileSync,existsSync} from 'node:fs';
import {spawn} from 'node:child_process';
import {resolveChromiumExecutable,createBrowserProfilePath} from './browser_test_runtime.mjs';
const out='../../artifacts/conversion-symbol-audit-20261002/cp254-font-scope-production20.json';
if(existsSync(out))throw new Error('Refusing overwrite');
const port=19800+process.pid%500, debug=port+1000, url='http://127.0.0.1:'+port;
const server=createServer((req,res)=>{if(req.url==='/mathlive.mjs'){res.setHeader('Content-Type','application/javascript');res.end(kernel);}else if(req.url?.startsWith('/fonts/')){const name=req.url.slice(7);if(!/^[a-zA-Z0-9._-]+$/.test(name)){res.writeHead(404).end();return;}try{res.end(readFileSync('node_modules/mathlive/fonts/'+name));}catch{res.writeHead(404).end();}}else{res.setHeader('Content-Type','text/html');res.end('<html><body><script type="module">const m=await import("/mathlive.mjs");m.MathfieldElement.fontsDirectory="/fonts";m.MathfieldElement.soundsDirectory=null;window.m=m;window.ready=true;</script></body></html>');}});
await new Promise(r=>server.listen(port,'127.0.0.1',r));
const browser=spawn(resolveChromiumExecutable(),['--headless=new','--disable-gpu','--no-first-run','--no-default-browser-check','--remote-debugging-address=127.0.0.1','--remote-debugging-port='+debug,'--user-data-dir='+createBrowserProfilePath('cp253-accent-editor-production'),url],{stdio:'ignore'});
let socket;const pending=new Map();let id=0;
async function wait(fn){const end=Date.now()+15000;while(Date.now()<end){try{const x=await fn();if(x)return x;}catch{}await new Promise(r=>setTimeout(r,100));}throw new Error('Timed out');}
try{
 const page=await wait(async()=>{const a=await(await fetch('http://127.0.0.1:'+debug+'/json/list')).json();return a.find(x=>x.type==='page'&&x.url.startsWith(url));});
 socket=new WebSocket(page.webSocketDebuggerUrl);
 await new Promise((r,j)=>{socket.addEventListener('open',r,{once:true});socket.addEventListener('error',j,{once:true});});
 socket.addEventListener('message',e=>{const x=JSON.parse(e.data);if(x.id){const p=pending.get(x.id);pending.delete(x.id);x.error?p.reject(x.error):p.resolve(x.result);}});
 const send=(method,params={})=>new Promise((resolve,reject)=>{const key=++id;pending.set(key,{resolve,reject});socket.send(JSON.stringify({id:key,method,params}));});
 const evaluate=async expression=>{const r=await send('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true});if(r.exceptionDetails)throw new Error(JSON.stringify(r.exceptionDetails));return r.result.value;};
 await wait(()=>evaluate('window.ready===true'));
 const rows=await evaluate(`(async()=>{
 const sources=${JSON.stringify(probeSources)};
 const rows=[];
 for(const source of sources){
  const field=new window.m.MathfieldElement({mathVirtualKeyboardPolicy:'manual'});document.body.append(field);
  field.setValue(source,{format:'latex',silenceNotifications:true});
  const initial=field.getValue('latex');const expanded=field.getValue('latex-expanded');
  const atoms=(field._mathfield?.model?.root?.children||[]).map(a=>({type:a.type,command:a.command,value:a.value,style:a.style,body:a.body?.map(b=>({type:b.type,command:b.command,value:b.value,style:b.style}))}));
  field.position=field.lastOffset;field.insert('+z',{format:'latex'});
  const edited=field.getValue('latex');
  const markup=field.shadowRoot?.innerHTML||'';
  rows.push({source,initial,expanded,edited,atoms,markerCount:[...initial].filter(c=>c==='⁡').length,markupHasMarker:markup.includes('⁡')});field.remove();
 }
 return rows;
 })()`);
 writeFileSync(out,JSON.stringify({browser:resolveChromiumExecutable(),rows},null,2));console.log(JSON.stringify({count:rows.length}));
 await send('Browser.close');
}finally{socket?.close();if(browser.exitCode===null)browser.kill();server.close();}
