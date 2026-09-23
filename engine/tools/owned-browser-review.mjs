// Test-only browser lifetime helper. It never attaches to the user's browser.
import {spawn} from 'node:child_process';
import {existsSync,mkdtempSync,rmSync} from 'node:fs';
import {createServer} from 'node:http';
import {tmpdir} from 'node:os';
import {join} from 'node:path';

class Client {
  next=1;pending=new Map();listeners=new Set();
  async connect(url) {
    this.socket=new WebSocket(url);
    await new Promise((ok,bad)=>{
      const t=setTimeout(()=>bad(new Error('Owned browser connection timeout')),15000);
      this.socket.addEventListener('open',()=>{clearTimeout(t);ok();},{once:true});
      this.socket.addEventListener('error',()=>{clearTimeout(t);bad(new Error('Owned browser connection error'));},{once:true});
    });
    this.socket.addEventListener('message',e=>{
      const m=JSON.parse(e.data);
      if(m.id){const p=this.pending.get(m.id);if(!p)return;this.pending.delete(m.id);clearTimeout(p.t);
        if(m.error)p.bad(new Error(m.error.message));else p.ok(m.result);
      }else for(const p of [...this.listeners])if(p.method===m.method&&p.sessionId===m.sessionId){
        this.listeners.delete(p);clearTimeout(p.t);p.ok(m.params);
      }
    });
    this.socket.addEventListener('close',()=>this.rejectAll());
  }
  send(method,params={},sessionId){
    const id=this.next++;
    return new Promise((ok,bad)=>{const t=setTimeout(()=>{this.pending.delete(id);bad(new Error(`CDP timeout: ${method}`));},20000);
      this.pending.set(id,{ok,bad,t});this.socket.send(JSON.stringify({id,method,params,sessionId}));});
  }
  wait(method,sessionId){return new Promise((ok,bad)=>{const p={method,sessionId,ok,bad};
    p.t=setTimeout(()=>{this.listeners.delete(p);bad(new Error(`CDP event timeout: ${method}`));},15000);this.listeners.add(p);});}
  rejectAll(){for(const p of [...this.pending.values(),...this.listeners]){clearTimeout(p.t);p.bad(new Error('Owned browser disposed'));}this.pending.clear();this.listeners.clear();}
  close(){this.rejectAll();this.socket?.close();}
}

/** assets is a fixed map of URL paths to {type,bytes}; no filesystem URL mapping. */
export async function withOwnedBrowser(assets,html,run,options={}) {
  const browser=[process.env.VISUALTEX_ENGINE_REVIEW_BROWSER,
    'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe','C:/Program Files/Microsoft/Edge/Application/msedge.exe',
    'C:/Program Files/Google/Chrome/Application/chrome.exe'].find(p=>p&&existsSync(p));
  if(!browser)throw new Error('No browser executable for isolated view tests');
  const requests=new Map();
  const server=createServer((req,res)=>{
    const path=new URL(req.url,'http://127.0.0.1').pathname;
    requests.set(path,(requests.get(path)??0)+1);res.setHeader('Cache-Control','no-store');res.setHeader('X-Content-Type-Options','nosniff');
    if(path==='/'){
      res.setHeader('Content-Type','text/html; charset=utf-8');
      res.setHeader('Content-Security-Policy',"default-src 'none'; script-src 'self' 'wasm-unsafe-eval'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; connect-src 'self'");res.end(html);
    }else if(assets.has(path)){const a=assets.get(path);res.setHeader('Content-Type',a.type);res.end(a.bytes);}
    else{res.statusCode=404;res.end('Not found');}
  });
  await new Promise((ok,bad)=>{server.once('error',bad);server.listen(0,'127.0.0.1',ok);});
  const profile=mkdtempSync(join(tmpdir(),'visualtex-owned-svg-review-'));
  let child,client;let stderr='';
  try{
    child=spawn(browser,[...(options.visible?['--window-size=1100,850','--window-position=20,20']:['--headless=new']),
      '--remote-debugging-port=0','--remote-debugging-address=127.0.0.1',
      '--no-first-run','--no-default-browser-check','--disable-background-networking',`--user-data-dir=${profile}`,'about:blank'],
      {stdio:['ignore','ignore','pipe'],windowsHide:!options.visible});
    const endpoint=await new Promise((ok,bad)=>{
      const t=setTimeout(()=>bad(new Error(`Owned browser startup timeout: ${stderr.slice(-1000)}`)),20000);
      child.once('error',e=>{clearTimeout(t);bad(e);});child.once('exit',c=>{clearTimeout(t);bad(new Error(`Owned browser exited ${c}`));});
      child.stderr.on('data',d=>{stderr=(stderr+d.toString()).slice(-16000);
        const match=stderr.match(/DevTools listening on (ws:\/\/127\.0\.0\.1:\d+\/devtools\/browser\/[^\s]+)/u);
        if(match){clearTimeout(t);ok(match[1]);}});
    });
    client=new Client();await client.connect(endpoint);
    const browserVersion=await client.send('Browser.getVersion');
    const {targetId}=await client.send('Target.createTarget',{url:'about:blank'});
    const {sessionId}=await client.send('Target.attachToTarget',{targetId,flatten:true});
    const send=(method,params={})=>client.send(method,params,sessionId);
    await send('Page.enable');const loaded=client.wait('Page.loadEventFired',sessionId);
    const origin=`http://127.0.0.1:${server.address().port}`;
    await send('Page.navigate',{url:origin});await loaded;
    const evaluate=async expression=>{
      const result=await send('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true});
      if(result.exceptionDetails)throw new Error(result.exceptionDetails.exception?.description??result.exceptionDetails.text);
      return result.result.value;
    };
    return await run({send,evaluate,browserVersion,origin,requests,ownedProcessId:child.pid});
  }finally{
    if(client){try{await client.send('Browser.close');}catch{}client.close();}
    if(child&&child.exitCode===null){const closed=await Promise.race([new Promise(ok=>child.once('exit',()=>ok(true))),new Promise(ok=>setTimeout(()=>ok(false),3000))]);if(!closed)child.kill();}
    server.closeAllConnections();await new Promise(ok=>server.close(ok));
    try{rmSync(profile,{recursive:true,force:true,maxRetries:3,retryDelay:100});}catch{/* only owned temporary evidence remains */}
  }
}
