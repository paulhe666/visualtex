import { spawn } from "node:child_process";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { resolveChromiumExecutable } from "./browser_test_runtime.mjs";

const offset = process.pid % 500;
const port = 22500 + offset;
const debug = 23500 + offset;
const origin = `http://127.0.0.1:${port}`;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
async function wait(fn, timeout=12000){const s=Date.now();while(Date.now()-s<timeout){try{const v=await fn();if(v)return v;}catch{}await sleep(60);}throw new Error("timeout");}
class Cdp {
  next=0; pending=new Map();
  async connect(url){this.ws=new WebSocket(url);await new Promise((res,rej)=>{this.ws.addEventListener("open",res,{once:true});this.ws.addEventListener("error",rej,{once:true});});this.ws.addEventListener("message",e=>{const m=JSON.parse(e.data);const p=this.pending.get(m.id);if(!p)return;this.pending.delete(m.id);m.error?p.reject(new Error(m.error.message)):p.resolve(m.result);});await this.send("Runtime.enable");await this.send("Page.enable");}
  send(method,params={}){const id=++this.next;return new Promise((resolve,reject)=>{this.pending.set(id,{resolve,reject});this.ws.send(JSON.stringify({id,method,params}));});}
  async eval(expression){const r=await this.send("Runtime.evaluate",{expression,awaitPromise:true,returnByValue:true});if(r.exceptionDetails)throw new Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;}
  close(){this.ws?.close();}
}
const profile=await mkdtemp(join(tmpdir(),"vt-persist-layout-"));
const preview=spawn(process.execPath,["node_modules/vite/bin/vite.js","preview","--host","127.0.0.1","--port",String(port),"--strictPort"],{cwd:process.cwd(),stdio:"ignore"});
let chrome; const c=new Cdp();
try{
  await wait(async()=>{try{return (await fetch(origin)).ok}catch{return false}});
  chrome=spawn(resolveChromiumExecutable(),["--headless=new","--disable-gpu","--no-first-run","--no-default-browser-check",`--remote-debugging-port=${debug}`,`--user-data-dir=${profile}`,"--window-size=1440,900","about:blank"],{stdio:"ignore"});
  const target=await wait(async()=>{try{return (await (await fetch(`http://127.0.0.1:${debug}/json/list`)).json()).find(x=>x.type==="page")}catch{return null}});
  await c.connect(target.webSocketDebuggerUrl);
  await c.send("Page.navigate",{url:origin});
  await sleep(500);
  await c.eval(`(() => {
    localStorage.setItem("visualtex.onboarding.v3.completed","true");
    localStorage.setItem("visualtex.office.macos.first-run.v1.completed","true");
    localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed","true");
    const p=JSON.parse(localStorage.getItem("visualtex-editor")||"{}");
    const id=crypto.randomUUID();
    p.state={...(p.state||{}),lines:[{id,latex:""}],activeLineId:id};
    localStorage.setItem("visualtex-editor",JSON.stringify(p));
  })()`);
  await c.send("Page.reload",{ignoreCache:true});
  await wait(()=>c.eval("Boolean(document.querySelector('math-field') && document.querySelector('[data-formula-persistent-bold]'))"));
  await c.eval(`(() => {
    const f=document.querySelector("math-field");
    f.setValue("\\\\frac{\\\\placeholder{}}{\\\\placeholder{}}",{mode:"math",format:"latex",insertionMode:"replaceAll",selectionMode:"placeholder",silenceNotifications:true});
    f.focus();
    if(f.selection?.ranges?.every(([a,b])=>a===b)){f.position=0;f.executeCommand("moveToNextPlaceholder");}
    f.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({preventScroll:true});
  })()`);
  await sleep(120);
  const snapshot=()=>c.eval(`(() => {
    const f=document.querySelector("math-field");
    const line=f?.closest(".formula-line");
    const host=f?.closest(".mathfield-host");
    const editor=document.querySelector(".editor-surface");
    const scroll=document.querySelector(".editor-pane-scroll");
    const body=document.querySelector(".classic-editor-pane-body");
    const dock=document.querySelector(".classic-bottom-dock");
    const tabs=document.querySelector(".classic-bottom-tabs");
    const workspace=document.querySelector(".workspace");
    const app=document.querySelector(".app-shell");
    const ph=f?.shadowRoot?.querySelector(".visualtex-structural-placeholder.ML__selected,.ML__selected .visualtex-structural-placeholder,.visualtex-structural-placeholder");
    const rect=(e)=>e?Object.fromEntries(["top","bottom","left","right","width","height"].map(k=>[k,e.getBoundingClientRect()[k]])):null;
    return {
      scrollY, docScroll:document.documentElement.scrollTop, bodyScroll:document.body.scrollTop,
      field:rect(f), line:rect(line), host:rect(host), editor:rect(editor), scroll:rect(scroll), body:rect(body), dock:rect(dock), tabs:rect(tabs), workspace:rect(workspace), app:rect(app), placeholder:rect(ph),
      scrollTop:{editor:editor?.scrollTop, scroll:scroll?.scrollTop, body:body?.scrollTop},
      grids:{workspace:workspace?getComputedStyle(workspace).gridTemplateRows:"",body:body?getComputedStyle(body).gridTemplateRows:"",app:app?getComputedStyle(app).gridTemplateRows:""},
      mounts:{header:document.querySelector(".editor-pane-header")?.getBoundingClientRect().height, bottom:document.querySelector(".classic-bottom-workspace-controls")?.getBoundingClientRect().height},
      persistent:f?._mathfield?.visualTexPersistentTypingStyle??null
    };
  })()`);
  const click=async(sel)=>{
    const p=await c.eval(`(()=>{const e=document.querySelector(${JSON.stringify(sel)});e.scrollIntoView({block:"nearest",inline:"nearest"});const r=e.getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()`);
    await c.send("Input.dispatchMouseEvent",{type:"mousePressed",x:p.x,y:p.y,button:"left",buttons:1,clickCount:1});
    await c.send("Input.dispatchMouseEvent",{type:"mouseReleased",x:p.x,y:p.y,button:"left",buttons:0,clickCount:1});
    await sleep(100);
  };
  console.log("BEFORE",JSON.stringify(await snapshot(),null,2));
  for(const sel of ["[data-formula-persistent-bold]","[data-formula-persistent-italic]","[data-formula-persistent-color]","[data-formula-persistent-background]"]){
    await click(sel); console.log(sel,JSON.stringify(await snapshot(),null,2));
  }
} finally {c.close();chrome?.kill("SIGTERM");preview.kill("SIGTERM");await sleep(150);await rm(profile,{recursive:true,force:true});}
