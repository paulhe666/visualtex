import assert from "node:assert/strict";
import { writeFile } from "node:fs/promises";
import { withBrowserTest } from "./browser_test_runtime.mjs";

// Exercise the shipped editor and kernel through public field APIs and real
// pointer/key events. Store access only supplies isolated document fixtures.
await withBrowserTest(async ({ baseUrl, send, evaluate }) => {
  const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
  await send("Page.navigate", { url: baseUrl });
  await evaluate(`(async () => {
    while (!document.querySelector('math-field')?.shadowRoot) await new Promise(r => setTimeout(r, 40));
    await document.fonts.ready;
  })()`);
  let fixture = 0;
  let lastPointerDown = 0;
  const seed = async values => {
    const id = `selection-${++fixture}-`;
    await evaluate(`(async () => {
      const {useEditorStore: store, createFormulaLine: line} = await import('/src/stores/editorStore.ts');
      store.setState({ editorLayout: 'classic', sourceOpen: false,
        lines: ${JSON.stringify(values)}.map((value, i) => line(value, ${JSON.stringify(id)} + i)),
        activeLineId: ${JSON.stringify(id + "0")} });
    })()`);
    await wait(120);
    await evaluate("document.fonts.ready");
  };
  const key = async (key, code, vk, modifiers = 0) => {
    await send("Input.dispatchKeyEvent", {
      type: "keyDown", key, code, windowsVirtualKeyCode: vk, modifiers,
      ...(key.length === 1 && !modifiers ? { text: key } : {}),
    });
    await send("Input.dispatchKeyEvent", { type: "keyUp", key, code, windowsVirtualKeyCode: vk, modifiers });
    await wait(50);
  };
  const drag = async (from, to, release = to, modifiers = 0) => {
    // MathLive also detects multiple clicks by location/time. Each case is a
    // separate gesture, even when a new fixture uses the same screen position.
    await wait(Math.max(0, 550 - (Date.now() - lastPointerDown)));
    lastPointerDown = Date.now();
    await send("Input.dispatchMouseEvent", { type: "mousePressed", ...from, button: "left", buttons: 1, clickCount: 1, modifiers });
    if (to !== from) await send("Input.dispatchMouseEvent", { type: "mouseMoved", ...to, button: "left", buttons: 1, modifiers });
    await send("Input.dispatchMouseEvent", { type: "mouseReleased", ...release, button: "left", buttons: 0, clickCount: 1, modifiers });
    await wait(70);
  };
  const values = () => evaluate("[...document.querySelectorAll('math-field')].map(f => f.value)");
  const fieldState = () => evaluate(`(() => {
    const f = document.querySelector('math-field');
    return { selection: f.selection, position: f.position, selected: f.getValue(f.selection, 'latex'),
      rects: [...f.shadowRoot.querySelectorAll('.ML__selection')].map(node => {
        const r = node.getBoundingClientRect(); return {left:r.left,right:r.right,top:r.top,bottom:r.bottom};
      }) };
  })()`);
  const crossRowDrag = async reverse => {
    const points = await evaluate(`(() => {
      const fields = [...document.querySelectorAll('math-field')];
      return [[0,1],[1,2]].map(([row,offset]) => {
        const r = fields[row].getElementInfo(offset).bounds;
        return {x:r.right-0.2,y:(r.top+r.bottom)/2};
      });
    })()`);
    await drag(points[reverse ? 1 : 0], points[reverse ? 0 : 1]);
  };

  for (const reverse of [false, true]) {
    await seed(["abc", "def", "ghi"]);
    await crossRowDrag(reverse);
    await key("X", "KeyX", 88);
    assert.deepEqual(await values(), ["aXf", "ghi"], `cross-row typing (${reverse})`);
    await key("Y", "KeyY", 89);
    assert.deepEqual(await values(), ["aXYf", "ghi"], "caret stays before preserved suffix");
    await key("z", "KeyZ", 90, 4);
    assert.deepEqual(await values(), ["aXf", "ghi"]);
    await key("z", "KeyZ", 90, 4);
    assert.deepEqual(await values(), ["abc", "def", "ghi"]);
    await key("X", "KeyX", 88);
    assert.deepEqual(await values(), ["aXf", "ghi"], "undo restores partial cross-row selection");
  }
  await seed(["abc", "def", "ghi"]);
  await crossRowDrag(false);
  await key("Backspace", "Backspace", 8);
  assert.deepEqual(await values(), ["af", "ghi"]);
  await key("X", "KeyX", 88);
  assert.deepEqual(await values(), ["aXf", "ghi"], "delete leaves caret at join");

  await seed(["abc", "def"]);
  const newline = await evaluate(`(() => {
    const [a,b]=document.querySelectorAll('math-field');
    const left=a.getElementInfo(3).bounds,right=b.getElementInfo(1).bounds;
    return [{x:left.right-0.2,y:(left.top+left.bottom)/2},{x:right.left+0.2,y:(right.top+right.bottom)/2}];
  })()`);
  await drag(...newline);
  await key("Backspace", "Backspace", 8);
  assert.deepEqual(await values(), ["abcdef"], "deleting only a row boundary preserves both formulas");
  await key("X", "KeyX", 88);
  assert.deepEqual(await values(), ["abcXdef"]);

  await seed(["abc", "def", "ghi"]);
  await crossRowDrag(false);
  await evaluate(`(() => {
    const data = new DataTransfer();
    data.setData('application/x-visualtex-multiline-latex', JSON.stringify({version:1,lines:['X','Y']}));
    document.activeElement.dispatchEvent(new ClipboardEvent('paste',{bubbles:true,cancelable:true,clipboardData:data}));
  })()`);
  await wait(80);
  assert.deepEqual(await values(), ["aX", "Yf", "ghi"], "multi-row paste preserves both edges");
  await key("Z", "KeyZ", 90);
  assert.deepEqual(await values(), ["aX", "YZf", "ghi"], "paste caret precedes old suffix");
  await key("z", "KeyZ", 90, 4);
  await key("z", "KeyZ", 90, 4);
  assert.deepEqual(await values(), ["abc", "def", "ghi"]);
  await key("X", "KeyX", 88);
  assert.deepEqual(await values(), ["aXf", "ghi"], "paste undo restores document selection");

  await seed([String.raw`\textcolor{#112233}{a}bc`, String.raw`de\textcolor{#445566}{f}`]);
  await crossRowDrag(false);
  await evaluate(`document.querySelectorAll('math-field').forEach(field => {
    field.visualTexPersistentTypingStyle={bold:false,italic:null,color:'#778899',backgroundColor:null};
  })`);
  await key("X", "KeyX", 88);
  const colors = await evaluate(`(() => {
    const f=document.querySelector('math-field');
    return Object.fromEntries(Array.from({length:f.lastOffset},(_,i)=>
      [f.getValue(i,i+1,'latex-unstyled'),f.getElementInfo(i+1)?.style.color]));
  })()`);
  assert.equal(colors.a, '#112233', 'untouched prefix retains its color');
  assert.equal(colors.f, '#445566', 'untouched suffix retains its color');
  assert.equal(colors.X, '#778899', 'new input receives current typing style');

  await seed(["abc", "def", "ghi"]);
  await evaluate(`(async () => {
    const field = document.querySelector('math-field');
    field.focus();
    const deadline = performance.now() + 2000;
    while (document.activeElement !== field) {
      if (performance.now() > deadline) throw new Error('Field did not acquire focus');
      await new Promise(resolve => requestAnimationFrame(resolve));
    }
  })()`);
  await key("a", "KeyA", 65, 4);
  await key("Backspace", "Backspace", 8);
  assert.deepEqual(await values(), [""]);
  await key("z", "KeyZ", 90, 4);
  const restored = await evaluate("[...document.querySelectorAll('math-field')].map(f=>f.selection.ranges[0])");
  assert.deepEqual(restored, [[0,3],[0,3],[0,3]]);
  await key("X", "KeyX", 88);
  assert.deepEqual(await values(), ["X"]);

  await seed(["abcdef"]);
  const blank = await evaluate(`(() => {
    const f=document.querySelector('math-field');f.focus();f.position=2;
    const r=f.shadowRoot.querySelector('[part="content"]').getBoundingClientRect();
    return {x:r.right+25,y:(r.top+r.bottom)/2};
  })()`);
  await drag(blank, blank, blank, 8);
  assert.deepEqual((await fieldState()).selection, {ranges:[[2,6]],direction:"forward"});

  await seed(["abcdef"]);
  const reversePoints = await evaluate(`(() => {
    const f=document.querySelector('math-field');return [5,2].map(i=>{
      const r=f.getElementInfo(i).bounds;return {x:r.right-0.2,y:(r.top+r.bottom)/2};
    });
  })()`);
  await drag(...reversePoints);
  const beforeStyle = await fieldState();
  const bold = await evaluate(`(() => {const r=document.querySelector('[data-formula-selection-bold]').getBoundingClientRect();return {x:r.left+r.width/2,y:r.top+r.height/2};})()`);
  await drag(bold, bold);
  const afterStyle = await fieldState();
  assert.equal(beforeStyle.selection.direction, "backward");
  assert.deepEqual(afterStyle.selection, beforeStyle.selection);
  assert.equal(afterStyle.position, beforeStyle.position);

  for (const command of ["overset", "underset", "stackrel"]) {
    await seed([String.raw`a+\frac{\placeholder{}}{\placeholder{}}`]);
    await evaluate(`(() => {
      const f=document.querySelector('math-field');f.focus();f.position=0;
      f.visualTexInsertToolbarTemplate(${JSON.stringify("\\" + command + "{\\placeholder{}}{\\placeholder{}}")}, {format:'latex',selectionMode:'placeholder'});
    })()`);
    await key("q", "KeyQ", 81);
    assert.equal((await values())[0], `\\${command}{q}{\\placeholder{}}a+\\frac{\\placeholder{}}{\\placeholder{}}`);
  }
  await seed([String.raw`a+\frac{\placeholder{}}{\placeholder{}}`]);
  await evaluate(`(() => {const f=document.querySelector('math-field');f.focus();f.position=0;
    document.querySelector('button[data-command-id="overset"]').click();})()`);
  await wait(70);
  await key("q", "KeyQ", 81);
  assert.equal((await values())[0], String.raw`\overset{q}{\placeholder{}}a+\frac{\placeholder{}}{\placeholder{}}`);
  await key("z", "KeyZ", 90, 4);
  await key("w", "KeyW", 87);
  assert.equal((await values())[0], String.raw`\overset{w}{\placeholder{}}a+\frac{\placeholder{}}{\placeholder{}}`);

  await seed([String.raw`\overset{\placeholder{}}{\placeholder{}}`]);
  const nested = await evaluate(`(() => {
    const f=document.querySelector('math-field');f.focus();
    const offset=Array.from({length:f.lastOffset},(_,i)=>i+1)
      .find(i=>f.getElementInfo(i)?.latex==='\\\\placeholder{}');
    f.selection={ranges:[[offset-1,offset]],direction:'forward'};
    f.visualTexInsertToolbarTemplate('\\\\overset{\\\\placeholder{}}{\\\\placeholder{}}', {selectionMode:'placeholder'});
    f.executeCommand('typedText','q');return f.value;
  })()`);
  assert.equal(nested, String.raw`\overset{\placeholder{}}{\overset{q}{\placeholder{}}}`);

  await seed(["abcdef"]);
  const cancelled = await evaluate(`(() => {
    const f=document.querySelector('math-field');f.focus();f.selection={ranges:[[1,4]],direction:'backward'};
    const before={value:f.value,selection:f.selection,position:f.position};
    f.dispatchEvent(new CompositionEvent('compositionstart',{bubbles:true,composed:true,data:''}));
    f.setValue('aef',{silenceNotifications:true});
    f.dispatchEvent(new CompositionEvent('compositionend',{bubbles:true,composed:true,data:''}));
    return {before,after:{value:f.value,selection:f.selection,position:f.position}};
  })()`);
  assert.deepEqual(cancelled.after, cancelled.before, "cancelled IME restores direction and range");

  const structureCases = [
    ["accent", String.raw`p+\vec{ab}+q`, String.raw`\vec{ab}`, String.raw`\vec{ab}`, String.raw`\vec{ab}`, "q"],
    ["box", String.raw`p+\boxed{ab}+q`, "a", "b", String.raw`\boxed{ab}`, "q"],
    ["overbrace", String.raw`p+\overbrace{ab}+q`, String.raw`\overbrace{ab}`, String.raw`\overbrace{ab}`, String.raw`\overbrace{ab}`, "q"],
    ["underbrace", String.raw`p+\underbrace{ab}+q`, String.raw`\underbrace{ab}`, String.raw`\underbrace{ab}`, String.raw`\underbrace{ab}`, "q"],
    ["fraction branch", String.raw`p+\frac{ab}{cd}+q`, "a", "a", "a", "c"],
    ["fraction crossing branches", String.raw`p+\frac{ab}{cd}+q`, "b", "c", String.raw`\frac{ab}{cd}`],
    ["fraction to outside", String.raw`p+\frac{ab}{cd}+q`, "b", "q", String.raw`\frac{ab}{cd}+q`],
    ["script branch", String.raw`p+x_{ab}^{cd}+q`, "c", "c", "c", "a"],
    ["script crossing branches", String.raw`p+x_{ab}^{cd}+q`, "a", "c", String.raw`x_{ab}^{cd}`],
    ["script tail", String.raw`p+x_{ab}+q`, "b", "b", "b", "a"],
    ["radical tail", String.raw`p+\sqrt{ab}+q`, "b", "b", "b", "a"],
    ["indexed radical branch", String.raw`p+\sqrt[3]{ab}+q`, "a", "a", "a", "3"],
    ["radical to outside", String.raw`p+\sqrt[3]{ab}+q`, "b", "q", String.raw`\sqrt[3]{ab}+q`],
    ["nested fence to outside", String.raw`p+\left(\frac{ab}{cd}+e\right)+q`, "b", "q", String.raw`\left(\frac{ab}{cd}+e\right)+q`],
    ["matrix cell", String.raw`p+\begin{pmatrix}ab&cd\\ef&gh\end{pmatrix}+q`, "a", "a", "a", "e"],
    ["matrix row", String.raw`p+\begin{pmatrix}ab&cd\\ef&gh\end{pmatrix}+q`, "a", "d", "abcd", "e"],
    ["matrix to outside", String.raw`p+\begin{pmatrix}ab&cd\\ef&gh\end{pmatrix}+q`, "b", "q", String.raw`\begin{pmatrix}ab&cd\\ef&gh\end{pmatrix}+q`],
  ];
  const normalize = value => value.replace(/\s/g, "");
  for (const [name, latex, first, last, expected, excluded] of structureCases) {
    for (const reverse of [false, true]) {
      await seed([latex]);
      const geometry = await evaluate(`(() => {
        const f=document.querySelector('math-field');
        const atoms=Array.from({length:f.lastOffset+1},(_,i)=>f.getElementInfo(i));
        const point=(latex,right)=>{const r=atoms.find(a=>a?.latex===latex)?.bounds;if(!r)throw new Error('No bounds for '+latex);return {x:right?r.right-0.2:r.left+0.2,y:(r.top+r.bottom)/2};};
        const r=atoms.find(a=>a?.latex===${JSON.stringify(excluded ?? "")})?.bounds;
        return {points:[point(${JSON.stringify(first)},false),point(${JSON.stringify(last)},true)],
          excluded:r&&{left:r.left,right:r.right,top:r.top,bottom:r.bottom},
          decorations:[...f.shadowRoot.querySelectorAll('.ML__latex svg, .ML__latex .ML__box')].map(node=>{
            const b=node.getBoundingClientRect();return {left:b.left,right:b.right,top:b.top,bottom:b.bottom};
          })};
      })()`);
      const [left, right] = geometry.points;
      await drag(reverse ? right : left, reverse ? left : right);
      const state = await fieldState();
      assert.equal(normalize(state.selected), normalize(expected), `${name}, reverse=${reverse}: ${JSON.stringify(state)}`);
      assert.equal(state.selection.direction, reverse ? "backward" : "forward", name);
      const [start, end] = state.selection.ranges[0];
      assert.equal(state.position, reverse ? start : end, "model endpoints match the painted selection");
      assert.ok(state.rects.length > 0, `${name} has visible selection`);
      assert.ok(state.rects.every(r => Object.values(r).every(Number.isFinite) && r.right > r.left && r.bottom > r.top));
      if (["accent", "box", "overbrace", "underbrace"].includes(name)) {
        assert.ok(geometry.decorations.length > 0, `${name} has a visible decoration`);
        for (const r of geometry.decorations) {
          assert.ok(state.rects.some(s=>s.left <= r.left+2 && s.right >= r.right-2 &&
            s.top <= r.top+2 && s.bottom >= r.bottom-2), `${name} highlight covers its decoration: ${JSON.stringify({r,state})}`);
        }
      }
      if (geometry.excluded) {
        const r = geometry.excluded;
        assert.ok(state.rects.every(s => Math.min(s.right,r.right)-Math.max(s.left,r.left) <= 2 ||
          Math.min(s.bottom,r.bottom)-Math.max(s.top,r.top) <= 2), `${name} must not highlight an unselected sibling`);
      }
      if (name === "matrix row" && !reverse && process.env.SELECTION_SCREENSHOT) {
        const result = await send("Page.captureScreenshot", {format:"png"});
        await writeFile(process.env.SELECTION_SCREENSHOT, Buffer.from(result.data,"base64"));
      }
    }
  }
  // Exercise delayed insertion through the editor's public ref: a later
  // selection must not become the target of an earlier OCR request.
  const delayed = await evaluate(`(async () => {
    const React=(await import('/node_modules/.vite/deps/react.js')).default;
    const {createRoot}=(await import('/node_modules/.vite/deps/react-dom_client.js')).default;
    const {MathEditor}=await import('/src/editor/MathEditor.tsx');
    const {EditorSessionProvider,useDocumentSession}=await import('/src/history/EditorSession.tsx');
    const {useEditorStore:store,createFormulaLine:line}=await import('/src/stores/editorStore.ts');
    const wait=()=>new Promise(r=>setTimeout(r,100));
    const ref=React.createRef(),host=document.createElement('div');document.body.append(host);
    function Harness(){const state=store();useDocumentSession(ref);return React.createElement(MathEditor,{
      ref,lines:state.lines,activeLineId:state.activeLineId,formulaAlignment:'left',latexCodeFormat:'raw',zoom:100,
    });}
    const root=createRoot(host);
    root.render(React.createElement(EditorSessionProvider,null,React.createElement(Harness)));
    await wait();
    const reset=async()=>{store.setState({lines:['abc','def','ghi'].map((v,i)=>line(v,'delay-'+i)),activeLineId:'delay-0'});await wait();};
    const select=(anchor,focus)=>{const snapshot=ref.current.captureDocumentSnapshot();
      snapshot.activeLineId=focus.lineId;snapshot.documentSelection={anchor,focus};ref.current.restoreDocumentSelection(snapshot);};
    const result=[];
    try {
      for(const text of ['X',['X','Y'].join(String.fromCharCode(10))]) {
        await reset();const fields=host.querySelectorAll('math-field');fields[0].position=1;
        const target=ref.current.captureInsertionTarget();
        select({lineId:'delay-1',offset:1},{lineId:'delay-2',offset:2});
        ref.current.insertLatexAt(target,text,'ocr');await wait();
        result.push(store.getState().lines.map(l=>l.latex));
      }
      await reset();select({lineId:'delay-0',offset:1},{lineId:'delay-1',offset:2});
      const target=ref.current.captureInsertionTarget();
      select({lineId:'delay-1',offset:1},{lineId:'delay-2',offset:2});
      ref.current.insertLatexAt(target,'X','ocr');await wait();
      result.push(store.getState().lines.map(l=>l.latex));
      return result;
    } finally {root.unmount();host.remove();}
  })()`);
  assert.deepEqual(delayed, [["aXbc","def","ghi"],["aX","Ybc","def","ghi"],["aXf","ghi"]]);
  console.log(`Selection transactions, history, toolbar, IME and ${structureCases.length * 2} structural drag cases passed.`);
});
