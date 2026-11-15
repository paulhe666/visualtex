import assert from "node:assert/strict";
import process from "node:process";
import {
  canonicalSourceWrapperCommand,
  sourceWrapperKinds,
} from "../../shared/visualtex-app/src/editor/sourceWrapperRegistry.ts";
import { withAppBrowser } from "../../shared/visualtex-app/testing/browser_test_runtime.mjs";

const requestedCommand = process.argv[2] || "";
const cases = [...sourceWrapperKinds.entries()]
  .map(([command, kind]) => ({
    command,
    canonical: canonicalSourceWrapperCommand(command),
    kind,
  }))
  .filter((testCase) => !requestedCommand || testCase.command === requestedCommand);

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

await withAppBrowser(async ({ baseUrl, client }) => {
  await client.send("Page.navigate", { url: baseUrl });
  await sleep(1800);

  const evaluate = async (expression) => {
    const result = await client.send("Runtime.evaluate", {
      expression,
      awaitPromise: true,
      returnByValue: true,
    });
    if (result.exceptionDetails) {
      throw new Error(
        result.exceptionDetails.exception?.description ??
          result.exceptionDetails.text ??
          "Runtime.evaluate failed",
      );
    }
    return result.result.value;
  };

  await evaluate(
    '(() => {' +
      'localStorage.setItem("visualtex.onboarding.v3.completed","true");' +
      'localStorage.setItem("visualtex.onboarding.windows.desktop.v1.1.0.completed","true");' +
      'const storageKey="visualtex-editor";' +
      'const persisted=JSON.parse(localStorage.getItem(storageKey)||"{}");' +
      'const line={id:crypto.randomUUID(),latex:""};' +
      'persisted.state={...(persisted.state||{}),lines:[line],activeLineId:line.id,personalize:false,suggestionCount:10,inputBehavior:{...(persisted.state?.inputBehavior||{}),autoExitAccent:false,autoExitWrapperCommand:false}};' +
      'localStorage.setItem(storageKey,JSON.stringify(persisted));' +
    '})()',
  );
  await client.send("Page.reload", { ignoreCache: true });
  await sleep(1400);

  const mountedStart = Date.now();
  let mounted = false;
  while (Date.now() - mountedStart < 12000) {
    mounted = await evaluate(
      'Boolean(document.querySelector("math-field.visual-mathfield")?.shadowRoot)',
    );
    if (mounted) break;
    await sleep(60);
  }
  assert.equal(mounted, true, "VisualTeX mathfield mounted");

  const key = async (value, code, virtualKeyCode) => {
    const common = {
      key: value,
      code,
      windowsVirtualKeyCode: virtualKeyCode,
      nativeVirtualKeyCode: virtualKeyCode,
    };
    await client.send("Input.dispatchKeyEvent", {
      type: "keyDown",
      ...common,
      text: value.length === 1 ? value : undefined,
      unmodifiedText: value.length === 1 ? value : undefined,
    });
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
    await sleep(12);
  };

  const typeText = async (text) => {
    for (const character of text) {
      const code =
        character === "\\"
          ? "Backslash"
          : /^[a-zA-Z]$/.test(character)
            ? "Key" + character.toUpperCase()
            : "";
      const virtualKeyCode =
        character === "\\" ? 220 : character.toUpperCase().charCodeAt(0);
      await key(character, code, virtualKeyCode);
    }
  };

  const clearAndFocus = async () => {
    const ok = await evaluate(
      '(() => {' +
        'const field=document.querySelector("math-field.visual-mathfield");' +
        'if(!field?.shadowRoot)return false;' +
        'field.setValue("",{mode:"math",format:"latex",insertionMode:"replaceAll",selectionMode:"after",silenceNotifications:true});' +
        'field.dataset.visualtexAutoExitAccent="false";' +
        'field.dataset.visualtexAutoExitWrapperCommand="false";' +
        'field.dispatchEvent(new InputEvent("input",{bubbles:true,composed:true,inputType:"insertText"}));' +
        'field.focus();field.position=field.lastOffset;' +
        'field.shadowRoot?.querySelector("[part=\\\"keyboard-sink\\\"]")?.focus({preventScroll:true});' +
        'return field.isConnected&&field.value===""&&field.hasFocus();' +
      '})()',
    );
    if (!ok) throw new Error("Could not prepare empty focused mathfield");
    await sleep(25);
  };

  const readState = () =>
    evaluate(
      '(() => {' +
        'const field=document.querySelector("math-field.visual-mathfield");' +
        'const root=field?.shadowRoot;' +
        'return {' +
          'value:field?.value??"",mode:field?.mode??"",' +
          'sourceCommand:field?.dataset.visualtexSourceWrapperCommand??"",' +
          'sourceKind:field?.dataset.visualtexSourceWrapperKind??"",' +
          'sourceContent:field?.dataset.visualtexSourceWrapperContent??"",' +
          'placeholderCount:root?.querySelectorAll(".visualtex-structural-placeholder, .ML__placeholder").length??-1,' +
          'raw:[...(root?.querySelectorAll(".ML__raw-latex")??[])].filter((node)=>!node.classList.contains("ML__suggestion")).map((node)=>node.textContent??"").join("")' +
        '};' +
      '})()',
    );

  const failures = [];
  const passed = [];

  for (const testCase of cases) {
    try {
      await clearAndFocus();
      await typeText(testCase.command);
      const completionBeforeSpace = await evaluate(
        '(() => {' +
          'const field=document.querySelector("math-field.visual-mathfield");' +
          'const completion=field?._mathfield?.visualTexCompletion;' +
          'return {raw:[...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex")??[])].filter((node)=>!node.classList.contains("ML__suggestion")).map((node)=>node.textContent??"").join(""),suggestionIndex:field?._mathfield?.suggestionIndex??-1,suggestions:completion?.suggestions??[],records:(completion?.suggestions??[]).map((latex)=>{const record=field?._mathfield?.visualTexCompletionRecords?.get(latex);return {latex,command:record?.command??"",id:record?.id??""};})};' +
        '})()',
      );
      await key(" ", "Space", 32);
      await sleep(55);
      const shell = await readState();
      const expectedShell = testCase.canonical + "{}";
      if (
        shell.value !== expectedShell ||
        shell.raw !== expectedShell ||
        shell.mode !== "latex" ||
        shell.sourceCommand !== testCase.canonical ||
        shell.sourceKind !== testCase.kind ||
        shell.sourceContent !== "" ||
        shell.placeholderCount !== 0
      ) {
        throw new Error(
          "source shell mismatch: " +
            JSON.stringify({ expectedShell, completionBeforeSpace, shell }),
        );
      }

      await key("x", "KeyX", 88);
      await sleep(35);
      const filled = await readState();
      if (
        filled.mode !== "latex" ||
        filled.sourceCommand !== testCase.canonical ||
        filled.sourceKind !== testCase.kind ||
        filled.sourceContent !== "x" ||
        filled.placeholderCount !== 0 ||
        !filled.raw.includes("{x}")
      ) {
        throw new Error("filled source shell mismatch: " + JSON.stringify(filled));
      }

      await key("Enter", "Enter", 13);
      await sleep(55);
      const committed = await readState();
      if (
        committed.mode !== "math" ||
        committed.sourceCommand ||
        committed.sourceKind ||
        committed.placeholderCount !== 0 ||
        committed.value.includes("\\placeholder")
      ) {
        throw new Error("Enter commit mismatch: " + JSON.stringify(committed));
      }
      passed.push(testCase.command);
    } catch (error) {
      failures.push({
        command: testCase.command,
        canonical: testCase.canonical,
        kind: testCase.kind,
        error: error instanceof Error ? error.message : String(error),
      });
    }
  }

  console.log(
    JSON.stringify(
      {
        total: cases.length,
        passed: passed.length,
        failed: failures.length,
        failures,
      },
      null,
      2,
    ),
  );
  assert.deepEqual(
    failures,
    [],
    "Some source-first commands still bypass the source shell",
  );
  console.log(
    "Source-wrapper browser regression passed: " +
      passed.length +
      "/" +
      cases.length +
      " commands.",
  );
});
