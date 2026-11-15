import assert from "node:assert/strict";
import { withAppBrowser } from "../../shared/visualtex-app/testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function main() {
  await withAppBrowser(async ({ baseUrl, client }) => {
    // Chromium may expose the page target before its initial navigation has
    // settled. Navigate explicitly so subsequent Runtime.evaluate calls keep a
    // stable same-origin execution context.
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(700);

    const evaluate = async (expression) => {
      const result = await client.send("Runtime.evaluate", {
        expression,
        awaitPromise: true,
        returnByValue: true,
      });
      if (result.exceptionDetails) {
        throw new Error(
          result.exceptionDetails.exception?.description ||
            result.exceptionDetails.text ||
            "Runtime.evaluate failed",
        );
      }
      return result.result.value;
    };

    const typeKey = async (key, code, text = key, modifiers = 0) => {
      const specialVirtualKey =
        key === "Enter"
          ? 13
          : key === "Tab"
            ? 9
            : key === "Escape"
              ? 27
              : key === "ArrowDown"
                ? 40
                : key === " "
                  ? 32
                  : key === "\\"
                    ? 220
                    : 0;
      const virtualKey =
        specialVirtualKey ||
        (key.length === 1 ? key.toUpperCase().charCodeAt(0) : 0);
      const common = {
        key,
        code,
        modifiers,
        windowsVirtualKeyCode: virtualKey,
        nativeVirtualKeyCode: virtualKey,
      };
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        ...common,
        text,
        unmodifiedText: text,
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(35);
    };
    // Candidates are ranked by corpus frequency (\begin{array} outranks cases):
    // walk down the list to a candidate the way a user does.
    const selectSuggestion = async (command) => {
      for (let step = 0; step < 40; step += 1) {
        const current = await evaluate(
          `document.querySelector('#visualtex-native-input-suggestion-popover li.ML__popover__current')?.dataset.command ?? ''`,
        );
        if (current === command) return;
        await typeKey("ArrowDown", "ArrowDown", "");
        await sleep(40);
      }
      throw new Error(`${command} could not be selected in the suggestion list`);
    };

    await evaluate(`new Promise((resolve) => {
      const done = () => document.querySelector("math-field") ? resolve(true) : setTimeout(done, 30);
      done();
    })`);

    const casesLatex = String.raw`\begin{cases}x & x>0 \\ 0 & x\le 0\end{cases}`;
    const environmentProbe = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.setValue(${JSON.stringify(casesLatex)}, {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
      });
      const positions = [];
      for (let position = 0; position <= field.lastOffset; position += 1) {
        field.position = position;
        positions.push({
          position,
          environment: field._mathfield?.model?.parentEnvironment?.environmentName ?? null,
        });
      }
      return { value: field.value, positions };
    })()`);
    assert.match(environmentProbe.value, /\\begin\{cases\}/);
    assert.ok(
      environmentProbe.positions.some((entry) => entry.environment === "cases"),
      `MathLive did not expose the active cases environment: ${JSON.stringify(environmentProbe.positions)}`,
    );

    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.dispatchEvent(new KeyboardEvent("keydown", {
        key: "Escape",
        code: "Escape",
        bubbles: true,
        composed: true,
      }));
      field.executeCommand(["complete", "reject"]);
      field.mode = "math";
      field.setValue("", {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      field.position = field.lastOffset;
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus();
      return true;
    })()`);
    // Match the production editor's delayed focus repair window before sending
    // a physical Backslash key, otherwise the 0/80 ms focus callbacks can turn
    // the command introducer into ordinary math text in headless Chromium.
    await sleep(140);

    await typeKey("\\", "Backslash", "\\");
    for (const letter of "beg") {
      await typeKey(letter, `Key${letter.toUpperCase()}`);
    }
    await sleep(120);

    const begSuggestion = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      const panel = document.querySelector('#visualtex-native-input-suggestion-popover');
      return {
        raw: Array.from(field.shadowRoot?.querySelectorAll('.ML__raw-latex') ?? [])
          .filter((node) => !node.classList.contains('ML__suggestion'))
          .map((node) => node.textContent ?? "")
          .join(""),
        visible: panel?.classList.contains('is-visible') ?? false,
        commands: Array.from(panel?.querySelectorAll('li[data-command]') ?? [])
          .map((item) => item.dataset.command ?? ''),
      };
    })()`);
    assert.equal(begSuggestion.raw.replace(/\s+/g, ""), "\\beg");
    assert.equal(
      begSuggestion.visible,
      true,
      "cases suggestion did not appear at the early \\beg prefix",
    );
    assert.ok(
      begSuggestion.commands.includes("\\begin{cases}"),
      `cases is not among the \\beg candidates: ${JSON.stringify(begSuggestion.commands)}`,
    );

    for (const letter of "in") {
      await typeKey(letter, `Key${letter.toUpperCase()}`);
    }
    await sleep(120);
    await selectSuggestion("\\begin{cases}");

    const casesSuggestion = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      const panel = document.querySelector('#visualtex-native-input-suggestion-popover');
      const current = panel?.querySelector('li.ML__popover__current');
      const preview = current?.querySelector('.ML__popover__command');
      const style = panel ? getComputedStyle(panel) : null;
      return {
        value: field.value,
        raw: Array.from(field.shadowRoot?.querySelectorAll('.ML__raw-latex') ?? [])
          .filter((node) => !node.classList.contains('ML__suggestion'))
          .map((node) => node.textContent ?? "")
          .join(""),
        visible: panel?.classList.contains('is-visible') ?? false,
        command: current?.dataset.command ?? '',
        previewText: preview?.textContent?.trim() ?? '',
        previewHtml: preview?.innerHTML ?? '',
        background: style?.backgroundColor ?? '',
        borderRadius: style?.borderRadius ?? '',
        width: panel?.getBoundingClientRect().width ?? 0,
      };
    })()`);
    assert.equal(casesSuggestion.raw.replace(/\s+/g, ""), "\\begin");
    assert.equal(
      casesSuggestion.visible,
      true,
      "cases did not show the VisualTeX native-style MathLive suggestion popover",
    );
    assert.equal(casesSuggestion.command, "\\begin{cases}");
    assert.ok(casesSuggestion.previewHtml.length > 0, "cases suggestion preview is empty");
    assert.ok(casesSuggestion.width >= 240, `cases suggestion popover is unexpectedly narrow: ${casesSuggestion.width}`);
    assert.doesNotMatch(
      casesSuggestion.value,
      /\\begin\{cases\}/,
      "cases committed before the user accepted the native suggestion",
    );

    await typeKey(" ", "Space", " ");
    await sleep(100);
    const typedCases = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        value: field.value,
        raw: Array.from(field.shadowRoot?.querySelectorAll('.ML__raw-latex') ?? [])
          .filter((node) => !node.classList.contains('ML__suggestion'))
          .map((node) => node.textContent ?? "")
          .join(""),
        environment: field._mathfield?.model?.parentEnvironment?.environmentName ?? null,
      };
    })()`);
    assert.match(
      typedCases.value,
      /\\begin\{cases\}\\placeholder\{\} & \\placeholder\{\}\\end\{cases\}/,
      "Space did not accept cases as a real editable environment",
    );
    assert.equal(typedCases.raw, "", "accepted cases input remained in raw-LaTeX mode");

    // The synthetic environment candidate is part of the same frequency-ranked
    // native pool as ordinary MathLive `\\b...` commands. (The one use recorded
    // above is a small boost; \\begin{array} and the common \\b commands still
    // come first.)
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.setValue("", { insertionMode: "replaceAll" });
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus();
      return true;
    })()`);
    await typeKey("\\", "Backslash", "\\");
    await typeKey("b", "KeyB");
    await sleep(140);

    const bPool = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      const panel = document.querySelector('#visualtex-native-input-suggestion-popover');
      const items = Array.from(panel?.querySelectorAll('li[data-command]') ?? []);
      return {
        raw: Array.from(field.shadowRoot?.querySelectorAll('.ML__raw-latex') ?? [])
          .filter((node) => !node.classList.contains('ML__suggestion'))
          .map((node) => node.textContent ?? "")
          .join(""),
        visible: panel?.classList.contains('is-visible') ?? false,
        selected: panel?.querySelector('li.ML__popover__current')?.dataset.command ?? '',
        commands: items.map((item) => item.dataset.command ?? ''),
      };
    })()`);
    assert.equal(bPool.raw.replace(/\s+/g, ""), "\\b");
    assert.equal(bPool.visible, true, "native candidate pool is not visible for \\b");
    assert.ok(
      bPool.commands.includes("\\begin{cases}"),
      `cases is missing from the normal \\b candidate pool: ${JSON.stringify(bPool.commands)}`,
    );
    assert.ok(
      bPool.commands.some((command) => command !== "\\begin{cases}" && command.startsWith("\\b")),
      `the \\b pool did not retain ordinary MathLive b-commands: ${JSON.stringify(bPool.commands)}`,
    );

    await typeKey("e", "KeyE");
    await sleep(100);
    const bePool = await evaluate(`(() => {
      const panel = document.querySelector('#visualtex-native-input-suggestion-popover');
      return {
        selected: panel?.querySelector('li.ML__popover__current')?.dataset.command ?? '',
        commands: Array.from(panel?.querySelectorAll('li[data-command]') ?? [])
          .map((item) => item.dataset.command ?? ''),
      };
    })()`);
    assert.ok(bePool.commands.includes("\\begin{cases}"), "cases disappeared at the \\be prefix");

    await typeKey("g", "KeyG");
    await sleep(100);
    const begPool = await evaluate(`(() => {
      const panel = document.querySelector('#visualtex-native-input-suggestion-popover');
      return {
        selected: panel?.querySelector('li.ML__popover__current')?.dataset.command ?? '',
        commands: Array.from(panel?.querySelectorAll('li[data-command]') ?? [])
          .map((item) => item.dataset.command ?? ''),
      };
    })()`);
    assert.ok(begPool.commands.includes("\\begin{cases}"), "cases disappeared at the \\beg prefix");

    await selectSuggestion("\\begin{cases}");
    await typeKey(" ", "Space", " ");
    await sleep(100);
    const casesFromRankedPrefix = await evaluate(`document.querySelector("math-field")?.value ?? ""`);
    assert.match(
      casesFromRankedPrefix,
      /\\begin\{cases\}\\placeholder\{\} & \\placeholder\{\}\\end\{cases\}/,
      "accepting the cases candidate from a short prefix failed",
    );

    await typeKey("\\", "Backslash", "\\");
    for (const letter of "frac") {
      await typeKey(letter, `Key${letter.toUpperCase()}`);
    }
    await typeKey("Enter", "Enter", "\r");
    const candidateCommit = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        value: field.value,
        rowSeparators: (field.value.match(/\\\\\\\\/g) ?? []).length,
      };
    })()`);
    assert.match(
      candidateCommit.value,
      /\\frac/,
      "Enter did not preserve command-candidate priority inside cases",
    );
    assert.equal(
      candidateCommit.rowSeparators,
      0,
      "confirming a command candidate unexpectedly added a cases row",
    );

    await typeKey("Enter", "Enter", "\r");
    const afterCasesEnter = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        value: field.value,
        rowSeparators: (field.value.match(/\\\\\\\\/g) ?? []).length,
      };
    })()`);
    assert.equal(
      afterCasesEnter.rowSeparators,
      1,
      `Enter inside cases did not add exactly one case row: ${afterCasesEnter.value}`,
    );

    const lineCountBeforeOuterEnter = await evaluate(
      `document.querySelectorAll(".formula-line").length`,
    );
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.position = field.lastOffset;
      field.selection = { ranges: [[field.lastOffset, field.lastOffset]], direction: "none" };
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus();
      return true;
    })()`);
    await typeKey("Enter", "Enter", "\r");
    const lineCountAfterOuterEnter = await evaluate(
      `document.querySelectorAll(".formula-line").length`,
    );
    assert.equal(
      lineCountAfterOuterEnter,
      lineCountBeforeOuterEnter + 1,
      "Enter outside the outermost cases did not create a new VisualTeX formula line",
    );

    const toolbarProbe = await evaluate(`(() => {
      const two = document.querySelector('.template-button[data-command-id="cases"]');
      const three = document.querySelector('.template-button[data-command-id="cases-three"]');
      return {
        two: Boolean(two),
        three: Boolean(three),
        twoEnlarged: two?.classList.contains('is-enlarged-cases-preview') ?? false,
        threeEnlarged: three?.classList.contains('is-enlarged-cases-preview') ?? false,
      };
    })()`);
    assert.deepEqual(toolbarProbe, {
      two: true,
      three: true,
      twoEnlarged: true,
      threeEnlarged: true,
    });

    console.log(
      `Windows cases regression passed: native popover ${Math.round(casesSuggestion.width)}px, Space acceptance, candidate priority, Enter row insertion, toolbar previews`,
    );
  });
}

await main();
