// candidates ranked by the user's history

export const names = ["usage-ranking"];

// state the page starts with (merged into the editor's persisted state)
export const storage = () => ({ editorLayout: "standard", sidebarOpen: true, personalize: true, usage: {
    sqrt: { commandId: "sqrt", useCount: 18, lastUsedAt: 1800, recentUses: [], acceptedPrefixes: {}, contextCounts: { toolbar: 18 }, pinned: false },
    "formula-tile-gaussian-integral": { commandId: "formula-tile-gaussian-integral", useCount: 12, lastUsedAt: 1700, recentUses: [], acceptedPrefixes: {}, contextCounts: { toolbar: 12 }, pinned: false },
    "mathlive-native:\\beth": { commandId: "mathlive-native:\\beth", useCount: 25, lastUsedAt: 1900, recentUses: [], acceptedPrefixes: { be: 25 }, contextCounts: { candidate: 25 }, pinned: false },
    beta: { commandId: "beta", useCount: 2, lastUsedAt: Date.now() - 400, recentUses: [Date.now() - 1200, Date.now() - 400], acceptedPrefixes: { be: 2 }, contextCounts: { candidate: 2 }, pinned: false },
  }, inputBehavior: { autoEscapeShortcuts: true, autoExitSuperscript: true, autoExitSubscript: true, autoExitAccent: true, autoExitWrapperCommand: true, showStructuredCommandSuggestions: false, showOtherCommandSuggestions: false } });

export async function run({ clearField, client, evaluate, key, sleep, typeText, waitForEvaluation }) {
  await evaluate(`(() => {
        if (!document.querySelector(".formula-toolbar")) {
          document.querySelector(".sidebar-toggle")?.click();
        }
        return true;
      })()`);
  const commonCommandState = await waitForEvaluation(`(() => {
        const buttons = [...document.querySelectorAll(
          '.template-strip > [data-command-id]',
        )];
        const persisted = JSON.parse(
          localStorage.getItem("visualtex-editor") || "{}",
        );
        return {
          ready:
            buttons.length === 45 &&
            buttons[0]?.dataset.commandId === "frac" &&
            persisted.state?.usage?.sqrt?.contextCounts?.toolbar === 18,
          firstCommandId: buttons[0]?.dataset.commandId ?? "",
          usage: persisted.state?.usage?.sqrt ?? null,
        };
      })()`, "toolbar common commands keep their fixed order despite usage");

  await evaluate(`document.querySelector('[data-toolbar-view="tiles"]').click()`);
  const commonTileState = await waitForEvaluation(`(() => {
        const tiles = [...document.querySelectorAll(
          '.formula-tile-list > .formula-tile-item > .formula-tile-button',
        )];
        return {
          ready:
            tiles.length === 10 &&
            tiles[0]?.dataset.formulaTileId === "quadratic-formula",
          firstTileId: tiles[0]?.dataset.formulaTileId ?? "",
        };
      })()`, "common formula tiles keep their fixed order despite usage");
  await evaluate(`document.querySelector(
        '[data-formula-tile-id="gaussian-integral"]',
      ).click()`);
  const tileUsageState = await waitForEvaluation(`(() => {
        const persisted = JSON.parse(
          localStorage.getItem("visualtex-editor") || "{}",
        );
        const usage = persisted.state?.usage?.["formula-tile-gaussian-integral"];
        return {
          ready:
            usage?.useCount === 13 &&
            usage?.contextCounts?.toolbar === 13,
          usage,
        };
      })()`, "formula tile click frequency persistence");

  await sleep(180);
  await clearField();
  await sleep(120);
  await clearField();
  await typeText("\\be");
  const nativeRankState = await waitForEvaluation(`(() => {
        const source = document.getElementById("mathlive-suggestion-popover");
        const stable = document.getElementById(
          "visualtex-native-input-suggestion-popover",
        );
        const sourceCommands = [...(source?.querySelectorAll(
          "li[data-command]",
        ) ?? [])].map((item) => item.dataset.command ?? "");
        const stableCommands = [...(stable?.querySelectorAll(
          "li[data-command]",
        ) ?? [])].map((item) => item.dataset.command ?? "");
        return {
          ready:
            source?.classList.contains("is-visible") &&
            sourceCommands[0] === "\\\\beta" &&
            source?.querySelector("li.ML__popover__current")?.dataset.command ===
              "\\\\beta",
          sourceCommands,
          stableCommands,
          selected:
            source?.querySelector("li.ML__popover__current")?.dataset.command ??
            "",
        };
      })()`, "MathLive native candidates prioritize a short-term repeated command");
  await key("Enter", "Enter", 13);
  const nativeUsageState = await waitForEvaluation(`(() => {
        const persisted = JSON.parse(
          localStorage.getItem("visualtex-editor") || "{}",
        );
        const usage = persisted.state?.usage?.beta;
        const field = document.querySelector("math-field");
        return {
          ready:
            usage?.useCount === 3 &&
            usage?.contextCounts?.candidate === 3 &&
            field?.value === "\\\\beta",
          usage,
          value: field?.value ?? "",
        };
      })()`, "native candidate usage persisted after commit");

  await client.send("Page.reload", { ignoreCache: true });
  await waitForEvaluation(
    `(() => ({ ready: Boolean(document.querySelector("math-field")) }))()`,
    "reloaded field for persistent ranking",
  );
  await clearField();
  await typeText("\\be");
  const reloadedNativeRankState = await waitForEvaluation(`(() => {
        const source = document.getElementById("mathlive-suggestion-popover");
        const first = source?.querySelector("li[data-command]");
        const persisted = JSON.parse(
          localStorage.getItem("visualtex-editor") || "{}",
        );
        return {
          ready:
            source?.classList.contains("is-visible") &&
            first?.dataset.command === "\\\\beta" &&
            persisted.state?.usage?.beta?.useCount === 3,
          firstCommand: first?.dataset.command ?? "",
          useCount: persisted.state?.usage?.beta?.useCount ?? 0,
        };
      })()`, "native frequency ranking survives reload");

  console.log(
    JSON.stringify(
      {
        commonCommandState,
        commonTileState,
        tileUsageState,
        nativeRankState,
        nativeUsageState,
        reloadedNativeRankState,
      },
      null,
      2,
    ),
  );
  console.log("Targeted persistent usage ranking regression passed");
  return;
}
