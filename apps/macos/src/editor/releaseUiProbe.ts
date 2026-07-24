import type { MathfieldElement } from "mathlive";

const RESULT_KEY = "visualtex.release-ui-probe.result";
const STABLE_NATIVE_POPOVER_ID = "visualtex-native-input-suggestion-popover";

interface WrapperProbeHook {
  startWrapper: (command: string) => boolean;
  inputWrapper: (text: string) => boolean;
  deleteWrapper: () => boolean;
  confirmWrapper: () => boolean;
}

interface NativeInputProbeHook {
  sync: () => void;
  move: (direction: 1 | -1) => string | null;
}

const wrapperProbeHook = (field: MathfieldElement) =>
  (
    field as MathfieldElement & {
      __visualtexReleaseProbe?: WrapperProbeHook;
    }
  ).__visualtexReleaseProbe;

const nativeInputProbeHook = () =>
  (
    globalThis as typeof globalThis & {
      __visualtexNativeInputProbe?: NativeInputProbeHook;
    }
  ).__visualtexNativeInputProbe;

const sleep = (milliseconds: number) =>
  new Promise<void>((resolve) => window.setTimeout(resolve, milliseconds));

const waitUntil = async (
  predicate: () => boolean,
  timeoutMilliseconds = 3000,
) => {
  const startedAt = performance.now();
  while (performance.now() - startedAt < timeoutMilliseconds) {
    if (predicate()) return true;
    await sleep(40);
  }
  return false;
};

const frames = async (count = 2) => {
  for (let index = 0; index < count; index += 1) {
    await new Promise<void>((resolve) => {
      let settled = false;
      const finish = () => {
        if (settled) return;
        settled = true;
        window.clearTimeout(timer);
        resolve();
      };
      const timer = window.setTimeout(finish, 120);
      window.requestAnimationFrame(finish);
    });
  }
};

const setProbeStatus = (status: string) => {
  localStorage.setItem("visualtex.release-ui-probe.status", status);
};

const key = (field: MathfieldElement, value: string, code: string) => {
  field.dispatchEvent(
    new KeyboardEvent("keydown", {
      key: value,
      code,
      bubbles: true,
      composed: true,
      cancelable: true,
    }),
  );
};

const beforeInput = (
  field: MathfieldElement,
  inputType: string,
  data: string | null = null,
) => {
  field.dispatchEvent(
    new InputEvent("beforeinput", {
      inputType,
      data,
      bubbles: true,
      composed: true,
      cancelable: true,
    }),
  );
};

const focusField = (field: MathfieldElement) => {
  field.focus();
  field.shadowRoot
    ?.querySelector<HTMLElement>('[part="keyboard-sink"]')
    ?.focus({ preventScroll: true });
};

const setFormulaAtMarker = async (
  field: MathfieldElement,
  latex: string,
  marker = "z",
) => {
  field.setValue(latex, {
    mode: "math",
    format: "latex",
    insertionMode: "replaceAll",
    selectionMode: "after",
    silenceNotifications: true,
  });
  await frames(2);
  await sleep(50);

  let markerEnd = -1;
  for (let end = 1; end <= field.lastOffset; end += 1) {
    const rangeLatex = field.getValue(end - 1, end, "latex").trim();
    const infoLatex = field.getElementInfo(end)?.latex?.trim() ?? "";
    if (rangeLatex === marker || infoLatex === marker) {
      markerEnd = end;
      break;
    }
  }
  if (markerEnd < 0) throw new Error(`Marker ${marker} not found in ${latex}`);

  field.selection = {
    ranges: [[markerEnd, markerEnd]],
    direction: "none",
  };
  field.position = markerEnd;
  focusField(field);
  await frames(2);
  return markerEnd;
};

const modelAnchor = (field: MathfieldElement) => {
  const host = field.closest<HTMLElement>(".mathfield-host");
  const bounds = field.getElementInfo(field.position)?.bounds;
  if (!host || !bounds) return null;
  const hostBounds = host.getBoundingClientRect();
  return {
    left: bounds.right - hostBounds.left,
    top: bounds.top - hostBounds.top + bounds.height / 2,
    height: bounds.height,
    depth: field.getElementInfo(field.position)?.depth ?? -1,
    latex: field.getElementInfo(field.position)?.latex ?? "",
  };
};

const pendingFrame = (field: MathfieldElement) => {
  const host = field.closest<HTMLElement>(".mathfield-host");
  if (!host) return null;
  return {
    visible: host.classList.contains("has-pending-wrapper-placeholder"),
    left: Number.parseFloat(
      host.style.getPropertyValue("--pending-wrapper-left") || "NaN",
    ),
    top: Number.parseFloat(
      host.style.getPropertyValue("--pending-wrapper-top") || "NaN",
    ),
    width: Number.parseFloat(
      host.style.getPropertyValue("--pending-wrapper-width") || "NaN",
    ),
    height: Number.parseFloat(
      host.style.getPropertyValue("--pending-wrapper-height") || "NaN",
    ),
    command: field.dataset.pendingWrapperCommand ?? "",
  };
};

const waitForPendingFrame = async (field: MathfieldElement) => {
  await waitUntil(() => {
    const frame = pendingFrame(field);
    return Boolean(
      frame?.visible &&
        Number.isFinite(frame.left) &&
        Number.isFinite(frame.top) &&
        Number.isFinite(frame.width) &&
        Number.isFinite(frame.height),
    );
  });
  return pendingFrame(field);
};

const enterRawCommand = async (field: MathfieldElement, command: string) => {
  key(field, "\\", "Backslash");
  field.mode = "latex";
  field.insert(command, {
    mode: "latex",
    format: "latex",
    insertionMode: "replaceSelection",
    selectionMode: "after",
    focus: true,
    scrollIntoView: false,
  });
  field.dispatchEvent(
    new InputEvent("input", {
      bubbles: true,
      composed: true,
      inputType: "insertText",
      data: command,
    }),
  );
  await frames(2);
};

const confirmRawWrapper = async (field: MathfieldElement) => {
  key(field, " ", "Space");
  await frames(4);
  await sleep(40);
};

interface WrapperCase {
  name: string;
  command: string;
  source: string;
}

const wrapperCases: WrapperCase[] = [
  {
    name: "fraction numerator",
    command: "\\mathbf",
    source: "p+\\frac{z+n}{d}+q+\\placeholder{}",
  },
  {
    name: "fraction denominator",
    command: "\\mathcal",
    source: "p+\\frac{n}{z+d}+q+\\placeholder{}",
  },
  {
    name: "integral upper limit",
    command: "\\mathfrak",
    source: "p+\\int_{l}^{z+u}f\\,dx+q+\\placeholder{}",
  },
  {
    name: "integral lower limit",
    command: "\\mathbb",
    source: "p+\\int_{z+l}^{u}f\\,dx+q+\\placeholder{}",
  },
];

const runWrapperMode = async (
  field: MathfieldElement,
  setAutoExit: (enabled: boolean) => void,
  enabled: boolean,
) => {
  setAutoExit(enabled);
  await frames(3);
  const results = [];

  for (const testCase of wrapperCases) {
    setProbeStatus(
      `wrapper:${enabled ? "auto" : "continuous"}:${testCase.name}`,
    );
    await setFormulaAtMarker(field, testCase.source);
    const expected = modelAnchor(field);
    const hook = wrapperProbeHook(field);
    if (!hook) throw new Error("Release wrapper probe hook is unavailable");
    hook.startWrapper(testCase.command);
    await frames(2);
    const empty = await waitForPendingFrame(field);

    hook.inputWrapper("A");
    await frames(3);
    const afterA = pendingFrame(field);

    if (!enabled) {
      hook.inputWrapper("B");
      await frames(3);
      hook.deleteWrapper();
      await frames(3);
      hook.inputWrapper("B");
      await frames(3);
    }
    const beforeConfirm = pendingFrame(field);
    if (!enabled) {
      hook.confirmWrapper();
      await frames(3);
    }
    const afterConfirm = pendingFrame(field);

    const emptyTopError =
      expected && empty ? Math.abs(empty.top - expected.top) : Number.POSITIVE_INFINITY;
    const persistentTopError =
      expected && beforeConfirm?.visible
        ? Math.abs(beforeConfirm.top - expected.top)
        : enabled
          ? 0
          : Number.POSITIVE_INFINITY;

    results.push({
      name: testCase.name,
      expected,
      empty,
      afterA,
      beforeConfirm,
      afterConfirm,
      emptyTopError,
      persistentTopError,
      pass:
        Boolean(expected && empty?.visible) &&
        emptyTopError <= 2 &&
        (enabled
          ? !afterA?.visible
          : Boolean(beforeConfirm?.visible) &&
            persistentTopError <= Math.max(6, (expected?.height ?? 0) * 0.35) &&
            !afterConfirm?.visible),
    });
  }

  return {
    enabled,
    pass: results.every((result) => result.pass),
    results,
  };
};

const nativePopoverMarkup = (commands: string[], selectedIndex = 0) =>
  `<ul>${commands
    .map(
      (command, index) =>
        `<li data-command="${command}"${index === selectedIndex ? ' class="ML__popover__current"' : ""}>` +
        `<span class="ML__popover__latex">${command}</span>` +
        `<span class="ML__popover__command">${command}</span>` +
        `</li>`,
    )
    .join("")}</ul>`;

const createNativePopoverSource = (
  commands: string[],
  selectedIndex = 0,
) => {
  document.getElementById("mathlive-suggestion-popover")?.remove();
  const source = document.createElement("div");
  source.id = "mathlive-suggestion-popover";
  source.className = "is-visible top-tip";
  source.style.left = "430px";
  source.style.top = "190px";
  source.innerHTML = nativePopoverMarkup(commands, selectedIndex);
  document.body.append(source);
  return source;
};

const runNativePopoverProbe = async (_field: MathfieldElement) => {
  const hook = nativeInputProbeHook();
  if (!hook) throw new Error("Native input popover probe hook is unavailable");
  createNativePopoverSource(["\\frac", "\\frak", "\\frown", "\\flat"]);
  hook.sync();
  await frames(4);

  const stable = document.getElementById(STABLE_NATIVE_POPOVER_ID);
  const initialBounds = stable?.getBoundingClientRect();
  const initialSelected = stable?.querySelector<HTMLElement>(
    "li.ML__popover__current",
  )?.dataset.command;
  const monitor = {
    removed: 0,
    hidden: 0,
    ariaHidden: 0,
  };
  const observer = new MutationObserver((records) => {
    for (const record of records) {
      if (
        record.type === "childList" &&
        stable &&
        [...record.removedNodes].some(
          (node) => node === stable || (node instanceof Element && node.contains(stable)),
        )
      ) {
        monitor.removed += 1;
      }
      if (record.target === stable && record.attributeName === "class") {
        if (!stable.classList.contains("is-visible")) monitor.hidden += 1;
      }
      if (record.target === stable && record.attributeName === "aria-hidden") {
        if (stable.getAttribute("aria-hidden") !== "false") monitor.ariaHidden += 1;
      }
    }
  });
  observer.observe(document.body, {
    childList: true,
    subtree: true,
    attributes: true,
    attributeFilter: ["class", "aria-hidden"],
  });

  hook.move(1);
  await frames(3);
  const afterArrowNode = document.getElementById(STABLE_NATIVE_POPOVER_ID);
  const afterArrowBounds = afterArrowNode?.getBoundingClientRect();
  const afterArrowSelected = afterArrowNode?.querySelector<HTMLElement>(
    "li.ML__popover__current",
  )?.dataset.command;

  document.getElementById("mathlive-suggestion-popover")?.remove();
  await sleep(12);
  createNativePopoverSource(["\\frac", "\\frak", "\\frown"]);
  hook.sync();
  await frames(4);
  const refinedNode = document.getElementById(STABLE_NATIVE_POPOVER_ID);
  const refinedCommands = Array.from(
    refinedNode?.querySelectorAll<HTMLElement>("li[data-command]") ?? [],
  ).map((item) => item.dataset.command ?? "");

  observer.disconnect();
  const sameGeometry = Boolean(
    initialBounds &&
      afterArrowBounds &&
      Math.abs(initialBounds.left - afterArrowBounds.left) <= 1 &&
      Math.abs(initialBounds.top - afterArrowBounds.top) <= 1 &&
      Math.abs(initialBounds.width - afterArrowBounds.width) <= 1 &&
      Math.abs(initialBounds.height - afterArrowBounds.height) <= 1,
  );
  const pass = Boolean(
    stable &&
      stable.classList.contains("is-visible") &&
      afterArrowNode === stable &&
      refinedNode === stable &&
      initialSelected &&
      afterArrowSelected &&
      initialSelected !== afterArrowSelected &&
      refinedCommands.some((command) => command === "\\frac") &&
      refinedCommands.every((command) => command.startsWith("\\fr")) &&
      sameGeometry &&
      monitor.removed === 0 &&
      monitor.hidden === 0 &&
      monitor.ariaHidden === 0 &&
      !document.querySelector(".suggestion-popup"),
  );
  document.getElementById("mathlive-suggestion-popover")?.remove();
  hook.sync();

  return {
    pass,
    sameNodeAfterArrow: afterArrowNode === stable,
    sameNodeAfterRefine: refinedNode === stable,
    sameGeometry,
    initialSelected,
    afterArrowSelected,
    refinedCommands,
    monitor,
  };
};

const runRawPlaceholderVisualProbe = async (field: MathfieldElement) => {
  field.setValue("\\frac{\\placeholder{}}{\\placeholder{}}", {
    mode: "math",
    format: "latex",
    insertionMode: "replaceAll",
    selectionMode: "placeholder",
    silenceNotifications: true,
  });
  focusField(field);
  await sleep(120);
  key(field, "\\", "Backslash");
  field.mode = "latex";
  field.insert("\\the", {
    mode: "latex",
    format: "latex",
    insertionMode: "replaceSelection",
    selectionMode: "after",
    focus: true,
    scrollIntoView: false,
  });
  field.dispatchEvent(
    new InputEvent("input", {
      bubbles: true,
      composed: true,
      inputType: "insertText",
      data: "\\the",
    }),
  );
  await sleep(180);

  const root = field.shadowRoot;
  const container = root?.querySelector<HTMLElement>(".ML__container");
  const rawText = Array.from(
    root?.querySelectorAll<HTMLElement>(".ML__raw-latex") ?? [],
  )
    .filter((node) => !node.classList.contains("ML__suggestion"))
    .map((node) => node.textContent ?? "")
    .join("");
  const isTransparent = (value: string) =>
    value === "transparent" ||
    value === "rgba(0, 0, 0, 0)" ||
    /rgba\([^)]*,\s*0(?:\.0+)?\)$/.test(value);
  const inspected = container
    ? [container, ...Array.from(container.querySelectorAll<HTMLElement>("*"))]
    : [];
  const offenders = inspected.flatMap((node) => {
    if (
      node.classList.contains("visualtex-structural-placeholder") ||
      node.classList.contains("visualtex-structural-placeholder-caret")
    ) {
      return [];
    }
    const bounds = node.getBoundingClientRect();
    const style = getComputedStyle(node);
    if (
      bounds.width < 8 ||
      bounds.height < 8 ||
      isTransparent(style.backgroundColor)
    ) {
      return [];
    }
    return [
      {
        classes: node.className,
        backgroundColor: style.backgroundColor,
        width: bounds.width,
        height: bounds.height,
      },
    ];
  });

  return {
    pass:
      rawText === "\\the" &&
      Boolean(container?.classList.contains("has-visualtex-raw-latex-command")) &&
      offenders.length === 0,
    rawText,
    rawClass:
      container?.classList.contains("has-visualtex-raw-latex-command") ?? false,
    offenders,
  };
};

const runNativeSpaceSelectionProbe = async (field: MathfieldElement) => {
  field.setValue("", {
    mode: "math",
    format: "latex",
    insertionMode: "replaceAll",
    selectionMode: "after",
    silenceNotifications: true,
  });
  focusField(field);
  key(field, "\\", "Backslash");
  field.mode = "latex";
  field.insert("\\the", {
    mode: "latex",
    format: "latex",
    insertionMode: "replaceSelection",
    selectionMode: "after",
    focus: true,
    scrollIntoView: false,
  });
  field.dispatchEvent(
    new InputEvent("input", {
      bubbles: true,
      composed: true,
      inputType: "insertText",
      data: "\\the",
    }),
  );
  createNativePopoverSource(
    ["\\the", "\\theta", "\\thetasym", "\\therefore"],
    0,
  );
  nativeInputProbeHook()?.sync();
  await sleep(120);

  key(field, "ArrowDown", "ArrowDown");
  await sleep(100);
  const sourceSelected = document
    .getElementById("mathlive-suggestion-popover")
    ?.querySelector<HTMLElement>("li.ML__popover__current[data-command]")
    ?.dataset.command;
  const stableSelected = document
    .getElementById(STABLE_NATIVE_POPOVER_ID)
    ?.querySelector<HTMLElement>("li.ML__popover__current[data-command]")
    ?.dataset.command;

  key(field, " ", "Space");
  const committed = await waitUntil(
    () =>
      field.value.replaceAll(" ", "") === "\\theta" &&
      !field.shadowRoot?.querySelector(".ML__raw-latex"),
    2500,
  );

  return {
    pass:
      sourceSelected === "\\theta" &&
      stableSelected === "\\theta" &&
      committed &&
      field.value.replaceAll(" ", "") === "\\theta",
    sourceSelected,
    stableSelected,
    committed,
    value: field.value,
  };
};

const runPartialWrapperProbe = async (field: MathfieldElement) => {
  await setFormulaAtMarker(field, "p+\\frac{z+n}{d}+q");
  const expected = modelAnchor(field);
  await enterRawCommand(field, "\\math");
  field.dataset.pendingNativeSuggestion = "\\mathbb";
  createNativePopoverSource(["\\mathbb", "\\mathbf", "\\mathcal"], 0);
  nativeInputProbeHook()?.sync();
  await frames(2);
  const selectedCommand = field.dataset.pendingNativeSuggestion;
  const selectedReady = selectedCommand === "\\mathbb";
  await confirmRawWrapper(field);
  const frame = await waitForPendingFrame(field);
  const topError =
    expected && frame ? Math.abs(frame.top - expected.top) : Number.POSITIVE_INFINITY;
  const hook = wrapperProbeHook(field);
  const inserted = Boolean(hook?.inputWrapper("A"));
  await frames(3);

  return {
    pass:
      selectedReady &&
      selectedCommand === "\\mathbb" &&
      frame?.visible === true &&
      frame.command === "\\mathbb" &&
      topError <= 2 &&
      inserted &&
      field.value.replaceAll(" ", "").includes("z\\mathbb{A}+n") &&
      !pendingFrame(field)?.visible,
    selectedReady,
    selectedCommand,
    expected,
    frame,
    topError,
    value: field.value,
  };
};

const runCandidateQueryResetProbe = async (field: MathfieldElement) => {
  field.setValue("\\int", {
    mode: "math",
    format: "latex",
    insertionMode: "replaceAll",
    selectionMode: "after",
    silenceNotifications: true,
  });
  focusField(field);
  field.dispatchEvent(
    new InputEvent("input", {
      bubbles: true,
      composed: true,
      inputType: "insertText",
    }),
  );
  const candidateOpened = await waitUntil(() =>
    Boolean(document.querySelector(".suggestion-popup")),
  );
  const confirmedQuery =
    document.querySelector<HTMLElement>(".editor-surface")?.dataset
      .commandQuery ?? "";

  key(field, "\\", "Backslash");
  field.mode = "latex";
  field.insert("\\", {
    mode: "latex",
    format: "latex",
    insertionMode: "replaceSelection",
    selectionMode: "after",
    focus: true,
    scrollIntoView: false,
  });
  field.dispatchEvent(
    new InputEvent("input", {
      bubbles: true,
      composed: true,
      inputType: "insertText",
      data: "\\",
    }),
  );
  await frames(4);
  const rawLatex = Array.from(
    field.shadowRoot?.querySelectorAll<HTMLElement>(".ML__raw-latex") ?? [],
  )
    .filter((node) => !node.classList.contains("ML__suggestion"))
    .map((node) => node.textContent ?? "")
    .join("");
  const resetQuery =
    document.querySelector<HTMLElement>(".editor-surface")?.dataset
      .commandQuery ?? "";
  const customCandidateVisible = Boolean(
    document.querySelector(".suggestion-popup"),
  );

  return {
    pass:
      candidateOpened &&
      confirmedQuery === "\\int" &&
      rawLatex === "\\" &&
      resetQuery === "" &&
      !customCandidateVisible,
    candidateOpened,
    confirmedQuery,
    rawLatex,
    resetQuery,
    customCandidateVisible,
  };
};

const runStructuralPlaceholderProbe = async (field: MathfieldElement) => {
  field.setValue("\\frac{\\placeholder{}}{\\placeholder{}}", {
    mode: "math",
    format: "latex",
    insertionMode: "replaceAll",
    selectionMode: "placeholder",
    silenceNotifications: true,
  });
  focusField(field);
  await frames(4);
  const initial = Array.from(
    field.shadowRoot?.querySelectorAll<HTMLElement>(
      ".visualtex-structural-placeholder",
    ) ?? [],
  );
  const styles = initial.map((node) => {
    const bounds = node.getBoundingClientRect();
    const style = getComputedStyle(node);
    return {
      width: bounds.width,
      height: bounds.height,
      ratio: bounds.height > 0 ? bounds.width / bounds.height : 99,
      background: style.backgroundColor,
      border: style.borderTopWidth,
      shadow: style.boxShadow,
      selected:
        node.classList.contains("ML__selected") ||
        Boolean(node.closest(".ML__selected")),
    };
  });
  const selected = initial.find(
    (node) =>
      node.classList.contains("ML__selected") ||
      Boolean(node.closest(".ML__selected")),
  );
  const selectedCaret = selected?.querySelector<HTMLElement>(
    ":scope > .visualtex-structural-placeholder-caret",
  );
  const selectedCaretStyle = selectedCaret
    ? getComputedStyle(selectedCaret)
    : null;

  field.setValue("x^{\\placeholder{}}", {
    mode: "math",
    format: "latex",
    insertionMode: "replaceAll",
    selectionMode: "placeholder",
    silenceNotifications: true,
  });
  focusField(field);
  await frames(4);
  const scriptPlaceholder = field.shadowRoot?.querySelector<HTMLElement>(
    ".visualtex-structural-placeholder",
  );
  const scriptBounds = scriptPlaceholder?.getBoundingClientRect();

  field.setValue("\\frac{x}{\\placeholder{}}", {
    mode: "math",
    format: "latex",
    insertionMode: "replaceAll",
    selectionMode: "after",
    silenceNotifications: true,
  });
  await frames(4);
  const afterTyping = field.shadowRoot?.querySelectorAll(
    ".visualtex-structural-placeholder",
  ).length;

  field.setValue("\\frac{\\placeholder{}}{\\placeholder{}}", {
    mode: "math",
    format: "latex",
    insertionMode: "replaceAll",
    selectionMode: "placeholder",
    silenceNotifications: true,
  });
  await frames(4);
  const restored = field.shadowRoot?.querySelectorAll(
    ".visualtex-structural-placeholder",
  ).length;
  const restoredValue = field.value;

  return {
    pass:
      initial.length === 2 &&
      styles.every(
        (style) =>
          style.width > 7 &&
          style.height > 6 &&
          style.ratio > 0.35 &&
          style.ratio < 0.75 &&
          style.border === "0px" &&
          ["rgb(217, 237, 249)", "rgb(207, 232, 247)"].includes(
            style.background,
          ) &&
          style.shadow === "none",
      ) &&
      Boolean(selected) &&
      Boolean(selectedCaret) &&
      Number.parseFloat(selectedCaretStyle?.borderLeftWidth ?? "0") >= 1 &&
      selectedCaretStyle?.left === "0px" &&
      selectedCaretStyle?.animationName.includes(
        "visualtex-placeholder-caret-blink",
      ) &&
      Boolean(
        scriptBounds &&
          styles[0] &&
          scriptBounds.height < styles[0].height &&
          scriptBounds.width < styles[0].width,
      ) &&
      afterTyping === 1 &&
      restored === 2 &&
      restoredValue.includes("\\placeholder{}"),
    styles,
    initialCount: initial.length,
    selectedCaret: {
      present: Boolean(selectedCaret),
      borderLeftWidth: selectedCaretStyle?.borderLeftWidth ?? "",
      left: selectedCaretStyle?.left ?? "",
      animationName: selectedCaretStyle?.animationName ?? "",
    },
    scriptSize: scriptBounds
      ? { width: scriptBounds.width, height: scriptBounds.height }
      : null,
    afterTyping,
    restored,
    restoredValue,
  };
};

export async function runReleasePlaceholderSelectionProbe(
  field: MathfieldElement,
) {
  const resultKey = "visualtex.release-placeholder-selection-probe.result";
  localStorage.removeItem(resultKey);
  const startedAt = Date.now();
  try {
    field.setValue("x+\\frac{\\alpha}{\\placeholder{}}+y", {
      mode: "math",
      format: "latex",
      insertionMode: "replaceAll",
      selectionMode: "placeholder",
      silenceNotifications: true,
    });
    focusField(field);
    await frames(4);
    await sleep(120);

    const root = field.shadowRoot;
    const container = root?.querySelector<HTMLElement>(".ML__container");
    const placeholder = root?.querySelector<HTMLElement>(
      ".visualtex-structural-placeholder",
    );
    const caret = placeholder?.querySelector<HTMLElement>(
      ":scope > .visualtex-structural-placeholder-caret",
    );
    const isTransparent = (value: string) =>
      value === "transparent" ||
      value === "rgba(0, 0, 0, 0)" ||
      /rgba\([^)]*,\s*0(?:\.0+)?\)$/.test(value);
    const offenders = Array.from(
      root?.querySelectorAll<HTMLElement>(
        ".ML__contains-highlight, .ML__highlight, .ML__selected",
      ) ?? [],
    ).flatMap((node) => {
      if (
        node === placeholder ||
        node.classList.contains("visualtex-structural-placeholder-caret")
      ) {
        return [];
      }
      const bounds = node.getBoundingClientRect();
      const style = getComputedStyle(node);
      if (
        bounds.width < 8 ||
        bounds.height < 8 ||
        isTransparent(style.backgroundColor)
      ) {
        return [];
      }
      return [
        {
          classes: node.className,
          backgroundColor: style.backgroundColor,
          width: bounds.width,
          height: bounds.height,
        },
      ];
    });
    const alphaPlaceholder = {
      pass:
        field.value.replaceAll(" ", "") ===
          "x+\\frac{\\alpha}{\\placeholder{}}+y" &&
        Boolean(
          container?.classList.contains(
            "has-visualtex-structural-placeholder-selection",
          ),
        ) &&
        Boolean(caret) &&
        offenders.length === 0,
      value: field.value,
      caretPresent: Boolean(caret),
      offenders,
    };

    field.setValue(
      "a+b+\\frac{\\placeholder{}}{\\placeholder{}}+c+d",
      {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      },
    );
    focusField(field);
    await frames(4);
    const host = field.closest<HTMLElement>(".mathfield-host");
    const bounds = field.getBoundingClientRect();
    host?.dispatchEvent(
      new PointerEvent("pointerdown", {
        bubbles: true,
        composed: true,
        cancelable: true,
        button: 0,
        buttons: 1,
        clientX: bounds.left + 12,
        clientY: bounds.top + bounds.height / 2,
        pointerId: 1,
        pointerType: "mouse",
        isPrimary: true,
      }),
    );
    field.selection = {
      ranges: [[0, field.lastOffset]],
      direction: "forward",
    };
    window.dispatchEvent(
      new PointerEvent("pointerup", {
        bubbles: true,
        composed: true,
        cancelable: true,
        button: 0,
        buttons: 0,
        clientX: bounds.right - 12,
        clientY: bounds.top + bounds.height / 2,
        pointerId: 1,
        pointerType: "mouse",
        isPrimary: true,
      }),
    );
    await frames(4);
    await sleep(120);

    const nextRoot = field.shadowRoot;
    const nextContainer = nextRoot?.querySelector<HTMLElement>(".ML__container");
    const [selectionStart, selectionEnd] = field.selection.ranges[0] ?? [-1, -1];
    const selectedLatex =
      selectionStart >= 0 && selectionEnd >= 0
        ? field.getValue(
            Math.min(selectionStart, selectionEnd),
            Math.max(selectionStart, selectionEnd),
            "latex",
          )
        : "";
    const selection = nextRoot?.querySelector<HTMLElement>(".ML__selection");
    const selectionBounds = selection?.getBoundingClientRect();
    const selectedAtoms = Array.from(
      nextRoot?.querySelectorAll<HTMLElement>(".ML__selected") ?? [],
    );
    const visibleSelectedAtoms = selectedAtoms.filter((node) => {
      const bounds = node.getBoundingClientRect();
      return bounds.width > 0 && bounds.height > 0;
    });
    const rangeSelection = {
      pass:
        !field.selectionIsCollapsed &&
        Math.abs(selectionEnd - selectionStart) > 2 &&
        selectedLatex.includes("\\placeholder{}") &&
        !field.classList.contains("visualtex-pointer-selecting") &&
        !nextContainer?.classList.contains(
          "has-visualtex-structural-placeholder-selection",
        ) &&
        (nextRoot?.querySelectorAll(
          ".visualtex-structural-placeholder-caret",
        ).length ?? -1) === 0 &&
        Boolean(
          (selection &&
            getComputedStyle(selection).display !== "none" &&
            selectionBounds &&
            selectionBounds.width > 5) ||
            visibleSelectedAtoms.length > 0,
        ),
      ranges: field.selection.ranges,
      selectedLatex,
      pointerSelectingClass: field.classList.contains(
        "visualtex-pointer-selecting",
      ),
      placeholderEditingClass: Boolean(
        nextContainer?.classList.contains(
          "has-visualtex-structural-placeholder-selection",
        ),
      ),
      placeholderCaretCount:
        nextRoot?.querySelectorAll(
          ".visualtex-structural-placeholder-caret",
        ).length ?? -1,
      selectionDisplay: selection
        ? getComputedStyle(selection).display
        : "missing",
      selectionWidth: selectionBounds?.width ?? 0,
      selectedAtomCount: selectedAtoms.length,
      visibleSelectedAtomCount: visibleSelectedAtoms.length,
    };

    const result = {
      pass: alphaPlaceholder.pass && rangeSelection.pass,
      startedAt,
      finishedAt: Date.now(),
      userAgent: navigator.userAgent,
      alphaPlaceholder,
      rangeSelection,
    };
    localStorage.setItem(resultKey, JSON.stringify(result));
    return result;
  } catch (error) {
    const result = {
      pass: false,
      startedAt,
      finishedAt: Date.now(),
      error: error instanceof Error ? `${error.name}: ${error.message}` : String(error),
    };
    localStorage.setItem(resultKey, JSON.stringify(result));
    return result;
  }
}

export async function runReleaseUiProbe(
  field: MathfieldElement,
  setAutoExit: (enabled: boolean) => void,
) {
  localStorage.removeItem(RESULT_KEY);
  const startedAt = Date.now();
  try {
    setProbeStatus("probe:wrapper-auto");
    const wrapperAuto = await runWrapperMode(field, setAutoExit, true);
    setProbeStatus("probe:wrapper-continuous");
    const wrapperContinuous = await runWrapperMode(field, setAutoExit, false);
    setProbeStatus("probe:native-popover");
    const nativePopover = await runNativePopoverProbe(field);
    setProbeStatus("probe:raw-placeholder-visual");
    const rawPlaceholderVisual = await runRawPlaceholderVisualProbe(field);
    setProbeStatus("probe:native-space-selection");
    const nativeSpaceSelection = await runNativeSpaceSelectionProbe(field);
    setProbeStatus("probe:partial-wrapper");
    const partialWrapper = await runPartialWrapperProbe(field);
    setProbeStatus("probe:candidate-query-reset");
    const candidateQueryReset = await runCandidateQueryResetProbe(field);
    setProbeStatus("probe:structural-placeholder");
    const structuralPlaceholder = await runStructuralPlaceholderProbe(field);
    setProbeStatus("probe:writing-result");
    const result = {
      pass:
        wrapperAuto.pass &&
        wrapperContinuous.pass &&
        nativePopover.pass &&
        rawPlaceholderVisual.pass &&
        nativeSpaceSelection.pass &&
        partialWrapper.pass &&
        candidateQueryReset.pass &&
        structuralPlaceholder.pass,
      startedAt,
      finishedAt: Date.now(),
      userAgent: navigator.userAgent,
      viewport: { width: innerWidth, height: innerHeight },
      wrapperAuto,
      wrapperContinuous,
      nativePopover,
      rawPlaceholderVisual,
      nativeSpaceSelection,
      partialWrapper,
      candidateQueryReset,
      structuralPlaceholder,
    };
    localStorage.setItem(RESULT_KEY, JSON.stringify(result));
    return result;
  } catch (error) {
    const result = {
      pass: false,
      startedAt,
      finishedAt: Date.now(),
      error: error instanceof Error ? `${error.name}: ${error.message}` : String(error),
    };
    localStorage.setItem(RESULT_KEY, JSON.stringify(result));
    return result;
  }
}
