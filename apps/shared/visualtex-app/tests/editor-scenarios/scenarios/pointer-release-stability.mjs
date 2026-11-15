// releasing the pointer keeps the selection and caret it made

export const names = ["pointer-release-stability"];

export async function run({ client, ensurePointerTarget, evaluate, sleep, waitForEvaluation }) {
  await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        if (!field?.isConnected) return { ready: false };
        field.setValue("a+b+c+d+e+f+g+h+i+j", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        field.position = field.lastOffset;
        field.dispatchEvent(new InputEvent("input", {
          bubbles: true,
          composed: true,
          inputType: "insertText",
        }));
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        const base = field.shadowRoot?.querySelector(".ML__base");
        const bounds = base?.getBoundingClientRect();
        return {
          ready: Boolean(bounds && bounds.width > 120 && bounds.height > 12),
          left: bounds?.left ?? 0,
          right: bounds?.right ?? 0,
          centerY: bounds ? bounds.top + bounds.height / 2 : 0,
        };
      })()`, "formula geometry for pointer release stability");
  await sleep(250);
  await waitForEvaluation(`(() => {
        const field = document.querySelector("math-field");
        const persisted = JSON.parse(
          localStorage.getItem("visualtex-editor") || "{}",
        );
        return {
          ready:
            field?.value === "a+b+c+d+e+f+g+h+i+j" &&
            persisted.state?.lines?.[0]?.latex === field.value,
          value: field?.value ?? "",
          storeValue: persisted.state?.lines?.[0]?.latex ?? "",
        };
      })()`, "settled formula and store before pointer drag");
  const geometry = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        if (!field) return null;
        const compact = (value) => value.replace(/\\s+/g, "");
        const offsetForPrefix = (prefix) => {
          for (let offset = 0; offset <= field.lastOffset; offset += 1) {
            if (compact(field.getValue(0, offset, "latex")) === compact(prefix)) {
              return offset;
            }
          }
          return -1;
        };
        const startOffset = offsetForPrefix("a+b");
        const endOffset = offsetForPrefix("a+b+c+d+e+f+g+h");
        const contentBounds = field.shadowRoot
          ?.querySelector('[part="content"]')
          ?.getBoundingClientRect();
        if (!contentBounds || startOffset < 0 || endOffset < 0) return null;
        const y = (contentBounds.top + contentBounds.bottom) / 2;
        const samples = [];
        for (
          let x = Math.ceil(contentBounds.left);
          x <= Math.floor(contentBounds.right);
          x += 1
        ) {
          samples.push({
            x,
            offset: field.getOffsetFromPoint(x, y, { bias: 0 }),
          });
        }
        if (!samples.length) return null;
        const sampleForOffset = (targetOffset) => {
          const exact = samples.filter(
            (sample) => sample.offset === targetOffset,
          );
          if (exact.length) return exact[Math.floor(exact.length / 2)];
          return samples.reduce((best, sample) =>
            Math.abs(sample.offset - targetOffset) <
            Math.abs(best.offset - targetOffset)
              ? sample
              : best
          );
        };
        const startSample = sampleForOffset(startOffset);
        const endSample = sampleForOffset(endOffset);
        return {
          startOffset,
          endOffset,
          mappedStartOffset: startSample.offset,
          mappedEndOffset: endSample.offset,
          startX: startSample.x,
          endX: endSample.x,
          leftX: contentBounds.left + 2,
          rightX: contentBounds.right - 2,
          y,
        };
      })()`);
  if (!geometry) throw new Error("Could not resolve pointer release geometry");

  await ensurePointerTarget(geometry.startX, geometry.y);

  await client.send("Input.dispatchMouseEvent", {
    type: "mousePressed",
    x: geometry.startX,
    y: geometry.y,
    button: "left",
    buttons: 1,
    clickCount: 1,
  });
  await client.send("Input.dispatchMouseEvent", {
    type: "mouseMoved",
    x: (geometry.startX + geometry.endX) / 2,
    y: geometry.y,
    button: "left",
    buttons: 1,
  });
  await client.send("Input.dispatchMouseEvent", {
    type: "mouseMoved",
    x: geometry.endX,
    y: geometry.y,
    button: "left",
    buttons: 1,
  });
  await sleep(100);
  const dragged = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        return {
          range: field?.selection?.ranges?.at(-1) ?? null,
          pointerSelecting:
            field?.classList.contains("visualtex-pointer-selecting") ?? false,
          multiLineSelecting:
            field?.classList.contains("has-visualtex-multi-line-selection") ?? false,
          mathLiveTracking:
            Boolean(field?.shadowRoot?.querySelector(".tracking")),
          contentHasPointerCapture:
            field?.shadowRoot
              ?.querySelector('[part="content"]')
              ?.hasPointerCapture?.(1) ?? false,
        };
      })()`);
  await client.send("Input.dispatchMouseEvent", {
    type: "mouseReleased",
    x: geometry.endX,
    y: geometry.y,
    button: "left",
    buttons: 0,
    clickCount: 1,
  });
  await sleep(120);

  const readSelection = async () => evaluate(`(() => {
        const field = document.querySelector("math-field");
        const range = field?.selection?.ranges?.at(-1) ?? null;
        const start = range ? Math.min(range[0], range[1]) : -1;
        const end = range ? Math.max(range[0], range[1]) : -1;
        return {
          value: field?.value ?? "",
          range,
          selectedLatex:
            field && start >= 0 && end >= 0
              ? field.getValue(start, end, "latex")
              : "",
          pointerSelecting:
            field?.classList.contains("visualtex-pointer-selecting") ?? false,
          mathLiveTracking:
            Boolean(field?.shadowRoot?.querySelector(".tracking")),
          contentHasPointerCapture:
            field?.shadowRoot
              ?.querySelector('[part="content"]')
              ?.hasPointerCapture?.(1) ?? false,
        };
      })()`);
  const released = await readSelection();
  if (
    !released.range ||
    released.range[0] === released.range[1] ||
    released.pointerSelecting ||
    released.mathLiveTracking ||
    released.contentHasPointerCapture
  ) {
    throw new Error(
      `Pointer drag did not finish with a stable range: ${JSON.stringify({ geometry, dragged, released })}`,
    );
  }

  await client.send("Input.dispatchMouseEvent", {
    type: "mouseMoved",
    x: geometry.leftX,
    y: geometry.y,
    button: "none",
    buttons: 0,
  });
  await sleep(100);
  const movedLeft = await readSelection();
  await client.send("Input.dispatchMouseEvent", {
    type: "mouseMoved",
    x: geometry.rightX,
    y: geometry.y,
    button: "none",
    buttons: 0,
  });
  await sleep(100);
  const movedRight = await readSelection();

  const stableSelection =
    JSON.stringify(released.range) === JSON.stringify(movedLeft.range) &&
    JSON.stringify(released.range) === JSON.stringify(movedRight.range) &&
    released.selectedLatex === movedLeft.selectedLatex &&
    released.selectedLatex === movedRight.selectedLatex &&
    !movedLeft.mathLiveTracking &&
    !movedRight.mathLiveTracking &&
    !movedLeft.contentHasPointerCapture &&
    !movedRight.contentHasPointerCapture;
  if (!stableSelection) {
    throw new Error(
      `Selection followed the pointer after mouse release: ${JSON.stringify({ released, movedLeft, movedRight })}`,
    );
  }

  console.log(JSON.stringify({ geometry, dragged, released, movedLeft, movedRight }, null, 2));
  console.log("Targeted pointer release stability regression passed");
  return;
}
