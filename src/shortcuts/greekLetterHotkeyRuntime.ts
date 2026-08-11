import type { MathfieldElement } from "mathlive";
import {
  greekLetterHotkeyCommandFromEvent,
  isGreekLetterHotkeyPrefix,
} from "./greekLetterHotkeys";

let installed = false;
let armedField: MathfieldElement | null = null;
let armedTimer = 0;

function activeMathfield(event: KeyboardEvent) {
  const target = event.composedPath()[0];
  if (target instanceof HTMLElement) {
    const direct = target.closest?.("math-field") as MathfieldElement | null;
    if (direct) return direct;
  }
  const active = document.activeElement;
  return active instanceof HTMLElement && active.matches("math-field")
    ? (active as MathfieldElement)
    : null;
}

function clearGreekMode() {
  if (armedTimer) window.clearTimeout(armedTimer);
  armedTimer = 0;
  if (armedField) delete (armedField as HTMLElement).dataset.visualtexGreekMode;
  armedField = null;
}

function armGreekMode(field: MathfieldElement) {
  clearGreekMode();
  armedField = field;
  (field as HTMLElement).dataset.visualtexGreekMode = "true";
  armedTimer = window.setTimeout(clearGreekMode, 5000);
}

function insertGreekCommand(field: MathfieldElement, latex: string) {
  field.insert(latex, {
    insertionMode: "replaceSelection",
    selectionMode: "after",
    focus: true,
  });
}

export function installGreekLetterHotkeyRuntime() {
  if (installed || typeof window === "undefined" || typeof document === "undefined") {
    return;
  }
  installed = true;

  document.addEventListener(
    "keydown",
    (event) => {
      const field = activeMathfield(event);
      if (field && isGreekLetterHotkeyPrefix(event)) {
        event.preventDefault();
        event.stopPropagation();
        armGreekMode(field);
        return;
      }

      if (!armedField) return;
      if (event.key === "Escape") {
        event.preventDefault();
        clearGreekMode();
        return;
      }
      if (event.key === "Shift" || event.key === "Control" || event.key === "Alt" || event.key === "Meta") {
        return;
      }

      const targetField = field ?? armedField;
      if (targetField !== armedField) {
        clearGreekMode();
        return;
      }
      const command = greekLetterHotkeyCommandFromEvent(event);
      clearGreekMode();
      if (!command) return;
      event.preventDefault();
      event.stopPropagation();
      insertGreekCommand(targetField, command.insertTemplate ?? command.command);
    },
    true,
  );

  window.addEventListener("blur", clearGreekMode);
}

installGreekLetterHotkeyRuntime();
