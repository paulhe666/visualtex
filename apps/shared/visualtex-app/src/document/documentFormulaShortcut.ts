// The formula input method's shortcuts (the app's own global hotkeys, anywhere/hotkey.rs) pressed while VisualTeX's own
// window has the keyboard: the backend opens no input method over the app's page (anywhere/mod.rs toggle) and hands
// the press to the main window, whose document makes the formula itself: an inline formula (the inline shortcut) or a
// display formula (the display shortcut) at the caret, opened for typing, as typing "$$" or "$$$$" there would.
import { invoke, isTauri } from "@tauri-apps/api/core";
import { listen } from "@tauri-apps/api/event";

/** The backend's event (anywhere/mod.rs OWN_FORMULA_SHORTCUT_EVENT); payload { display }. */
export const FORMULA_SHORTCUT_EVENT = "visualtex-formula-shortcut";

/** Calls `handle(display)` for each press; returns the undoing. Nothing outside the desktop app. */
export function watchFormulaShortcut(handle: (display: boolean) => void): () => void {
  if (!isTauri()) return () => undefined;
  let stop: (() => void) | null = null;
  let stopped = false;
  void listen<{ display?: boolean }>(FORMULA_SHORTCUT_EVENT, (event) => handle(Boolean(event.payload?.display))).then(
    (unlisten) => {
      if (stopped) unlisten();
      else stop = unlisten;
    },
    () => undefined,
  );
  return () => {
    stopped = true;
    stop?.();
  };
}

/** How often the formulas are looked at while English is on; one never opened within ENGLISH_OPEN_WAIT gives it back. */
const ENGLISH_CHECK_MS = 100;
const ENGLISH_OPEN_WAIT_MS = 1500;
let englishOn = false;

/**
 * A formula the shortcut opened is typed in English when Settings > Keyboard shortcuts > Formula input method says so:
 * the backend switches the input method to English (anywhere_own_formula_input; nothing when it types Latin letters
 * already) and gives the one it had back once no formula of the document is open (`formulaOpen`; an arrow from one
 * formula into the next keeps English) or the window loses the keyboard.
 */
export function englishWhileFormulaOpen(formulaOpen: () => boolean): void {
  if (!isTauri() || englishOn) return;
  englishOn = true;
  void invoke<boolean>("anywhere_own_formula_input", { english: true }).then(
    (switched) => {
      if (!switched) {
        englishOn = false;
        return;
      }
      const started = performance.now();
      let opened = false;
      const end = () => {
        window.clearInterval(timer);
        window.removeEventListener("blur", end);
        englishOn = false;
        void invoke("anywhere_own_formula_input", { english: false }).catch(() => undefined);
      };
      const timer = window.setInterval(() => {
        if (formulaOpen()) opened = true;
        else if (opened || performance.now() - started > ENGLISH_OPEN_WAIT_MS) end();
      }, ENGLISH_CHECK_MS);
      window.addEventListener("blur", end);
    },
    () => {
      englishOn = false;
    },
  );
}

/** Another text field of the window has the keyboard (the source panel, a search box, a dialog's input). */
export function typingOutside(documentRoot: HTMLElement | null | undefined): boolean {
  const active = window.document.activeElement as HTMLElement | null;
  if (!active || active === window.document.body || (documentRoot && documentRoot.contains(active))) return false;
  return active.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(active.tagName);
}
