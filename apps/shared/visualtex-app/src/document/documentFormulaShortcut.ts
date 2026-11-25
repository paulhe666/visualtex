// The formula input method's shortcuts (the app's own global hotkeys, anywhere/hotkey.rs) pressed while VisualTeX's own
// window has the keyboard: the backend opens no input method over the app's page (anywhere/mod.rs toggle) and hands
// the press to the main window, whose document makes the formula itself: an inline formula (the inline shortcut) or a
// display formula (the display shortcut) at the caret, opened for typing, as typing "$$" or "$$$$" there would.
import { isTauri } from "@tauri-apps/api/core";
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

/** Another text field of the window has the keyboard (the source panel, a search box, a dialog's input). */
export function typingOutside(documentRoot: HTMLElement | null | undefined): boolean {
  const active = window.document.activeElement as HTMLElement | null;
  if (!active || active === window.document.body || (documentRoot && documentRoot.contains(active))) return false;
  return active.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(active.tagName);
}
